"""Read-only resolution of the character formatting effective for a run."""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml.ns import qn
from docx.oxml.simpletypes import ST_HexColor, ST_HpsMeasure, ST_OnOff, ST_String

if TYPE_CHECKING:
    from docx.text.run import Run


TOGGLE_TAGS = frozenset(
    {
        "b",
        "bCs",
        "caps",
        "emboss",
        "i",
        "iCs",
        "imprint",
        "outline",
        "shadow",
        "smallCaps",
        "strike",
        "vanish",
    }
)

LENGTH_TAGS = frozenset({"sz", "szCs"})
FONT_SLOTS = (
    ("ascii", "asciiTheme"),
    ("hAnsi", "hAnsiTheme"),
    ("eastAsia", "eastAsiaTheme"),
    ("cs", "cstheme"),
)


class FormatSource:
    """Source of a single resolved character-formatting value."""

    def __init__(
        self,
        kind: str,
        *,
        path: str = "",
        style_id: str | None = None,
        style_name: str | None = None,
        condition: str | None = None,
        mode: str = "explicit",
        inherited_from: FormatSource | None = None,
        overridden_from: FormatSource | None = None,
        xml_value: Any = None,
    ):
        self.kind = kind
        self.path = path
        self.style_id = style_id
        self.style_name = style_name
        self.condition = condition
        self.mode = mode
        self.inherited_from = inherited_from
        self.overridden_from = overridden_from
        self.xml_value = xml_value

    def describe(self) -> str:
        if self.kind == "document-default":
            label = "document default run properties"
        elif self.kind == "style":
            label = "style %r" % (self.style_name or self.style_id or "<unnamed>")
        elif self.kind == "table-conditional-style":
            label = "table style %r conditional formatting %r" % (
                self.style_name or self.style_id or "<unnamed>",
                self.condition,
            )
        elif self.kind == "direct":
            label = "direct run formatting"
        elif self.kind == "specification-default":
            label = "WordprocessingML default"
        else:
            label = self.kind

        if self.mode == "toggle":
            label += " (toggle reversing inherited value)"
        elif self.mode == "explicit":
            label += " (explicit value)"

        return f"{label}; {self.path}" if self.path else label


class FormatDiagnostic:
    """Diagnostic recorded while resolving formatting."""

    def __init__(self, code: str, message: str):
        self.code = code
        self.message = message


class ResolvedFontFace:
    """Effective explicit and theme typefaces for one character range."""

    def __init__(
        self,
        *,
        name: str | None,
        theme: str | None,
        name_source: FormatSource | None,
        theme_source: FormatSource | None,
    ):
        self.name = name
        self.theme = theme
        self.name_source = name_source
        self.theme_source = theme_source

    @property
    def value(self) -> str | None:
        return self.name if self.name is not None else self.theme

    @property
    def is_theme(self) -> bool:
        return self.name is None and self.theme is not None

    @property
    def source(self) -> FormatSource | None:
        return self.name_source if self.name is not None else self.theme_source


class ResolvedCharacterFormat:
    """Effective character properties and the source of each resolved value."""

    def __init__(
        self,
        values: dict[str, Any],
        sources: dict[str, FormatSource],
        diagnostics: list[FormatDiagnostic],
    ):
        self._values = values
        self._sources = sources
        self.diagnostics = diagnostics

    def __getitem__(self, key: str) -> Any:
        return self._values[key]

    def __iter__(self):
        return iter(self._values)

    def __len__(self) -> int:
        return len(self._values)

    def __contains__(self, key: object) -> bool:
        return key in self._values

    def get(self, key: str, default: Any = None) -> Any:
        return self._values.get(key, default)

    def source(self, key: str) -> FormatSource | None:
        return self._sources.get(key)

    @property
    def fonts(self) -> dict[str, ResolvedFontFace]:
        return {
            slot: ResolvedFontFace(
                name=self._values.get(f"rFonts.{slot}.name"),
                theme=self._values.get(f"rFonts.{slot}.theme"),
                name_source=self._sources.get(f"rFonts.{slot}.name"),
                theme_source=self._sources.get(f"rFonts.{slot}.theme"),
            )
            for slot in ("ascii", "hAnsi", "eastAsia", "cs")
            if (f"rFonts.{slot}.name" in self._values or f"rFonts.{slot}.theme" in self._values)
        }


class _CharacterFormatResolver:
    def __init__(self, run: Run):
        self._run = run
        self._r = run._r
        self._values: dict[str, Any] = {}
        self._sources: dict[str, FormatSource] = {}
        self._diagnostics: list[FormatDiagnostic] = []

    def resolve(self) -> ResolvedCharacterFormat:
        styles_root = self._styles_root()

        if styles_root is not None:
            rpr = self._first(styles_root.xpath("./w:docDefaults/w:rPrDefault/w:rPr"))
            if rpr is not None:
                self._apply_rpr(
                    rpr,
                    FormatSource(
                        "document-default",
                        path="w:styles/w:docDefaults/w:rPrDefault/w:rPr",
                    ),
                    is_style_layer=False,
                )

            self._apply_styles(styles_root)

        direct_rpr = self._r.find(qn("w:rPr"))
        if direct_rpr is not None:
            self._apply_rpr(
                direct_rpr,
                FormatSource("direct", path="w:r/w:rPr"),
                is_style_layer=False,
            )

        for key in TOGGLE_TAGS:
            if key not in self._values:
                self._set(key, False, FormatSource("specification-default", mode="default"))

        return ResolvedCharacterFormat(self._values, self._sources, self._diagnostics)

    def _apply_styles(self, styles_root: Any) -> None:
        p = self._first(self._r.xpath("./ancestor::w:p[1]"))
        tbl = self._first(self._r.xpath("./ancestor::w:tbl[1]"))

        if tbl is not None:
            tbl_pr = tbl.find(qn("w:tblPr"))
            table_style_id = None
            if tbl_pr is not None:
                tbl_style = tbl_pr.find(qn("w:tblStyle"))
                if tbl_style is not None:
                    table_style_id = tbl_style.get(qn("w:val"))
            for style in self._style_chain(
                styles_root, table_style_id, "table", "w:tbl/w:tblPr/w:tblStyle"
            ):
                self._apply_table_style(styles_root, style, tbl)

        paragraph_style_id = None
        if p is not None:
            p_style = self._first(p.xpath("./w:pPr/w:pStyle"))
            if p_style is not None:
                paragraph_style_id = p_style.get(qn("w:val"))
        for style in self._style_chain(
            styles_root, paragraph_style_id, "paragraph", "w:p/w:pPr/w:pStyle"
        ):
            self._apply_style_rpr(style)

        r_style = self._first(self._r.xpath("./w:rPr/w:rStyle"))
        run_style_id = r_style.get(qn("w:val")) if r_style is not None else None
        for style in self._style_chain(
            styles_root, run_style_id, "character", "w:r/w:rPr/w:rStyle"
        ):
            self._apply_style_rpr(style)

    def _apply_table_style(self, styles_root: Any, style: Any, tbl: Any) -> None:
        self._apply_style_rpr(style)
        conditions = self._table_conditions(styles_root, style, tbl)
        conditions_by_type = {
            condition: self._first(style.xpath(f"./w:tblStylePr[@w:type='{condition}']/w:rPr"))
            for condition in conditions
        }

        for condition in (
            "nwCell",
            "neCell",
            "swCell",
            "seCell",
            "firstRow",
            "lastRow",
            "firstCol",
            "lastCol",
            "band1Vert",
            "band2Vert",
            "band1Horz",
            "band2Horz",
        ):
            rpr = conditions_by_type.get(condition)
            if rpr is None:
                continue
            self._apply_rpr(
                rpr,
                FormatSource(
                    "table-conditional-style",
                    path=self._style_path(style, f"w:tblStylePr[@w:type='{condition}']/w:rPr"),
                    style_id=self._style_id(style),
                    style_name=self._style_name(style),
                    condition=condition,
                ),
                is_style_layer=True,
            )

    def _apply_style_rpr(self, style: Any) -> None:
        rpr = style.find(qn("w:rPr"))
        if rpr is None:
            return
        self._apply_rpr(
            rpr,
            FormatSource(
                "style",
                path=self._style_path(style, "w:rPr"),
                style_id=self._style_id(style),
                style_name=self._style_name(style),
            ),
            is_style_layer=True,
        )

    def _apply_rpr(self, rpr: Any, source: FormatSource, *, is_style_layer: bool) -> None:
        for child in rpr:
            local_name = self._local_name(child.tag)
            if local_name in {"rStyle", "rPrChange"}:
                continue
            if local_name == "rFonts":
                self._apply_rfonts(child, source)
            elif local_name in TOGGLE_TAGS:
                self._apply_toggle(child, source, is_style_layer)
            elif local_name in LENGTH_TAGS:
                self._apply_length(child, source)
            elif local_name == "color":
                self._apply_color(child, source)
            elif local_name in {"highlight", "u", "vertAlign"}:
                self._apply_attr(child, source, local_name, "val")
            else:
                self._apply_generic(child, source)

    def _apply_rfonts(self, element: Any, source: FormatSource) -> None:
        for explicit_attr, theme_attr in FONT_SLOTS:
            explicit = element.get(qn(f"w:{explicit_attr}"))
            if explicit is not None:
                self._set_font_face(
                    f"rFonts.{explicit_attr}.name",
                    explicit,
                    source,
                    element,
                    explicit_attr,
                )
                continue
            theme = element.get(qn(f"w:{theme_attr}"))
            if theme is not None:
                self._set_font_face(
                    f"rFonts.{explicit_attr}.theme",
                    theme,
                    source,
                    element,
                    theme_attr,
                )

        for attr_name, value in element.attrib.items():
            local_attr = self._local_name(attr_name)
            if local_attr in {
                "ascii",
                "hAnsi",
                "eastAsia",
                "cs",
                "asciiTheme",
                "hAnsiTheme",
                "eastAsiaTheme",
                "cstheme",
            }:
                continue
            self._set(
                f"rFonts.{local_attr}",
                ST_String.from_xml(value),
                self._child_source(source, element, local_attr),
            )

    def _apply_toggle(self, element: Any, source: FormatSource, is_style_layer: bool) -> None:
        key = self._local_name(element.tag)
        explicit_value = ST_OnOff.from_xml(element.get(qn("w:val"), "true"))
        if is_style_layer:
            inherited = self._values.get(key, False)
            self._set(
                key,
                not inherited,
                FormatSource(
                    source.kind,
                    path=source.path,
                    style_id=source.style_id,
                    style_name=source.style_name,
                    condition=source.condition,
                    mode="toggle",
                    inherited_from=self._sources.get(key),
                    xml_value=explicit_value,
                ),
            )
        else:
            self._set(key, explicit_value, source)

    def _apply_length(self, element: Any, source: FormatSource) -> None:
        key = self._local_name(element.tag)
        value = element.get(qn("w:val"))
        if value is not None:
            self._set(
                key,
                ST_HpsMeasure.from_xml(value),
                self._child_source(source, element, "val"),
            )

    def _apply_color(self, element: Any, source: FormatSource) -> None:
        value = element.get(qn("w:val"))
        if value is not None:
            self._set(
                "color.val",
                ST_HexColor.from_xml(value),
                self._child_source(source, element, "val"),
            )
        for attr_name, value in element.attrib.items():
            local_attr = self._local_name(attr_name)
            if local_attr == "val":
                continue
            self._set(
                f"color.{local_attr}",
                ST_String.from_xml(value),
                self._child_source(source, element, local_attr),
            )

    def _apply_attr(self, element: Any, source: FormatSource, key: str, attr: str) -> None:
        value = element.get(qn(f"w:{attr}"))
        if value is not None:
            self._set(key, ST_String.from_xml(value), self._child_source(source, element, attr))

    def _apply_generic(self, element: Any, source: FormatSource) -> None:
        tag = self._local_name(element.tag)
        for attr_name, value in element.attrib.items():
            local_attr = self._local_name(attr_name)
            self._set(
                f"{tag}.{local_attr}",
                ST_String.from_xml(value),
                self._child_source(source, element, local_attr),
            )

    def _table_conditions(self, styles_root: Any, table_style: Any, tbl: Any) -> set[str]:
        flags = self._cnf_flags()
        has_cnf_marker = self._has_cnf_marker()
        if not has_cnf_marker:
            flags = self._inferred_cnf_flags(tbl)
            if not any(flags[name] for name in ("oddHBand", "evenHBand", "oddVBand", "evenVBand")):
                row_index = self._row_index(tbl)
                col_index = self._column_index(tbl)
                if row_index is not None and col_index is not None:
                    row_band = row_index // self._band_size(
                        styles_root, table_style, "tblStyleRowBandSize", 1
                    )
                    col_band = col_index // self._band_size(
                        styles_root, table_style, "tblStyleColBandSize", 1
                    )
                    flags["oddHBand"] = row_band % 2 == 0
                    flags["evenHBand"] = row_band % 2 == 1
                    flags["oddVBand"] = col_band % 2 == 0
                    flags["evenVBand"] = col_band % 2 == 1

        conditions: set[str] = set()
        look = self._first(tbl.xpath("./w:tblPr/w:tblLook"))
        if look is None:
            return conditions

        first_row = flags["firstRow"] and self._edge_look_enabled(look, "firstRow")
        last_row = flags["lastRow"] and self._edge_look_enabled(look, "lastRow")
        first_column = flags["firstColumn"] and self._edge_look_enabled(look, "firstColumn")
        last_column = flags["lastColumn"] and self._edge_look_enabled(look, "lastColumn")
        h_band = (flags["oddHBand"] or flags["evenHBand"]) and not self._on_off(
            look.get(qn("w:noHBand"), "0")
        )
        v_band = (flags["oddVBand"] or flags["evenVBand"]) and not self._on_off(
            look.get(qn("w:noVBand"), "0")
        )

        if first_row and first_column and flags["firstRowFirstColumn"]:
            conditions.add("nwCell")
        if first_row and last_column and flags["firstRowLastColumn"]:
            conditions.add("neCell")
        if last_row and first_column and flags["lastRowFirstColumn"]:
            conditions.add("swCell")
        if last_row and last_column and flags["lastRowLastColumn"]:
            conditions.add("seCell")

        if first_row:
            conditions.add("firstRow")
        if last_row:
            conditions.add("lastRow")
        if first_column:
            conditions.add("firstCol")
        if last_column:
            conditions.add("lastCol")

        row_band = self._band_size(styles_root, table_style, "tblStyleRowBandSize", 1)
        column_band = self._band_size(styles_root, table_style, "tblStyleColBandSize", 1)
        if h_band:
            row_band_index = self._row_band_index(tbl, row_band)
            if row_band_index is not None:
                conditions.add("band1Horz" if row_band_index % 2 == 0 else "band2Horz")
        if v_band:
            column_band_index = self._column_band_index(tbl, column_band)
            if column_band_index is not None:
                conditions.add("band1Vert" if column_band_index % 2 == 0 else "band2Vert")

        return conditions

    def _cnf_flags(self) -> dict[str, bool]:
        flags = {name: False for name in self._cnf_flag_names()}
        nodes = self._r.xpath(
            "./ancestor::w:tr[1]/w:trPr/w:cnfStyle | "
            "./ancestor::w:tc[1]/w:tcPr/w:cnfStyle | "
            "./ancestor::w:p[1]/w:pPr/w:cnfStyle"
        )
        for node in nodes:
            val = node.get(qn("w:val"))
            if val and len(val) == len(self._cnf_flag_names()):
                for flag, char in zip(self._cnf_flag_names(), val):
                    flags[flag] = flags[flag] or char == "1"
            for flag in self._cnf_flag_names():
                attr = node.get(qn(f"w:{flag}"))
                if attr is not None:
                    flags[flag] = flags[flag] or self._on_off(attr)
        return flags

    def _has_cnf_marker(self) -> bool:
        return bool(
            self._r.xpath(
                "./ancestor::w:tr[1]/w:trPr/w:cnfStyle | "
                "./ancestor::w:tc[1]/w:tcPr/w:cnfStyle | "
                "./ancestor::w:p[1]/w:pPr/w:cnfStyle"
            )
        )

    def _inferred_cnf_flags(self, tbl: Any) -> dict[str, bool]:
        flags = {name: False for name in self._cnf_flag_names()}
        row_index = self._row_index(tbl)
        col_index = self._column_index(tbl)
        row_count = len(tbl.tr_lst)
        col_count = tbl.col_count
        if row_index == 0:
            flags["firstRow"] = True
        if row_index == row_count - 1:
            flags["lastRow"] = True
        if col_index == 0:
            flags["firstColumn"] = True
        if col_index == col_count - 1:
            flags["lastColumn"] = True
        if row_index == 0 and col_index == 0:
            flags["firstRowFirstColumn"] = True
        if row_index == 0 and col_index == col_count - 1:
            flags["firstRowLastColumn"] = True
        if row_index == row_count - 1 and col_index == 0:
            flags["lastRowFirstColumn"] = True
        if row_index == row_count - 1 and col_index == col_count - 1:
            flags["lastRowLastColumn"] = True
        return flags

    def _edge_look_enabled(self, look: Any, attr: str) -> bool:
        return self._on_off(look.get(qn(f"w:{attr}"), "0"))

    def _row_band_index(self, tbl: Any, band_size: int) -> int | None:
        row_index = self._row_index(tbl)
        return None if row_index is None else row_index // band_size

    def _column_band_index(self, tbl: Any, band_size: int) -> int | None:
        column_index = self._column_index(tbl)
        return None if column_index is None else column_index // band_size

    def _band_size(self, styles_root: Any, style: Any, tag: str, default: int) -> int:
        seen: set[str] = set()
        current = style
        while current is not None and self._style_id(current) not in seen:
            seen.add(self._style_id(current))
            value = self._first(current.xpath(f"./w:tblPr/w:{tag}/@w:val"))
            if value is not None:
                try:
                    return int(str(value))
                except ValueError:
                    pass
            based_on = current.find(qn("w:basedOn"))
            if based_on is None:
                break
            if styles_root is None:
                break
            current = styles_root.get_by_id(based_on.get(qn("w:val")))
        return default

    def _row_index(self, tbl: Any) -> int | None:
        row = self._first(self._r.xpath("./ancestor::w:tr[1]"))
        if row is None:
            return None
        try:
            return tbl.tr_lst.index(row)
        except ValueError:
            return None

    def _column_index(self, tbl: Any) -> int | None:
        cell = self._first(self._r.xpath("./ancestor::w:tc[1]"))
        row_index = self._row_index(tbl)
        if cell is None or row_index is None:
            return None
        try:
            return cell.grid_offset
        except (ValueError, IndexError, AttributeError):
            return None

    def _style_chain(
        self,
        styles_root: Any | None,
        style_id: str | None,
        style_type: str,
        reference_path: str,
    ) -> list[Any]:
        if styles_root is None:
            return []
        if style_id is None:
            default = self._default_style(styles_root, style_type)
            return [default] if default is not None else []
        chain: list[Any] = []
        seen: list[str] = []
        current_id: str | None = style_id
        while current_id is not None:
            if current_id in seen:
                cycle = " -> ".join([*seen, current_id])
                self._diagnostics.append(
                    FormatDiagnostic("cycle", f"cyclic {style_type} style reference: {cycle}")
                )
                break
            seen.append(current_id)
            style = styles_root.get_by_id(current_id)
            if style is None:
                self._diagnostics.append(
                    FormatDiagnostic(
                        "missing-style",
                        f"{style_type} style {current_id!r} referenced by "
                        f"{reference_path} is not defined",
                    )
                )
                default = self._default_style(styles_root, style_type)
                if default is not None and default not in chain:
                    chain.insert(0, default)
                break
            actual_type = style.get(qn("w:type"))
            if actual_type is not None and actual_type != style_type:
                self._diagnostics.append(
                    FormatDiagnostic(
                        "wrong-style-type",
                        f"style {current_id!r} has type {actual_type!r}, "
                        f"expected {style_type!r}",
                    )
                )
                default = self._default_style(styles_root, style_type)
                if default is not None and default not in chain:
                    chain.insert(0, default)
                break
            chain.append(style)
            based_on = style.find(qn("w:basedOn"))
            current_id = based_on.get(qn("w:val")) if based_on is not None else None
        return list(reversed(chain))

    def _default_style(self, styles_root: Any, style_type: str) -> Any | None:
        matches = styles_root.xpath(
            f"./w:style[@w:type='{style_type}']"
            "[(@w:default='true') or (@w:default='1') or (@w:default='on')]"
        )
        return matches[-1] if matches else None

    def _styles_root(self) -> Any | None:
        try:
            part = self._run.part
        except (AttributeError, ValueError):
            return None
        try:
            styles_part = part.part_related_by(RT.STYLES)
        except KeyError:
            return None
        return styles_part.element

    def _set_font_face(
        self, key: str, value: str, source: FormatSource, element: Any, attr: str
    ) -> None:
        self._set(key, value, self._child_source(source, element, attr))

    def _with_overridden(
        self, source: FormatSource, overridden: FormatSource | None
    ) -> FormatSource:
        if overridden is None:
            return source
        return FormatSource(
            source.kind,
            path=source.path,
            style_id=source.style_id,
            style_name=source.style_name,
            condition=source.condition,
            mode=source.mode,
            inherited_from=None if source.kind == "direct" else overridden,
            overridden_from=overridden if source.kind == "direct" else None,
            xml_value=source.xml_value,
        )

    def _set(self, key: str, value: Any, source: FormatSource) -> None:
        if (
            source.kind == "direct"
            and key in self._sources
            and source.inherited_from is None
            and source.overridden_from is None
        ):
            source = self._with_overridden(source, self._sources[key])
        self._values[key] = value
        self._sources[key] = source

    def _child_source(self, parent: FormatSource, element: Any, attr: str | None) -> FormatSource:
        return FormatSource(
            parent.kind,
            path=self._child_path(parent.path, element, attr),
            style_id=parent.style_id,
            style_name=parent.style_name,
            condition=parent.condition,
            mode=parent.mode,
            inherited_from=parent.inherited_from,
            overridden_from=parent.overridden_from,
            xml_value=parent.xml_value,
        )

    def _child_path(self, parent_path: str, element: Any, attr: str | None) -> str:
        parts = [part for part in parent_path.split("/") if part]
        tag = f"w:{self._local_name(element.tag)}"
        if tag not in parts:
            parts.append(tag)
        if attr:
            parts[-1] = f"{parts[-1]}/@w:{attr}"
        return "/".join(parts)

    def _style_path(self, style: Any, suffix: str) -> str:
        return f"w:styles/w:style[@w:styleId='{self._style_id(style)}']/{suffix}"

    def _style_id(self, style: Any) -> str:
        return str(style.get(qn("w:styleId")) or "")

    def _style_name(self, style: Any) -> str | None:
        name = style.find(qn("w:name"))
        return name.get(qn("w:val")) if name is not None else None

    def _first(self, values: Any) -> Any | None:
        return next(iter(values), None)

    def _on_off(self, value: str) -> bool:
        return ST_OnOff.from_xml(value)

    def _local_name(self, tag: str) -> str:
        return tag.rsplit("}", 1)[-1]

    def _cnf_flag_names(self) -> tuple[str, ...]:
        return (
            "firstRow",
            "lastRow",
            "firstColumn",
            "lastColumn",
            "oddVBand",
            "evenVBand",
            "oddHBand",
            "evenHBand",
            "firstRowFirstColumn",
            "firstRowLastColumn",
            "lastRowFirstColumn",
            "lastRowLastColumn",
        )


def effective_character_format(run: Run) -> ResolvedCharacterFormat:
    """Return character formatting that is effective for `run`, with provenance."""
    return _CharacterFormatResolver(run).resolve()
