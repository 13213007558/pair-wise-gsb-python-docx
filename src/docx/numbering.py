"""List-numbering resolution and the |ListItem| proxy object.

This module computes the *rendered* number of each list paragraph in a document,
expanding abstract numbering definitions (<w:abstractNum>), concrete numbering
instances (<w:num>), level overrides (<w:lvlOverride> / <w:startOverride>), restart
behavior (<w:lvlRestart>), and style links into a concrete label such as "2.3)" or
"第十二条" for each numbered paragraph, in document order.

Resolution is strictly read-only; no numbering XML or paragraph content is modified
and automatic numbering is never converted to literal text.
"""

from __future__ import annotations

import re
from typing import TYPE_CHECKING, Dict, Iterator, List, NamedTuple, Optional, Tuple

if TYPE_CHECKING:
    from docx.document import Document
    from docx.text.paragraph import Paragraph

#: Number formats (w:numFmt/@w:val) for which an expanded label can be rendered.
#: Any other format is reported via ListItem.is_num_fmt_supported and its original
#: definition is preserved (in particular, it is never coerced to decimal).
SUPPORTED_NUM_FMTS = frozenset(
    {
        "bullet",
        "none",
        "decimal",
        "decimalZero",
        "lowerLetter",
        "upperLetter",
        "lowerRoman",
        "upperRoman",
        "chineseCounting",
        "chineseCountingThousand",
    }
)

#: Matches an upper-level number reference like %1 .. %9 in w:lvlText.
_LVL_REF_RE = re.compile(r"%([1-9])")

_CHINESE_DIGITS = "零一二三四五六七八九"
_CHINESE_UNITS = ("", "十", "百", "千")

_ROMAN_NUMERALS = (
    (1000, "M"),
    (900, "CM"),
    (500, "D"),
    (400, "CD"),
    (100, "C"),
    (90, "XC"),
    (50, "L"),
    (40, "XL"),
    (10, "X"),
    (9, "IX"),
    (5, "V"),
    (4, "IV"),
    (1, "I"),
)


def _to_letters(value: int) -> str:
    """Return spreadsheet-style letters for value, e.g. 1 -> "a", 27 -> "aa"."""
    letters = ""
    remaining = value
    while remaining > 0:
        remaining, offset = divmod(remaining - 1, 26)
        letters = chr(ord("a") + offset) + letters
    return letters


def _to_roman(value: int) -> str:
    """Return Roman numerals for value in range 1..3999, e.g. 4 -> "iv"."""
    if not 0 < value < 4000:
        return str(value)
    result = []
    remainder = value
    for arabic, numeral in _ROMAN_NUMERALS:
        count, remainder = divmod(remainder, arabic)
        result.append(numeral * count)
    return "".join(result).lower()


def _to_chinese(value: int) -> str:
    """Return Chinese counting numerals for value in range 1..9999, e.g. 12 -> "十二"."""
    if not 0 < value < 10000:
        return str(value)
    digits = str(value)
    result = []
    zero_pending = False
    for idx, char in enumerate(digits):
        digit = int(char)
        unit = _CHINESE_UNITS[len(digits) - 1 - idx]
        if digit == 0:
            zero_pending = True
            continue
        if zero_pending:
            result.append("零")
            zero_pending = False
        result.append(_CHINESE_DIGITS[digit] + unit)
    text = "".join(result)
    # -- "一十二" is conventionally written "十二" --
    if text.startswith("一十"):
        text = text[1:]
    return text


#: Renderers for each supported numbered (non-bullet) format.
_RENDERERS = {
    "decimal": str,
    "decimalZero": lambda value: "%02d" % value,
    "lowerLetter": _to_letters,
    "upperLetter": lambda value: _to_letters(value).upper(),
    "lowerRoman": _to_roman,
    "upperRoman": lambda value: _to_roman(value).upper(),
    "chineseCounting": _to_chinese,
    "chineseCountingThousand": _to_chinese,
}


class ListItem:
    """A single numbered (or bulleted) paragraph occurrence in a document.

    Produced by Document.iter_list_items(). Each instance pairs a list paragraph with
    its resolved numbering position: the concrete numbering instance (numId), the
    zero-based level, the computed counter value, and the fully expanded label text
    (for example "2.3)" or "第十二条").
    """

    def __init__(
        self,
        paragraph: Paragraph,
        num_id: int,
        level: int,
        value: Optional[int],
        label: Optional[str],
        num_fmt: str,
        is_num_fmt_supported: bool,
    ):
        self._paragraph = paragraph
        self._num_id = num_id
        self._level = level
        self._value = value
        self._label = label
        self._num_fmt = num_fmt
        self._is_num_fmt_supported = is_num_fmt_supported

    def __repr__(self) -> str:
        return "ListItem(num_id=%r, level=%r, value=%r, label=%r)" % (
            self._num_id,
            self._level,
            self._value,
            self._label,
        )

    @property
    def is_num_fmt_supported(self) -> bool:
        """True when the number format of this item can be rendered by this library.

        When False, the original w:numFmt definition is preserved in .num_fmt (it is
        not coerced to decimal) and .label is None. The set of supported formats is
        available as docx.numbering.SUPPORTED_NUM_FMTS.
        """
        return self._is_num_fmt_supported

    @property
    def label(self) -> Optional[str]:
        """The fully expanded numbering label for this paragraph, e.g. "2.3)".

        Upper-level references in w:lvlText (like %1) are expanded using the current
        counter values of the referenced levels. None when the number format of this
        item (or of a level it references) is not supported; see
        .is_num_fmt_supported.
        """
        return self._label

    @property
    def level(self) -> int:
        """Zero-based numbering level of this paragraph (w:ilvl), in range 0-8."""
        return self._level

    @property
    def num_fmt(self) -> str:
        """The raw w:numFmt value of the level of this item, e.g. "decimal".

        The original definition is always preserved here, even for formats this
        library cannot render.
        """
        return self._num_fmt

    @property
    def num_id(self) -> int:
        """The concrete numbering instance id (w:numId) this paragraph belongs to.

        Each distinct numId maintains its own independent set of level counters.
        """
        return self._num_id

    @property
    def paragraph(self) -> Paragraph:
        """The |Paragraph| object this list item describes."""
        return self._paragraph

    @property
    def value(self) -> Optional[int]:
        """The computed counter value of this item at its own level.

        None for bullet and "none" formats, which carry no counter.
        """
        return self._value


class _LevelProps(NamedTuple):
    """Effective properties of a single level of a concrete numbering instance."""

    start: int
    num_fmt: str
    lvl_text: str
    restart_after: Optional[int]


def iter_list_items(document: Document) -> Iterator[ListItem]:
    """Generate a |ListItem| for each list paragraph in document, in document order."""
    return _ListItemEnumerator(document).iter_list_items()


class _ListItemEnumerator:
    """Computes resolved numbering for each list paragraph in a document.

    Counters are kept per numId, so two different numbering instances never share
    counters, and paragraphs are visited in document order, including paragraphs
    inside (possibly nested) tables.
    """

    def __init__(self, document: Document):
        self._document = document
        self._numbering = self._numbering_element(document)
        self._counters: Dict[int, List[Optional[int]]] = {}
        self._level_props_cache: Dict[Tuple[int, int], Optional[_LevelProps]] = {}
        self._style_link_map: Optional[Dict[str, Tuple[int, int]]] = None

    def iter_list_items(self) -> Iterator[ListItem]:
        if self._numbering is None:
            return
        for paragraph in self._iter_paragraphs():
            item = self._list_item_for(paragraph)
            if item is not None:
                yield item

    # -- paragraph enumeration ------------------------------------------

    def _iter_paragraphs(self) -> Iterator[Paragraph]:
        yield from self._iter_block_paragraphs(self._document.iter_inner_content())

    def _iter_block_paragraphs(self, inner_content) -> Iterator[Paragraph]:
        from docx.table import Table

        for item in inner_content:
            if isinstance(item, Table):
                yield from self._iter_table_paragraphs(item)
            else:
                yield item

    def _iter_table_paragraphs(self, table) -> Iterator[Paragraph]:
        visited_cells = set()
        for row in table.rows:
            for cell in row.cells:
                if id(cell._tc) in visited_cells:
                    continue
                visited_cells.add(id(cell._tc))
                yield from self._iter_block_paragraphs(cell.iter_inner_content())

    # -- per-paragraph resolution ----------------------------------------

    def _list_item_for(self, paragraph: Paragraph) -> Optional[ListItem]:
        resolved = self._resolve_num_id_and_ilvl(paragraph)
        if resolved is None:
            return None
        num_id, ilvl = resolved
        props = self._level_props(num_id, ilvl)
        if props is None:
            return None
        counters = self._counters.setdefault(num_id, [None] * 9)
        if props.num_fmt in ("bullet", "none"):
            value = None
        else:
            value = self._advance(counters, num_id, ilvl)
        label, is_supported = self._expand_label(num_id, counters, props)
        return ListItem(paragraph, num_id, ilvl, value, label, props.num_fmt, is_supported)

    def _resolve_num_id_and_ilvl(self, paragraph: Paragraph) -> Optional[Tuple[int, int]]:
        pPr = paragraph._p.pPr
        num_id, ilvl = None, None
        if pPr is not None and pPr.numPr is not None:
            numPr = pPr.numPr
            if numPr.numId is not None:
                num_id = numPr.numId.val
            if numPr.ilvl is not None:
                ilvl = numPr.ilvl.val
        if num_id is None:
            style_num_pr = self._style_num_pr(paragraph)
            if style_num_pr is not None:
                num_id, style_ilvl = style_num_pr
                if ilvl is None:
                    ilvl = style_ilvl
        if num_id is None:
            linked = self._style_linked_num(paragraph)
            if linked is not None:
                num_id, ilvl = linked
        if not num_id:
            return None
        return num_id, ilvl if ilvl is not None else 0

    def _style_num_pr(self, paragraph: Paragraph) -> Optional[Tuple[int, Optional[int]]]:
        for style in self._style_chain(paragraph):
            pPr = style.pPr
            if pPr is None or pPr.numPr is None:
                continue
            numPr = pPr.numPr
            if numPr.numId is None:
                continue
            ilvl = numPr.ilvl.val if numPr.ilvl is not None else None
            return numPr.numId.val, ilvl
        return None

    def _style_linked_num(self, paragraph: Paragraph) -> Optional[Tuple[int, int]]:
        link_map = self._get_style_link_map()
        for style in self._style_chain(paragraph):
            linked = link_map.get(style.styleId)
            if linked is not None:
                return linked
        return None

    def _style_chain(self, paragraph: Paragraph):
        from docx.enum.style import WD_STYLE_TYPE

        styles_elm = self._document.part.styles.element
        style_id = paragraph._p.style
        if style_id is not None:
            style = styles_elm.get_by_id(style_id)
        else:
            style = styles_elm.default_for(WD_STYLE_TYPE.PARAGRAPH)
        visited = set()
        while style is not None and style.styleId not in visited:
            visited.add(style.styleId)
            yield style
            style = style.base_style

    def _get_style_link_map(self) -> Dict[str, Tuple[int, int]]:
        if self._style_link_map is None:
            link_map: Dict[str, Tuple[int, int]] = {}
            for num in self._numbering.num_lst:
                try:
                    abstract_num = self._numbering.abstractNum_having_abstractNumId(
                        num.abstractNumId.val
                    )
                except KeyError:
                    continue
                for lvl in abstract_num.lvl_lst:
                    if lvl.pStyle is not None and lvl.pStyle.val not in link_map:
                        link_map[lvl.pStyle.val] = (num.numId, lvl.ilvl)
            self._style_link_map = link_map
        return self._style_link_map

    # -- counter state ----------------------------------------------------

    def _advance(self, counters: List[Optional[int]], num_id: int, ilvl: int) -> int:
        if counters[ilvl] is None:
            # -- entering a level directly (e.g. jumping into a third-level list)
            #    initializes any untouched ancestor counters to their start values --
            for lvl in range(ilvl):
                if counters[lvl] is None:
                    counters[lvl] = self._start_of(num_id, lvl)
            counters[ilvl] = self._start_of(num_id, ilvl)
        else:
            counters[ilvl] += 1
        for lvl in range(ilvl + 1, 9):
            restart_after = self._restart_after(num_id, lvl)
            if restart_after is None or (restart_after > 0 and ilvl < restart_after):
                counters[lvl] = None
        return counters[ilvl]

    def _start_of(self, num_id: int, ilvl: int) -> int:
        props = self._level_props(num_id, ilvl)
        return props.start if props is not None else 1

    def _restart_after(self, num_id: int, ilvl: int) -> Optional[int]:
        props = self._level_props(num_id, ilvl)
        return props.restart_after if props is not None else None

    # -- numbering-definition lookups -------------------------------------

    @staticmethod
    def _numbering_element(document: Document):
        from docx.opc.constants import RELATIONSHIP_TYPE as RT

        try:
            return document.part.part_related_by(RT.NUMBERING).element
        except KeyError:
            return None

    def _level_props(self, num_id: int, ilvl: int) -> Optional[_LevelProps]:
        key = (num_id, ilvl)
        if key not in self._level_props_cache:
            self._level_props_cache[key] = self._compute_level_props(num_id, ilvl)
        return self._level_props_cache[key]

    def _compute_level_props(self, num_id: int, ilvl: int) -> Optional[_LevelProps]:
        try:
            num = self._numbering.num_having_numId(num_id)
            abstract_num = self._numbering.abstractNum_having_abstractNumId(
                num.abstractNumId.val
            )
        except KeyError:
            return None
        start_override, lvl = None, None
        for lvl_override in num.lvlOverride_lst:
            if lvl_override.ilvl == ilvl:
                if lvl_override.startOverride is not None:
                    start_override = lvl_override.startOverride.val
                lvl = lvl_override.lvl
                break
        if lvl is None:
            lvl = abstract_num.lvl_having_ilvl(ilvl)
        if lvl is None:
            return None
        if start_override is not None:
            start = start_override
        elif lvl.start_val is not None:
            start = lvl.start_val
        else:
            start = 1
        restart_after = lvl.lvlRestart.val if lvl.lvlRestart is not None else None
        return _LevelProps(
            start=start,
            num_fmt=lvl.numFmt_val or "decimal",
            lvl_text=lvl.lvlText_val or "",
            restart_after=restart_after,
        )

    # -- label rendering ---------------------------------------------------

    def _expand_label(
        self, num_id: int, counters: List[Optional[int]], props: _LevelProps
    ) -> Tuple[Optional[str], bool]:
        if props.num_fmt == "none":
            return "", True
        if props.num_fmt == "bullet":
            return props.lvl_text, True
        if props.num_fmt not in _RENDERERS:
            return None, False
        is_supported = True

        def replace(match) -> str:
            nonlocal is_supported
            lvl = int(match.group(1)) - 1
            value = counters[lvl]
            if value is None:
                return ""
            lvl_props = self._level_props(num_id, lvl)
            num_fmt = lvl_props.num_fmt if lvl_props is not None else "decimal"
            renderer = _RENDERERS.get(num_fmt)
            if renderer is None:
                is_supported = False
                return match.group(0)
            return renderer(value)

        label = _LVL_REF_RE.sub(replace, props.lvl_text)
        return (label, True) if is_supported else (None, False)

