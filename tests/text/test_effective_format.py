"""Tests for read-only effective character-format resolution."""

from __future__ import annotations

from types import SimpleNamespace
from typing import cast

from docx.oxml.text.run import CT_R
from docx.text.run import Run

from ..unitutil.cxml import element


class DescribeEffectiveCharacterFormat:
    def it_resolves_document_defaults_and_marks_sources(self):
        run = self._run(
            'w:r/w:t"foo"',
            "w:styles/("
            "w:docDefaults/w:rPrDefault/w:rPr/("
            "w:rFonts{w:ascii=Arial,w:eastAsia=SimSun,w:hAnsi=Arial,w:cs=Traditional Arabic},"
            "w:sz{w:val=24},w:szCs{w:val=21}),"
            "w:style{w:type=paragraph,w:default=1,w:styleId=Normal},"
            "w:style{w:type=character,w:default=1,w:styleId=DefaultChar},"
            "w:style{w:type=table,w:default=1,w:styleId=TableNormal})",
        )

        effective = run.effective_character_format

        assert effective["b"] is False
        assert effective["sz"].pt == 12
        assert effective["szCs"].pt == 10.5
        assert effective.fonts["ascii"].value == "Arial"
        assert effective.fonts["ascii"].is_theme is False
        assert effective.fonts["eastAsia"].value == "SimSun"
        assert effective.fonts["cs"].value == "Traditional Arabic"
        assert effective.source("b").kind == "specification-default"
        assert effective.source("sz").kind == "document-default"
        assert "rFonts/@w:ascii" in effective.fonts["ascii"].source.path

    def it_keeps_theme_font_mappings_symbolic(self):
        run = self._run(
            'w:r/w:t"foo"',
            "w:styles/(w:docDefaults/w:rPrDefault/w:rPr/w:rFonts{w:asciiTheme=majorHAnsi,w:eastAsiaTheme=majorEastAsia,w:hAnsiTheme=majorHAnsi,w:cstheme=majorBidi},w:style{w:type=paragraph,w:default=1,w:styleId=Normal},w:style{w:type=character,w:default=1,w:styleId=DefaultChar},w:style{w:type=table,w:default=1,w:styleId=TableNormal})",
        )
        effective = run.effective_character_format
        assert effective.fonts["ascii"].value == "majorHAnsi"
        assert effective.fonts["ascii"].is_theme is True
        assert effective.fonts["eastAsia"].value == "majorEastAsia"
        assert effective.fonts["eastAsia"].is_theme is True
        assert effective.fonts["cs"].value == "majorBidi"
        assert effective.fonts["cs"].is_theme is True

    def it_distinguishes_a_toggle_reversal_from_direct_explicit_off(self):
        run = self._run(
            'w:r/(w:rPr/(w:rStyle{w:val=Unbold},w:b{w:val=0}),w:t"x")',
            "w:styles/(w:style{w:type=paragraph,w:default=1,w:styleId=Normal}/w:rPr/w:b,w:style{w:type=character,w:default=1,w:styleId=DefaultChar},w:style{w:type=table,w:default=1,w:styleId=TableNormal},w:style{w:type=character,w:styleId=Unbold}/(w:basedOn{w:val=DefaultChar},w:rPr/w:b{w:val=0}))",
        )
        effective = run.effective_character_format
        toggle_source = effective.source("b").overridden_from
        assert effective["b"] is False
        assert effective.source("b").kind == "direct"
        assert effective.source("b").mode == "explicit"
        assert toggle_source.kind == "style"
        assert toggle_source.style_id == "Unbold"
        assert toggle_source.mode == "toggle"
        assert toggle_source.inherited_from.kind == "style"
        assert toggle_source.inherited_from.style_id == "Normal"

    def it_applies_table_first_row_conditional_toggle_and_character_style_toggle(self):
        run = self._run(
            'w:tbl/(w:tblPr/(w:tblStyle{w:val=Table},w:tblLook{w:firstRow=1}),w:tblGrid/w:gridCol,w:tr/w:tc/w:p/w:r/(w:rPr/w:rStyle{w:val=Unbold},w:t"x"))',
            "w:styles/(w:style{w:type=paragraph,w:default=1,w:styleId=Normal},w:style{w:type=character,w:default=1,w:styleId=DefaultChar},w:style{w:type=table,w:default=1,w:styleId=TableNormal},w:style{w:type=table,w:styleId=Table}/w:tblStylePr{w:type=firstRow}/w:rPr/w:b,w:style{w:type=character,w:styleId=Unbold}/(w:basedOn{w:val=DefaultChar},w:rPr/w:b{w:val=0}))",
        )
        effective = run.effective_character_format
        source = effective.source("b")
        assert effective["b"] is False
        assert source.kind == "style"
        assert source.style_id == "Unbold"
        assert source.mode == "toggle"
        assert source.inherited_from.kind == "table-conditional-style"
        assert source.inherited_from.condition == "firstRow"

    def it_uses_recorded_table_cnf_style_markers(self):
        run = self._run(
            "w:tbl/(w:tblPr/(w:tblStyle{w:val=Table},w:tblLook{w:firstRow=0}),"
            "w:tblGrid/w:gridCol,w:tr/(w:trPr/w:cnfStyle{w:firstRow=1},"
            'w:tc/w:p/w:r/w:t"x"))',
            "w:styles/(w:style{w:type=paragraph,w:default=1,w:styleId=Normal},"
            "w:style{w:type=character,w:default=1,w:styleId=DefaultChar},"
            "w:style{w:type=table,w:default=1,w:styleId=TableNormal},"
            "w:style{w:type=table,w:styleId=Table}/"
            "w:tblStylePr{w:type=firstRow}/w:rPr/w:b)",
        )

        effective = run.effective_character_format

        assert effective["b"] is False
        assert effective.source("b").kind == "specification-default"

    def it_lets_child_table_style_unconditional_override_parent_condition(self):
        run = self._run(
            "w:tbl/(w:tblPr/(w:tblStyle{w:val=Child},w:tblLook{w:firstRow=1}),"
            'w:tblGrid/w:gridCol,w:tr/w:tc/w:p/w:r/w:t"x")',
            "w:styles/(w:style{w:type=paragraph,w:default=1,w:styleId=Normal},"
            "w:style{w:type=character,w:default=1,w:styleId=DefaultChar},"
            "w:style{w:type=table,w:default=1,w:styleId=TableNormal},"
            "w:style{w:type=table,w:styleId=Parent}/"
            "w:tblStylePr{w:type=firstRow}/w:rPr/w:b,"
            "w:style{w:type=table,w:styleId=Child}/(w:basedOn{w:val=Parent},w:rPr/w:b{w:val=0}))",
        )

        effective = run.effective_character_format

        assert effective["b"] is False
        assert effective.source("b").style_id == "Child"
        assert effective.source("b").condition is None

    def it_reports_style_reference_cycles_without_raising(self):
        run = self._run(
            'w:r/(w:rPr/w:rStyle{w:val=A},w:t"x")',
            "w:styles/(w:style{w:type=paragraph,w:default=1,w:styleId=Normal},w:style{w:type=table,w:default=1,w:styleId=TableNormal},w:style{w:type=character,w:default=1,w:styleId=DefaultChar},w:style{w:type=character,w:styleId=A}/(w:basedOn{w:val=B},w:rPr/w:i),w:style{w:type=character,w:styleId=B}/w:basedOn{w:val=A})",
        )
        effective = run.effective_character_format
        assert effective.diagnostics
        assert effective.diagnostics[0].code == "cycle"
        assert "A -> B -> A" in effective.diagnostics[0].message

    def it_is_read_only(self):
        r_cxml = 'w:r/(w:rPr/w:rStyle{w:val=Strong},w:t"x")'
        styles_cxml = (
            "w:styles/(w:style{w:type=paragraph,w:default=1,w:styleId=Normal},"
            "w:style{w:type=character,w:default=1,w:styleId=DefaultChar},"
            "w:style{w:type=table,w:default=1,w:styleId=TableNormal},"
            "w:style{w:type=character,w:styleId=Strong}/w:rPr/w:b)"
        )
        r = cast(CT_R, element(r_cxml))
        styles = element(styles_cxml)
        run = Run(r, self._parent(styles))
        before = r.xml
        run.effective_character_format
        assert r.xml == before
        assert styles.xml == element(styles_cxml).xml

    def _run(self, run_cxml: str, styles_cxml: str) -> Run:
        root = element(run_cxml)
        r = root if root.tag.endswith("}r") else root.xpath(".//w:r")[0]
        return Run(cast(CT_R, r), self._parent(element(styles_cxml)))

    def _parent(self, styles):
        part = SimpleNamespace(part_related_by=lambda reltype: SimpleNamespace(element=styles))
        return SimpleNamespace(part=part)
