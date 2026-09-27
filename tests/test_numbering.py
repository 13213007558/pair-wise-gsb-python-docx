"""Test suite for the docx.numbering module."""

from __future__ import annotations

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.numbering import iter_list_items
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml.ns import nsdecls, qn
from docx.oxml.parser import parse_xml

# ---------------------------------------------------------------------------
# fixture helpers
# ---------------------------------------------------------------------------


def _numbering_elm(document):
    return document.part.part_related_by(RT.NUMBERING).element


def _add_abstract_num(document, abstract_num_id, levels):
    """Append an abstract numbering definition built from levels.

    levels is a list of (ilvl, num_fmt, lvl_text, start) tuples. Uses
    abstractNumIds in the 100+ range to avoid clashing with the default
    template definitions.
    """
    lvl_xml = "".join(
        '<w:lvl w:ilvl="%d">'
        '<w:start w:val="%d"/>'
        '<w:numFmt w:val="%s"/>'
        '<w:lvlText w:val="%s"/>'
        "</w:lvl>" % (ilvl, start, num_fmt, lvl_text)
        for ilvl, num_fmt, lvl_text, start in levels
    )
    abstract_elm = parse_xml(
        '<w:abstractNum %s w:abstractNumId="%d">%s</w:abstractNum>'
        % (nsdecls("w"), abstract_num_id, lvl_xml)
    )
    numbering_elm = _numbering_elm(document)
    first_num = numbering_elm.find(qn("w:num"))
    if first_num is not None:
        first_num.addprevious(abstract_elm)
    else:
        numbering_elm.append(abstract_elm)


def _add_num(document, num_id, abstract_num_id, start_overrides=None):
    """Append a concrete w:num list instance, optionally with startOverride values.

    start_overrides is a dict of {ilvl: start_value} rendered as
    w:lvlOverride/w:startOverride elements.
    """
    overrides_xml = "".join(
        '<w:lvlOverride w:ilvl="%d"><w:startOverride w:val="%d"/></w:lvlOverride>' % (ilvl, start)
        for ilvl, start in (start_overrides or {}).items()
    )
    num_elm = parse_xml(
        '<w:num %s w:numId="%d">'
        '<w:abstractNumId w:val="%d"/>'
        "%s"
        "</w:num>" % (nsdecls("w"), num_id, abstract_num_id, overrides_xml)
    )
    _numbering_elm(document).append(num_elm)


def _set_num_pr(paragraph, num_id=None, ilvl=None):
    """Apply direct numbering properties to paragraph."""
    num_pr = paragraph._p.get_or_add_pPr().get_or_add_numPr()
    if ilvl is not None:
        num_pr.get_or_add_ilvl().val = ilvl
    if num_id is not None:
        num_pr.get_or_add_numId().val = num_id


def _add_numbered_paragraph(container, text, num_id, ilvl=None, style=None):
    paragraph = container.add_paragraph(text, style=style)
    if num_id is not None or ilvl is not None:
        _set_num_pr(paragraph, num_id=num_id, ilvl=ilvl)
    return paragraph


def _labels(document):
    return [item.label for item in iter_list_items(document)]


# ---------------------------------------------------------------------------
# shared abstract definitions (ids 100+ avoid the default template's 0-8)
# ---------------------------------------------------------------------------

# multilevel: "1." / "1.1." / "1.1.a"
MULTI_ABSTRACT = 100
MULTI_LEVELS = [
    (0, "decimal", "%1.", 1),
    (1, "decimal", "%1.%2.", 1),
    (2, "lowerLetter", "%1.%2.%3", 1),
]

# single level upper-letter: "A)"
LETTER_ABSTRACT = 101
LETTER_LEVELS = [(0, "upperLetter", "%1)", 1)]

# single level bullet
BULLET_ABSTRACT = 102
BULLET_LEVELS = [(0, "bullet", "\u2022", 1)]

# two-level roman: "I." / "i."
ROMAN_ABSTRACT = 103
ROMAN_LEVELS = [
    (0, "upperRoman", "%1.", 1),
    (1, "lowerRoman", "%2.", 1),
]

# unsupported format
CARDINAL_ABSTRACT = 104
CARDINAL_LEVELS = [(0, "cardinalText", "%1.", 1)]


def _document_with_shared_definitions():
    document = Document()
    _add_abstract_num(document, MULTI_ABSTRACT, MULTI_LEVELS)
    _add_abstract_num(document, LETTER_ABSTRACT, LETTER_LEVELS)
    _add_abstract_num(document, BULLET_ABSTRACT, BULLET_LEVELS)
    _add_abstract_num(document, ROMAN_ABSTRACT, ROMAN_LEVELS)
    _add_abstract_num(document, CARDINAL_ABSTRACT, CARDINAL_LEVELS)
    return document


# numIds used across tests
NUM_MULTI = 100
NUM_MULTI_RESTART = 101
NUM_LETTER = 102
NUM_BULLET = 103
NUM_ROMAN = 104
NUM_CARDINAL = 105


class DescribeIterListItems:
    def it_keeps_independent_counters_for_interleaved_lists(self):
        """Two lists interleaved paragraph-by-paragraph each keep their own count."""
        document = _document_with_shared_definitions()
        _add_num(document, NUM_MULTI, MULTI_ABSTRACT)
        _add_num(document, NUM_LETTER, LETTER_ABSTRACT)

        _add_numbered_paragraph(document, "multi-1", NUM_MULTI, ilvl=0)
        _add_numbered_paragraph(document, "letter-1", NUM_LETTER, ilvl=0)
        _add_numbered_paragraph(document, "multi-2", NUM_MULTI, ilvl=0)
        _add_numbered_paragraph(document, "multi-2.1", NUM_MULTI, ilvl=1)
        _add_numbered_paragraph(document, "letter-2", NUM_LETTER, ilvl=0)
        _add_numbered_paragraph(document, "multi-2.2", NUM_MULTI, ilvl=1)
        _add_numbered_paragraph(document, "multi-3", NUM_MULTI, ilvl=0)
        _add_numbered_paragraph(document, "letter-3", NUM_LETTER, ilvl=0)

        items = list(iter_list_items(document))
        assert [item.paragraph.text for item in items] == [
            "multi-1",
            "letter-1",
            "multi-2",
            "multi-2.1",
            "letter-2",
            "multi-2.2",
            "multi-3",
            "letter-3",
        ]
        assert [item.label for item in items] == [
            "1.",
            "A)",
            "2.",
            "2.1.",
            "B)",
            "2.2.",
            "3.",
            "C)",
        ]
        assert [item.level for item in items] == [0, 0, 0, 1, 0, 1, 0, 0]
        assert [item.num_id for item in items] == [
            NUM_MULTI,
            NUM_LETTER,
            NUM_MULTI,
            NUM_MULTI,
            NUM_LETTER,
            NUM_MULTI,
            NUM_MULTI,
            NUM_LETTER,
        ]
        assert {item.abstract_num_id for item in items} == {MULTI_ABSTRACT, LETTER_ABSTRACT}

    def it_restarts_deeper_levels_when_a_higher_level_advances(self):
        """Advancing level 0 restarts the level-1 and level-2 counters."""
        document = _document_with_shared_definitions()
        _add_num(document, NUM_MULTI, MULTI_ABSTRACT)

        _add_numbered_paragraph(document, "p1", NUM_MULTI, ilvl=0)
        _add_numbered_paragraph(document, "p1.1", NUM_MULTI, ilvl=1)
        _add_numbered_paragraph(document, "p1.1.a", NUM_MULTI, ilvl=2)
        _add_numbered_paragraph(document, "p1.1.b", NUM_MULTI, ilvl=2)
        _add_numbered_paragraph(document, "p1.2", NUM_MULTI, ilvl=1)
        _add_numbered_paragraph(document, "p2", NUM_MULTI, ilvl=0)
        _add_numbered_paragraph(document, "p2.1", NUM_MULTI, ilvl=1)
        _add_numbered_paragraph(document, "p2.2", NUM_MULTI, ilvl=1)
        _add_numbered_paragraph(document, "p2.2.a", NUM_MULTI, ilvl=2)
        _add_numbered_paragraph(document, "p2.3", NUM_MULTI, ilvl=1)
        _add_numbered_paragraph(document, "p2.3.a", NUM_MULTI, ilvl=2)

        assert _labels(document) == [
            "1.",
            "1.1.",
            "1.1.a",
            "1.1.b",
            "1.2.",
            "2.",
            "2.1.",
            "2.2.",
            "2.2.a",
            "2.3.",
            "2.3.a",
        ]

    def it_restarts_a_list_mid_document_via_start_override(self):
        """A second numId with lvlOverride/startOverride restarts the abstract list."""
        document = _document_with_shared_definitions()
        _add_num(document, NUM_MULTI, MULTI_ABSTRACT)
        _add_num(document, NUM_MULTI_RESTART, MULTI_ABSTRACT, start_overrides={0: 1})

        _add_numbered_paragraph(document, "first-1", NUM_MULTI, ilvl=0)
        _add_numbered_paragraph(document, "first-2", NUM_MULTI, ilvl=0)
        _add_numbered_paragraph(document, "first-3", NUM_MULTI, ilvl=0)
        # mid-document restart: same abstract definition, fresh numId + startOverride
        _add_numbered_paragraph(document, "restarted-1", NUM_MULTI_RESTART, ilvl=0)
        _add_numbered_paragraph(document, "restarted-2", NUM_MULTI_RESTART, ilvl=0)
        # the original list is unaffected and continues where it left off
        _add_numbered_paragraph(document, "first-4", NUM_MULTI, ilvl=0)

        assert _labels(document) == ["1.", "2.", "3.", "1.", "2.", "4."]

    def it_inherits_numbering_through_the_style_basedOn_chain(self):
        """A paragraph with no direct numPr is numbered via its style's base style."""
        document = _document_with_shared_definitions()
        _add_num(document, NUM_LETTER, LETTER_ABSTRACT)

        styles = document.styles
        base_style = styles.add_style("ListBase", WD_STYLE_TYPE.PARAGRAPH)
        base_num_pr = base_style.element.get_or_add_pPr().get_or_add_numPr()
        base_num_pr.get_or_add_numId().val = NUM_LETTER
        derived_style = styles.add_style("ListDerived", WD_STYLE_TYPE.PARAGRAPH)
        derived_style.base_style = base_style

        _add_numbered_paragraph(document, "base-1", None, style="ListBase")
        _add_numbered_paragraph(document, "derived-1", None, style="ListDerived")
        _add_numbered_paragraph(document, "derived-2", None, style="ListDerived")

        items = list(iter_list_items(document))
        assert [item.paragraph.text for item in items] == ["base-1", "derived-1", "derived-2"]
        assert [item.label for item in items] == ["A)", "B)", "C)"]
        assert all(item.num_id == NUM_LETTER for item in items)

    def it_renders_bullet_labels_as_literal_lvl_text(self):
        document = _document_with_shared_definitions()
        _add_num(document, NUM_BULLET, BULLET_ABSTRACT)

        _add_numbered_paragraph(document, "bullet-1", NUM_BULLET, ilvl=0)
        _add_numbered_paragraph(document, "bullet-2", NUM_BULLET, ilvl=0)

        items = list(iter_list_items(document))
        assert [item.label for item in items] == ["\u2022", "\u2022"]
        assert all(item.is_bullet for item in items)
        assert all(item.num_fmt == "bullet" for item in items)

    def it_renders_roman_numeral_formats(self):
        document = _document_with_shared_definitions()
        _add_num(document, NUM_ROMAN, ROMAN_ABSTRACT)

        for ilvl in [0, 0, 1, 1, 0, 1]:
            _add_numbered_paragraph(document, "roman", NUM_ROMAN, ilvl=ilvl)

        assert _labels(document) == ["I.", "II.", "i.", "ii.", "III.", "i."]

    def it_reports_a_reason_for_unsupported_num_fmt_instead_of_using_decimal(self):
        document = _document_with_shared_definitions()
        _add_num(document, NUM_CARDINAL, CARDINAL_ABSTRACT)

        _add_numbered_paragraph(document, "cardinal-1", NUM_CARDINAL, ilvl=0)
        _add_numbered_paragraph(document, "cardinal-2", NUM_CARDINAL, ilvl=0)

        items = list(iter_list_items(document))
        assert len(items) == 2
        for item in items:
            assert item.label is None
            assert "cardinalText" in item.unsupported_reason
            assert "not supported" in item.unsupported_reason

    def it_treats_numId_zero_as_cancelled_without_advancing_counters(self):
        document = _document_with_shared_definitions()
        _add_num(document, NUM_MULTI, MULTI_ABSTRACT)

        _add_numbered_paragraph(document, "item-1", NUM_MULTI, ilvl=0)
        _add_numbered_paragraph(document, "cancelled", 0, ilvl=0)
        _add_numbered_paragraph(document, "item-2", NUM_MULTI, ilvl=0)

        items = list(iter_list_items(document))
        assert [item.paragraph.text for item in items] == ["item-1", "item-2"]
        assert [item.label for item in items] == ["1.", "2."]

    def it_walks_table_paragraphs_in_document_order(self):
        document = _document_with_shared_definitions()
        _add_num(document, NUM_MULTI, MULTI_ABSTRACT)

        _add_numbered_paragraph(document, "before-table", NUM_MULTI, ilvl=0)
        table = document.add_table(rows=2, cols=1)
        _set_num_pr(table.cell(0, 0).paragraphs[0], num_id=NUM_MULTI, ilvl=0)
        table.cell(0, 0).paragraphs[0].text = "in-table-1"
        _set_num_pr(table.cell(1, 0).paragraphs[0], num_id=NUM_MULTI, ilvl=1)
        table.cell(1, 0).paragraphs[0].text = "in-table-1.1"
        _add_numbered_paragraph(document, "after-table", NUM_MULTI, ilvl=0)

        items = list(iter_list_items(document))
        assert [item.paragraph.text for item in items] == [
            "before-table",
            "in-table-1",
            "in-table-1.1",
            "after-table",
        ]
        assert [item.label for item in items] == ["1.", "2.", "2.1.", "3."]

    def it_reports_a_reason_for_an_unknown_num_id(self):
        document = _document_with_shared_definitions()
        _add_numbered_paragraph(document, "dangling", 999, ilvl=0)

        (item,) = list(iter_list_items(document))
        assert item.label is None
        assert "999" in item.unsupported_reason

    def it_skips_unnumbered_paragraphs(self):
        document = _document_with_shared_definitions()
        _add_num(document, NUM_MULTI, MULTI_ABSTRACT)

        document.add_paragraph("plain paragraph")
        _add_numbered_paragraph(document, "item-1", NUM_MULTI, ilvl=0)
        document.add_paragraph("another plain one")
        _add_numbered_paragraph(document, "item-2", NUM_MULTI, ilvl=0)

        assert _labels(document) == ["1.", "2."]

    def it_does_not_modify_numbering_xml_or_the_document(self):
        document = _document_with_shared_definitions()
        _add_num(document, NUM_MULTI, MULTI_ABSTRACT)
        _add_num(document, NUM_LETTER, LETTER_ABSTRACT)
        _add_numbered_paragraph(document, "item-1", NUM_MULTI, ilvl=0)
        _add_numbered_paragraph(document, "item-a", NUM_LETTER, ilvl=0)

        numbering_part = document.part.part_related_by(RT.NUMBERING)
        numbering_blob_before = numbering_part.blob
        document_xml_before = document.element.xml

        assert _labels(document) == ["1.", "A)"]

        assert numbering_part.blob == numbering_blob_before
        assert document.element.xml == document_xml_before

    def it_yields_nothing_for_a_document_without_a_numbering_part(self):
        document = Document()
        rels = document.part.rels
        numbering_rIds = [rId for rId, rel in rels.items() if rel.reltype == RT.NUMBERING]
        for rId in numbering_rIds:
            document.part.drop_rel(rId)
        document.add_paragraph("plain")

        assert list(iter_list_items(document)) == []
