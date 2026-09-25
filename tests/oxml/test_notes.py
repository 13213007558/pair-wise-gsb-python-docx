# pyright: reportPrivateUsage=false

"""Unit-test suite for `docx.oxml.notes` module."""

from __future__ import annotations

from typing import cast

import pytest

from docx.oxml.notes import CT_Endnotes, CT_Footnotes

from ..unitutil.cxml import element


class DescribeCT_Footnotes:
    """Unit-test suite for `docx.oxml.notes.CT_Footnotes`."""

    @pytest.mark.parametrize(
        ("cxml", "expected_value"),
        [
            ("w:footnotes", 1),
            ("w:footnotes/w:footnote{w:id=-1,w:type=separator}", 1),
            (
                "w:footnotes/(w:footnote{w:id=-1,w:type=separator},"
                "w:footnote{w:id=0,w:type=continuationSeparator})",
                1,
            ),
            ("w:footnotes/w:footnote{w:id=1}", 2),
            ("w:footnotes/(w:footnote{w:id=1},w:footnote{w:id=7})", 8),
            ("w:footnotes/w:footnote{w:id=2147483647}", 1),
        ],
    )
    def it_finds_the_next_available_footnote_id_to_help(self, cxml: str, expected_value: int):
        footnotes_elm = cast(CT_Footnotes, element(cxml))
        assert footnotes_elm._next_available_note_id() == expected_value

    def it_adds_a_footnote_with_a_unique_id_and_reference_mark_to_help(self):
        footnotes_elm = cast(CT_Footnotes, element("w:footnotes/w:footnote{w:id=1}"))

        footnote_elm = footnotes_elm.add_note()

        assert footnote_elm.id == 2
        assert footnote_elm.type is None
        assert footnote_elm.xpath("./w:p/w:r/w:noteRef")
        assert footnote_elm.xpath("./w:p/w:pPr/w:pStyle/@w:val") == ["FootnoteText"]
        assert footnote_elm.xpath("./w:p/w:r/w:rPr/w:rStyle/@w:val") == ["FootnoteReference"]

    def it_finds_a_footnote_by_id_to_help(self):
        footnotes_elm = cast(CT_Footnotes, element("w:footnotes/w:footnote{w:id=42}"))

        assert footnotes_elm.get_note_by_id(42) is footnotes_elm.note_lst[0]
        assert footnotes_elm.get_note_by_id(43) is None


class DescribeCT_Endnotes:
    """Unit-test suite for `docx.oxml.notes.CT_Endnotes`."""

    @pytest.mark.parametrize(
        ("cxml", "expected_value"),
        [
            ("w:endnotes", 1),
            (
                "w:endnotes/(w:endnote{w:id=-1,w:type=separator},"
                "w:endnote{w:id=0,w:type=continuationSeparator})",
                1,
            ),
            ("w:endnotes/w:endnote{w:id=3}", 4),
        ],
    )
    def it_finds_the_next_available_endnote_id_to_help(self, cxml: str, expected_value: int):
        endnotes_elm = cast(CT_Endnotes, element(cxml))
        assert endnotes_elm._next_available_note_id() == expected_value

    def it_adds_an_endnote_with_a_unique_id_and_reference_mark_to_help(self):
        endnotes_elm = cast(CT_Endnotes, element("w:endnotes/w:endnote{w:id=1}"))

        endnote_elm = endnotes_elm.add_note()

        assert endnote_elm.id == 2
        assert endnote_elm.xpath("./w:p/w:r/w:noteRef")
        assert endnote_elm.xpath("./w:p/w:pPr/w:pStyle/@w:val") == ["EndnoteText"]
        assert endnote_elm.xpath("./w:p/w:r/w:rPr/w:rStyle/@w:val") == ["EndnoteReference"]
