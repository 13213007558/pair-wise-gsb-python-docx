# pyright: reportPrivateUsage=false

"""Unit-test suite for the `docx.oxml.notes` module."""

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
            ("w:footnotes/w:footnote{w:id=-1}", 1),
            ("w:footnotes/(w:footnote{w:id=-1},w:footnote{w:id=0})", 1),
            ("w:footnotes/(w:footnote{w:id=-1},w:footnote{w:id=0},w:footnote{w:id=1})", 2),
            ("w:footnotes/(w:footnote{w:id=2},w:footnote{w:id=7})", 8),
            ("w:footnotes/w:footnote{w:id=2147483646}", 2147483647),
            ("w:footnotes/w:footnote{w:id=2147483647}", 1),
            ("w:footnotes/(w:footnote{w:id=1},w:footnote{w:id=2147483647})", 2),
        ],
    )
    def it_finds_the_next_available_footnote_id_to_help(
        self, cxml: str, expected_value: int
    ):
        footnotes_elm = cast(CT_Footnotes, element(cxml))
        assert footnotes_elm._next_available_note_id() == expected_value

    def it_can_add_a_footnote(self):
        footnotes_elm = cast(
            CT_Footnotes,
            element(
                "w:footnotes/(w:footnote{w:type=separator,w:id=-1},"
                "w:footnote{w:type=continuationSeparator,w:id=0})"
            ),
        )

        footnote_elm = footnotes_elm.add_footnote()

        # -- id starts at 1, skipping reserved ids -1 and 0 --
        assert footnote_elm.id == 1
        assert footnote_elm is footnotes_elm.footnote_lst[-1]
        # -- contains a single FootnoteText paragraph with the reference mark --
        assert len(footnote_elm.p_lst) == 1
        p = footnote_elm.p_lst[0]
        assert p.style == "FootnoteText"
        assert p.r_lst[0].style == "FootnoteReference"
        assert p.r_lst[0][-1].tag.endswith("}footnoteRef")

    def it_can_get_a_footnote_by_id(self):
        footnotes_elm = cast(
            CT_Footnotes,
            element("w:footnotes/(w:footnote{w:id=1},w:footnote{w:id=2})"),
        )

        assert footnotes_elm.get_footnote_by_id(2) is footnotes_elm.footnote_lst[1]
        assert footnotes_elm.get_footnote_by_id(3) is None


class DescribeCT_Endnotes:
    """Unit-test suite for `docx.oxml.notes.CT_Endnotes`."""

    @pytest.mark.parametrize(
        ("cxml", "expected_value"),
        [
            ("w:endnotes", 1),
            ("w:endnotes/(w:endnote{w:id=-1},w:endnote{w:id=0})", 1),
            ("w:endnotes/(w:endnote{w:id=3},w:endnote{w:id=4})", 5),
        ],
    )
    def it_finds_the_next_available_endnote_id_to_help(self, cxml: str, expected_value: int):
        endnotes_elm = cast(CT_Endnotes, element(cxml))
        assert endnotes_elm._next_available_note_id() == expected_value

    def it_can_add_an_endnote(self):
        endnotes_elm = cast(
            CT_Endnotes,
            element(
                "w:endnotes/(w:endnote{w:type=separator,w:id=-1},"
                "w:endnote{w:type=continuationSeparator,w:id=0})"
            ),
        )

        endnote_elm = endnotes_elm.add_endnote()

        assert endnote_elm.id == 1
        assert endnote_elm is endnotes_elm.endnote_lst[-1]
        assert len(endnote_elm.p_lst) == 1
        p = endnote_elm.p_lst[0]
        assert p.style == "EndnoteText"
        assert p.r_lst[0].style == "EndnoteReference"
        assert p.r_lst[0][-1].tag.endswith("}endnoteRef")

    def it_can_get_an_endnote_by_id(self):
        endnotes_elm = cast(
            CT_Endnotes,
            element("w:endnotes/(w:endnote{w:id=1},w:endnote{w:id=2})"),
        )

        assert endnotes_elm.get_endnote_by_id(1) is endnotes_elm.endnote_lst[0]
        assert endnotes_elm.get_endnote_by_id(9) is None


class DescribeCT_Note:
    """Unit-test suite for shared `w:footnote`/`w:endnote` element behavior."""

    @pytest.mark.parametrize(
        ("cxml", "expected_value"),
        [
            ("w:footnote{w:id=1}", False),
            ("w:footnote{w:id=1,w:type=normal}", False),
            ("w:footnote{w:id=-1,w:type=separator}", True),
            ("w:footnote{w:id=0,w:type=continuationSeparator}", True),
            ("w:footnote{w:id=2,w:type=continuationNotice}", True),
        ],
    )
    def it_knows_whether_it_is_a_special_note(self, cxml: str, expected_value: bool):
        assert element(cxml).is_special is expected_value
