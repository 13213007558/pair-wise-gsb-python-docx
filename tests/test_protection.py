# pyright: reportPrivateUsage=false

"""Unit test suite for the docx.protection module."""

from __future__ import annotations

import pytest

from docx import Document
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.protection import (
    IssueKind,
    ProtectionMode,
    SegmentStatus,
    analyze_protection,
)

_W = nsdecls("w")

READ_ONLY_ENFORCED = '<w:documentProtection %s w:edit="readOnly" w:enforcement="1"/>' % _W
READ_ONLY_NOT_ENFORCED = '<w:documentProtection %s w:edit="readOnly" w:enforcement="0"/>' % _W


def _run(text: str) -> str:
    return '<w:r><w:t xml:space="preserve">%s</w:t></w:r>' % text


def _p(inner: str) -> str:
    return "<w:p %s>%s</w:p>" % (_W, inner)


def _perm_start(perm_id: str, attrs: str = "") -> str:
    return '<w:permStart w:id="%s"%s/>' % (perm_id, (" " + attrs) if attrs else "")


def _perm_end(perm_id: str) -> str:
    return '<w:permEnd w:id="%s"/>' % perm_id


def _document(*body_fragments: str, protection: "str | None" = None) -> Document:
    document = Document()
    body = document.element.body
    sect_pr = body.find(qn("w:sectPr"))
    for fragment in body_fragments:
        el = parse_xml(fragment)
        if sect_pr is not None:
            sect_pr.addprevious(el)
        else:
            body.append(el)
    if protection is not None:
        document.settings.element.append(parse_xml(protection))
    return document


def _table_middle_column_open() -> str:
    """3-column table whose middle column is an everyone-editable range."""
    return (
        "<w:tbl %s>"
        "<w:tblGrid>"
        '<w:gridCol w:w="1440"/><w:gridCol w:w="1440"/><w:gridCol w:w="1440"/>'
        "</w:tblGrid>"
        "<w:tr>"
        "<w:tc>%s</w:tc>"
        "<w:tc>%s</w:tc>"
        "<w:tc>%s</w:tc>"
        "</w:tr>"
        "</w:tbl>"
    ) % (
        _W,
        _p(_run("A0")),
        _p(
            _perm_start("5", 'w:edGrp="everyone" w:colFirst="1" w:colLast="1"')
            + _run("B0")
            + _perm_end("5")
        ),
        _p(_run("C0")),
    )


class DescribeDocumentProtectionParsing:
    def it_reports_no_protection_when_settings_lacks_documentProtection(self):
        document = _document(_p(_run("hello")))

        report = analyze_protection(document, "alice")

        assert report.protection.mode is ProtectionMode.NONE
        assert report.read_only_enforced is False
        assert [s.status for s in report.segments] == [SegmentStatus.NOT_PROTECTED]

    def it_reports_not_enforced_when_enforcement_is_off(self):
        document = _document(_p(_run("hello")), protection=READ_ONLY_NOT_ENFORCED)

        report = analyze_protection(document, "alice")

        assert report.protection.mode is ProtectionMode.READ_ONLY
        assert report.protection.enforced is False
        assert report.read_only_enforced is False
        assert [s.status for s in report.segments] == [SegmentStatus.NOT_PROTECTED]

    def it_distinguishes_unprotected_from_editable_by_exception(self):
        document = _document(
            _p(_perm_start("1", 'w:edGrp="everyone"') + _run("mid") + _perm_end("1")),
            protection=READ_ONLY_ENFORCED,
        )

        enforced = analyze_protection(document, "alice")
        assert [s.status for s in enforced.segments] == [SegmentStatus.EDITABLE_EXCEPTION]

        not_enforced = analyze_protection(
            _document(
                _p(_perm_start("1", 'w:edGrp="everyone"') + _run("mid") + _perm_end("1")),
                protection=READ_ONLY_NOT_ENFORCED,
            ),
            "alice",
        )
        assert [s.status for s in not_enforced.segments] == [SegmentStatus.NOT_PROTECTED]


class DescribeSegmentation:
    def it_partitions_paragraph_text_around_a_permission_range(self):
        document = _document(
            _p(
                _run("before ")
                + _perm_start("1", 'w:ed="alice"')
                + _run("middle")
                + _perm_end("1")
                + _run(" after")
            ),
            protection=READ_ONLY_ENFORCED,
        )

        report = analyze_protection(document, "alice")

        assert [(s.text, s.status) for s in report.segments] == [
            ("before ", SegmentStatus.RESTRICTED),
            ("middle", SegmentStatus.EDITABLE_EXCEPTION),
            (" after", SegmentStatus.RESTRICTED),
        ]
        middle = report.segments[1]
        assert middle.active_perm_ids == ("1",)
        assert middle.granting_perm_ids == ("1",)
        assert (middle.start_offset, middle.end_offset) == (7, 13)

    def it_handles_interleaved_ranges_spanning_paragraphs(self):
        document = _document(
            _p(_perm_start("1", 'w:edGrp="editors"') + _run("alpha")),
            _p(_perm_start("2", 'w:ed="alice"') + _run("bravo") + _perm_end("1")),
            _p(_run("charlie") + _perm_end("2")),
            protection=READ_ONLY_ENFORCED,
        )

        alice = analyze_protection(document, "alice")
        assert [(s.text, s.status, s.granting_perm_ids) for s in alice.segments] == [
            ("alpha", SegmentStatus.RESTRICTED, ()),
            ("bravo", SegmentStatus.EDITABLE_EXCEPTION, ("2",)),
            ("charlie", SegmentStatus.EDITABLE_EXCEPTION, ("2",)),
        ]
        assert [s.active_perm_ids for s in alice.segments] == [("1",), ("1", "2"), ("2",)]

        editor = analyze_protection(document, "carol", ("editors",))
        assert [(s.text, s.status) for s in editor.segments] == [
            ("alpha", SegmentStatus.EDITABLE_EXCEPTION),
            ("bravo", SegmentStatus.EDITABLE_EXCEPTION),
            ("charlie", SegmentStatus.RESTRICTED),
        ]

    def it_limits_a_table_permission_to_the_middle_column(self):
        document = _document(_table_middle_column_open(), protection=READ_ONLY_ENFORCED)

        report = analyze_protection(document, "alice")

        assert [(s.location, s.text, s.status) for s in report.segments] == [
            ("body/tbl[0]/tr[0]/tc[0]/p[0]", "A0", SegmentStatus.RESTRICTED),
            ("body/tbl[0]/tr[0]/tc[1]/p[0]", "B0", SegmentStatus.EDITABLE_EXCEPTION),
            ("body/tbl[0]/tr[0]/tc[2]/p[0]", "C0", SegmentStatus.RESTRICTED),
        ]
        assert report.issues == []

    def it_distinguishes_adjacent_paragraphs_with_different_permissions(self):
        document = _document(
            _p(_perm_start("1", 'w:ed="bob"') + _run("first") + _perm_end("1")),
            _p(_perm_start("2", 'w:edGrp="editors"') + _run("second") + _perm_end("2")),
            protection=READ_ONLY_ENFORCED,
        )

        bob = analyze_protection(document, "bob")
        assert [(s.text, s.status) for s in bob.segments] == [
            ("first", SegmentStatus.EDITABLE_EXCEPTION),
            ("second", SegmentStatus.RESTRICTED),
        ]

        editor = analyze_protection(document, "carol", ("editors",))
        assert [(s.text, s.status) for s in editor.segments] == [
            ("first", SegmentStatus.RESTRICTED),
            ("second", SegmentStatus.EDITABLE_EXCEPTION),
        ]

    def it_grants_everyone_and_never_grants_none_group(self):
        document = _document(
            _p(_perm_start("1", 'w:edGrp="everyone"') + _run("all") + _perm_end("1")),
            _p(_perm_start("2", 'w:edGrp="none"') + _run("nobody") + _perm_end("2")),
            protection=READ_ONLY_ENFORCED,
        )

        report = analyze_protection(document, "alice")

        assert [(s.text, s.status) for s in report.segments] == [
            ("all", SegmentStatus.EDITABLE_EXCEPTION),
            ("nobody", SegmentStatus.RESTRICTED),
        ]


class DescribeTraceability:
    def it_traces_ranges_and_segments_back_to_their_markers(self):
        document = _document(
            _p(_perm_start("7", 'w:ed="alice"') + _run("x") + _perm_end("7")),
            protection=READ_ONLY_ENFORCED,
        )

        report = analyze_protection(document, "alice")

        assert len(report.permission_ranges) == 1
        pr = report.permission_ranges[0]
        assert pr.perm_id == "7"
        assert pr.ed == "alice"
        assert pr.start.location == "body/p[0]"
        assert pr.end.location == "body/p[0]"
        segment = report.segments[0]
        assert segment.active_perm_ids == ("7",)
        assert segment.granting_perm_ids == ("7",)


class DescribeIssueReporting:
    def it_reports_an_unpaired_perm_start_and_does_not_grant(self):
        document = _document(
            _p(_perm_start("9", 'w:edGrp="everyone"') + _run("tail")),
            protection=READ_ONLY_ENFORCED,
        )

        report = analyze_protection(document, "alice")

        assert [i.kind for i in report.issues] == [IssueKind.UNPAIRED_PERM_START]
        assert report.issues[0].perm_id == "9"
        assert report.issues[0].location == "body/p[0]"
        assert [s.status for s in report.segments] == [SegmentStatus.RESTRICTED]

    def it_reports_an_unpaired_perm_end(self):
        document = _document(
            _p(_run("text") + _perm_end("9")),
            protection=READ_ONLY_ENFORCED,
        )

        report = analyze_protection(document, "alice")

        assert [i.kind for i in report.issues] == [IssueKind.UNPAIRED_PERM_END]
        assert report.issues[0].perm_id == "9"

    def it_reports_an_unknown_editor_group_and_does_not_grant(self):
        document = _document(
            _p(_perm_start("3", 'w:edGrp="bogus"') + _run("x") + _perm_end("3")),
            protection=READ_ONLY_ENFORCED,
        )

        report = analyze_protection(document, "alice", ("bogus",))

        assert [i.kind for i in report.issues] == [IssueKind.UNKNOWN_EDITOR_GROUP]
        assert report.permission_ranges[0].valid is False
        assert [s.status for s in report.segments] == [SegmentStatus.RESTRICTED]

    def it_reports_a_perm_start_without_any_editor(self):
        document = _document(
            _p(_perm_start("4") + _run("x") + _perm_end("4")),
            protection=READ_ONLY_ENFORCED,
        )

        report = analyze_protection(document, "alice")

        assert [i.kind for i in report.issues] == [IssueKind.MISSING_EDITOR]
        assert [s.status for s in report.segments] == [SegmentStatus.RESTRICTED]

    @pytest.mark.parametrize(
        "attrs",
        [
            'w:edGrp="everyone" w:colFirst="2" w:colLast="1"',
            'w:edGrp="everyone" w:colFirst="0" w:colLast="5"',
            'w:edGrp="everyone" w:colFirst="1"',
            'w:edGrp="everyone" w:colFirst="-1" w:colLast="1"',
        ],
    )
    def it_reports_invalid_column_ranges_in_tables(self, attrs: str):
        table = (
            "<w:tbl %s>"
            "<w:tblGrid>"
            '<w:gridCol w:w="1440"/><w:gridCol w:w="1440"/><w:gridCol w:w="1440"/>'
            "</w:tblGrid>"
            "<w:tr><w:tc>%s</w:tc></w:tr>"
            "</w:tbl>"
        ) % (_W, _p(_perm_start("6", attrs) + _run("x") + _perm_end("6")))
        document = _document(table, protection=READ_ONLY_ENFORCED)

        report = analyze_protection(document, "alice")

        assert [i.kind for i in report.issues] == [IssueKind.INVALID_COLUMN_RANGE]
        assert report.permission_ranges[0].valid is False
        assert [s.status for s in report.segments] == [SegmentStatus.RESTRICTED]

    def it_reports_column_attributes_outside_a_table(self):
        document = _document(
            _p(
                _perm_start("8", 'w:edGrp="everyone" w:colFirst="0" w:colLast="1"')
                + _run("x")
                + _perm_end("8")
            ),
            protection=READ_ONLY_ENFORCED,
        )

        report = analyze_protection(document, "alice")

        assert [i.kind for i in report.issues] == [IssueKind.INVALID_COLUMN_RANGE]
        assert [s.status for s in report.segments] == [SegmentStatus.RESTRICTED]


class DescribeReadOnlyAnalysis:
    def it_does_not_modify_the_document(self):
        document = _document(
            _p(_perm_start("1", 'w:ed="alice"') + _run("x") + _perm_end("1")),
            protection=READ_ONLY_ENFORCED,
        )
        document_xml_before = document.element.xml
        settings_xml_before = document.settings.element.xml

        analyze_protection(document, "alice")

        assert document.element.xml == document_xml_before
        assert document.settings.element.xml == settings_xml_before

    def it_reports_editable_and_restricted_segment_views(self):
        document = _document(
            _p(_run("locked")),
            _p(_perm_start("1", 'w:ed="alice"') + _run("open") + _perm_end("1")),
            protection=READ_ONLY_ENFORCED,
        )

        report = analyze_protection(document, "alice")

        assert [s.text for s in report.editable_segments] == ["open"]
        assert [s.text for s in report.restricted_segments] == ["locked"]
