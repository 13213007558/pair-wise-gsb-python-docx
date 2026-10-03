"""Unit test suite for the docx.protection module."""

from __future__ import annotations

from docx import Document
from docx.protection import (
    ISSUE_INVALID_COLUMN_RANGE,
    ISSUE_MISSING_END,
    ISSUE_MISSING_PRINCIPAL,
    ISSUE_MISSING_START,
    ISSUE_UNKNOWN_GROUP,
    REASON_EXCEPTION,
    REASON_NO_EXCEPTION,
    REASON_PROTECTION_NOT_ENFORCED,
    analyze_editability,
)

from .unitutil.cxml import element

ENFORCED = "w:documentProtection{w:edit=readOnly,w:enforcement=on}"


def para(text):
    return 'w:p/w:r/w:t"%s"' % text


def make_document(body_cxml=None, settings_cxml=None):
    document = Document()
    if body_cxml:
        body = document.element.body
        for child in element("w:body/(%s)" % body_cxml):
            body.insert_element_before(child, "w:sectPr")
    if settings_cxml:
        settings = document.settings.element
        for child in element("w:settings/(%s)" % settings_cxml):
            settings.append(child)
    return document


def enforced_doc(body_cxml):
    return make_document(body_cxml, ENFORCED)


class DescribeProtectionStatus:
    def it_treats_absent_protection_as_not_enforced(self):
        document = make_document(para("hello"))

        report = analyze_editability(document, user="alice")

        assert report.status == "not-enforced"
        assert report.enforced is False
        assert [(s.editable, s.reason) for s in report.segments] == [
            (True, REASON_PROTECTION_NOT_ENFORCED)
        ]

    def it_treats_unenforced_protection_as_not_enforced(self):
        document = make_document(para("hello"), "w:documentProtection{w:edit=readOnly}")

        report = analyze_editability(document, user="alice")

        assert report.status == "not-enforced"
        assert report.segments[0].reason == REASON_PROTECTION_NOT_ENFORCED

    def it_treats_comments_mode_as_not_applicable_to_range_permissions(self):
        document = make_document(
            para("hello"), "w:documentProtection{w:edit=comments,w:enforcement=on}"
        )

        report = analyze_editability(document, user="alice")

        assert report.status == "not-enforced"
        assert report.segments[0].reason == REASON_PROTECTION_NOT_ENFORCED

    def it_enforces_readOnly_and_forms_modes(self):
        for edit in ("readOnly", "forms"):
            document = make_document(
                para("locked"),
                "w:documentProtection{w:edit=%s,w:enforcement=on}" % edit,
            )

            report = analyze_editability(document, user="alice")

            assert report.status == "enforced"
            assert [(s.editable, s.reason) for s in report.segments] == [
                (False, REASON_NO_EXCEPTION)
            ]


class DescribeRangePairing:
    def it_grants_a_range_to_the_named_user(self):
        document = enforced_doc(
            ",".join(
                [
                    "w:permStart{w:id=1,w:ed=alice}",
                    para("for alice"),
                    "w:permEnd{w:id=1}",
                ]
            )
        )

        report = analyze_editability(document, user="alice")

        assert [(s.editable, s.reason, s.granted_by) for s in report.segments] == [
            (True, REASON_EXCEPTION, "ed")
        ]
        assert report.issues == []

    def it_denies_a_range_to_another_user(self):
        document = enforced_doc(
            ",".join(
                [
                    "w:permStart{w:id=1,w:ed=alice}",
                    para("for alice"),
                    "w:permEnd{w:id=1}",
                ]
            )
        )

        report = analyze_editability(document, user="bob")

        assert [(s.editable, s.reason) for s in report.segments] == [(False, REASON_NO_EXCEPTION)]

    def it_grants_a_range_to_a_known_group(self):
        document = enforced_doc(
            ",".join(
                [
                    "w:permStart{w:id=1,w:edGrp=editors}",
                    para("for editors"),
                    "w:permEnd{w:id=1}",
                ]
            )
        )

        report = analyze_editability(document, user="carol", groups=["editors"])

        assert [(s.editable, s.reason, s.granted_by) for s in report.segments] == [
            (True, REASON_EXCEPTION, "edGrp")
        ]
        assert report.issues == []

    def it_reports_an_unknown_group_without_granting(self):
        document = enforced_doc(
            ",".join(
                [
                    "w:permStart{w:id=1,w:edGrp=editors}",
                    para("for editors"),
                    "w:permEnd{w:id=1}",
                ]
            )
        )

        report = analyze_editability(document, user="carol")

        assert [(s.editable, s.reason) for s in report.segments] == [(False, REASON_NO_EXCEPTION)]
        assert [i.kind for i in report.issues] == [ISSUE_UNKNOWN_GROUP]
        assert report.issues[0].marker.ed_grp == "editors"

    def it_grants_the_reserved_everyone_principal_to_any_user(self):
        document = enforced_doc(
            ",".join(
                [
                    "w:permStart{w:id=1,w:ed=everyone}",
                    para("for all"),
                    "w:permEnd{w:id=1}",
                ]
            )
        )

        report = analyze_editability(document, user="nobody")

        assert [(s.editable, s.granted_by) for s in report.segments] == [(True, "everyone")]

    def it_reports_a_marker_without_principal_without_granting(self):
        document = enforced_doc(
            ",".join(["w:permStart{w:id=1}", para("orphan"), "w:permEnd{w:id=1}"])
        )

        report = analyze_editability(document, user="alice")

        assert [(s.editable, s.reason) for s in report.segments] == [(False, REASON_NO_EXCEPTION)]
        assert [i.kind for i in report.issues] == [ISSUE_MISSING_PRINCIPAL]

    def it_reports_a_permStart_without_matching_permEnd(self):
        document = enforced_doc(",".join(["w:permStart{w:id=1,w:ed=alice}", para("open")]))

        report = analyze_editability(document, user="alice")

        assert [i.kind for i in report.issues] == [ISSUE_MISSING_END]
        assert report.issues[0].marker.perm_id == 1
        # the unclosed range still applies to the following content
        assert [(s.editable, s.reason) for s in report.segments] == [(True, REASON_EXCEPTION)]
        assert report.segments[0].end_marker is None

    def it_reports_a_permEnd_without_matching_permStart(self):
        document = enforced_doc(",".join([para("closed"), "w:permEnd{w:id=9}"]))

        report = analyze_editability(document, user="alice")

        assert [i.kind for i in report.issues] == [ISSUE_MISSING_START]
        assert report.issues[0].marker.perm_id == 9
        assert [(s.editable, s.reason) for s in report.segments] == [(False, REASON_NO_EXCEPTION)]

    def it_pairs_interleaved_ranges_by_id(self):
        document = enforced_doc(
            ",".join(
                [
                    para("p0"),
                    "w:permStart{w:id=1,w:ed=alice}",
                    para("p1"),
                    "w:permStart{w:id=2,w:ed=bob}",
                    para("p2"),
                    "w:permEnd{w:id=1}",
                    para("p3"),
                    "w:permEnd{w:id=2}",
                    para("p4"),
                ]
            )
        )

        alice = analyze_editability(document, user="alice")
        bob = analyze_editability(document, user="bob")

        assert [(s.text, s.editable) for s in alice.segments] == [
            ("p0", False),
            ("p1", True),
            ("p2", True),
            ("p3", False),
            ("p4", False),
        ]
        assert [(s.text, s.editable) for s in bob.segments] == [
            ("p0", False),
            ("p1", False),
            ("p2", True),
            ("p3", True),
            ("p4", False),
        ]
        assert alice.issues == []
        assert bob.issues == []

    def it_distinguishes_adjacent_paragraphs_with_different_permissions(self):
        document = enforced_doc(
            ",".join(
                [
                    "w:permStart{w:id=1,w:ed=alice}",
                    para("for-alice"),
                    "w:permEnd{w:id=1}",
                    "w:permStart{w:id=2,w:edGrp=editors}",
                    para("for-editors"),
                    "w:permEnd{w:id=2}",
                    para("locked"),
                ]
            )
        )

        alice = analyze_editability(document, user="alice")
        carol = analyze_editability(document, user="carol", groups=["editors"])

        assert [(s.text, s.editable) for s in alice.segments] == [
            ("for-alice", True),
            ("for-editors", False),
            ("locked", False),
        ]
        assert [(s.text, s.editable) for s in carol.segments] == [
            ("for-alice", False),
            ("for-editors", True),
            ("locked", False),
        ]

    def it_splits_a_paragraph_at_inline_markers(self):
        document = enforced_doc(
            "w:p/("
            + ",".join(
                [
                    'w:r/w:t"locked-"',
                    "w:permStart{w:id=1,w:ed=alice}",
                    'w:r/w:t"editable"',
                    "w:permEnd{w:id=1}",
                    'w:r/w:t"-locked"',
                ]
            )
            + ")"
        )

        report = analyze_editability(document, user="alice")

        assert [(s.text, s.editable, s.location) for s in report.segments] == [
            ("locked-", False, "body/p[0]"),
            ("editable", True, "body/p[0]"),
            ("-locked", False, "body/p[0]"),
        ]


class DescribeTableColumnRanges:
    def table_doc(self, perm_start):
        return enforced_doc(
            "w:tbl/w:tr/("
            + ",".join(
                [
                    perm_start,
                    'w:tc/w:p/w:r/w:t"c0"',
                    'w:tc/w:p/w:r/w:t"c1"',
                    'w:tc/w:p/w:r/w:t"c2"',
                    "w:permEnd{w:id=1}",
                ]
            )
            + ")"
        )

    def it_opens_only_the_middle_column(self):
        document = self.table_doc("w:permStart{w:id=1,w:ed=alice,w:colFirst=1,w:colLast=1}")

        report = analyze_editability(document, user="alice")

        assert [(s.text, s.editable, s.location) for s in report.segments] == [
            ("c0", False, "body/tbl[0]/tr[0]/tc[0]/p[0]"),
            ("c1", True, "body/tbl[0]/tr[0]/tc[1]/p[0]"),
            ("c2", False, "body/tbl[0]/tr[0]/tc[2]/p[0]"),
        ]
        assert report.issues == []
        assert report.segments[1].perm_id == 1

    def it_reports_an_inverted_column_range_without_granting(self):
        document = self.table_doc("w:permStart{w:id=1,w:ed=alice,w:colFirst=2,w:colLast=1}")

        report = analyze_editability(document, user="alice")

        assert [i.kind for i in report.issues] == [ISSUE_INVALID_COLUMN_RANGE]
        assert [s.editable for s in report.segments] == [False, False, False]

    def it_reports_a_column_range_exceeding_the_row(self):
        document = self.table_doc("w:permStart{w:id=1,w:ed=alice,w:colFirst=1,w:colLast=3}")

        report = analyze_editability(document, user="alice")

        assert [i.kind for i in report.issues] == [ISSUE_INVALID_COLUMN_RANGE]
        assert [s.editable for s in report.segments] == [False, False, False]

    def it_reports_a_column_range_outside_a_table_row(self):
        document = enforced_doc(
            ",".join(
                [
                    "w:permStart{w:id=1,w:ed=alice,w:colFirst=0,w:colLast=0}",
                    para("body"),
                    "w:permEnd{w:id=1}",
                ]
            )
        )

        report = analyze_editability(document, user="alice")

        assert [i.kind for i in report.issues] == [ISSUE_INVALID_COLUMN_RANGE]
        assert [(s.editable, s.reason) for s in report.segments] == [(False, REASON_NO_EXCEPTION)]


class DescribeTraceability:
    def it_links_an_editable_segment_to_its_markers(self):
        document = enforced_doc(
            ",".join(
                [
                    para("before"),
                    "w:permStart{w:id=7,w:ed=alice}",
                    para("inside"),
                    "w:permEnd{w:id=7}",
                    para("after"),
                ]
            )
        )

        report = analyze_editability(document, user="alice")

        segment = report.segments[1]
        assert segment.perm_id == 7
        assert segment.start_marker is not None
        assert segment.start_marker.kind == "permStart"
        assert segment.start_marker.location == "body"
        assert segment.end_marker is not None
        assert segment.end_marker.kind == "permEnd"
        assert segment.end_marker.location == "body"
        restricted = report.restricted_segments
        assert [s.location for s in restricted] == ["body/p[0]", "body/p[2]"]
        assert report.editable_segments == [segment]


class DescribeReadOnlyAnalysis:
    def it_does_not_modify_the_document(self):
        document = enforced_doc(
            ",".join(
                [
                    "w:permStart{w:id=1,w:ed=alice}",
                    para("for alice"),
                    "w:permEnd{w:id=1}",
                ]
            )
        )
        xml_before = document.element.xml
        settings_before = document.settings.element.xml

        analyze_editability(document, user="alice")

        assert document.element.xml == xml_before
        assert document.settings.element.xml == settings_before


class DescribeSettingsDocumentProtection:
    def it_reports_recorded_protection_settings(self):
        document = make_document(None, ENFORCED)

        info = document.settings.document_protection

        assert info.present is True
        assert info.edit == "readOnly"
        assert info.enforcement is True
        assert info.enforced is True

    def it_reports_absent_protection_settings(self):
        document = make_document()

        info = document.settings.document_protection

        assert info.present is False
        assert info.enforced is False
