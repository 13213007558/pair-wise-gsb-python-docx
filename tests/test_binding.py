# pyright: reportPrivateUsage=false

"""Test suite for content-control data binding (docx.binding module)."""

from __future__ import annotations

import pytest

from docx import Document
from docx.exceptions import BindingRefreshError
from docx.oxml import parse_xml
from docx.oxml.ns import qn

ITEM_ID = "{EF278816-EC6F-A645-907D-7F25AECB1D4A}"

DATA_XML = (
    b"<?xml version='1.0' encoding='UTF-8' standalone='yes'?>"
    b"<root><name>Alice</name><active>true</active><items>"
    b"<item><n>one</n><done>false</done></item>"
    b"<item><n>two</n><done>true</done></item>"
    b"<item><n>three</n><done>yes</done></item>"
    b"</items></root>"
)

NS_DATA_XML = (
    b"<?xml version='1.0' encoding='UTF-8' standalone='yes'?>"
    b'<inv:invoice xmlns:inv="http://example.com/inv">'
    b"<inv:customer>Acme</inv:customer></inv:invoice>"
)

NSDECLS = (
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" '
    'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"'
)

CHECKBOX_PR = (
    "<w14:checkbox>"
    '<w14:checked w14:val="0"/>'
    '<w14:checkedState w14:val="2612" w14:font="MS Gothic"/>'
    '<w14:uncheckedState w14:val="2610" w14:font="MS Gothic"/>'
    "</w14:checkbox>"
)


def text_sdt(id_, tag="name", xpath="/root[1]/name[1]", store=ITEM_ID, prefix=""):
    return (
        f"<w:p><w:sdt><w:sdtPr>"
        f'<w:id w:val="{id_}"/><w:tag w:val="{tag}"/>'
        f'<w:dataBinding w:storeItemID="{store}" w:xpath="{xpath}"'
        f' w:prefixMappings="{prefix}"/>'
        f"<w:text/></w:sdtPr>"
        f"<w:sdtContent><w:r><w:t>OLD-{tag}</w:t></w:r></w:sdtContent></w:sdt></w:p>"
    )


def checkbox_sdt(id_, tag="active", xpath="/root[1]/active[1]", store=ITEM_ID, prefix=""):
    return (
        f"<w:p><w:sdt><w:sdtPr>"
        f'<w:id w:val="{id_}"/><w:tag w:val="{tag}"/>'
        f'<w:dataBinding w:storeItemID="{store}" w:xpath="{xpath}"'
        f' w:prefixMappings="{prefix}"/>'
        f"{CHECKBOX_PR}</w:sdtPr>"
        f"<w:sdtContent><w:r><w:t>☐</w:t></w:r></w:sdtContent></w:sdt></w:p>"
    )


def section_sdt(
    id_=12, xpath="/root[1]/items[1]/item[1]", child_xpath="n[1]", done_xpath="done[1]"
):
    return (
        f"<w:sdt><w:sdtPr>"
        f'<w:id w:val="{id_}"/><w:tag w:val="items"/>'
        f'<w:dataBinding w:storeItemID="{ITEM_ID}" w:xpath="{xpath}" w:prefixMappings=""/>'
        f"<w15:repeatingSection/></w:sdtPr><w:sdtContent>"
        f"<w:sdt><w:sdtPr>"
        f'<w:id w:val="{id_ + 1}"/><w15:repeatingSectionItem/>'
        f"</w:sdtPr><w:sdtContent><w:p>"
        f"<w:sdt><w:sdtPr>"
        f'<w:id w:val="{id_ + 2}"/><w:tag w:val="n"/>'
        f'<w:dataBinding w:storeItemID="{ITEM_ID}" w:xpath="{child_xpath}"'
        f' w:prefixMappings=""/>'
        f"<w:text/></w:sdtPr>"
        f"<w:sdtContent><w:r><w:t>X</w:t></w:r></w:sdtContent></w:sdt>"
        f"<w:sdt><w:sdtPr>"
        f'<w:id w:val="{id_ + 3}"/><w:tag w:val="done"/>'
        f'<w:dataBinding w:storeItemID="{ITEM_ID}" w:xpath="{done_xpath}"'
        f' w:prefixMappings=""/>'
        f"{CHECKBOX_PR}</w:sdtPr>"
        f"<w:sdtContent><w:r><w:t>☐</w:t></w:r></w:sdtContent></w:sdt>"
        f"</w:p></w:sdtContent></w:sdt>"
        f"</w:sdtContent></w:sdt>"
    )


UNBOUND_SDT = (
    "<w:p><w:sdt><w:sdtPr>"
    '<w:id w:val="90"/><w:tag w:val="keep"/><w:text/>'
    "</w:sdtPr><w:sdtContent><w:r><w:t>KEEP</w:t></w:r></w:sdtContent></w:sdt></w:p>"
)

RICHTEXT_SDT = (
    "<w:p><w:sdt><w:sdtPr>"
    '<w:id w:val="91"/><w:tag w:val="rich"/>'
    f'<w:dataBinding w:storeItemID="{ITEM_ID}" w:xpath="/root[1]/name[1]"'
    ' w:prefixMappings=""/>'
    "<w:richText/></w:sdtPr>"
    "<w:sdtContent><w:r><w:t>RICH</w:t></w:r></w:sdtContent></w:sdt></w:p>"
)


def document_with(*fragments, data=DATA_XML):
    """A Document whose body contains 'fragments' and whose custom XML is 'data'."""
    doc = Document()
    for part in doc.part.package.iter_parts():
        if str(part.partname) == "/customXml/item1.xml":
            part._blob = data
    body = doc.element.body
    sectPr = body.find(qn("w:sectPr"))
    for child in list(body):
        if child is not sectPr:
            body.remove(child)
    fragment = parse_xml(f"<w:body {NSDECLS}>" + "".join(fragments) + "</w:body>")
    for child in list(fragment):
        body.insert(list(body).index(sectPr), child)
    return doc


def body_texts(doc):
    return [t.text for t in doc.element.body.iter(qn("w:t"))]


def sdt_ids(doc):
    return [int(e.get(qn("w:val"))) for e in doc.element.body.iter(qn("w:id"))]


def sdt_by_tag(doc, tag):
    for sdt in doc.element.body.iter(qn("w:sdt")):
        sdtPr = sdt.find(qn("w:sdtPr"))
        tag_el = sdtPr.find(qn("w:tag"))
        if tag_el is not None and tag_el.get(qn("w:val")) == tag:
            return sdt
    raise KeyError(tag)


class DescribeDocumentBindings:
    def it_provides_access_to_custom_xml_stores(self):
        doc = document_with(text_sdt(10))
        stores = doc.bindings.stores
        assert len(stores) == 1
        assert stores[0].item_id == ITEM_ID

    def it_looks_up_a_store_by_item_id_case_insensitively(self):
        doc = document_with(text_sdt(10))
        assert doc.bindings.store(ITEM_ID.lower()).item_id == ITEM_ID

    def it_raises_when_no_store_matches_item_id(self):
        doc = document_with(text_sdt(10))
        with pytest.raises(KeyError):
            doc.bindings.store("{NO-SUCH-ITEM}")

    def it_allows_updating_custom_xml_data(self):
        doc = document_with(text_sdt(10))
        store = doc.bindings.store(ITEM_ID)
        store.set_text("/root[1]/name[1]", "Zoe")
        assert store.find_text("/root/name") == "Zoe"


class DescribePlanRefresh:
    def it_plans_each_bound_control(self):
        doc = document_with(text_sdt(10), checkbox_sdt(11), section_sdt(12))
        plan = doc.bindings.plan_refresh()
        assert plan.ok
        assert plan.problems == ()
        assert [(e.control_id, e.kind) for e in plan.entries] == [
            (10, "plain-text"),
            (11, "checkbox"),
            (12, "repeating-section"),
        ]
        text_entry, checkbox_entry, section_entry = plan.entries
        assert text_entry.action == "set text to 'Alice'"
        assert text_entry.value == "Alice"
        assert checkbox_entry.action == "set checked"
        assert checkbox_entry.value == "true"
        assert section_entry.action == "expand to 3 item(s)"
        assert section_entry.item_count == 3

    def it_does_not_modify_the_document(self):
        doc = document_with(text_sdt(10), checkbox_sdt(11), section_sdt(12))
        before = body_texts(doc)
        doc.bindings.plan_refresh()
        assert body_texts(doc) == before

    @pytest.mark.parametrize(
        "xpath",
        ["/root//name", "/root/name[last()]", "/root/@name", "/root/*"],
    )
    def it_reports_unsupported_xpath(self, xpath):
        doc = document_with(text_sdt(10, xpath=xpath))
        plan = doc.bindings.plan_refresh()
        assert not plan.ok
        assert len(plan.problems) == 1
        problem = plan.problems[0]
        assert problem.control_id == 10
        assert problem.tag == "name"
        assert problem.kind == "plain-text"

    def it_reports_when_xpath_matches_no_node(self):
        doc = document_with(text_sdt(10, xpath="/root[1]/missing[1]"))
        plan = doc.bindings.plan_refresh()
        assert len(plan.problems) == 1
        assert "matched 0 nodes" in plan.problems[0].reason

    def it_reports_when_xpath_matches_too_many_nodes(self):
        doc = document_with(text_sdt(10, xpath="/root[1]/items[1]/item"))
        plan = doc.bindings.plan_refresh()
        assert len(plan.problems) == 1
        assert "matched 3 nodes" in plan.problems[0].reason

    def it_reports_checkbox_value_type_mismatch(self):
        doc = document_with(checkbox_sdt(11, xpath="/root[1]/name[1]"))
        plan = doc.bindings.plan_refresh()
        assert len(plan.problems) == 1
        problem = plan.problems[0]
        assert problem.control_id == 11
        assert "checkbox state" in problem.reason

    def it_reports_unknown_custom_xml_store(self):
        doc = document_with(text_sdt(10, store="{UNKNOWN-STORE}"))
        plan = doc.bindings.plan_refresh()
        assert len(plan.problems) == 1
        assert "not found" in plan.problems[0].reason

    def it_reports_unsupported_control_type(self):
        doc = document_with(RICHTEXT_SDT)
        plan = doc.bindings.plan_refresh()
        assert len(plan.problems) == 1
        problem = plan.problems[0]
        assert problem.control_id == 91
        assert "not supported" in problem.reason

    def it_reports_repeating_section_with_no_matching_nodes(self):
        doc = document_with(section_sdt(12, xpath="/root[1]/missing[1]"))
        plan = doc.bindings.plan_refresh()
        assert len(plan.problems) == 1
        assert "at least 1" in plan.problems[0].reason

    def it_reports_section_child_problems_with_item_index(self):
        data = (
            b"<root><items>"
            b"<item><n>one</n><done>true</done></item>"
            b"<item><n>two</n></item>"
            b"</items></root>"
        )
        doc = document_with(section_sdt(12), data=data)
        plan = doc.bindings.plan_refresh()
        assert len(plan.problems) == 1
        problem = plan.problems[0]
        assert problem.tag == "done"
        assert problem.item_index == 1
        assert "matched 0 nodes" in problem.reason


class DescribeApply:
    def it_updates_plain_text_controls(self):
        doc = document_with(text_sdt(10))
        doc.bindings.refresh()
        assert body_texts(doc) == ["Alice"]

    def it_updates_checkbox_state_and_glyph(self):
        doc = document_with(checkbox_sdt(11))
        doc.bindings.refresh()
        sdt = sdt_by_tag(doc, "active")
        checked = sdt.find(f"{qn('w:sdtPr')}/{qn('w14:checkbox')}/{qn('w14:checked')}")
        assert checked.get(qn("w14:val")) == "1"
        assert body_texts(doc) == ["☒"]

    def it_expands_a_repeating_section_with_unique_ids(self):
        doc = document_with(section_sdt(12))
        doc.bindings.refresh()
        assert body_texts(doc) == ["one", "☐", "two", "☒", "three", "☒"]
        items = doc.element.body.findall(f".//{qn('w15:repeatingSectionItem')}")
        assert len(items) == 3
        ids = sdt_ids(doc)
        assert len(ids) == len(set(ids))

    def it_does_not_duplicate_items_when_refreshed_repeatedly(self):
        doc = document_with(section_sdt(12))
        doc.bindings.refresh()
        doc.bindings.refresh()
        assert body_texts(doc) == ["one", "☐", "two", "☒", "three", "☒"]
        items = doc.element.body.findall(f".//{qn('w15:repeatingSectionItem')}")
        assert len(items) == 3
        ids = sdt_ids(doc)
        assert len(ids) == len(set(ids))

    def it_updates_all_controls_bound_to_the_same_data(self):
        doc = document_with(text_sdt(10), text_sdt(20, tag="name2"))
        doc.bindings.refresh()
        assert body_texts(doc) == ["Alice", "Alice"]

    def it_leaves_unbound_controls_untouched(self):
        doc = document_with(UNBOUND_SDT, text_sdt(10))
        doc.bindings.refresh()
        assert body_texts(doc) == ["KEEP", "Alice"]

    def it_raises_and_leaves_document_unmodified_when_problems(self):
        doc = document_with(checkbox_sdt(11, xpath="/root[1]/name[1]"), text_sdt(10))
        plan = doc.bindings.plan_refresh()
        with pytest.raises(BindingRefreshError) as excinfo:
            plan.apply()
        assert "id=11" in str(excinfo.value)
        assert body_texts(doc) == ["☐", "OLD-name"]

    def it_raises_on_refresh_when_problems(self):
        doc = document_with(text_sdt(10, xpath="/root[1]/missing[1]"))
        with pytest.raises(BindingRefreshError):
            doc.bindings.refresh()
        assert body_texts(doc) == ["OLD-name"]

    def it_refreshes_display_after_custom_xml_data_is_updated(self):
        doc = document_with(text_sdt(10))
        doc.bindings.store(ITEM_ID).set_text("/root[1]/name[1]", "Zoe")
        doc.bindings.refresh()
        assert body_texts(doc) == ["Zoe"]

    def it_supports_prefix_mappings(self):
        doc = document_with(
            text_sdt(
                10,
                xpath="/inv:invoice[1]/inv:customer[1]",
                prefix="xmlns:inv='http://example.com/inv'",
            ),
            data=NS_DATA_XML,
        )
        doc.bindings.refresh()
        assert body_texts(doc) == ["Acme"]

    def it_persists_data_and_display_changes_across_save_and_reload(self):
        from io import BytesIO

        doc = document_with(text_sdt(10), section_sdt(12))
        doc.bindings.store(ITEM_ID).set_text("/root[1]/name[1]", "Zoe")
        doc.bindings.refresh()
        stream = BytesIO()
        doc.save(stream)
        stream.seek(0)

        reloaded = Document(stream)
        assert body_texts(reloaded)[:1] == ["Zoe"]
        assert reloaded.bindings.store(ITEM_ID).find_text("/root/name") == "Zoe"
        items = reloaded.element.body.findall(f".//{qn('w15:repeatingSectionItem')}")
        assert len(items) == 3
