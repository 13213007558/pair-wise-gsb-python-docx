# pyright: reportPrivateUsage=false

"""Unit test suite for the docx.customxml module."""

from __future__ import annotations

import pytest

from docx.customxml import (
    CustomXmlStore,
    evaluate_xpath_subset,
    parse_prefix_mappings,
    parse_xpath_subset,
)
from docx.exceptions import UnsupportedXPathError
from docx.opc.packuri import PackURI
from docx.opc.part import Part
from docx.oxml import parse_xml

DATA_XML = (
    b"<?xml version='1.0' encoding='UTF-8' standalone='yes'?>"
    b"<root><name>Alice</name><items>"
    b"<item><n>one</n></item><item><n>two</n></item><item><n>three</n></item>"
    b"</items></root>"
)

NS_DATA_XML = (
    b"<?xml version='1.0' encoding='UTF-8' standalone='yes'?>"
    b'<inv:invoice xmlns:inv="http://example.com/inv">'
    b"<inv:customer>Acme</inv:customer></inv:invoice>"
)


def make_store(data: bytes = DATA_XML, item_id: str = "{TEST-ITEM-ID}") -> CustomXmlStore:
    part = Part(PackURI("/customXml/item9.xml"), "application/xml", data, None)
    return CustomXmlStore(part, item_id)


class DescribeParsePrefixMappings:
    def it_parses_word_style_prefix_mappings(self):
        mappings = parse_prefix_mappings(
            "xmlns:ns0='http://example.com/a' xmlns:ns1='http://example.com/b'"
        )
        assert mappings == {
            "ns0": "http://example.com/a",
            "ns1": "http://example.com/b",
        }

    def it_accepts_double_quoted_uris(self):
        assert parse_prefix_mappings('xmlns:a="http://example.com/a"') == {
            "a": "http://example.com/a"
        }

    def it_returns_empty_mapping_for_empty_input(self):
        assert parse_prefix_mappings("") == {}
        assert parse_prefix_mappings(None) == {}


class DescribeParseXpathSubset:
    @pytest.mark.parametrize(
        ("xpath", "expected"),
        [
            ("/root", (True, [("root", None)])),
            ("/root[1]/name[2]", (True, [("root", 1), ("name", 2)])),
            ("name", (False, [("name", None)])),
            ("ns0:item/ns0:n[1]", (False, [("ns0:item", None), ("ns0:n", 1)])),
        ],
    )
    def it_parses_supported_expressions(self, xpath, expected):
        is_absolute, steps = parse_xpath_subset(xpath)
        assert is_absolute == expected[0]
        assert [(s.name, s.index) for s in steps] == expected[1]

    @pytest.mark.parametrize(
        "xpath",
        ["", "  ", "/", "//name", "/root/@attr", "/root/*", "/root/name[last()]", "/root/a-b!"],
    )
    def it_rejects_unsupported_expressions(self, xpath):
        with pytest.raises(UnsupportedXPathError):
            parse_xpath_subset(xpath)


class DescribeEvaluateXpathSubset:
    def it_selects_elements_by_absolute_path(self):
        root = parse_xml(DATA_XML)
        nodes = evaluate_xpath_subset(root, "/root[1]/name[1]", {})
        assert [n.text for n in nodes] == ["Alice"]

    def it_applies_positional_predicates(self):
        root = parse_xml(DATA_XML)
        nodes = evaluate_xpath_subset(root, "/root/items/item[2]", {})
        assert len(nodes) == 1
        assert nodes[0].find("n").text == "two"

    def it_evaluates_relative_paths_against_a_context_node(self):
        root = parse_xml(DATA_XML)
        item = evaluate_xpath_subset(root, "/root/items/item[1]", {})[0]
        nodes = evaluate_xpath_subset(item, "n[1]", {})
        assert [n.text for n in nodes] == ["one"]

    def it_resolves_namespace_prefixes(self):
        root = parse_xml(NS_DATA_XML)
        nsmap = {"inv": "http://example.com/inv"}
        nodes = evaluate_xpath_subset(root, "/inv:invoice[1]/inv:customer[1]", nsmap)
        assert [n.text for n in nodes] == ["Acme"]

    def it_raises_for_undefined_prefixes(self):
        root = parse_xml(NS_DATA_XML)
        with pytest.raises(UnsupportedXPathError):
            evaluate_xpath_subset(root, "/inv:invoice", {})

    def it_ignores_final_predicate_when_expanding(self):
        root = parse_xml(DATA_XML)
        nodes = evaluate_xpath_subset(
            root, "/root[1]/items[1]/item[1]", {}, ignore_final_predicate=True
        )
        assert len(nodes) == 3


class DescribeCustomXmlStore:
    def it_finds_elements_by_xpath(self):
        store = make_store()
        nodes = store.find("/root/items/item")
        assert len(nodes) == 3

    def it_finds_text_of_a_single_element(self):
        store = make_store()
        assert store.find_text("/root[1]/name[1]") == "Alice"

    def it_raises_when_find_text_match_count_is_not_one(self):
        store = make_store()
        with pytest.raises(UnsupportedXPathError):
            store.find_text("/root/items/item")

    def it_can_replace_the_text_of_a_node(self):
        store = make_store()
        store.set_text("/root[1]/name[1]", "Bob")
        assert store.find_text("/root/name") == "Bob"

    def it_writes_changes_through_to_the_part(self):
        store = make_store()
        store.set_text("/root/name", "Carol")
        assert b"<name>Carol</name>" in store._part.blob

    def it_raises_when_set_text_match_count_is_not_one(self):
        store = make_store()
        with pytest.raises(UnsupportedXPathError):
            store.set_text("/root/items/item", "x")

    def it_knows_its_item_id_and_partname(self):
        store = make_store()
        assert store.item_id == "{TEST-ITEM-ID}"
        assert store.partname == "/customXml/item9.xml"
