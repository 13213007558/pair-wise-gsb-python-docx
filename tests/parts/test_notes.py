# pyright: reportPrivateUsage=false

"""Unit test suite for the docx.parts.notes module."""

from __future__ import annotations

from typing import cast

import pytest

from docx.notes import Endnotes, Footnotes
from docx.opc.constants import CONTENT_TYPE as CT
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.packuri import PackURI
from docx.opc.part import PartFactory
from docx.oxml.notes import CT_Endnotes, CT_Footnotes
from docx.package import Package
from docx.parts.notes import EndnotesPart, FootnotesPart

from ..unitutil.cxml import element
from ..unitutil.mock import FixtureRequest, Mock, class_mock, instance_mock


class DescribeFootnotesPart:
    """Unit test suite for `docx.parts.notes.FootnotesPart` objects."""

    def it_is_used_by_the_part_loader_to_construct_a_footnotes_part(
        self, package_: Mock, FootnotesPart_load_: Mock, footnotes_part_: Mock
    ):
        partname = PackURI("/word/footnotes.xml")
        content_type = CT.WML_FOOTNOTES
        reltype = RT.FOOTNOTES
        blob = b"<w:footnotes/>"
        FootnotesPart_load_.return_value = footnotes_part_

        part = PartFactory(partname, content_type, reltype, blob, package_)

        FootnotesPart_load_.assert_called_once_with(partname, content_type, blob, package_)
        assert part is footnotes_part_

    def it_provides_access_to_its_footnotes_collection(
        self, Footnotes_: Mock, footnotes_: Mock, package_: Mock
    ):
        Footnotes_.return_value = footnotes_
        footnotes_elm = cast(CT_Footnotes, element("w:footnotes"))
        footnotes_part = FootnotesPart(
            PackURI("/word/footnotes.xml"), CT.WML_FOOTNOTES, footnotes_elm, package_
        )

        footnotes = footnotes_part.footnotes

        Footnotes_.assert_called_once_with(footnotes_part.element, footnotes_part)
        assert footnotes is footnotes_

    def it_constructs_a_default_footnotes_part_to_help(self):
        package = Package()

        footnotes_part = FootnotesPart.default(package)

        assert isinstance(footnotes_part, FootnotesPart)
        assert footnotes_part.partname == "/word/footnotes.xml"
        assert footnotes_part.content_type == CT.WML_FOOTNOTES
        # -- a default footnotes part contains the separator and continuation-separator --
        # -- footnotes, but no user-visible footnotes --
        footnotes = footnotes_part.footnotes
        assert len(footnotes) == 0
        w = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
        special_types = [f.get(w + "type") for f in footnotes_part.element.footnote_lst]
        assert special_types == ["separator", "continuationSeparator"]

    # -- fixtures --------------------------------------------------------------------------------

    @pytest.fixture
    def Footnotes_(self, request: FixtureRequest):
        return class_mock(request, "docx.parts.notes.Footnotes")

    @pytest.fixture
    def FootnotesPart_load_(self, request: FixtureRequest):
        return class_mock(request, "docx.parts.notes.FootnotesPart.load")

    @pytest.fixture
    def footnotes_(self, request: FixtureRequest) -> Mock:
        return instance_mock(request, Footnotes)

    @pytest.fixture
    def footnotes_part_(self, request: FixtureRequest) -> Mock:
        return instance_mock(request, FootnotesPart)

    @pytest.fixture
    def package_(self, request: FixtureRequest) -> Mock:
        return instance_mock(request, Package)


class DescribeEndnotesPart:
    """Unit test suite for `docx.parts.notes.EndnotesPart` objects."""

    def it_is_used_by_the_part_loader_to_construct_an_endnotes_part(
        self, package_: Mock, EndnotesPart_load_: Mock, endnotes_part_: Mock
    ):
        partname = PackURI("/word/endnotes.xml")
        content_type = CT.WML_ENDNOTES
        reltype = RT.ENDNOTES
        blob = b"<w:endnotes/>"
        EndnotesPart_load_.return_value = endnotes_part_

        part = PartFactory(partname, content_type, reltype, blob, package_)

        EndnotesPart_load_.assert_called_once_with(partname, content_type, blob, package_)
        assert part is endnotes_part_

    def it_provides_access_to_its_endnotes_collection(
        self, Endnotes_: Mock, endnotes_: Mock, package_: Mock
    ):
        Endnotes_.return_value = endnotes_
        endnotes_elm = cast(CT_Endnotes, element("w:endnotes"))
        endnotes_part = EndnotesPart(
            PackURI("/word/endnotes.xml"), CT.WML_ENDNOTES, endnotes_elm, package_
        )

        endnotes = endnotes_part.endnotes

        Endnotes_.assert_called_once_with(endnotes_part.element, endnotes_part)
        assert endnotes is endnotes_

    def it_constructs_a_default_endnotes_part_to_help(self):
        package = Package()

        endnotes_part = EndnotesPart.default(package)

        assert isinstance(endnotes_part, EndnotesPart)
        assert endnotes_part.partname == "/word/endnotes.xml"
        assert endnotes_part.content_type == CT.WML_ENDNOTES
        endnotes = endnotes_part.endnotes
        assert len(endnotes) == 0
        w = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
        special_types = [f.get(w + "type") for f in endnotes_part.element.endnote_lst]
        assert special_types == ["separator", "continuationSeparator"]

    # -- fixtures --------------------------------------------------------------------------------

    @pytest.fixture
    def Endnotes_(self, request: FixtureRequest):
        return class_mock(request, "docx.parts.notes.Endnotes")

    @pytest.fixture
    def EndnotesPart_load_(self, request: FixtureRequest):
        return class_mock(request, "docx.parts.notes.EndnotesPart.load")

    @pytest.fixture
    def endnotes_(self, request: FixtureRequest) -> Mock:
        return instance_mock(request, Endnotes)

    @pytest.fixture
    def endnotes_part_(self, request: FixtureRequest) -> Mock:
        return instance_mock(request, EndnotesPart)

    @pytest.fixture
    def package_(self, request: FixtureRequest) -> Mock:
        return instance_mock(request, Package)
