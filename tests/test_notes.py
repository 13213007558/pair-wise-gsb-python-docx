# pyright: reportPrivateUsage=false

"""Unit test suite for the `docx.notes` module."""

from __future__ import annotations

import io
import zipfile
from typing import cast

import pytest

from docx import Document
from docx.notes import Endnote, Endnotes, Footnote, FootnoteReference, Footnotes
from docx.opc.constants import CONTENT_TYPE as CT
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.packuri import PackURI
from docx.oxml.notes import CT_Endnotes, CT_Footnotes
from docx.oxml.parser import OxmlElement
from docx.oxml.ns import qn
from docx.package import Package
from docx.parts.notes import EndnotesPart, FootnotesPart
from docx.shared import Inches

from .unitutil.cxml import element
from .unitutil.file import absjoin, test_file_dir
from .unitutil.mock import FixtureRequest, Mock, instance_mock


class DescribeFootnotes:
    """Unit-test suite for `docx.notes.Footnotes` objects."""

    @pytest.mark.parametrize(
        ("cxml", "count"),
        [
            ("w:footnotes", 0),
            ("w:footnotes/w:footnote{w:type=separator,w:id=-1}", 0),
            ("w:footnotes/(w:footnote{w:type=separator,w:id=-1},w:footnote{w:id=1})", 1),
            ("w:footnotes/(w:footnote{w:id=1},w:footnote{w:id=2},w:footnote{w:id=3})", 3),
            ("w:footnotes/(w:footnote{w:id=1,w:type=normal})", 1),
            ("w:footnotes/w:footnote{w:id=2,w:type=continuationNotice}", 0),
        ],
    )
    def it_counts_only_normal_footnotes(self, cxml: str, count: int, package_: Mock):
        footnotes = self._footnotes_for(cxml, package_)
        assert len(footnotes) == count

    def it_is_iterable_over_the_normal_footnotes_it_contains(self, package_: Mock):
        footnotes = self._footnotes_for(
            "w:footnotes/(w:footnote{w:type=separator,w:id=-1},w:footnote{w:id=1},"
            "w:footnote{w:id=2})",
            package_,
        )

        footnote_iter = iter(footnotes)

        footnote1 = next(footnote_iter)
        assert type(footnote1) is Footnote, "expected a `Footnote` object"
        assert footnote1.footnote_id == 1
        footnote2 = next(footnote_iter)
        assert type(footnote2) is Footnote, "expected a `Footnote` object"
        with pytest.raises(StopIteration):
            next(footnote_iter)

    def it_can_get_a_footnote_by_id(self, package_: Mock):
        footnotes = self._footnotes_for(
            "w:footnotes/(w:footnote{w:id=1},w:footnote{w:id=2},w:footnote{w:id=3})",
            package_,
        )

        footnote = footnotes.get(2)

        assert type(footnote) is Footnote, "expected a `Footnote` object"
        assert footnote._note_elm is footnotes._notes_elm.footnote_lst[1]

    def but_it_returns_None_when_no_footnote_with_that_id_exists(self, package_: Mock):
        footnotes = self._footnotes_for("w:footnotes/w:footnote{w:id=1}", package_)
        assert footnotes.get(42) is None

    def and_it_returns_None_when_the_id_belongs_to_a_special_footnote(
        self, package_: Mock
    ):
        footnotes = self._footnotes_for(
            "w:footnotes/w:footnote{w:type=separator,w:id=-1}", package_
        )
        assert footnotes.get(-1) is None

    def it_can_add_a_new_footnote(self, package_: Mock):
        footnotes = self._footnotes_for(
            "w:footnotes/(w:footnote{w:type=separator,w:id=-1},"
            "w:footnote{w:type=continuationSeparator,w:id=0})",
            package_,
        )

        footnote = footnotes.add_footnote()

        assert isinstance(footnote, Footnote)
        # -- it is "linked" to the footnotes part so it can add images, hyperlinks, etc. --
        assert footnote.part is footnotes._notes_part
        # -- footnote numbering starts at 1, skipping the reserved ids -1 and 0 --
        assert footnote.footnote_id == 1
        # -- by default, a new footnote contains a single empty paragraph --
        assert [p.text for p in footnote.paragraphs] == [""]
        # -- that paragraph has the "FootnoteText" style, same as Word applies --
        footnote_elm = footnote._note_elm
        assert len(footnote_elm.p_lst) == 1
        p = footnote_elm.p_lst[0]
        assert p.style == "FootnoteText"
        # -- and that paragraph contains a single run with the reference mark --
        assert len(p.r_lst) == 1
        r = p.r_lst[0]
        assert r.style == "FootnoteReference"
        assert r[-1].tag == qn("w:footnoteRef")

    def and_it_can_add_text_to_the_footnote_when_adding_it(self, package_: Mock):
        footnotes = self._footnotes_for("w:footnotes", package_)

        footnote = footnotes.add_footnote(text="para 1\n\npara 2")

        assert len(footnote.paragraphs) == 3
        assert [p.text for p in footnote.paragraphs] == ["para 1", "", "para 2"]
        assert all(p._p.style == "FootnoteText" for p in footnote.paragraphs)

    # -- fixtures --------------------------------------------------------------------------------

    @pytest.fixture
    def package_(self, request: FixtureRequest):
        return instance_mock(request, Package)

    def _footnotes_for(self, cxml: str, package_: Mock) -> Footnotes:
        footnotes_elm = cast(CT_Footnotes, element(cxml))
        footnotes_part = FootnotesPart(
            PackURI("/word/footnotes.xml"), CT.WML_FOOTNOTES, footnotes_elm, package_
        )
        return Footnotes(footnotes_elm, footnotes_part)


class DescribeEndnotes:
    """Unit-test suite for `docx.notes.Endnotes` objects."""

    def it_counts_and_iterates_only_normal_endnotes(self, package_: Mock):
        endnotes = self._endnotes_for(
            "w:endnotes/(w:endnote{w:type=separator,w:id=-1},"
            "w:endnote{w:type=continuationSeparator,w:id=0},w:endnote{w:id=1})",
            package_,
        )

        assert len(endnotes) == 1
        (endnote,) = list(endnotes)
        assert type(endnote) is Endnote, "expected an `Endnote` object"
        assert endnote.endnote_id == 1

    def it_can_get_an_endnote_by_id_but_not_a_special_one(self, package_: Mock):
        endnotes = self._endnotes_for(
            "w:endnotes/(w:endnote{w:type=separator,w:id=-1},w:endnote{w:id=1})",
            package_,
        )

        assert type(endnotes.get(1)) is Endnote
        assert endnotes.get(-1) is None
        assert endnotes.get(42) is None

    def it_can_add_a_new_endnote(self, package_: Mock):
        endnotes = self._endnotes_for(
            "w:endnotes/(w:endnote{w:type=separator,w:id=-1},"
            "w:endnote{w:type=continuationSeparator,w:id=0})",
            package_,
        )

        endnote = endnotes.add_endnote(text="note text")

        assert isinstance(endnote, Endnote)
        assert endnote.part is endnotes._notes_part
        assert endnote.endnote_id == 1
        assert [p.text for p in endnote.paragraphs] == ["note text"]
        p = endnote._note_elm.p_lst[0]
        assert p.style == "EndnoteText"
        assert p.r_lst[0].style == "EndnoteReference"
        assert p.r_lst[0][-1].tag == qn("w:endnoteRef")

    # -- fixtures --------------------------------------------------------------------------------

    @pytest.fixture
    def package_(self, request: FixtureRequest):
        return instance_mock(request, Package)

    def _endnotes_for(self, cxml: str, package_: Mock) -> Endnotes:
        endnotes_elm = cast(CT_Endnotes, element(cxml))
        return Endnotes(
            endnotes_elm,
            EndnotesPart(PackURI("/word/endnotes.xml"), CT.WML_ENDNOTES, endnotes_elm, package_),
        )


class DescribeDocumentNotes:
    """Integration tests for footnote/endnote behaviors on a real document."""

    def it_can_add_a_footnote_referenced_after_a_run(self):
        document = Document()
        paragraph = document.add_paragraph("some text")

        footnote = document.add_footnote(paragraph.runs[0], text="footnote text")

        assert isinstance(footnote, Footnote)
        assert footnote.text == "footnote text"
        # -- the reference run directly follows the anchored run --
        reference_run = paragraph.runs[0]._r.getnext()
        assert reference_run.tag == qn("w:r")
        reference_elm = reference_run.find(qn("w:footnoteReference"))
        assert reference_elm is not None
        assert int(reference_elm.get(qn("w:id"))) == footnote.footnote_id
        # -- and the footnote appears in the document collection --
        assert [f.footnote_id for f in document.footnotes] == [footnote.footnote_id]

    def it_can_add_an_endnote_referenced_after_a_run(self):
        document = Document()
        paragraph = document.add_paragraph("some text")

        endnote = document.add_endnote(paragraph.runs[0], text="endnote text")

        assert isinstance(endnote, Endnote)
        assert endnote.text == "endnote text"
        reference_run = paragraph.runs[0]._r.getnext()
        reference_elm = reference_run.find(qn("w:endnoteReference"))
        assert reference_elm is not None
        assert int(reference_elm.get(qn("w:id"))) == endnote.endnote_id
        assert [e.endnote_id for e in document.endnotes] == [endnote.endnote_id]

    def it_distinguishes_note_id_from_display_number(self):
        document = Document()
        paragraph = document.add_paragraph("a b c")
        footnote = document.add_footnote(paragraph.runs[0], text="note")

        # -- id 1 because separators hold -1 and 0; number 1 because it is the --
        # -- first referenced footnote in the document --
        assert footnote.footnote_id == 1
        assert footnote.number == 1

        # -- a footnote with no reference has no display number --
        unreferenced = document.footnotes.add_footnote(text="orphan")
        assert unreferenced.number is None

    def it_can_reference_the_same_footnote_more_than_once(self):
        document = Document()
        p1 = document.add_paragraph("first")
        p2 = document.add_paragraph("second")
        footnote = document.add_footnote(p1.runs[0], text="shared note")

        reference = footnote.add_reference(p2.runs[0])

        assert isinstance(reference, FootnoteReference)
        assert reference.footnote_id == footnote.footnote_id
        assert len(footnote.references) == 2
        # -- both references display the same number --
        assert footnote.number == 1

    def it_can_delete_a_single_reference_and_the_note_when_orphaned(self):
        document = Document()
        p1 = document.add_paragraph("first")
        p2 = document.add_paragraph("second")
        footnote = document.add_footnote(p1.runs[0], text="shared note")
        footnote.add_reference(p2.runs[0])

        # -- deleting one of two references leaves the footnote in place --
        footnote.references[0].delete()
        assert len(footnote.references) == 1
        assert [f.footnote_id for f in document.footnotes] == [footnote.footnote_id]

        # -- deleting the last reference removes the footnote as well --
        footnote.references[0].delete()
        assert len(document.footnotes) == 0

    def it_can_delete_a_footnote_with_all_its_references(self):
        document = Document()
        p1 = document.add_paragraph("first")
        p2 = document.add_paragraph("second")
        footnote = document.add_footnote(p1.runs[0], text="doomed")
        footnote.add_reference(p2.runs[0])

        footnote.delete()

        assert len(document.footnotes) == 0
        assert document.element.body.xpath(".//w:footnoteReference") == []

    def it_can_prune_unreferenced_notes_but_keeps_special_ones(self):
        document = Document()
        paragraph = document.add_paragraph("text")
        footnote = document.add_footnote(paragraph.runs[0], text="referenced")
        document.footnotes.add_footnote(text="orphan footnote")
        document.endnotes.add_endnote(text="orphan endnote")

        pruned_footnotes = document.footnotes.prune()
        pruned_endnotes = document.endnotes.prune()

        assert pruned_footnotes == [2]
        assert pruned_endnotes == [1]
        assert [f.footnote_id for f in document.footnotes] == [footnote.footnote_id]
        # -- separator and continuation-separator notes survive pruning --
        footnotes_elm = document.part._footnotes_part.element
        assert [f.type for f in footnotes_elm.footnote_lst] == [
            "separator",
            "continuationSeparator",
            None,
        ]

    def it_continues_note_ids_without_reusing_existing_ones_after_load(self):
        document = Document()
        paragraph = document.add_paragraph("text")
        first = document.add_footnote(paragraph.runs[0], text="first")
        buffer = io.BytesIO()
        document.save(buffer)

        reloaded = Document(io.BytesIO(buffer.getvalue()))
        paragraph2 = reloaded.add_paragraph("more")
        second = reloaded.add_footnote(paragraph2.runs[0], text="second")

        assert second.footnote_id > first.footnote_id
        assert [f.footnote_id for f in reloaded.footnotes] == [1, 2]

    def it_places_image_relationships_for_notes_in_the_notes_part(self):
        document = Document()
        paragraph = document.add_paragraph("text")
        footnote = document.add_footnote(paragraph.runs[0], text="see ")
        image_path = absjoin(test_file_dir, "python-icon.png")
        footnote.paragraphs[0].add_run().add_picture(image_path)

        buffer = io.BytesIO()
        document.save(buffer)
        archive = zipfile.ZipFile(io.BytesIO(buffer.getvalue()))

        # -- the image relationship belongs to the footnotes part, not the document part --
        footnotes_rels = archive.read("word/_rels/footnotes.xml.rels").decode()
        footnotes_xml = archive.read("word/footnotes.xml").decode()
        document_rels = archive.read("word/_rels/document.xml.rels").decode()
        embed_rId = footnotes_xml.split('r:embed="')[1].split('"')[0]
        assert '%s"' % embed_rId in footnotes_rels
        assert "relationships/image" in footnotes_rels
        assert "relationships/image" not in document_rels

    def it_places_hyperlink_relationships_for_notes_in_the_notes_part(self):
        document = Document()
        paragraph = document.add_paragraph("text")
        footnote = document.add_footnote(paragraph.runs[0], text="see ")
        self._add_hyperlink(footnote.paragraphs[0], "https://example.com", "example")

        buffer = io.BytesIO()
        document.save(buffer)
        archive = zipfile.ZipFile(io.BytesIO(buffer.getvalue()))

        footnotes_rels = archive.read("word/_rels/footnotes.xml.rels").decode()
        document_rels = archive.read("word/_rels/document.xml.rels").decode()
        assert "https://example.com" in footnotes_rels
        assert "https://example.com" not in document_rels

        # -- and the hyperlink is readable from the reloaded footnote --
        reloaded = Document(io.BytesIO(buffer.getvalue()))
        footnote2 = reloaded.footnotes.get(1)
        assert footnote2.paragraphs[0].hyperlinks[0].address == "https://example.com"

    def it_preserves_unmodified_notes_and_relationships_on_round_trip(self):
        # -- build a document with an image-bearing footnote and a table-bearing endnote --
        document = Document()
        p1 = document.add_paragraph("first")
        p2 = document.add_paragraph("second")
        footnote = document.add_footnote(p1.runs[0], text="with image ")
        image_path = absjoin(test_file_dir, "python-icon.png")
        footnote.paragraphs[0].add_run().add_picture(image_path)
        footnote.add_reference(p2.runs[0])
        endnote = document.add_endnote(p2.runs[0], text="with table")
        table = endnote.add_table(2, 2, Inches(3))
        table.cell(0, 0).text = "cell text"
        buffer = io.BytesIO()
        document.save(buffer)
        original_footnotes_xml = zipfile.ZipFile(io.BytesIO(buffer.getvalue())).read(
            "word/footnotes.xml"
        )

        # -- reload, add another note, and save again --
        reloaded = Document(io.BytesIO(buffer.getvalue()))
        p3 = reloaded.add_paragraph("third")
        reloaded.add_footnote(p3.runs[0], text="added later")
        buffer2 = io.BytesIO()
        reloaded.save(buffer2)

        # -- unmodified content and relationships survive the round trip --
        final = Document(io.BytesIO(buffer2.getvalue()))
        assert [f.footnote_id for f in final.footnotes] == [1, 2]
        footnote2 = final.footnotes.get(1)
        assert len(footnote2.references) == 2
        assert footnote2.text.startswith("with image")
        endnote2 = final.endnotes.get(1)
        assert endnote2.tables[0].cell(0, 0).text == "cell text"
        archive2 = zipfile.ZipFile(io.BytesIO(buffer2.getvalue()))
        assert "word/_rels/footnotes.xml.rels" in archive2.namelist()
        assert any("media" in name for name in archive2.namelist())
        # -- the original footnotes XML (separators, note 1, its image) is carried --
        # -- through unchanged; the new footnote is appended before the closing tag --
        final_footnotes_xml = archive2.read("word/footnotes.xml")
        original_without_closing = original_footnotes_xml.rsplit(b"</w:footnotes>", 1)[0]
        assert final_footnotes_xml.startswith(original_without_closing)
        # -- separator notes are preserved --
        assert b'w:type="separator"' in final_footnotes_xml
        assert b'w:type="continuationSeparator"' in final_footnotes_xml

    # -- helpers ---------------------------------------------------------------------------------

    def _add_hyperlink(self, paragraph, url: str, text: str):
        """Add an external hyperlink with `text` to `paragraph`.

        python-docx has no public add-hyperlink API; this helper uses the same
        low-level mechanics Word produces. The relationship is added to the
        part containing the paragraph, which is the footnotes part for a
        footnote paragraph.
        """
        part = paragraph.part
        r_id = part.relate_to(url, RT.HYPERLINK, is_external=True)
        hyperlink = OxmlElement("w:hyperlink")
        hyperlink.set(qn("r:id"), r_id)
        run = OxmlElement("w:r")
        t = OxmlElement("w:t")
        t.text = text
        run.append(t)
        hyperlink.append(run)
        paragraph._p.append(hyperlink)
        return hyperlink
