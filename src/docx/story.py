# pyright: reportPrivateUsage=false

"""Unified traversal of Word "stories".

A Word document is a collection of *stories*. The main document body is one story
and each header and footer part is its own story. Every story contains block items
(paragraphs and tables), and those block items can contain further block items,
such as a table nested in a table cell or content held inside a text box.

This module provides a single traversal interface, :meth:`Story.iter_items`, that
visits every paragraph and table of a story, including content nested in table
cells and in both DrawingML (``<wps:txbx>``) and legacy VML (``<v:textbox>``)
text boxes. Each yielded :class:`StoryItem` identifies the package part that
owns the item, the chain of nested containers from the story root to the item,
and -- for items inside a text box -- the paragraph that anchors the text box.

Markup-compatibility branches (``<mc:AlternateContent>``) are resolved against an
explicit *capability set*: the first ``<mc:Choice>`` whose required capabilities
are all available is used, otherwise ``<mc:Fallback>`` is used. Exactly one
branch is traversed, so content appearing in both a DrawingML choice and a VML
fallback is never extracted twice.

Floating text boxes are visited at the position of their XML anchor (the
paragraph containing the drawing), in the order the anchors appear in the XML.
This is **not** Word's visual layout order -- floating shapes can be rendered
anywhere on a page -- it is only a deterministic document-order traversal.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Iterable, Iterator, Sequence, Tuple, cast

from docx.blkcntnr import BlockItemContainer
from docx.oxml.ns import qn
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.shared import StoryChild
from docx.table import Table, _Cell
from docx.text.paragraph import Paragraph

if TYPE_CHECKING:
    from docx.document import Document
    from docx.oxml.section import CT_HdrFtrRef
    from docx.oxml.xmlchemy import BaseOxmlElement
    from docx.parts.story import StoryPart

# -- tag constants (Clark notation) ----------------------------------------

_W_P = qn("w:p")
_W_TBL = qn("w:tbl")
_W_PICT = qn("w:pict")
_W_TXBX_CONTENT = qn("w:txbxContent")
_WP_INLINE = qn("wp:inline")
_WP_ANCHOR = qn("wp:anchor")
_MC_ALTERNATE_CONTENT = qn("mc:AlternateContent")
_MC_CHOICE = qn("mc:Choice")
_MC_FALLBACK = qn("mc:Fallback")


#: Capabilities available by default when resolving ``<mc:AlternateContent>``.
#:
#: ``"wps"`` is the capability required by DrawingML text-box content
#: (``<wps:wsp>`` shapes). With this default a DrawingML ``<mc:Choice>`` is
#: selected and its VML ``<mc:Fallback>`` is ignored. Pass an explicit set,
#: such as ``frozenset()``, to force the VML fallback branch.
DEFAULT_CAPABILITIES: frozenset[str] = frozenset({"wps"})


class StoryReference(tuple):
    """Reference to a header or footer story from a section.

    ``section_index`` is the index of the section within
    :attr:`~.Document.sections` and ``header_footer_type`` is the
    :class:`~docx.enum.section.WD_HEADER_FOOTER` member identifying the slot
    (primary, first-page, or even-page) the part is referenced from.
    """

    __slots__ = ()

    def __new__(cls, section_index: int, header_footer_type):
        return tuple.__new__(cls, (section_index, header_footer_type))

    @property
    def section_index(self) -> int:
        return self[0]

    @property
    def header_footer_type(self):
        return self[1]


class StoryContainer:
    """One level in the chain of nested containers leading to a story item.

    ``kind`` is one of ``"body"``, ``"header"``, ``"footer"``, ``"cell"``, or
    ``"textbox"`` and ``proxy`` is the python-docx object representing that
    container (a block-item container, |_Cell|, or |TextBox|).
    ``index`` records ``(row_index, column_index)`` for a table cell and is
    |None| otherwise.
    """

    def __init__(self, kind: str, proxy: object, index: Tuple[int, int] | None = None):
        self.kind = kind
        self.proxy = proxy
        self.index = index

    def __repr__(self) -> str:
        if self.kind == "cell" and self.index is not None:
            return "StoryContainer(cell %d,%d)" % self.index
        return "StoryContainer(%s)" % self.kind


class TextBoxAnchor:
    """Association between a text box and the paragraph that anchors it in XML.

    ``paragraph`` is the |Paragraph| containing the ``<w:drawing>`` or
    ``<w:pict>`` that holds the text box and ``is_floating`` is |True| for a
    floating DrawingML ``<wp:anchor>`` or a legacy VML shape and |False| for an
    inline DrawingML ``<wp:inline>``. ``order`` is the zero-based position of
    the text-box anchor in the story's XML traversal.

    ``order`` reflects XML anchor order, which is **not** Word's visual layout
    order; floating text boxes may be rendered at a position unrelated to their
    anchor paragraph.
    """

    def __init__(self, paragraph: Paragraph, is_floating: bool, order: int):
        self.paragraph = paragraph
        self.is_floating = is_floating
        self.order = order

    def __repr__(self) -> str:
        placement = "floating" if self.is_floating else "inline"
        return "TextBoxAnchor(%s #%d)" % (placement, self.order)


class StoryItem:
    """One paragraph or table produced by :meth:`Story.iter_items`.

    ``content`` is the |Paragraph| or |Table|, ``part`` is the package part
    that owns it (the document part for body content, the relevant header or
    footer part for header/footer content), ``path`` is the chain of nested
    :class:`StoryContainer` objects from the story root to the item, and
    ``anchor`` is the :class:`TextBoxAnchor` of the nearest enclosing text box
    or |None| when the item sits directly in story flow content.
    """

    def __init__(
        self,
        content: Paragraph | Table,
        part: StoryPart,
        path: Tuple[StoryContainer, ...],
        anchor: TextBoxAnchor | None,
    ):
        self.content = content
        self.part = part
        self.path = path
        self.anchor = anchor

    @property
    def kind(self) -> str:
        """``\"paragraph\"`` or ``\"table\"``."""
        return "table" if isinstance(self.content, Table) else "paragraph"

    def __repr__(self) -> str:
        text = getattr(self.content, "text", None)
        detail = "%r" % text if isinstance(text, str) and text else self.kind
        return "StoryItem(%s in %s)" % (detail, " > ".join(c.kind for c in self.path))


class TextBox(StoryChild):
    """Proxy for a ``<w:txbxContent>`` text-box element.

    A text box behaves like a block-item container (it can contain paragraphs
    and tables, including nested tables and further text boxes). Its ``.part``
    -- and therefore relationship resolution for images and hyperlinks inside
    the text box -- follows the parent chain, so a text box in a header resolves
    its relationships against the header part, not the document part.
    """

    def __init__(
        self,
        txbxContent: BaseOxmlElement,
        parent: BlockItemContainer | TextBox,
        anchor: TextBoxAnchor | None = None,
    ):
        super(TextBox, self).__init__(parent)
        self._element = self._txbxContent = txbxContent
        self.anchor = anchor

    def iter_inner_content(self) -> Iterator[Paragraph | Table]:
        """Generate each |Paragraph| or |Table| directly in this text box.

        Items appear in document order. Nested tables and nested text boxes are
        not expanded here; use :meth:`Story.iter_items` for a full traversal.
        """
        for element in self._txbxContent.xpath("./w:p | ./w:tbl"):
            if element.tag == _W_P:
                yield Paragraph(cast("CT_P", element), self)
            else:
                yield Table(cast("CT_Tbl", element), self)

    @property
    def paragraphs(self) -> list[Paragraph]:
        """List of the paragraphs directly in this text box, in document order."""
        return [Paragraph(cast("CT_P", p), self) for p in self._txbxContent.xpath("./w:p")]

    @property
    def tables(self) -> list[Table]:
        """List of the tables directly in this text box, in document order."""
        return [Table(cast("CT_Tbl", tbl), self) for tbl in self._txbxContent.xpath("./w:tbl")]


class _StoryPartProvider:
    """Minimal parent that locates content directly on a story part."""

    def __init__(self, part: StoryPart):
        self._part = part

    @property
    def part(self) -> StoryPart:
        return self._part


class Story:
    """A single Word story: the document body or one header/footer part.

    ``kind`` is ``\"body\"``, ``\"header\"``, or ``\"footer\"``, ``part`` is the
    owning package part, and ``references`` is a tuple of
    :class:`StoryReference` values describing the sections that reference a
    header/footer part (empty for the body story). Use :meth:`iter_items` for
    the unified paragraph/table traversal.
    """

    def __init__(
        self,
        kind: str,
        part: StoryPart,
        root_container: BlockItemContainer,
        references: Sequence[StoryReference] = (),
    ):
        self.kind = kind
        self.part = part
        self._root_container = root_container
        self.references: Tuple[StoryReference, ...] = tuple(references)

    @property
    def sections(self) -> Tuple[int, ...]:
        """Indices of sections referencing this story (headers/footers only).

        Multiple sections referencing the same header or footer part share one
        |Story| and each appears here.
        """
        return tuple(r.section_index for r in self.references)

    def iter_items(self, capabilities: Iterable[str] | None = None) -> Iterator[StoryItem]:
        """Generate each paragraph and table in this story, including nested content.

        Every item is a :class:`StoryItem` carrying its owning part, the chain
        of nested containers, and the text-box anchor when applicable.

        ``capabilities`` is the set of markup-compatibility capability tokens
        (values of ``<mc:Choice @Requires>``) the consumer supports; it
        defaults to :data:`DEFAULT_CAPABILITIES` (DrawingML ``wps``). Only one
        branch of each ``<mc:AlternateContent>`` is traversed.

        Content of floating text boxes is emitted at the XML anchor position in
        XML anchor order, not in Word's visual layout order.
        """
        capability_set = DEFAULT_CAPABILITIES if capabilities is None else frozenset(capabilities)
        path = (StoryContainer(self.kind, self._root_container),)
        state = _TraversalState()
        yield from _iter_block_items(
            self._root_container,
            self._root_container._element,
            self.part,
            path,
            None,
            capability_set,
            state,
        )

    def _add_reference(self, reference: StoryReference) -> None:
        self.references = self.references + (reference,)


class Stories(Sequence):
    """The stories of a document: the body plus one per header/footer part.

    Header and footer parts are returned as separate stories; a part referenced
    by multiple sections appears exactly once, with each referencing section
    recorded in :attr:`Story.sections`.
    """

    def __init__(self, document: Document):
        super(Stories, self).__init__()
        self._document = document
        main_story = Story("body", document.part, document._body)
        self._stories: list[Story] = [main_story]
        self._by_partname = {str(document.part.partname): main_story}
        self._collect_header_footer_stories()

    @property
    def main(self) -> Story:
        """The |Story| for the main document body."""
        return self._stories[0]

    def for_part(self, part: StoryPart) -> Story | None:
        """Return the |Story| whose content lives in `part`, or |None|.

        Matches by part reference/partname, so a header or footer part obtained
        through a section (for example ``section.header.part``) locates the
        single story shared by every section referencing that part.
        """
        return self._by_partname.get(str(part.partname))

    def iter_items(self, capabilities: Iterable[str] | None = None) -> Iterator[StoryItem]:
        """Generate items from every story, body first then headers and footers."""
        for story in self._stories:
            yield from story.iter_items(capabilities)

    def __getitem__(self, key):  # type: ignore[override]
        return self._stories[key]

    def __len__(self) -> int:
        return len(self._stories)

    def _collect_header_footer_stories(self) -> None:
        document_part = self._document.part
        for section_index, section in enumerate(self._document.sections):
            sectPr = section._sectPr
            xpath = "./w:headerReference | ./w:footerReference"
            for reference in sectPr.xpath(xpath):
                reference = cast("CT_HdrFtrRef", reference)
                if reference.tag == qn("w:headerReference"):
                    kind = "header"
                    part = document_part.header_part(reference.rId)
                else:
                    kind = "footer"
                    part = document_part.footer_part(reference.rId)
                partname = str(part.partname)
                story_reference = StoryReference(section_index, reference.type_)
                story = self._by_partname.get(partname)
                if story is None:
                    root_container = BlockItemContainer(part.element, _StoryPartProvider(part))
                    story = Story(kind, part, root_container, [story_reference])
                    self._stories.append(story)
                    self._by_partname[partname] = story
                else:
                    story._add_reference(story_reference)


class _TraversalState:
    """Mutable per-story traversal state tracking text-box anchor order."""

    def __init__(self) -> None:
        self._anchor_count = 0

    def next_anchor_order(self) -> int:
        order = self._anchor_count
        self._anchor_count += 1
        return order


# -- traversal primitives --------------------------------------------------


def _select_ac_branch(element, capabilities: frozenset[str]):
    """Return the single child branch of a `mc:AlternateContent` to traverse.

    The first `mc:Choice` whose space-delimited `Requires` capabilities are all
    available wins; otherwise `mc:Fallback` is returned (or |None| when the
    element has no usable branch).
    """
    fallback = None
    for child in element:
        if child.tag == _MC_CHOICE:
            required = child.get("Requires", "").split()
            if all(token in capabilities for token in required):
                return child
        elif child.tag == _MC_FALLBACK and fallback is None:
            fallback = child
    return fallback


def _iter_selected_children(element, capabilities: frozenset[str]):
    """Generate non-`AlternateContent` children of `element` in document order.

    For an `mc:AlternateContent` child, only the selected branch is descended
    into, so content in a rejected Choice or Fallback is never visited.
    """
    for child in element:
        if child.tag == _MC_ALTERNATE_CONTENT:
            branch = _select_ac_branch(child, capabilities)
            if branch is not None:
                # -- the branch holds the elements the Choice/Fallback wraps --
                yield from _iter_selected_children(branch, capabilities)
        else:
            yield child


def _iter_block_elements(element, capabilities: frozenset[str]):
    """Generate direct `w:p` and `w:tbl` block children of `element`.

    Markup-compatibility wrappers at block level are resolved first.
    """
    for child in _iter_selected_children(element, capabilities):
        if child.tag == _W_P or child.tag == _W_TBL:
            yield child


def _iter_txbx_contents(element, capabilities: frozenset[str]):
    """Generate the outermost `w:txbxContent` elements anchored under `element`.

    `element` is a block container such as a paragraph or table cell. Text
    boxes are yielded in XML document order. Only AlternateContent branches
    selected against `capabilities` are searched, so the same text-box content
    is never produced twice from a Choice and its Fallback. Text boxes nested
    inside another text box are not yielded here; they are reached by
    traversing the outer text box's own content.
    """
    for child in _iter_selected_children(element, capabilities):
        if child.tag == _W_TXBX_CONTENT:
            yield child
        else:
            yield from _iter_txbx_contents(child, capabilities)


def _is_floating_txbx(txbxContent) -> bool:
    """|True| if `txbxContent` is hosted by a floating drawing anchor.

    A DrawingML ``<wp:anchor>`` and a legacy VML ``<w:pict>`` shape are
    treated as floating; a DrawingML ``<wp:inline>`` is inline. Nearest
    ancestor wins, which correctly classifies a nested text box by its own
    host drawing.
    """
    for ancestor in txbxContent.iterancestors():
        if ancestor.tag in (_WP_ANCHOR, _W_PICT):
            return True
        if ancestor.tag == _WP_INLINE:
            return False
    return False


def _iter_block_items(container, element, part, path, anchor, capabilities, state):
    """Recursively generate :class:`StoryItem` for paragraphs and tables.

    Each top-level paragraph is yielded and then any text boxes it anchors are
    traversed (in XML anchor order); each top-level table is yielded and then
    its cells are traversed with a ``cell`` container added to the path.
    """
    for block_element in _iter_block_elements(element, capabilities):
        if block_element.tag == _W_P:
            paragraph = Paragraph(cast("CT_P", block_element), container)
            yield StoryItem(paragraph, part, path, anchor)
            yield from _iter_paragraph_textboxes(
                container, paragraph, part, path, anchor, capabilities, state
            )
        else:
            table = Table(cast("CT_Tbl", block_element), container)
            yield StoryItem(table, part, path, anchor)
            yield from _iter_table_cells(table, part, path, anchor, capabilities, state)


def _iter_paragraph_textboxes(container, paragraph, part, path, anchor, capabilities, state):
    """Traverse each text box anchored by `paragraph` in XML anchor order."""
    for txbxContent in _iter_txbx_contents(paragraph._p, capabilities):
        text_box = TextBox(txbxContent, container)
        text_box_anchor = TextBoxAnchor(
            paragraph,
            _is_floating_txbx(txbxContent),
            state.next_anchor_order(),
        )
        text_box.anchor = text_box_anchor
        text_box_path = path + (StoryContainer("textbox", text_box),)
        yield from _iter_block_items(
            text_box,
            txbxContent,
            part,
            text_box_path,
            text_box_anchor,
            capabilities,
            state,
        )


def _iter_table_cells(table, part, path, anchor, capabilities, state):
    """Traverse the block items in each cell of `table`, row by row."""
    for row_index, tr in enumerate(table._tbl.tr_lst):
        for column_index, tc in enumerate(tr.tc_lst):
            cell = _Cell(tc, table)
            cell_path = path + (StoryContainer("cell", cell, (row_index, column_index)),)
            yield from _iter_block_items(cell, tc, part, cell_path, anchor, capabilities, state)
