"""Field objects, a unified view of simple and complex fields in a story.

A *simple* field is a ``w:fldSimple`` element; its instruction is the
``w:instr`` attribute and its cached result is the run content it contains.

A *complex* field is a sequence of runs delimited by ``w:fldChar`` elements:
``begin``, instruction runs (``w:instrText``), an optional ``separate``, the
cached result content, and ``end``. These runs may be spread across multiple
runs and even multiple paragraphs of the same story, and fields may nest: the
instruction or the cached result of an outer field can itself contain fields.
For example the instruction of an ``IF`` field may contain a nested ``REF``
field.

Each story (document body, table cell, header, footer, ...) is parsed
independently; a field never spans a story boundary.
"""

from __future__ import annotations

import copy
from typing import List, NamedTuple, Optional

from docx.exceptions import UnclosedFieldError
from docx.oxml.ns import nsdecls, qn
from docx.oxml.parser import parse_xml


class FieldPosition(NamedTuple):
    """Location of a field within its story.

    ``paragraph_index`` is the 0-based index of the ``w:p`` element containing
    the start of the field, counting only paragraphs of the field own story.
    ``run_index`` is the 0-based index of the run containing the ``begin``
    field-char within that paragraph, or |None| for a simple field (a
    ``w:fldSimple`` element is a sibling of runs, not a run itself).
    """

    paragraph_index: int
    run_index: Optional[int]


class Field:
    """A single field in a story, either simple (``w:fldSimple``) or complex.

    Not constructed directly; obtained from ``BlockItemContainer.fields``,
    ``Document.fields``, or ``Field.children``.
    """

    def __init__(self) -> None:
        self._story_elm = None
        self._parent: Optional[Field] = None
        self._position: Optional[FieldPosition] = None
        # -- content items in the instruction and result regions; each item --
        # -- is a "w:r" or "w:hyperlink" element or a nested |Field| --
        self._instr_items: list = []
        self._result_items: list = []
        self._instr_text_elms: list = []
        # -- simple-field state --
        self._fld_simple_elm = None
        # -- complex-field state --
        self._separate_r = None
        self._end_r = None
        self._is_broken = False

    @classmethod
    def _new_simple(cls, fld_simple_elm, position, story_elm) -> "Field":
        fld = cls()
        fld._fld_simple_elm = fld_simple_elm
        fld._position = position
        fld._story_elm = story_elm
        return fld

    @classmethod
    def _new_complex(cls, begin_r, position, story_elm) -> "Field":
        fld = cls()
        fld._position = position
        fld._story_elm = story_elm
        return fld

    def __repr__(self) -> str:
        kind = "simple" if self.is_simple else "complex"
        broken = ", broken" if self.is_broken else ""
        return f"Field({kind}, instruction={self.instruction!r}{broken})"

    @property
    def children(self) -> "List[Field]":
        """Fields nested directly inside this one, in document order.

        A nested field may appear in the instruction region (before the
        ``separate`` field-char) or in the cached-result region of this field.
        """
        return [item for item in self._instr_items + self._result_items if isinstance(item, Field)]

    @property
    def instruction(self) -> str:
        """The field instruction, stripped of surrounding whitespace.

        For a complex field this is the concatenation of its own
        ``w:instrText`` runs; instructions of nested fields are not included
        (see ``children``). For a simple field it is the ``w:instr`` attribute.
        """
        if self.is_simple:
            return (self._fld_simple_elm.get(qn("w:instr")) or "").strip()
        return "".join(elm.text or "" for elm in self._instr_text_elms).strip()

    @property
    def is_broken(self) -> bool:
        """|True| if this complex field has no matching ``end`` in its story.

        A broken field has no result; the content following it in the story is
        regular story content and is never treated as its result.
        """
        return self._is_broken

    @property
    def is_complex(self) -> bool:
        """|True| if this field is delimited by ``w:fldChar`` runs."""
        return self._fld_simple_elm is None

    @property
    def is_simple(self) -> bool:
        """|True| if this field is a ``w:fldSimple`` element."""
        return self._fld_simple_elm is not None

    @property
    def parent(self) -> "Optional[Field]":
        """The field this one is nested in, or |None| for a top-level field."""
        return self._parent

    @property
    def position(self) -> Optional[FieldPosition]:
        """The location of the start of this field within its story."""
        return self._position

    @property
    def result(self) -> Optional[str]:
        """The cached result text of this field.

        Line breaks appear as newlines and tabs as tab characters; text inside
        hyperlinks and the cached results of nested fields are included.
        |None| if the field has no cached result (no ``separate`` field-char)
        or is unclosed.
        """
        if self.is_simple:
            return "".join(
                _content_text(child)
                for child in self._fld_simple_elm
                if child.tag in (qn("w:r"), qn("w:hyperlink"))
            )
        if self._is_broken or self._separate_r is None:
            return None
        return "".join(self._item_text(item) for item in self._result_items)

    def replace_result(self, text: str) -> None:
        """Replace the cached result of this field with ``text``.

        The instruction is not evaluated and Word is never invoked; ``text``
        becomes the new cached result exactly as provided. Newline and tab
        characters become line-break and tab elements. The formatting of the
        first existing result run, if any, is carried over. Content outside
        the field, bookmarks, and fields nested in the instruction region are
        left untouched. Raises ``UnclosedFieldError`` on a broken field.
        """
        if self.is_simple:
            self._replace_simple_result(text)
        else:
            self._replace_complex_result(text)

    def _add_item(self, item) -> None:
        if self._separate_r is None:
            self._instr_items.append(item)
        else:
            self._result_items.append(item)

    def _replace_complex_result(self, text: str) -> None:
        if self._is_broken or self._end_r is None:
            pos = self._position
            where = (
                f"paragraph {pos.paragraph_index}, run {pos.run_index}"
                if pos is not None
                else "unknown location"
            )
            raise UnclosedFieldError(
                f"cannot replace the result of an unclosed field (begin at {where})"
            )
        end_r = self._end_r
        separate_r = self._separate_r
        if separate_r is None:
            # -- no cached result yet; add a separator so instruction and --
            # -- result are properly delimited --
            separate_r = _new_fld_char_run("separate")
            end_r.addprevious(separate_r)
            self._separate_r = separate_r
        between = _elements_between(self._story_elm, separate_r, end_r)
        r_pr = _first_r_pr(between)
        self._remove_result_content(between)
        new_r = _new_result_run(text, r_pr)
        end_r.addprevious(new_r)
        self._result_items = [new_r]

    def _replace_simple_result(self, text: str) -> None:
        fld_simple = self._fld_simple_elm
        r_pr = _first_r_pr(list(fld_simple))
        for child in list(fld_simple):
            if child.tag in (qn("w:bookmarkStart"), qn("w:bookmarkEnd")):
                continue
            fld_simple.remove(child)
        fld_simple.append(_new_result_run(text, r_pr))

    @staticmethod
    def _remove_result_content(between) -> None:
        removed = set()
        for elm in between:
            if any(ancestor in removed for ancestor in elm.iterancestors()):
                continue
            if elm.tag == qn("w:p") or elm.tag in _RESULT_CONTENT_TAGS:
                elm.getparent().remove(elm)
                removed.add(elm)

    @staticmethod
    def _item_text(item) -> str:
        if isinstance(item, Field):
            return item.result or ""
        return _content_text(item)


def iter_fields(story_elm) -> List[Field]:
    """Return the top-level fields in ``story_elm``, in document order.

    ``story_elm`` is a story root element such as ``w:body``, ``w:tc``,
    ``w:hdr`` or ``w:ftr``. Nested stories (table cells, text boxes) are not
    descended into; each is parsed as its own story. Unclosed complex fields
    are included with ``Field.is_broken`` |True|.
    """
    return _StoryFieldParser(story_elm).parse()


# -- tags removed along with a cached result; bookmarks and other --
# -- non-content elements are deliberately preserved --
_RESULT_CONTENT_TAGS = frozenset(
    qn(tag) for tag in ("w:r", "w:hyperlink", "w:fldSimple", "w:smartTag", "w:sdt")
)

# -- tags that delimit a nested story; never descended into while parsing --
_NESTED_STORY_TAGS = frozenset(qn(tag) for tag in ("w:tc", "w:txbxContent"))


def _content_text(elm) -> str:
    """Text equivalent of a ``w:r`` or ``w:hyperlink`` element."""
    if elm.tag == qn("w:r"):
        return elm.text
    if elm.tag == qn("w:hyperlink"):
        return "".join(r.text for r in elm.iter(qn("w:r")))
    return ""


def _elements_between(story_elm, start_elm, end_elm) -> list:
    """Elements strictly between ``start_elm`` and ``end_elm``, in document order."""
    between = []
    in_range = False
    for elm in story_elm.iter():
        if elm is end_elm:
            break
        if in_range:
            between.append(elm)
        elif elm is start_elm:
            in_range = True
    return between


def _first_r_pr(elms):
    """A deep copy of the ``w:rPr`` of the first run in ``elms``, or None."""
    for elm in elms:
        if elm.tag == qn("w:r"):
            r_pr = elm.find(qn("w:rPr"))
            return copy.deepcopy(r_pr) if r_pr is not None else None
    return None


def _new_fld_char_run(fld_char_type: str):
    return parse_xml(
        '<w:r %s><w:fldChar w:fldCharType="%s"/></w:r>' % (nsdecls("w"), fld_char_type)
    )


def _new_result_run(text: str, r_pr):
    r = parse_xml("<w:r %s/>" % nsdecls("w"))
    if r_pr is not None:
        r.append(r_pr)
    # -- the CT_R text setter keeps w:rPr and maps newline/tab to w:br/w:tab --
    r.text = text
    return r


class _StoryFieldParser:
    """Parses the fields of a single story element into a tree of Fields."""

    def __init__(self, story_elm) -> None:
        self._story_elm = story_elm
        self._roots: List[Field] = []
        self._stack: List[Field] = []
        self._para_idx = -1
        self._run_idx = -1

    def parse(self) -> List[Field]:
        self._walk(self._story_elm)
        # -- any field still open at the end of the story is unclosed; it is --
        # -- reported as broken and does not capture the remaining content --
        for fld in self._stack:
            fld._is_broken = True
        return self._roots

    def _walk(self, elm) -> None:
        for child in elm:
            tag = child.tag
            if tag == qn("w:p"):
                self._para_idx += 1
                self._run_idx = -1
                self._walk(child)
            elif tag == qn("w:r"):
                self._run_idx += 1
                self._on_run(child)
            elif tag == qn("w:hyperlink"):
                self._on_content_item(child)
            elif tag == qn("w:fldSimple"):
                fld = Field._new_simple(child, self._position(None), self._story_elm)
                self._attach(fld)
            elif tag in _NESTED_STORY_TAGS:
                continue
            else:
                self._walk(child)

    def _position(self, run_idx) -> FieldPosition:
        return FieldPosition(self._para_idx, run_idx)

    def _attach(self, fld: Field) -> None:
        if self._stack:
            parent = self._stack[-1]
            fld._parent = parent
            parent._add_item(fld)
        else:
            self._roots.append(fld)

    def _on_run(self, r) -> None:
        fld_char = r.find(qn("w:fldChar"))
        if fld_char is not None:
            self._on_fld_char(fld_char.get(qn("w:fldCharType")), r)
            return
        instr_text = r.find(qn("w:instrText"))
        if instr_text is not None:
            if self._stack:
                self._stack[-1]._instr_text_elms.append(instr_text)
            return
        self._on_content_item(r)

    def _on_content_item(self, elm) -> None:
        if self._stack:
            self._stack[-1]._add_item(elm)

    def _on_fld_char(self, fld_char_type, r) -> None:
        if fld_char_type == "begin":
            fld = Field._new_complex(r, self._position(self._run_idx), self._story_elm)
            self._attach(fld)
            self._stack.append(fld)
        elif fld_char_type == "separate":
            if self._stack:
                self._stack[-1]._separate_r = r
        elif fld_char_type == "end":
            if self._stack:
                self._stack.pop()._end_r = r
