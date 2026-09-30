"""Custom element classes and transformations for tracked revisions."""

from __future__ import annotations

from copy import deepcopy
from typing import Dict, Iterator, Sequence

from docx.exceptions import InvalidXmlError, UnsupportedRevisionError
from docx.oxml.ns import qn
from docx.oxml.simpletypes import ST_DateTime, ST_DecimalNumber, ST_String
from docx.oxml.text.paragraph import CT_P
from docx.oxml.xmlchemy import BaseOxmlElement, OptionalAttribute, RequiredAttribute


class CT_TrackChange(BaseOxmlElement):
    """Base class for revision elements identified by id, author, and optional date."""

    revision_id: int = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "w:id", ST_DecimalNumber
    )
    author: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:author", ST_String
    )
    date = OptionalAttribute("w:date", ST_DateTime)


class CT_RunTrackChange(CT_TrackChange):
    """A run-level insertion, deletion, or move revision wrapper."""


class CT_PPrChange(CT_TrackChange):
    """A paragraph-property revision."""


class CT_RPrChange(CT_TrackChange):
    """A run- or paragraph-mark character-property revision."""


def iter_revision_elements(root: BaseOxmlElement) -> Iterator[BaseOxmlElement]:
    """Yield recognized revision elements in document order.

    Revision snapshots inside `w:pPrChange` and `w:rPrChange` are history, not
    independently selectable revisions.
    """
    xpath = (
        ".//w:ins | .//w:del | .//w:moveFrom | .//w:moveTo"
        " | .//w:pPrChange | .//w:rPrChange"
        " | .//w:cellIns | .//w:cellDel | .//w:cellMerge"
        " | .//w:tblPrChange | .//w:tblPrExChange | .//w:tcPrChange"
        " | .//w:trPrChange | .//w:sectPrChange | .//w:tblGridChange"
    )
    elements = [element for element in root.iter() if element.tag in _REVISION_TAGS]
    for element in elements:
        if not _has_change_ancestor(element):
            yield element


_REVISION_TAGS = frozenset(
    qn(name)
    for name in (
        "w:ins",
        "w:del",
        "w:moveFrom",
        "w:moveTo",
        "w:pPrChange",
        "w:rPrChange",
        "w:cellIns",
        "w:cellDel",
        "w:cellMerge",
        "w:tblPrChange",
        "w:tblPrExChange",
        "w:tcPrChange",
        "w:trPrChange",
        "w:sectPrChange",
        "w:tblGridChange",
    )
)


def classify_revision(element: BaseOxmlElement) -> str:
    """Return the machine-readable kind of `element`."""
    local_name = element.tag.rsplit("}", 1)[-1]

    if local_name in ("ins", "del"):
        parent = element.getparent()
        if (
            parent is not None
            and parent.tag == qn("w:rPr")
            and parent.getparent() is not None
            and parent.getparent().tag == qn("w:pPr")
        ):
            return (
                "paragraph_mark_insertion"
                if local_name == "ins"
                else "paragraph_mark_deletion"
            )
        if parent is not None and parent.tag == qn("w:numPr"):
            return "unsupported"
        if parent is not None and not _is_run_revision_parent(element, parent):
            return "unsupported"
        return "text_insertion" if local_name == "ins" else "text_deletion"

    if local_name == "pPrChange":
        return "paragraph_property"
    if local_name == "rPrChange":
        return "character_property"
    return "unsupported"


def apply_revision_actions(root: BaseOxmlElement, actions: Dict[int, bool]) -> None:
    """Accept (`True`) or reject (`False`) revisions by id.

    Targets are selected from the unmodified tree. This prevents accepting an outer
    revision from implicitly changing the identity of adjacent or nested revisions that
    were not included in the batch.
    """
    elements = {}
    for element in iter_revision_elements(root):
        revision_id = int(element.get(qn("w:id")))
        if revision_id in actions:
            elements[revision_id] = element

    missing = sorted(revision_id for revision_id in actions if revision_id not in elements)
    if missing:
        raise KeyError(missing[0])

    selected = tuple(elements[revision_id] for revision_id in actions)
    for element in selected:
        if classify_revision(element) == "unsupported":
            raise UnsupportedRevisionError(
                "revision %d (%s) is not supported"
                % (element.revision_id, element.tag.rsplit("}", 1)[-1])
            )

    for element in selected:
        if element.getparent() is None:
            continue
        kind = classify_revision(element)
        if kind in (
            "text_insertion",
            "text_deletion",
            "paragraph_property",
            "character_property",
        ):
            _apply_content_or_property_revision(element, actions[element.revision_id])

    _apply_paragraph_mark_revisions(selected, actions)


def _has_change_ancestor(element: BaseOxmlElement) -> bool:
    ancestor = element.getparent()
    while ancestor is not None:
        if ancestor.tag in (qn("w:pPrChange"), qn("w:rPrChange")):
            return True
        ancestor = ancestor.getparent()
    return False


def _is_run_revision_parent(element: BaseOxmlElement, parent: BaseOxmlElement) -> bool:
    if parent.tag in frozenset(
        qn(name)
        for name in (
            "w:p",
            "w:hyperlink",
            "w:customXml",
            "w:smartTag",
            "w:dir",
            "w:bdo",
            "w:fldSimple",
            "w:ins",
            "w:del",
            "w:moveFrom",
            "w:moveTo",
        )
    ):
        return True
    if parent.tag == qn("w:sdtContent"):
        return bool(element.xpath("ancestor::w:p"))
    return False


def _apply_content_or_property_revision(element: BaseOxmlElement, accept: bool) -> None:
    kind = classify_revision(element)
    if kind == "text_insertion":
        if accept:
            _unwrap(element)
        else:
            element.getparent().remove(element)
    elif kind == "text_deletion":
        if accept:
            _accept_text_deletion(element)
        else:
            _restore_deleted_text(element)
            _unwrap(element)
    elif accept:
        element.getparent().remove(element)
    else:
        _restore_original_properties(element)


def _unwrap(element: BaseOxmlElement) -> None:
    parent = element.getparent()
    if parent is None:
        raise InvalidXmlError("revision element has no parent")
    position = parent.index(element)
    for child in list(element):
        parent.insert(position, child)
        position += 1
    parent.remove(element)


def _restore_deleted_text(deletion: BaseOxmlElement) -> None:
    """Restore deleted text without touching text inside nested pending revisions."""
    for element in deletion.xpath(
        ".//w:delText[not(ancestor::w:ins) and not(ancestor::w:del[2])]"
    ):
        element.tag = qn("w:t")


def _accept_text_deletion(deletion: BaseOxmlElement) -> None:
    """Remove deleted runs while retaining independently-pending nested insertions."""
    for child in list(deletion):
        if child.tag == qn("w:ins"):
            continue
        _remove_deleted_content(child)
        if child.getparent() is not None and len(child) == 0:
            deletion.remove(child)
    if len(deletion):
        _unwrap(deletion)
    else:
        deletion.getparent().remove(deletion)


def _remove_deleted_content(element: BaseOxmlElement) -> None:
    for child in list(element):
        if child.tag == qn("w:ins"):
            continue
        if child.tag in (qn("w:r"), qn("w:del")):
            element.remove(child)
            continue
        _remove_deleted_content(child)
        if len(child) == 0:
            element.remove(child)


def _restore_original_properties(change: BaseOxmlElement) -> None:
    original = change[0]
    container = change.getparent()
    if container is None or original is None:
        return

    if original.tag == qn("w:pPr"):
        _restore_original_paragraph_properties(container, change, original)
    elif original.tag == qn("w:rPr"):
        _restore_original_run_properties(container, change, original)


def _restore_original_paragraph_properties(
    container: BaseOxmlElement, change: BaseOxmlElement, original: BaseOxmlElement
) -> None:
    preserved = []
    for child in container:
        if child.tag == qn("w:rPr"):
            rPr = deepcopy(child)
            for nested_change in rPr.xpath("./w:rPrChange"):
                rPr.remove(nested_change)
            preserved.append(rPr)
        elif child.tag == qn("w:sectPr"):
            preserved.append(deepcopy(child))
    for child in list(container):
        container.remove(child)
    for child in original:
        container.append(deepcopy(child))
    for child in preserved:
        container.append(child)


def _restore_original_run_properties(
    container: BaseOxmlElement, change: BaseOxmlElement, original: BaseOxmlElement
) -> None:
    is_paragraph_mark_rpr = (
        container.tag == qn("w:rPr")
        and container.getparent() is not None
        and container.getparent().tag == qn("w:pPr")
    )
    preserved_markers = []
    if is_paragraph_mark_rpr:
        preserved_markers = [
            deepcopy(child)
            for child in container
            if child.tag
            in (qn("w:ins"), qn("w:del"), qn("w:moveFrom"), qn("w:moveTo"))
        ]

    for child in list(container):
        container.remove(child)
    for child in preserved_markers:
        container.append(child)
    for child in original:
        if is_paragraph_mark_rpr and child.tag in (
            qn("w:ins"),
            qn("w:del"),
            qn("w:moveFrom"),
            qn("w:moveTo"),
        ):
            continue
        container.append(deepcopy(child))


def _apply_paragraph_mark_revisions(
    selected: Sequence[BaseOxmlElement], actions: Dict[int, bool]
) -> None:
    marks = [
        element
        for element in selected
        if classify_revision(element).startswith("paragraph_mark_")
    ]
    if not marks:
        return

    snapshots = []

    for mark in marks:
        if mark.getparent() is None:
            continue
        paragraph_result = mark.xpath("ancestor::w:p[1]")
        if not paragraph_result:
            continue
        paragraph = paragraph_result[0]
        kind = classify_revision(mark)
        left = paragraph if kind == "paragraph_mark_deletion" else _preceding_paragraph(paragraph)
        right = _following_paragraph(paragraph) if left is paragraph else paragraph
        snapshots.append((mark, left, right))

    # -- Process deletions first, using the pre-mutation paragraph identities above. --
    for deletion, left, right in snapshots:
        if classify_revision(deletion) != "paragraph_mark_deletion":
            continue
        _remove_paragraph_mark(deletion)
        if not actions[deletion.revision_id] or left is None or right is None:
            continue
        if left.getparent() is not right.getparent():
            continue
        insertion = _paragraph_mark(right, "ins")
        insertion_id = int(insertion.get(qn("w:id"))) if insertion is not None else None
        use_right_properties = (
            insertion_id is None or insertion_id not in actions or actions[insertion_id]
        )
        _merge_paragraphs(left, right, use_right_properties=use_right_properties)

    # -- Rejecting an insertion is also a boundary removal, unless already merged above. --
    for insertion, left, right in snapshots:
        if classify_revision(insertion) != "paragraph_mark_insertion":
            continue
        _remove_paragraph_mark(insertion)
        if actions[insertion.revision_id] or left is None or right is None:
            continue
        if left.getparent() is not right.getparent() or right.getparent() is None:
            continue
        deletion = _paragraph_mark(left, "del")
        if deletion is not None:
            deletion_id = int(deletion.get(qn("w:id")))
            if deletion_id not in actions or not actions[deletion_id]:
                continue
        _merge_paragraphs(left, right, use_right_properties=False)


def _remove_paragraph_mark(mark: BaseOxmlElement) -> None:
    rPr = mark.getparent()
    if rPr is not None:
        rPr.remove(mark)
        if rPr.tag == qn("w:rPr") and len(rPr) == 0:
            pPr = rPr.getparent()
            if pPr is not None:
                pPr.remove(rPr)


def _paragraph_mark(paragraph: CT_P, name: str) -> BaseOxmlElement | None:
    markers = paragraph.xpath(f"./w:pPr/w:rPr/w:{name}")
    return markers[-1] if markers else None


def _following_paragraph(paragraph: CT_P) -> CT_P | None:
    candidate = paragraph.getnext()
    return candidate if candidate is not None and candidate.tag == qn("w:p") else None


def _preceding_paragraph(paragraph: CT_P) -> CT_P | None:
    candidate = paragraph.getprevious()
    return candidate if candidate is not None and candidate.tag == qn("w:p") else None


def _merge_paragraphs(left: CT_P, right: CT_P, use_right_properties: bool) -> None:
    if right.getparent() is not left.getparent():
        return

    if use_right_properties:
        if left.pPr is not None:
            left.remove(left.pPr)
        if right.pPr is not None:
            right_pPr = right.pPr
            right.remove(right_pPr)
            left.insert(0, right_pPr)

    if right.pPr is not None:
        right.remove(right.pPr)
    for child in list(right):
        left.append(child)
    right.getparent().remove(right)
