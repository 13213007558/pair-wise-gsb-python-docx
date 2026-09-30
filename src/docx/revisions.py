"""Proxy types for tracked revisions in a Word document."""

from __future__ import annotations

from datetime import datetime
from typing import TYPE_CHECKING, Iterable, Iterator, Sequence

from docx.enum.revision import WD_REVISION_TYPE
from docx.oxml.revisions import (
    apply_revision_actions,
    classify_revision,
    iter_revision_elements,
)

if TYPE_CHECKING:
    from docx.oxml.document import CT_Document


_REVISION_TYPE_BY_KIND = {
    "text_insertion": WD_REVISION_TYPE.TEXT_INSERTION,
    "text_deletion": WD_REVISION_TYPE.TEXT_DELETION,
    "paragraph_mark_insertion": WD_REVISION_TYPE.PARAGRAPH_MARK_INSERTION,
    "paragraph_mark_deletion": WD_REVISION_TYPE.PARAGRAPH_MARK_DELETION,
    "paragraph_property": WD_REVISION_TYPE.PARAGRAPH_PROPERTY,
    "character_property": WD_REVISION_TYPE.CHARACTER_PROPERTY,
    "unsupported": WD_REVISION_TYPE.UNSUPPORTED,
}


class Revision:
    """A single tracked revision."""

    def __init__(self, element, revisions: Revisions):
        self._element = element
        self._revisions = revisions

    def accept(self) -> None:
        """Apply this revision to the document."""
        self._revisions.accept(self.id)

    @property
    def author(self) -> str | None:
        """Author recorded for this revision, or |None| if absent."""
        return self._element.author

    @property
    def date(self) -> datetime | None:
        """Date recorded for this revision, or |None| if absent."""
        return self._element.date

    @property
    def id(self) -> int:
        """Stable integer revision ID used by accept and reject operations."""
        return self._element.revision_id

    def reject(self) -> None:
        """Restore the document content as it existed before this revision."""
        self._revisions.reject(self.id)

    @property
    def type(self) -> WD_REVISION_TYPE:
        """Kind of this revision."""
        return _REVISION_TYPE_BY_KIND[classify_revision(self._element)]


class Revisions:
    """Collection of tracked revisions in a document."""

    def __init__(self, document_element: CT_Document):
        self._element = document_element

    def accept(self, revision_ids: int | Revision | Iterable[int | Revision]) -> None:
        """Accept one revision or all revisions in an iterable."""
        self._apply(revision_ids, True)

    def accept_all(self, author: str | None = None) -> None:
        """Accept every revision, optionally limited to revisions by `author`."""
        self._apply(self._ids(author), True)

    def __getitem__(self, index: int) -> Revision:
        return list(self)[index]

    def __iter__(self) -> Iterator[Revision]:
        for element in iter_revision_elements(self._element):
            yield Revision(element, self)

    def __len__(self) -> int:
        return sum(1 for _ in iter_revision_elements(self._element))

    def reject(self, revision_ids: int | Revision | Iterable[int | Revision]) -> None:
        """Reject one revision or all revisions in an iterable."""
        self._apply(revision_ids, False)

    def reject_all(self, author: str | None = None) -> None:
        """Reject every revision, optionally limited to revisions by `author`."""
        self._apply(self._ids(author), False)

    def _apply(
        self, revision_ids: int | Revision | Iterable[int | Revision], accept: bool
    ) -> None:
        ids = _normalise_revision_ids(revision_ids)
        apply_revision_actions(
            self._element, {revision_id: accept for revision_id in ids}
        )

    def _ids(self, author: str | None) -> Sequence[int]:
        revisions = list(self)
        if author is None:
            return [revision.id for revision in revisions]
        return [revision.id for revision in revisions if revision.author == author]


def _normalise_revision_ids(
    revision_ids: int | Revision | Iterable[int | Revision],
) -> tuple[int, ...]:
    if isinstance(revision_ids, int):
        return (revision_ids,)
    if isinstance(revision_ids, Revision):
        return (revision_ids.id,)
    return tuple(
        revision_id if isinstance(revision_id, int) else revision_id.id
        for revision_id in revision_ids
    )
