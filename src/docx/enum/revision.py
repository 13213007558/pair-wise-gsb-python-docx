"""Enumerations related to tracked revisions."""

from __future__ import annotations

from docx.enum.base import BaseEnum


class WD_REVISION_TYPE(BaseEnum):
    """Specifies the kind of tracked revision represented by a revision element."""

    TEXT_INSERTION = (1, "Inserted run content.")
    """Run content inserted by an author."""

    TEXT_DELETION = (2, "Deleted run content.")
    """Run content deleted by an author."""

    PARAGRAPH_MARK_INSERTION = (3, "Inserted paragraph mark.")
    """Paragraph boundary inserted by an author."""

    PARAGRAPH_MARK_DELETION = (4, "Deleted paragraph mark.")
    """Paragraph boundary deleted by an author."""

    PARAGRAPH_PROPERTY = (5, "Changed paragraph properties.")
    """Revision to paragraph formatting."""

    CHARACTER_PROPERTY = (6, "Changed character properties.")
    """Revision to run or paragraph-mark character formatting."""

    UNSUPPORTED = (0, "Revision kind that python-docx cannot accept or reject.")
    """A recognized revision outside the supported subset."""


WdRevisionType = WD_REVISION_TYPE
