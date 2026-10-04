"""Custom element classes for content controls ('w:sdt' and related elements)."""

from __future__ import annotations

import copy
from typing import Iterator, Set

from docx.oxml.ns import qn
from docx.oxml.parser import OxmlElement
from docx.oxml.xmlchemy import BaseOxmlElement


class CT_Sdt(BaseOxmlElement):
    """'<w:sdt>' element, a structured document tag (content control)."""

    @property
    def pr(self):
        """The 'w:sdtPr' (properties) child, or |None| if not present."""
        return self.find(qn("w:sdtPr"))

    @property
    def content(self):
        """The 'w:sdtContent' child, or |None| if not present."""
        return self.find(qn("w:sdtContent"))

    @property
    def sdt_id(self) -> int | None:
        """Integer value of 'w:sdtPr/w:id', or |None| if not present."""
        pr = self.pr
        if pr is None:
            return None
        id_elm = pr.find(qn("w:id"))
        if id_elm is None:
            return None
        val = id_elm.get(qn("w:val"))
        return int(val) if val is not None else None

    @property
    def alias(self) -> str | None:
        """Value of 'w:sdtPr/w:alias/@w:val', the control title, if present."""
        return self._pr_child_val("w:alias")

    @property
    def sdt_tag(self) -> str | None:
        """Value of 'w:sdtPr/w:tag/@w:val', if present."""
        return self._pr_child_val("w:tag")

    @property
    def binding(self):
        """The 'w:sdtPr/w:dataBinding' element, or |None| if control is unbound."""
        pr = self.pr
        if pr is None:
            return None
        return pr.find(qn("w:dataBinding"))

    @property
    def is_checkbox(self) -> bool:
        """True when this control is a 'w14:checkbox' content control."""
        pr = self.pr
        return pr is not None and pr.find(qn("w14:checkbox")) is not None

    @property
    def is_plain_text(self) -> bool:
        """True when this control is a plain-text ('w:text') content control."""
        pr = self.pr
        return pr is not None and pr.find(qn("w:text")) is not None

    @property
    def is_repeating_section(self) -> bool:
        """True when this control is a 'w15:repeatingSection' template control."""
        pr = self.pr
        return pr is not None and pr.find(qn("w15:repeatingSection")) is not None

    @property
    def is_repeating_section_item(self) -> bool:
        """True when this control is a generated 'w15:repeatingSectionItem' copy."""
        pr = self.pr
        return pr is not None and pr.find(qn("w15:repeatingSectionItem")) is not None

    def iter_sdt(self) -> "Iterator[CT_Sdt]":
        """Generate each descendant 'w:sdt' element in document order."""
        return (e for e in self.iter(qn("w:sdt")) if e is not self)

    def display_text(self) -> str:
        """Concatenated text of all 'w:t' descendants of 'w:sdtContent'."""
        content = self.content
        if content is None:
            return ""
        return "".join(t.text or "" for t in content.iter(qn("w:t")))

    def _pr_child_val(self, tag: str) -> str | None:
        pr = self.pr
        if pr is None:
            return None
        elm = pr.find(qn(tag))
        if elm is None:
            return None
        return elm.get(qn("w:val"))

    def set_display_text(self, text: str) -> None:
        """Replace the displayed content of this control with a single run of text.

        Formatting of the first existing run is preserved when present. Block-level
        controls (content holds paragraphs) receive a single new paragraph; inline
        controls receive a single new run.
        """
        content = self.content
        if content is None:
            return
        first_r = content.find(".//" + qn("w:r"))
        rPr = first_r.find(qn("w:rPr")) if first_r is not None else None
        is_block = content.find(qn("w:p")) is not None
        for child in list(content):
            content.remove(child)
        r = OxmlElement("w:r")
        if rPr is not None:
            r.append(copy.deepcopy(rPr))
        t = OxmlElement("w:t")
        t.text = text
        if text != text.strip():
            t.set(qn("xml:space"), "preserve")
        r.append(t)
        if is_block:
            p = OxmlElement("w:p")
            p.append(r)
            content.append(p)
        else:
            content.append(r)

    def set_checkbox_state(self, checked: bool) -> None:
        """Set checked state of a 'w14:checkbox' control and its displayed glyph."""
        pr = self.pr
        if pr is None:
            return
        checkbox = pr.find(qn("w14:checkbox"))
        if checkbox is None:
            return
        checked_elm = checkbox.find(qn("w14:checked"))
        if checked_elm is None:
            checked_elm = OxmlElement("w14:checked")
            checkbox.insert(0, checked_elm)
        checked_elm.set(qn("w14:val"), "1" if checked else "0")
        state_tag = "w14:checkedState" if checked else "w14:uncheckedState"
        state_elm = checkbox.find(qn(state_tag))
        glyph = "☒" if checked else "☐"
        if state_elm is not None:
            hex_val = state_elm.get(qn("w14:val"))
            if hex_val:
                glyph = chr(int(hex_val, 16))
        self.set_display_text(glyph)

    def mark_as_repeating_item(self) -> None:
        """Convert this repeating-section template into a generated item copy."""
        pr = self.pr
        if pr is None:
            return
        repeating = pr.find(qn("w15:repeatingSection"))
        if repeating is not None:
            pr.remove(repeating)
        if pr.find(qn("w15:repeatingSectionItem")) is None:
            pr.append(OxmlElement("w15:repeatingSectionItem"))

    def renew_ids(self, used_ids: "Set[int]") -> None:
        """Assign a fresh unique id to this control and each descendant control.

        New ids are one greater than the largest id in used_ids, which is updated
        as ids are assigned so ids remain unique across the document.
        """
        for pr in self.xpath(".//w:sdtPr"):
            id_elm = pr.find(qn("w:id"))
            new_id = max(used_ids) + 1 if used_ids else 1
            used_ids.add(new_id)
            if id_elm is None:
                id_elm = OxmlElement("w:id")
                pr.append(id_elm)
            id_elm.set(qn("w:val"), str(new_id))

    def clone(self) -> "CT_Sdt":
        """A deep copy of this element, detached from the tree."""
        return copy.deepcopy(self)
