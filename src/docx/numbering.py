"""Read-only resolution of list-numbering labels for document paragraphs.

This module answers the question "what number/bullet label would Word show next to
this paragraph?" without modifying the document in any way. In particular,
`numbering.xml` is only ever read, never created or mutated.

The resolver understands:

- numbering applied directly on a paragraph (`w:pPr/w:numPr`),
- numbering inherited through the paragraph-style `basedOn` chain,
- concrete list instances (`w:num`) and their abstract definitions
  (`w:abstractNum`), including `w:lvlOverride/w:startOverride`,
- independent counters per `numId` and restart of deeper levels whenever a
  higher level advances,
- `decimal`, `upperLetter`, `lowerLetter`,
  `upperRoman`, `lowerRoman` and `bullet`
  formats, plus `%1`..`%9` level placeholders in
  `w:lvlText`.

A paragraph whose effective `numId` is `0` has its numbering
cancelled: it produces no item and does not advance any list counter. Levels
using an unsupported `w:numFmt` are reported with a reason instead of
being silently rendered as decimal.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Dict, Iterator, List, Optional, Tuple

from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

_MAX_LEVELS = 9

SUPPORTED_NUM_FMTS = ("decimal", "upperLetter", "lowerLetter", "upperRoman", "lowerRoman", "bullet")

_LEVEL_PLACEHOLDER_RE = re.compile(r"%([1-9])")

_ROMAN_VALUES = (
    (1000, "M"),
    (900, "CM"),
    (500, "D"),
    (400, "CD"),
    (100, "C"),
    (90, "XC"),
    (50, "L"),
    (40, "XL"),
    (10, "X"),
    (9, "IX"),
    (5, "V"),
    (4, "IV"),
    (1, "I"),
)


@dataclass
class ListItem:
    """The resolved numbering of a single numbered paragraph, in document order.

    `label` is the text Word would render for the item marker, e.g.
    `"2.3.a"` or a bullet character. It is |None| when the label cannot
    be produced, in which case `unsupported_reason` explains why.
    `level` is the zero-based `ilvl`. `num_id`
    and `abstract_num_id` identify the numbering definition the
    paragraph draws its label from.
    """

    paragraph: Paragraph
    label: Optional[str]
    level: int
    num_id: int
    abstract_num_id: Optional[int]
    num_fmt: Optional[str]
    lvl_text: Optional[str]
    is_bullet: bool
    unsupported_reason: Optional[str] = None


@dataclass
class _LevelDef:
    """Effective definition of a single numbering level."""

    start: int
    num_fmt: str
    lvl_text: str


@dataclass
class _NumInstance:
    """A concrete `w:num` list instance: abstract reference plus overrides."""

    num_id: int
    abstract_num_id: int
    start_overrides: Dict[int, int]


class _ListCounter:
    """Running counter state for one `numId`, across its nine levels."""

    def __init__(self, levels: Dict[int, _LevelDef], start_overrides: Dict[int, int]):
        self._levels = levels
        self._starts = {
            ilvl: start_overrides.get(ilvl, levels[ilvl].start if ilvl in levels else 1)
            for ilvl in range(_MAX_LEVELS)
        }
        self._values = dict(self._starts)
        self._started = [False] * _MAX_LEVELS

    def advance(self, ilvl: int) -> None:
        """Move the counter for `ilvl` to its next value, restarting deeper levels."""
        if self._started[ilvl]:
            self._values[ilvl] += 1
        else:
            self._started[ilvl] = True
        for deeper in range(ilvl + 1, _MAX_LEVELS):
            self._values[deeper] = self._starts[deeper]
            self._started[deeper] = False

    def value(self, ilvl: int) -> int:
        return self._values[ilvl]


def iter_list_items(document) -> Iterator[ListItem]:
    """Generate a |ListItem| for each numbered paragraph in `document`.

    Paragraphs are visited in explicit document order, body paragraphs and table
    cell paragraphs (including nested tables) interleaved as they appear.
    Paragraphs without numbering, or with numbering cancelled by `numId`
    of 0, are skipped and do not affect any list counter. The document is not
    modified; in particular `numbering.xml` is treated as read-only.
    """
    numbering_elm = _numbering_element(document)
    if numbering_elm is None:
        return
    abstract_nums, num_instances = _parse_numbering(numbering_elm)
    counters: Dict[int, _ListCounter] = {}
    for paragraph in _iter_paragraphs(document):
        reference = _effective_num_reference(paragraph)
        if reference is None:
            continue
        num_id, ilvl = reference
        if num_id == 0:
            # numbering explicitly cancelled; does not advance any list
            continue
        instance = num_instances.get(num_id)
        if instance is None:
            yield ListItem(
                paragraph=paragraph,
                label=None,
                level=ilvl,
                num_id=num_id,
                abstract_num_id=None,
                num_fmt=None,
                lvl_text=None,
                is_bullet=False,
                unsupported_reason="no <w:num> element with numId %d" % num_id,
            )
            continue
        levels = abstract_nums.get(instance.abstract_num_id, {})
        if num_id not in counters:
            counters[num_id] = _ListCounter(levels, instance.start_overrides)
        counter = counters[num_id]
        counter.advance(ilvl)
        label, reason = _render_label(levels, counter, ilvl)
        level_def = levels.get(ilvl)
        yield ListItem(
            paragraph=paragraph,
            label=label,
            level=ilvl,
            num_id=num_id,
            abstract_num_id=instance.abstract_num_id,
            num_fmt=level_def.num_fmt if level_def else None,
            lvl_text=level_def.lvl_text if level_def else None,
            is_bullet=bool(level_def and level_def.num_fmt == "bullet"),
            unsupported_reason=reason,
        )


def _numbering_element(document):
    """The root `w:numbering` element, or |None| if there is no numbering part.

    Read-only: unlike `DocumentPart.numbering_part`, this never creates
    a numbering part when one is absent.
    """
    try:
        numbering_part = document.part.part_related_by(RT.NUMBERING)
    except KeyError:
        return None
    return numbering_part.element


def _parse_numbering(
    numbering_elm,
) -> Tuple[Dict[int, Dict[int, _LevelDef]], Dict[int, _NumInstance]]:
    """Extract abstract level definitions and concrete num instances from numbering XML."""
    abstract_nums: Dict[int, Dict[int, _LevelDef]] = {}
    for abstract_num_elm in numbering_elm.findall(qn("w:abstractNum")):
        abstract_num_id = int(abstract_num_elm.get(qn("w:abstractNumId")))
        levels: Dict[int, _LevelDef] = {}
        for lvl_elm in abstract_num_elm.findall(qn("w:lvl")):
            ilvl = int(lvl_elm.get(qn("w:ilvl")))
            levels[ilvl] = _LevelDef(
                start=_child_val(lvl_elm, "w:start", 1),
                num_fmt=_child_attr(lvl_elm, "w:numFmt") or "decimal",
                lvl_text=_child_attr(lvl_elm, "w:lvlText") or "",
            )
        abstract_nums[abstract_num_id] = levels

    num_instances: Dict[int, _NumInstance] = {}
    for num in numbering_elm.num_lst:
        start_overrides = {
            lvl_override.ilvl: lvl_override.startOverride.val
            for lvl_override in num.lvlOverride_lst
            if lvl_override.startOverride is not None
        }
        num_instances[num.numId] = _NumInstance(
            num_id=num.numId,
            abstract_num_id=num.abstractNumId.val,
            start_overrides=start_overrides,
        )
    return abstract_nums, num_instances


def _child_val(lvl_elm, tag: str, default: int) -> int:
    child = lvl_elm.find(qn(tag))
    if child is None:
        return default
    val = child.get(qn("w:val"))
    return int(val) if val is not None else default


def _child_attr(lvl_elm, tag: str) -> Optional[str]:
    child = lvl_elm.find(qn(tag))
    if child is None:
        return None
    return child.get(qn("w:val"))


def _iter_paragraphs(document) -> Iterator[Paragraph]:
    """Generate every paragraph of the document body in explicit document order.

    Table rows and cells are walked in place, so paragraphs inside (possibly
    nested) tables appear exactly where the table occurs in the body flow.
    """

    def walk(container_elm) -> Iterator[Paragraph]:
        for child in container_elm.iterchildren():
            if child.tag == qn("w:p"):
                yield Paragraph(child, document)
            elif child.tag == qn("w:tbl"):
                for tr in child.findall(qn("w:tr")):
                    for tc in tr.findall(qn("w:tc")):
                        yield from walk(tc)

    return walk(document.element.body)


def _effective_num_reference(paragraph: Paragraph) -> Optional[Tuple[int, int]]:
    """Effective `(numId, ilvl)` for `paragraph`, or |None| if not numbered.

    Direct paragraph numbering properties take precedence; any component not
    specified directly is inherited through the paragraph-style `basedOn`
    chain. `ilvl` defaults to 0 when unspecified.
    """
    num_id: Optional[int] = None
    ilvl: Optional[int] = None

    pPr = paragraph._p.pPr
    if pPr is not None and pPr.numPr is not None:
        num_pr = pPr.numPr
        if num_pr.numId is not None:
            num_id = num_pr.numId.val
        if num_pr.ilvl is not None:
            ilvl = num_pr.ilvl.val

    if num_id is None or ilvl is None:
        for style_num_pr in _style_num_prs(paragraph):
            if num_id is None and style_num_pr.numId is not None:
                num_id = style_num_pr.numId.val
            if ilvl is None and style_num_pr.ilvl is not None:
                ilvl = style_num_pr.ilvl.val
            if num_id is not None and ilvl is not None:
                break

    if num_id is None:
        return None
    return num_id, ilvl if ilvl is not None else 0


def _style_num_prs(paragraph: Paragraph) -> Iterator:
    """Generate each `w:numPr` found walking the paragraph-style basedOn chain."""
    seen_style_ids = set()
    style = paragraph.style
    while style is not None and style.style_id not in seen_style_ids:
        seen_style_ids.add(style.style_id)
        pPr = style.element.pPr
        if pPr is not None and pPr.numPr is not None:
            yield pPr.numPr
        style = style.base_style


def _render_label(
    levels: Dict[int, _LevelDef], counter: _ListCounter, ilvl: int
) -> Tuple[Optional[str], Optional[str]]:
    """The `(label, unsupported_reason)` pair for the item at `ilvl`."""
    level_def = levels.get(ilvl)
    if level_def is None:
        return None, "no <w:lvl> defined for level %d" % ilvl
    if level_def.num_fmt == "bullet":
        return level_def.lvl_text, None

    def replace_placeholder(match) -> str:
        ref_ilvl = int(match.group(1)) - 1
        return _format_counter(levels, counter, ref_ilvl)

    try:
        return _LEVEL_PLACEHOLDER_RE.sub(replace_placeholder, level_def.lvl_text), None
    except _UnsupportedNumberFormat as exc:
        return None, str(exc)


def _format_counter(levels: Dict[int, _LevelDef], counter: _ListCounter, ilvl: int) -> str:
    """Formatted counter value for level `ilvl`, honoring that level's own numFmt."""
    level_def = levels.get(ilvl)
    if level_def is None:
        raise _UnsupportedNumberFormat("no <w:lvl> defined for level %d" % ilvl)
    num_fmt = level_def.num_fmt
    if num_fmt == "bullet":
        return level_def.lvl_text
    value = counter.value(ilvl)
    if num_fmt == "decimal":
        return str(value)
    if num_fmt == "upperLetter":
        return _to_letters(value).upper()
    if num_fmt == "lowerLetter":
        return _to_letters(value)
    if num_fmt == "upperRoman":
        return _to_roman(value)
    if num_fmt == "lowerRoman":
        return _to_roman(value).lower()
    raise _UnsupportedNumberFormat(
        "numbering format '%s' is not supported (supported formats: %s)"
        % (num_fmt, ", ".join(SUPPORTED_NUM_FMTS))
    )


def _to_letters(value: int) -> str:
    """Excel-style letter sequence: 1->a, 26->z, 27->aa."""
    if value < 1:
        return str(value)
    letters: List[str] = []
    while value > 0:
        value, remainder = divmod(value - 1, 26)
        letters.append(chr(ord("a") + remainder))
    return "".join(reversed(letters))


def _to_roman(value: int) -> str:
    if value < 1 or value > 3999:
        return str(value)
    numerals: List[str] = []
    remaining = value
    for arabic, numeral in _ROMAN_VALUES:
        while remaining >= arabic:
            numerals.append(numeral)
            remaining -= arabic
    return "".join(numerals)


class _UnsupportedNumberFormat(Exception):
    """Raised internally when a level uses a numFmt this resolver does not render."""
