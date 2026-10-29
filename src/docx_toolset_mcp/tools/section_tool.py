"""add_section tool: insert a named section into a DOCX document."""

from __future__ import annotations

import copy
from pathlib import Path
from typing import Literal

from docx import Document  # type: ignore[import-untyped]

# ---------------------------------------------------------------------------
# Public type alias — exact Word style names accepted by python-docx.
# Pylance enforces these at call sites; any other string is a type error.
# ---------------------------------------------------------------------------
DocxStyle = Literal[
    # Body
    "Normal",
    # Headings
    "Heading 1",
    "Heading 2",
    "Heading 3",
    "Heading 4",
    "Heading 5",
    "Heading 6",
    "Heading 7",
    "Heading 8",
    "Heading 9",
    # Title / Subtitle
    "Title",
    "Subtitle",
    # Quotes
    "Quote",
    "Intense Quote",
    # Lists
    "List Bullet",
    "List Number",
    # Inline emphasis (character styles)
    "Emphasis",
    "Strong",
]


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _find_paragraph(doc: object, text: str) -> object | None:
    """Return the first paragraph whose stripped text equals *text*, or None."""
    for para in doc.paragraphs:  # type: ignore[attr-defined]
        if para.text.strip() == text.strip():
            return para
    return None


def _heading_level(style_name: str) -> int | None:
    """Return the numeric level for a heading style (1-9), or None if not a heading."""
    if style_name.startswith("Heading "):
        try:
            return int(style_name.split()[-1])
        except ValueError:
            pass
    return None


def _find_section_end(doc: object, anchor: object) -> object:
    """Return the last paragraph that belongs to *anchor*'s section.

    Scans forward from *anchor* through all sub-headings and body paragraphs
    and stops just before the next heading whose level is <= the anchor's level
    (i.e. a sibling or parent section).  If the anchor is not a heading style,
    returns *anchor* itself so insertion happens immediately after it.
    """
    anchor_level = _heading_level(anchor.style.name)  # type: ignore[attr-defined]

    # Non-heading anchor: insert right after it.
    if anchor_level is None:
        return anchor

    paragraphs: list[object] = doc.paragraphs  # type: ignore[attr-defined]
    anchor_idx = next(
        i for i, p in enumerate(paragraphs)
        if p._element is anchor._element  # type: ignore[attr-defined]
    )

    last_in_section = anchor
    for para in paragraphs[anchor_idx + 1:]:
        level = _heading_level(para.style.name)  # type: ignore[attr-defined]
        if level is not None and level <= anchor_level:
            # Reached a sibling or parent heading — stop before it.
            break
        last_in_section = para

    return last_in_section


def _clone_para(doc: object, text: str, style: str) -> object:
    """Create a scratch paragraph, deep-copy its XML element, then remove it."""
    tmp = doc.add_paragraph(text, style=style)  # type: ignore[attr-defined]
    element = copy.deepcopy(tmp._element)        # type: ignore[attr-defined]
    tmp._element.getparent().remove(tmp._element)  # type: ignore[attr-defined]
    return element


# ---------------------------------------------------------------------------
# Public tool function
# ---------------------------------------------------------------------------

def add_section(
    *,
    new_section_name: str,
    add_after_section_name: str,
    file_path: str,
    section_text: str,
    style_name: DocxStyle | None = None,
) -> str:
    """Insert a new section at the end of an existing section in a DOCX file.

    The new heading is placed after *all* content that belongs to the anchor
    section (body paragraphs and any sub-headings), so it appears as a sibling
    of the anchor rather than a child.

    Parameters
    ----------
    new_section_name:
        Text of the new heading paragraph.
    add_after_section_name:
        Exact text of the existing heading after whose *section* (including all
        its sub-headings and body text) the new section will be appended.
    file_path:
        Absolute or relative path to the ``.docx`` file.
    section_text:
        Body text placed immediately below the new heading.
    style_name:
        Exact Word style name to apply to the new heading paragraph.
        When omitted (``None``) the anchor paragraph's own style is reused,
        so the new section is created at the same level automatically.
        Must be one of the ``DocxStyle`` values when supplied explicitly.
        Examples: ``"Heading 2"``, ``"Title"``, ``"Quote"``, ``"List Bullet"``.

    Returns
    -------
    str
        Human-readable success message including the resolved style name.

    Raises
    ------
    FileNotFoundError
        If *file_path* does not point to an existing file.
    ValueError
        If the file extension is not ``.docx`` or the anchor section is absent.
    """
    path = Path(file_path)

    if not path.exists():
        raise FileNotFoundError(f"File not found: {file_path!r}")
    if path.suffix.lower() != ".docx":
        raise ValueError(f"Expected a .docx file, got: {path.suffix!r}")

    doc = Document(str(path))

    anchor = _find_paragraph(doc, add_after_section_name)
    if anchor is None:
        raise ValueError(
            f"Anchor section {add_after_section_name!r} not found in {file_path!r}."
        )

    # Resolve the effective style — fall back to the anchor's own style.
    effective_style: str = style_name if style_name is not None else anchor.style.name  # type: ignore[attr-defined]

    # Walk to the last paragraph of the anchor's section so the new section
    # is inserted as a sibling, not as the first child.
    insertion_point = _find_section_end(doc, anchor)
    insertion_el = insertion_point._element  # type: ignore[attr-defined]

    # Clone both elements before touching the tree to avoid index drift.
    heading_el = _clone_para(doc, new_section_name, effective_style)
    body_el    = _clone_para(doc, section_text, "Normal")

    # addnext() inserts immediately after the reference element.
    # Push body first so that heading ends up on top: … → heading → body.
    insertion_el.addnext(body_el)
    insertion_el.addnext(heading_el)

    doc.save(str(path))

    return (
        f"Section {new_section_name!r} (style: {effective_style!r}) appended after "
        f"the end of section {add_after_section_name!r} in {file_path!r}."
    )
