"""update_section tool: replace text in a named DOCX section."""



from __future__ import annotations



import copy

from pathlib import Path



from docx import Document  # type: ignore[import-untyped]

from docx.oxml.ns import qn  # type: ignore[import-untyped]

from docx.table import Table  # type: ignore[import-untyped]

from docx.text.paragraph import Paragraph  # type: ignore[import-untyped]



from docx_toolset_mcp.tools.list_sections_tool import _TBL_TAG, _WP_TAG

from docx_toolset_mcp.tools.section_tool import DocxStyle, _clone_para, _heading_level



# Tags that indicate an embedded image inside a paragraph run

_DRAWING_TAG: str = qn("w:drawing")

_PICT_TAG: str = qn("w:pict")





# ---------------------------------------------------------------------------

# Internal helpers

# ---------------------------------------------------------------------------



def _para_has_image(para_el: object) -> bool:

    """Return True if the paragraph XML contains an inline image."""

    return bool(

        para_el.findall(f".//{_DRAWING_TAG}")  # type: ignore[attr-defined]

        or para_el.findall(f".//{_PICT_TAG}")  # type: ignore[attr-defined]

    )





# ---------------------------------------------------------------------------

# Public tool function

# ---------------------------------------------------------------------------



def update_section(

    *,

    file_path: str,

    section_name: str,

    section_body: str,

    header_style: DocxStyle = "Heading 2",

) -> str:

    """Replace the text content of a named section in a DOCX document.



    Finds the section whose heading matches *section_name*, removes all

    plain-text paragraphs that belong to it, then inserts *section_body*

    as a new paragraph immediately after the heading.



    Tables and image paragraphs inside the section are preserved in place;

    only text-only paragraphs are removed.



    Parameters

    ----------

    file_path:

        Absolute or relative path to the ``.docx`` file.

    section_name:

        Exact (case-sensitive, whitespace-stripped) heading text of the

        target section.

    section_body:

        Replacement text inserted right after the heading paragraph.

    header_style:

        Word style name used to identify section headings.

        Defaults to ``"Heading 2"``.



    Returns

    -------

    str

        Human-readable success message.



    Raises

    ------

    FileNotFoundError

        If *file_path* does not point to an existing file.

    ValueError

        If the file extension is not ``.docx``, or the section is not found.

    """

    path = Path(file_path)



    if not path.exists():

        raise FileNotFoundError(f"File not found: {file_path!r}")

    if path.suffix.lower() != ".docx":

        raise ValueError(f"Expected a .docx file, got: {path.suffix!r}")



    doc = Document(str(path))

    body = doc.element.body



    # ── 1. Locate the heading element ────────────────────────────────────────

    heading_el: object | None = None

    anchor_level: int | None = None



    for child in body:

        if child.tag != _WP_TAG:

            continue

        para = Paragraph(child, doc)

        if para.style.name == header_style and para.text.strip() == section_name.strip():  # type: ignore[attr-defined]

            heading_el = child

            anchor_level = _heading_level(header_style)

            break



    if heading_el is None:

        raise ValueError(

            f"Section {section_name!r} (style: {header_style!r}) "

            f"not found in {file_path!r}."

        )



    # ── 2. Collect elements that belong to this section ──────────────────────

    # Walk siblings after the heading until we hit another heading at the same

    # or higher level (lower number = higher level).

    to_remove: list[object] = []

    sibling = heading_el.getnext()  # type: ignore[attr-defined]



    while sibling is not None:

        tag: str = sibling.tag



        if tag == _WP_TAG:

            para = Paragraph(sibling, doc)

            sib_level = _heading_level(para.style.name)  # type: ignore[attr-defined]

            if sib_level is not None and (

                anchor_level is None or sib_level <= anchor_level

            ):

                # Reached sibling / parent heading -- stop.

                break

            # Text-only paragraph: mark for removal.

            if not _para_has_image(sibling):

                to_remove.append(sibling)



        elif tag != _TBL_TAG:

            # Skip non-paragraph, non-table elements (e.g. w:sectPr).

            pass

        # Tables: skip (preserve).



        sibling = sibling.getnext()  # type: ignore[attr-defined]



    # ── 3. Remove text paragraphs ────────────────────────────────────────────

    for el in to_remove:

        body.remove(el)



    # ── 4. Insert new body paragraph right after the heading ─────────────────

    new_body_el = _clone_para(doc, section_body, "Normal")

    heading_el.addnext(new_body_el)  # type: ignore[attr-defined]



    doc.save(str(path))



    return (

        f"Section {section_name!r} in {file_path!r} updated: "

        f"{len(to_remove)} text paragraph(s) replaced."

    )

