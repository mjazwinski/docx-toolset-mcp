"""get_section tool: return the full text of a named section in a DOCX."""



from __future__ import annotations



from pathlib import Path



from docx import Document  # type: ignore[import-untyped]

from docx.table import Table  # type: ignore[import-untyped]

from docx.text.paragraph import Paragraph  # type: ignore[import-untyped]



from docx_toolset_mcp.tools.list_sections_tool import (

    _TBL_TAG,

    _WP_TAG,

    _table_to_text,

)

from docx_toolset_mcp.tools.section_tool import DocxStyle





# ---------------------------------------------------------------------------

# Public tool function

# ---------------------------------------------------------------------------



def get_section(

    *,

    file_path: str,

    section_name: str,

    header_style: DocxStyle = "Heading 2",

) -> str:

    """Return the full text of a named section in a DOCX document.



    Walks the document body and collects every paragraph and table that

    belongs to the section whose heading text exactly matches *section_name*.

    Unlike ``list_sections``, no summarisation is applied -- the complete

    content is returned verbatim.



    Table text is rendered with columns tab-separated and rows

    newline-separated, preserving reading order.



    Parameters

    ----------

    file_path:

        Absolute or relative path to the ``.docx`` file.

    section_name:

        Exact text of the heading paragraph that starts the target section.

        Comparison is case-sensitive and whitespace-stripped.

    header_style:

        Word style name used to identify section headings.

        Defaults to ``"Heading 2"``.



    Returns

    -------

    str

        Full concatenated text of all content (paragraphs and tables) that

        belongs to the named section, with blocks separated by newlines.



    Raises

    ------

    FileNotFoundError

        If *file_path* does not point to an existing file.

    ValueError

        If the file extension is not ``.docx``, or no section with

        *section_name* is found in the document.

    """

    path = Path(file_path)



    if not path.exists():

        raise FileNotFoundError(f"File not found: {file_path!r}")

    if path.suffix.lower() != ".docx":

        raise ValueError(f"Expected a .docx file, got: {path.suffix!r}")



    doc = Document(str(path))



    inside_target = False

    lines: list[str] = []



    for child in doc.element.body:

        tag: str = child.tag



        if tag == _WP_TAG:

            para = Paragraph(child, doc)

            style_name: str = para.style.name  # type: ignore[attr-defined]

            text: str = para.text.strip()



            if style_name == header_style:

                if inside_target:

                    # Reached the next section of the same level -- stop.

                    break

                if text == section_name.strip():

                    # Found the target heading -- start collecting.

                    inside_target = True

            elif inside_target and text:

                lines.append(text)



        elif tag == _TBL_TAG:

            if inside_target:

                table = Table(child, doc)

                table_text = _table_to_text(table)

                if table_text.strip():

                    lines.append(table_text)



    if not inside_target:

        raise ValueError(

            f"Section {section_name!r} (style: {header_style!r}) "

            f"not found in {file_path!r}."

        )



    return "\n".join(lines).strip()

