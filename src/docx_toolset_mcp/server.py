"""Entry point for the DOCX Toolset MCP server."""



from __future__ import annotations



from mcp.server.mcpserver import MCPServer



from docx_toolset_mcp.tools.get_section_tool import get_section as _get_section
from docx_toolset_mcp.tools.update_section_tool import update_section as _update_section
from docx_toolset_mcp.tools.list_sections_tool import (

    SectionEntry,

    list_sections as _list_sections,

)

from docx_toolset_mcp.tools.section_tool import DocxStyle, add_section as _add_section



# ---------------------------------------------------------------------------

# Server instance

# ---------------------------------------------------------------------------

mcp: MCPServer = MCPServer(

    name="docx-toolset",

    instructions=(

        "A toolset for reading and modifying Microsoft Word (.docx) documents. "

        "Use these tools to programmatically add, edit, or query document sections."

    ),

)





# ---------------------------------------------------------------------------

# Tools

# ---------------------------------------------------------------------------



@mcp.tool()

def add_section(

    new_section_name: str,

    add_after_section_name: str,

    file_path: str,

    section_text: str,

    style_name: DocxStyle | None = None,

) -> str:

    """Insert a new section at the end of an existing section in a DOCX document.



    The new heading is placed after *all* content belonging to the anchor section

    (body paragraphs and sub-headings), making it a sibling -- not a child -- of

    the anchor.



    Parameters

    ----------

    new_section_name:

        Text of the new heading to create.

    add_after_section_name:

        Exact text of the existing heading after whose full section content the

        new section will be appended.

    file_path:

        Absolute or relative path to the target ``.docx`` file.

    section_text:

        Body text placed immediately below the new heading.

    style_name:

        Exact Word style name for the new heading (optional).

        When omitted the anchor paragraph's own style is reused automatically.

        Accepted values: ``"Normal"``, ``"Heading 1"``--``"Heading 9"``,

        ``"Title"``, ``"Subtitle"``, ``"Quote"``, ``"Intense Quote"``,

        ``"List Bullet"``, ``"List Number"``, ``"Emphasis"``, ``"Strong"``.

    """

    return _add_section(

        new_section_name=new_section_name,

        add_after_section_name=add_after_section_name,

        file_path=file_path,

        section_text=section_text,

        style_name=style_name,

    )





@mcp.tool()

def list_sections(

    file_path: str,

    header_style: DocxStyle = "Heading 2",

) -> list[SectionEntry]:

    """Extract all sections from a DOCX document, grouped by a header style.



    Walks the document body in reading order.  Every paragraph whose style

    matches *header_style* is treated as the start of a new section; all

    subsequent paragraphs and tables (until the next matching heading) become

    that section's content.



    Table text is included with columns tab-separated and rows

    newline-separated, preserving reading order.



    Parameters

    ----------

    file_path:

        Absolute or relative path to the target ``.docx`` file.

    header_style:

        Word style name that marks section boundaries.

        Defaults to ``"Heading 2"``.



    Returns

    -------

    list[SectionEntry]

        JSON array of ``{"name": string, "data": string}`` objects, one per

        section found in document order.

    """

    return _list_sections(

        file_path=file_path,

        header_style=header_style,

    )





# ---------------------------------------------------------------------------

# CLI entry point

# ---------------------------------------------------------------------------



def main() -> None:

    """Start the MCP server (stdio transport)."""

    mcp.run(transport="stdio")





if __name__ == "__main__":

    main()

