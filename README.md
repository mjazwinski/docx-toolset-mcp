
# docx-toolset-mcp



MCP server for reading and modifying Microsoft Word (.docx) documents.

Built with the [MCP Python SDK](https://py.sdk.modelcontextprotocol.io/) (v2.x) and [python-docx](https://python-docx.readthedocs.io/).



---



## Tools



### `add_section`

Insert a new section into a document as a sibling of an existing section (placed after all of the anchor section's content).



| Parameter | Type | Required | Description |

|---|---|---|---|

| `file_path` | string | yes | Path to the `.docx` file |

| `new_section_name` | string | yes | Heading text for the new section |

| `add_after_section_name` | string | yes | Exact heading text of the anchor section |

| `section_text` | string | yes | Body text placed below the new heading |

| `style_name` | DocxStyle | no | Word style for the heading; defaults to the anchor section's own style |



---



### `list_sections`

Extract all sections from a document and return an **LSA summary** of each.

Language is detected automatically via `langdetect`. Content before the first matching heading is ignored.



| Parameter | Type | Required | Description |

|---|---|---|---|

| `file_path` | string | yes | Path to the `.docx` file |

| `header_style` | DocxStyle | no | Style that marks section starts (default: `Heading 2`) |



Returns `[{ "name": string, "data": string }]` — one entry per section, in document order.



---



### `get_section`

Return the **full verbatim text** of a single named section (no summarisation).

Tables are included with columns tab-separated and rows newline-separated.



| Parameter | Type | Required | Description |

|---|---|---|---|

| `file_path` | string | yes | Path to the `.docx` file |

| `section_name` | string | yes | Exact (case-sensitive) heading text of the target section |

| `header_style` | DocxStyle | no | Style that marks section starts (default: `Heading 2`) |



Returns `string`. Raises `ValueError` if the section is not found.



---



## Accepted style names (`DocxStyle`)



`Normal` | `Heading 1` - `Heading 9` | `Title` | `Subtitle` | `Quote` | `Intense Quote` | `List Bullet` | `List Number` | `Emphasis` | `Strong`



---



## Setup



```bash

python -m venv .venv

.venv\Scripts\activate        # Windows

# .venv/bin/activate            # Linux / macOS

pip install -e .

```



## Running



```bash

docx-toolset-mcp               # stdio transport

```



## Cursor MCP configuration



**.cursor/mcp.json** (project-level) or **%APPDATA%/Cursor/mcp.json** (global):



```json

{

  "mcpServers": {

    "docx-toolset": {

      "command": "<absolute-path>/.venv/Scripts/docx-toolset-mcp.exe",

      "args": []

    }

  }

}

```

