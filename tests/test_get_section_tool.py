"""Tests for the get_section tool."""



from __future__ import annotations



import pytest

from pathlib import Path



from docx import Document  # type: ignore[import-untyped]



from docx_toolset_mcp.tools.get_section_tool import get_section





# ---------------------------------------------------------------------------

# Fixtures

# ---------------------------------------------------------------------------



@pytest.fixture()

def rich_docx(tmp_path: Path) -> Path:

    """Document with three Heading 2 sections; the second has sub-content."""

    doc = Document()



    doc.add_heading("Alpha", level=2)

    doc.add_paragraph("Alpha first paragraph.")

    doc.add_paragraph("Alpha second paragraph.")



    doc.add_heading("Beta", level=2)

    doc.add_paragraph("Beta body text.")

    table = doc.add_table(rows=2, cols=2)

    table.cell(0, 0).text = "Col1"

    table.cell(0, 1).text = "Col2"

    table.cell(1, 0).text = "Val1"

    table.cell(1, 1).text = "Val2"



    doc.add_heading("Gamma", level=2)

    doc.add_paragraph("Gamma body text.")



    out = tmp_path / "rich.docx"

    doc.save(str(out))

    return out





# ---------------------------------------------------------------------------

# Content correctness

# ---------------------------------------------------------------------------



def test_returns_full_paragraph_text(rich_docx: Path) -> None:

    result = get_section(file_path=str(rich_docx), section_name="Alpha")

    assert "Alpha first paragraph." in result

    assert "Alpha second paragraph." in result





def test_returns_all_paragraphs_in_order(rich_docx: Path) -> None:

    result = get_section(file_path=str(rich_docx), section_name="Alpha")

    assert result.index("Alpha first paragraph.") < result.index("Alpha second paragraph.")





def test_does_not_include_other_sections(rich_docx: Path) -> None:

    result = get_section(file_path=str(rich_docx), section_name="Alpha")

    assert "Beta body text." not in result

    assert "Gamma body text." not in result





def test_includes_table_content(rich_docx: Path) -> None:

    result = get_section(file_path=str(rich_docx), section_name="Beta")

    for val in ("Col1", "Col2", "Val1", "Val2"):

        assert val in result





def test_table_columns_tab_separated(rich_docx: Path) -> None:

    result = get_section(file_path=str(rich_docx), section_name="Beta")

    table_lines = [line for line in result.splitlines() if "\t" in line]

    assert len(table_lines) == 2

    assert table_lines[0].split("\t") == ["Col1", "Col2"]

    assert table_lines[1].split("\t") == ["Val1", "Val2"]





def test_last_section_collected(rich_docx: Path) -> None:

    """Section with no following sibling must still be fully collected."""

    result = get_section(file_path=str(rich_docx), section_name="Gamma")

    assert "Gamma body text." in result





def test_returns_plain_string(rich_docx: Path) -> None:

    result = get_section(file_path=str(rich_docx), section_name="Alpha")

    assert isinstance(result, str)





# ---------------------------------------------------------------------------

# Custom header_style

# ---------------------------------------------------------------------------



def test_custom_header_style(tmp_path: Path) -> None:

    doc = Document()

    doc.add_heading("Main Section", level=1)

    doc.add_paragraph("Main section body.")

    path = tmp_path / "h1.docx"

    doc.save(str(path))



    result = get_section(

        file_path=str(path),

        section_name="Main Section",

        header_style="Heading 1",

    )

    assert "Main section body." in result





# ---------------------------------------------------------------------------

# Error handling

# ---------------------------------------------------------------------------



def test_section_not_found(rich_docx: Path) -> None:

    with pytest.raises(ValueError, match="not found"):

        get_section(file_path=str(rich_docx), section_name="Nonexistent Section")





def test_file_not_found() -> None:

    with pytest.raises(FileNotFoundError):

        get_section(file_path="/no/such/file.docx", section_name="X")





def test_wrong_extension(tmp_path: Path) -> None:

    bad = tmp_path / "doc.txt"

    bad.write_text("hello")

    with pytest.raises(ValueError, match=r"\.docx"):

        get_section(file_path=str(bad), section_name="X")

