"""Tests for the list_sections tool."""



from __future__ import annotations



import pytest

from pathlib import Path



from docx import Document  # type: ignore[import-untyped]



from docx_toolset_mcp.tools.list_sections_tool import list_sections





# ---------------------------------------------------------------------------

# Fixtures

# ---------------------------------------------------------------------------



@pytest.fixture()

def two_section_docx(tmp_path: Path) -> Path:

    """Simple document: two Heading 2 sections each with one body paragraph."""

    doc = Document()

    doc.add_heading("Overview", level=2)

    doc.add_paragraph("Overview body text.")

    doc.add_heading("Details", level=2)

    doc.add_paragraph("Details body text.")

    out = tmp_path / "two_sections.docx"

    doc.save(str(out))

    return out





@pytest.fixture()

def table_section_docx(tmp_path: Path) -> Path:

    """Document with a Heading 2 section that contains a table."""

    doc = Document()

    doc.add_heading("Data", level=2)

    doc.add_paragraph("Intro text.")

    table = doc.add_table(rows=2, cols=3)

    table.cell(0, 0).text = "Alpha"

    table.cell(0, 1).text = "Beta"

    table.cell(0, 2).text = "Gamma"

    table.cell(1, 0).text = "One"

    table.cell(1, 1).text = "Two"

    table.cell(1, 2).text = "Three"

    out = tmp_path / "table_section.docx"

    doc.save(str(out))

    return out





@pytest.fixture()

def preamble_docx(tmp_path: Path) -> Path:

    """Document with content before the first matching heading (preamble)."""

    doc = Document()

    doc.add_heading("Title", level=1)

    doc.add_paragraph("Preamble paragraph.")

    doc.add_heading("Section A", level=2)

    doc.add_paragraph("Section A content.")

    out = tmp_path / "preamble.docx"

    doc.save(str(out))

    return out





@pytest.fixture()

def multi_para_section_docx(tmp_path: Path) -> Path:

    """Section with several body paragraphs (enough for LSA to work with)."""

    doc = Document()

    doc.add_heading("Chapter", level=2)

    # Use distinct meaningful sentences so LSA can rank them

    doc.add_paragraph("The quick brown fox jumps over the lazy dog.")

    doc.add_paragraph("Machine learning models require large amounts of training data.")

    doc.add_paragraph("Python is a versatile programming language used in many domains.")

    doc.add_paragraph("Neural networks are inspired by the structure of the human brain.")

    doc.add_paragraph("Data preprocessing is a critical step in any data science pipeline.")

    out = tmp_path / "multi_para.docx"

    doc.save(str(out))

    return out





# ---------------------------------------------------------------------------

# Basic extraction tests

# ---------------------------------------------------------------------------



def test_returns_correct_section_count(two_section_docx: Path) -> None:

    result = list_sections(file_path=str(two_section_docx))

    assert len(result) == 2





def test_section_names(two_section_docx: Path) -> None:

    result = list_sections(file_path=str(two_section_docx))

    assert result[0]["name"] == "Overview"

    assert result[1]["name"] == "Details"





def test_section_data_is_non_empty_string(two_section_docx: Path) -> None:

    """data must be a non-empty string for sections that have content."""

    result = list_sections(file_path=str(two_section_docx))

    assert isinstance(result[0]["data"], str) and result[0]["data"]

    assert isinstance(result[1]["data"], str) and result[1]["data"]





def test_summarised_data_shorter_than_full_text(multi_para_section_docx: Path) -> None:

    """LSA summary should be shorter than the concatenation of all paragraphs."""

    result = list_sections(file_path=str(multi_para_section_docx))

    assert len(result) == 1

    full_len = len(

        "The quick brown fox jumps over the lazy dog. "

        "Machine learning models require large amounts of training data. "

        "Python is a versatile programming language used in many domains. "

        "Neural networks are inspired by the structure of the human brain. "

        "Data preprocessing is a critical step in any data science pipeline."

    )

    summary = result[0]["data"]

    assert summary  # non-empty

    assert len(summary) <= full_len  # at most the same length (LSA won't expand)





def test_preamble_content_ignored(preamble_docx: Path) -> None:

    """Content before the first matching heading must not appear in output."""

    result = list_sections(file_path=str(preamble_docx))

    assert len(result) == 1

    assert result[0]["name"] == "Section A"

    assert "Preamble" not in result[0]["data"]

    assert not any(s["name"] == "Title" for s in result)





def test_empty_document_returns_empty_list(tmp_path: Path) -> None:

    doc = Document()

    path = tmp_path / "empty.docx"

    doc.save(str(path))

    result = list_sections(file_path=str(path))

    assert result == []





def test_no_matching_headings_returns_empty_list(tmp_path: Path) -> None:

    """If no paragraph matches header_style, the result should be empty."""

    doc = Document()

    doc.add_heading("Only H1", level=1)

    doc.add_paragraph("Some text.")

    path = tmp_path / "no_h2.docx"

    doc.save(str(path))

    result = list_sections(file_path=str(path))

    assert result == []





# ---------------------------------------------------------------------------

# Table extraction tests

# ---------------------------------------------------------------------------



def test_table_content_present_in_data(table_section_docx: Path) -> None:

    """All table cell values must appear somewhere in the summary data."""

    result = list_sections(file_path=str(table_section_docx))

    assert len(result) == 1

    data = result[0]["data"]

    for val in ("Alpha", "Beta", "Gamma", "One", "Two", "Three"):

        assert val in data, f"Expected {val!r} in data"





def test_table_columns_tab_separated(table_section_docx: Path) -> None:

    """Table rows in the data must use tab as column separator."""

    result = list_sections(file_path=str(table_section_docx))

    data = result[0]["data"]

    table_lines = [line for line in data.splitlines() if "\t" in line]

    assert len(table_lines) == 2, "Expected 2 table rows with tab separators"

    assert table_lines[0].split("\t") == ["Alpha", "Beta", "Gamma"]

    assert table_lines[1].split("\t") == ["One", "Two", "Three"]





def test_table_rows_newline_separated(table_section_docx: Path) -> None:

    """First table row must appear before second row in the data string."""

    result = list_sections(file_path=str(table_section_docx))

    data = result[0]["data"]

    assert data.index("Alpha") < data.index("One")





# ---------------------------------------------------------------------------

# Alternative header_style parameter

# ---------------------------------------------------------------------------



def test_custom_header_style(tmp_path: Path) -> None:

    doc = Document()

    doc.add_heading("Top Level", level=1)

    doc.add_paragraph("H1 body.")

    doc.add_heading("Sub", level=2)

    doc.add_paragraph("H2 body.")

    path = tmp_path / "custom_style.docx"

    doc.save(str(path))



    result = list_sections(file_path=str(path), header_style="Heading 1")

    assert len(result) == 1

    assert result[0]["name"] == "Top Level"

    assert result[0]["data"]  # non-empty summary





# ---------------------------------------------------------------------------

# Error-handling tests

# ---------------------------------------------------------------------------



def test_file_not_found() -> None:

    with pytest.raises(FileNotFoundError):

        list_sections(file_path="/nonexistent/doc.docx")





def test_wrong_extension(tmp_path: Path) -> None:

    bad = tmp_path / "doc.txt"

    bad.write_text("hello")

    with pytest.raises(ValueError, match=r"\.docx"):

        list_sections(file_path=str(bad))

