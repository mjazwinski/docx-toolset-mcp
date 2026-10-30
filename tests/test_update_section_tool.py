"""Tests for the update_section tool."""



from __future__ import annotations



import pytest

from pathlib import Path



from docx import Document  # type: ignore[import-untyped]

from docx.oxml.ns import qn  # type: ignore[import-untyped]



from docx_toolset_mcp.tools.update_section_tool import update_section





# ---------------------------------------------------------------------------

# Fixtures

# ---------------------------------------------------------------------------



@pytest.fixture()

def basic_docx(tmp_path: Path) -> Path:

    """Document with two Heading 2 sections, each with body text."""

    doc = Document()

    doc.add_heading("Alpha", level=2)

    doc.add_paragraph("Original Alpha text.")

    doc.add_paragraph("Second Alpha paragraph.")

    doc.add_heading("Beta", level=2)

    doc.add_paragraph("Original Beta text.")

    out = tmp_path / "basic.docx"

    doc.save(str(out))

    return out





@pytest.fixture()

def table_docx(tmp_path: Path) -> Path:

    """Section that contains text, a table, and more text."""

    doc = Document()

    doc.add_heading("Data", level=2)

    doc.add_paragraph("Before table.")

    tbl = doc.add_table(rows=2, cols=2)

    tbl.cell(0, 0).text = "A"

    tbl.cell(0, 1).text = "B"

    tbl.cell(1, 0).text = "C"

    tbl.cell(1, 1).text = "D"

    doc.add_paragraph("After table.")

    out = tmp_path / "table.docx"

    doc.save(str(out))

    return out





# ---------------------------------------------------------------------------

# Core behaviour

# ---------------------------------------------------------------------------



def test_new_body_replaces_old_text(basic_docx: Path) -> None:

    update_section(

        file_path=str(basic_docx),

        section_name="Alpha",

        section_body="Replacement text.",

    )

    doc = Document(str(basic_docx))

    texts = [p.text for p in doc.paragraphs]

    assert "Replacement text." in texts

    assert "Original Alpha text." not in texts

    assert "Second Alpha paragraph." not in texts





def test_new_body_inserted_right_after_heading(basic_docx: Path) -> None:

    update_section(

        file_path=str(basic_docx),

        section_name="Alpha",

        section_body="New text.",

    )

    doc = Document(str(basic_docx))

    texts = [p.text for p in doc.paragraphs]

    alpha_idx = texts.index("Alpha")

    new_idx = texts.index("New text.")

    assert new_idx == alpha_idx + 1





def test_other_sections_untouched(basic_docx: Path) -> None:

    update_section(

        file_path=str(basic_docx),

        section_name="Alpha",

        section_body="New Alpha.",

    )

    doc = Document(str(basic_docx))

    texts = [p.text for p in doc.paragraphs]

    assert "Original Beta text." in texts





def test_table_preserved(table_docx: Path) -> None:

    update_section(

        file_path=str(table_docx),

        section_name="Data",

        section_body="New intro.",

    )

    doc = Document(str(table_docx))

    # Table must still exist

    assert len(doc.tables) == 1

    assert doc.tables[0].cell(0, 0).text == "A"





def test_text_paragraphs_around_table_removed(table_docx: Path) -> None:

    update_section(

        file_path=str(table_docx),

        section_name="Data",

        section_body="New intro.",

    )

    doc = Document(str(table_docx))

    texts = [p.text for p in doc.paragraphs]

    assert "Before table." not in texts

    assert "After table." not in texts





def test_new_body_is_first_after_heading(table_docx: Path) -> None:

    """New body paragraph must immediately follow the heading, before the table."""

    update_section(

        file_path=str(table_docx),

        section_name="Data",

        section_body="New intro.",

    )

    doc = Document(str(table_docx))

    texts = [p.text for p in doc.paragraphs]

    data_idx = texts.index("Data")

    new_idx = texts.index("New intro.")

    assert new_idx == data_idx + 1





def test_returns_success_string(basic_docx: Path) -> None:

    result = update_section(

        file_path=str(basic_docx),

        section_name="Alpha",

        section_body="New.",

    )

    assert isinstance(result, str)

    assert "Alpha" in result





def test_custom_header_style(tmp_path: Path) -> None:

    doc = Document()

    doc.add_heading("MainSec", level=1)

    doc.add_paragraph("Old H1 text.")

    path = tmp_path / "h1.docx"

    doc.save(str(path))



    update_section(

        file_path=str(path),

        section_name="MainSec",

        section_body="New H1 text.",

        header_style="Heading 1",

    )

    doc2 = Document(str(path))

    texts = [p.text for p in doc2.paragraphs]

    assert "New H1 text." in texts

    assert "Old H1 text." not in texts





# ---------------------------------------------------------------------------

# Error handling

# ---------------------------------------------------------------------------



def test_section_not_found(basic_docx: Path) -> None:

    with pytest.raises(ValueError, match="not found"):

        update_section(

            file_path=str(basic_docx),

            section_name="Ghost",

            section_body="text",

        )





def test_file_not_found() -> None:

    with pytest.raises(FileNotFoundError):

        update_section(

            file_path="/no/such/file.docx",

            section_name="X",

            section_body="y",

        )





def test_wrong_extension(tmp_path: Path) -> None:

    bad = tmp_path / "f.txt"

    bad.write_text("x")

    with pytest.raises(ValueError, match=r"\.docx"):

        update_section(

            file_path=str(bad),

            section_name="X",

            section_body="y",

        )

