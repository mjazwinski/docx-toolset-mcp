"""Tests for the add_section tool."""

from __future__ import annotations

import pytest
from pathlib import Path

from docx import Document  # type: ignore[import-untyped]

from docx_toolset_mcp.tools.section_tool import add_section


# ---------------------------------------------------------------------------
# Fixtures
# ---------------------------------------------------------------------------

@pytest.fixture()
def sample_docx(tmp_path: Path) -> Path:
    """DOCX with two top-level sections, the first having a sub-heading.

    Structure:
        Introduction        (Heading 1)
          Sub Topic         (Heading 2)
          Sub-topic body.   (Normal)
        Conclusion          (Heading 1)
        Conclusion body.    (Normal)
    """
    doc = Document()
    doc.add_heading("Introduction", level=1)
    doc.add_heading("Sub Topic", level=2)
    doc.add_paragraph("Sub-topic body.")
    doc.add_heading("Conclusion", level=1)
    doc.add_paragraph("Conclusion body.")
    out = tmp_path / "sample.docx"
    doc.save(str(out))
    return out


# ---------------------------------------------------------------------------
# Insertion-position tests
# ---------------------------------------------------------------------------

def test_new_section_inserted_after_full_anchor_content(sample_docx: Path) -> None:
    """New section must appear after all of the anchor section's content.

    Expected result:
        Introduction   (H1)
          Sub Topic    (H2)
          Sub-topic body.
        New Section    (H1)   ← inserted here, NOT right after Introduction
        New body.
        Conclusion     (H1)
        Conclusion body.
    """
    add_section(
        new_section_name="New Section",
        add_after_section_name="Introduction",
        file_path=str(sample_docx),
        section_text="New body.",
    )

    doc = Document(str(sample_docx))
    texts = [p.text for p in doc.paragraphs]

    sub_idx        = texts.index("Sub Topic")
    sub_body_idx   = texts.index("Sub-topic body.")
    new_idx        = texts.index("New Section")
    conclusion_idx = texts.index("Conclusion")

    # New section must come after the sub-heading and its body
    assert new_idx > sub_idx,      "New section must be after Sub Topic"
    assert new_idx > sub_body_idx, "New section must be after sub-topic body"
    # Body text of new section immediately follows its heading
    assert texts[new_idx + 1] == "New body."
    # Conclusion must still follow the new section
    assert conclusion_idx == new_idx + 2


def test_new_section_not_inserted_as_child(sample_docx: Path) -> None:
    """New section must NOT appear between the anchor heading and its sub-heading."""
    add_section(
        new_section_name="New Section",
        add_after_section_name="Introduction",
        file_path=str(sample_docx),
        section_text="New body.",
    )

    doc = Document(str(sample_docx))
    texts = [p.text for p in doc.paragraphs]

    intro_idx   = texts.index("Introduction")
    sub_idx     = texts.index("Sub Topic")
    new_idx     = texts.index("New Section")

    assert new_idx != intro_idx + 1, (
        "New section must NOT be inserted immediately after the anchor heading"
    )
    assert sub_idx == intro_idx + 1, "Sub Topic must still follow Introduction"


# ---------------------------------------------------------------------------
# Style-inference tests
# ---------------------------------------------------------------------------

def test_style_inherited_from_anchor_when_omitted(sample_docx: Path) -> None:
    """When style_name is omitted the new section must use the anchor's style."""
    result = add_section(
        new_section_name="New Section",
        add_after_section_name="Introduction",
        file_path=str(sample_docx),
        section_text="body",
    )
    # Anchor is Heading 1 → effective style must be Heading 1
    assert "Heading 1" in result

    doc = Document(str(sample_docx))
    new_para = next(p for p in doc.paragraphs if p.text == "New Section")
    assert new_para.style.name == "Heading 1"


def test_explicit_style_overrides_anchor(sample_docx: Path) -> None:
    """Explicitly provided style_name must override the anchor's style."""
    result = add_section(
        new_section_name="New Section",
        add_after_section_name="Introduction",
        file_path=str(sample_docx),
        section_text="body",
        style_name="Heading 2",
    )
    assert "Heading 2" in result

    doc = Document(str(sample_docx))
    new_para = next(p for p in doc.paragraphs if p.text == "New Section")
    assert new_para.style.name == "Heading 2"


def test_add_section_title_style(sample_docx: Path) -> None:
    """'Title' is a valid Word style and must not raise."""
    result = add_section(
        new_section_name="Document Title",
        add_after_section_name="Introduction",
        file_path=str(sample_docx),
        section_text="Title body text.",
        style_name="Title",
    )
    assert "Document Title" in result


def test_add_section_quote_style(sample_docx: Path) -> None:
    """'Quote' is a valid Word style and must not raise."""
    result = add_section(
        new_section_name="A Quote Block",
        add_after_section_name="Introduction",
        file_path=str(sample_docx),
        section_text="Quote body text.",
        style_name="Quote",
    )
    assert "A Quote Block" in result


# ---------------------------------------------------------------------------
# Error-handling tests
# ---------------------------------------------------------------------------

def test_add_section_file_not_found() -> None:
    with pytest.raises(FileNotFoundError):
        add_section(
            new_section_name="X",
            add_after_section_name="Y",
            file_path="/nonexistent/file.docx",
            section_text="text",
        )


def test_add_section_wrong_extension(tmp_path: Path) -> None:
    bad = tmp_path / "doc.txt"
    bad.write_text("hello")
    with pytest.raises(ValueError, match=r"\.docx"):
        add_section(
            new_section_name="X",
            add_after_section_name="Y",
            file_path=str(bad),
            section_text="text",
        )


def test_add_section_anchor_not_found(sample_docx: Path) -> None:
    with pytest.raises(ValueError, match="not found"):
        add_section(
            new_section_name="New Section",
            add_after_section_name="Ghost Section",
            file_path=str(sample_docx),
            section_text="text",
        )
