"""list_sections tool: extract and summarise sections from a DOCX document."""



from __future__ import annotations



from pathlib import Path

from typing import TypedDict



# NLTK data download (required by sumy's sentence tokenizer)

import nltk

nltk.download("punkt",     quiet=True)

nltk.download("punkt_tab", quiet=True)



from docx import Document  # type: ignore[import-untyped]

from docx.oxml.ns import qn  # type: ignore[import-untyped]

from docx.table import Table  # type: ignore[import-untyped]

from docx.text.paragraph import Paragraph  # type: ignore[import-untyped]

from langdetect import LangDetectException, detect  # type: ignore[import-untyped]

from sumy.nlp.stemmers import Stemmer  # type: ignore[import-untyped]

from sumy.nlp.tokenizers import Tokenizer  # type: ignore[import-untyped]

from sumy.parsers.plaintext import PlaintextParser  # type: ignore[import-untyped]

from sumy.summarizers.lsa import LsaSummarizer  # type: ignore[import-untyped]

from sumy.utils import get_stop_words  # type: ignore[import-untyped]



from docx_toolset_mcp.tools.section_tool import DocxStyle



# ---------------------------------------------------------------------------

# Return-type contract (also exported so server.py can re-use it)

# ---------------------------------------------------------------------------



class SectionEntry(TypedDict):

    """A single named section extracted from the document."""



    name: str  # heading text that starts the section

    data: str  # LSA summary of all body content that follows the heading





# ---------------------------------------------------------------------------

# XML tag constants

# ---------------------------------------------------------------------------



_WP_TAG: str = qn("w:p")    # paragraph element

_TBL_TAG: str = qn("w:tbl") # table element



# ---------------------------------------------------------------------------

# Language support

# ---------------------------------------------------------------------------



# ISO 639-1 code -> sumy language name

_LANG_MAP: dict[str, str] = {

    "af": "english",   # Afrikaans -- fall back to english

    "ar": "arabic",

    "cs": "czech",

    "da": "danish",

    "de": "german",

    "el": "greek",

    "en": "english",

    "es": "spanish",

    "fi": "finnish",

    "fr": "french",

    "he": "hebrew",

    "hu": "hungarian",

    "it": "italian",

    "ja": "japanese",

    "ko": "korean",

    "nb": "norwegian",

    "nl": "dutch",

    "pl": "polish",

    "pt": "portuguese",

    "ro": "romanian",

    "ru": "russian",

    "sk": "slovak",

    "sv": "swedish",

    "tr": "turkish",

    "uk": "ukrainian",

    "zh-cn": "chinese",

    "zh-tw": "chinese",

}

_DEFAULT_LANGUAGE = "english"

_SUMMARY_SENTENCES = 3





# ---------------------------------------------------------------------------

# Summarization helpers

# ---------------------------------------------------------------------------



def _detect_language(text: str) -> str:

    """Return the sumy language name for *text*, defaulting to 'english'."""

    try:

        code = detect(text)

        return _LANG_MAP.get(code, _DEFAULT_LANGUAGE)

    except (LangDetectException, Exception):

        return _DEFAULT_LANGUAGE





def _summarize(text: str, sentences_count: int = _SUMMARY_SENTENCES) -> str:

    """Return an LSA summary of *text*.



    Falls back to the original text when:

    * the input is empty / whitespace-only, or

    * the summarizer produces no output (text too short for the algorithm).

    """

    stripped = text.strip()

    if not stripped:

        return stripped



    language = _detect_language(stripped)

    try:

        parser = PlaintextParser.from_string(stripped, Tokenizer(language))

        stemmer = Stemmer(language)

        summarizer = LsaSummarizer(stemmer)

        summarizer.stop_words = get_stop_words(language)

        sentences = summarizer(parser.document, sentences_count)

        summary = " ".join(str(s) for s in sentences).strip()

        # Sumy returns nothing for very short texts; keep original in that case.

        return summary if summary else stripped

    except Exception:

        return stripped





# ---------------------------------------------------------------------------

# Table extraction helper

# ---------------------------------------------------------------------------



def _table_to_text(table: Table) -> str:

    """Render a Word table to plain text.



    * Columns are separated by ``\\t``.

    * Rows are separated by ``\\n``.

    * Merged cells are de-duplicated (python-docx repeats them in row.cells).

    * Multi-paragraph cells have their paragraphs joined with ``\\n``.

    """

    rendered_rows: list[str] = []



    for row in table.rows:  # type: ignore[attr-defined]

        rendered_cells: list[str] = []

        seen_ids: set[int] = set()



        for cell in row.cells:  # type: ignore[attr-defined]

            cell_key = id(cell._tc)  # type: ignore[attr-defined]

            if cell_key in seen_ids:

                continue

            seen_ids.add(cell_key)



            lines = [

                p.text

                for p in cell.paragraphs  # type: ignore[attr-defined]

                if p.text.strip()

            ]

            rendered_cells.append("\n".join(lines))



        rendered_rows.append("\t".join(rendered_cells))



    return "\n".join(rendered_rows)





# ---------------------------------------------------------------------------

# Public tool function

# ---------------------------------------------------------------------------



def list_sections(

    *,

    file_path: str,

    header_style: DocxStyle = "Heading 2",

) -> list[SectionEntry]:

    """Extract all sections from a DOCX file and return an LSA summary of each.



    The function walks the document body in reading order, treating every

    paragraph whose style matches *header_style* as the start of a new

    section.  All content between two such headings (paragraphs and tables)

    is collected, then summarised using the LSA algorithm from the sumy

    library.  The document language is detected automatically via langdetect.



    Table text is rendered with columns tab-separated and rows

    newline-separated before being fed into the summariser.



    Parameters

    ----------

    file_path:

        Absolute or relative path to the ``.docx`` file.

    header_style:

        Word style name used to identify section headings.

        Defaults to ``"Heading 2"``.



    Returns

    -------

    list[SectionEntry]

        Ordered list of ``{"name": ..., "data": ...}`` dicts, one per

        section found.  ``data`` holds the LSA summary of the section

        content.  Empty list if no heading of *header_style* is present.



    Raises

    ------

    FileNotFoundError

        If *file_path* does not point to an existing file.

    ValueError

        If the file extension is not ``.docx``.

    """

    path = Path(file_path)



    if not path.exists():

        raise FileNotFoundError(f"File not found: {file_path!r}")

    if path.suffix.lower() != ".docx":

        raise ValueError(f"Expected a .docx file, got: {path.suffix!r}")



    doc = Document(str(path))



    sections: list[SectionEntry] = []

    current_name: str | None = None

    current_lines: list[str] = []



    def _flush() -> None:

        """Summarise the current buffer and append to results."""

        if current_name is not None:

            full_text = "\n".join(current_lines).strip()

            sections.append(

                SectionEntry(

                    name=current_name,

                    data=_summarize(full_text),

                )

            )



    for child in doc.element.body:

        tag: str = child.tag



        if tag == _WP_TAG:

            para = Paragraph(child, doc)

            style_name: str = para.style.name  # type: ignore[attr-defined]

            text: str = para.text.strip()



            if style_name == header_style:

                _flush()

                current_name = text

                current_lines = []

            elif current_name is not None and text:

                current_lines.append(text)



        elif tag == _TBL_TAG:

            if current_name is not None:

                table = Table(child, doc)

                table_text = _table_to_text(table)

                if table_text.strip():

                    current_lines.append(table_text)



    _flush()  # flush the final open section



    return sections

