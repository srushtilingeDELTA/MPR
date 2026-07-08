"""Update PowerPoint text while preserving template run formatting."""

from __future__ import annotations

from pptx.text.text import _Paragraph, _Run


def set_paragraph_text_preserve(paragraph: _Paragraph, text: str) -> None:
    """Replace paragraph text without resetting font, color, or size."""
    if paragraph.runs:
        paragraph.runs[0].text = text
        for run in paragraph.runs[1:]:
            run.text = ""
        return
    paragraph.text = text


def set_text_frame_preserve(text_frame, text: str) -> None:
    """Replace all text in a text frame, keeping the first run's formatting."""
    if not text_frame.paragraphs:
        text_frame.text = text
        return
    set_paragraph_text_preserve(text_frame.paragraphs[0], text)
    for paragraph in text_frame.paragraphs[1:]:
        for run in paragraph.runs:
            run.text = ""


def set_cell_text_preserve(cell, text: str) -> None:
    """Update a table cell value without changing template styling."""
    if cell.text_frame.paragraphs:
        set_text_frame_preserve(cell.text_frame, text)
    else:
        cell.text = text


def clear_text_frame_content(text_frame, *, keep_first_paragraph: bool = True) -> None:
    """Clear visible text but keep paragraph/run structure for styling."""
    if not text_frame.paragraphs:
        return
    start = 0 if keep_first_paragraph else 1
    for idx, paragraph in enumerate(text_frame.paragraphs):
        if idx < start:
            set_paragraph_text_preserve(paragraph, "")
        else:
            for run in paragraph.runs:
                run.text = ""


def clear_table_data_rows(table, *, header_rows: int = 1) -> None:
    """Blank table body cells below the header rows."""
    for row_idx in range(header_rows, len(table.rows)):
        for col_idx in range(len(table.columns)):
            set_cell_text_preserve(table.cell(row_idx, col_idx), "")