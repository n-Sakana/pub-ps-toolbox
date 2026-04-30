from __future__ import annotations

import dataclasses
import json
from pathlib import Path
from typing import Iterable

import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill

from .models import ErrorRecord, LinkRecord, PageRecord, PdfRecord

EXCEL_TEXT_LIMIT = 32767
SAFE_TEXT_LIMIT = 32000


def excel_safe(value: object) -> object:
    if isinstance(value, str) and len(value) > SAFE_TEXT_LIMIT:
        return value[:SAFE_TEXT_LIMIT] + "\n... [truncated for Excel cell limit]"
    return value


def records_to_frame(records: Iterable[object]) -> pd.DataFrame:
    rows = []
    for record in records:
        if dataclasses.is_dataclass(record):
            row = dataclasses.asdict(record)
        else:
            row = dict(record)  # type: ignore[arg-type]
        rows.append({key: excel_safe(value) for key, value in row.items()})
    return pd.DataFrame(rows)


def summary_frame(*, pages: list[PageRecord], pdfs: list[PdfRecord], links: list[LinkRecord], errors: list[ErrorRecord]) -> pd.DataFrame:
    unique_pdf_urls = {pdf.pdf_url for pdf in pdfs}
    downloaded = [pdf for pdf in pdfs if pdf.downloaded]
    downloaded_unique = {pdf.pdf_url for pdf in downloaded}
    summary = [
        {"metric": "pages", "value": len(pages)},
        {"metric": "page_links", "value": len([link for link in links if link.category == "internal_page"])},
        {"metric": "pdf_references", "value": len(pdfs)},
        {"metric": "unique_pdf_urls", "value": len(unique_pdf_urls)},
        {"metric": "downloaded_pdf_references", "value": len(downloaded)},
        {"metric": "downloaded_unique_pdf_urls", "value": len(downloaded_unique)},
        {"metric": "errors", "value": len(errors)},
    ]
    return pd.DataFrame(summary)


def write_workbook(output: Path, *, pages: list[PageRecord], pdfs: list[PdfRecord], links: list[LinkRecord], errors: list[ErrorRecord]) -> None:
    output.parent.mkdir(parents=True, exist_ok=True)
    sheets = {
        "Pages": records_to_frame(pages),
        "PDFs": records_to_frame(pdfs),
        "Links": records_to_frame(links),
        "Errors": records_to_frame(errors),
        "Summary": summary_frame(pages=pages, pdfs=pdfs, links=links, errors=errors),
    }
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet_name, frame in sheets.items():
            if frame.empty:
                frame = pd.DataFrame([{}])
            frame.to_excel(writer, sheet_name=sheet_name, index=False)
            worksheet = writer.sheets[sheet_name]
            worksheet.freeze_panes = "A2"
            header_fill = PatternFill("solid", fgColor="1F4E78")
            for cell in worksheet[1]:
                cell.font = Font(color="FFFFFF", bold=True)
                cell.fill = header_fill
                cell.alignment = Alignment(vertical="top", wrap_text=True)
            for column_cells in worksheet.columns:
                header = str(column_cells[0].value or "")
                width = min(80, max(12, len(header) + 2))
                for cell in column_cells[1:80]:
                    value = cell.value
                    if value is not None:
                        width = min(80, max(width, min(80, len(str(value)) + 2)))
                    cell.alignment = Alignment(vertical="top", wrap_text=True)
                worksheet.column_dimensions[column_cells[0].column_letter].width = width
