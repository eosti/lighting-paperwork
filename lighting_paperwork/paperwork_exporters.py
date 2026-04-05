"""Paperwork exporters to various filetypes."""

from abc import ABC, abstractmethod
from pathlib import Path

import openpyxl
import weasyprint
from openpyxl.workbook import Workbook

from lighting_paperwork.paperwork import PaperworkGenerator


class PaperworkExporter(ABC):
    """Virtual class for paperwork exports."""

    file_extension = ""

    def __init__(self, file_slug: str, paperwork: list[PaperworkGenerator]) -> None:
        """Initialize filename and paperwork list."""
        self.filename = Path(file_slug + "." + self.file_extension)
        self.paperwork = paperwork
        # TODO(eosti): verify file doesn't already exist
        # https://github.com/eosti/lighting-paperwork/issues/14

    @abstractmethod
    def make(self) -> Path:
        """Generate and save paperwork to self.filename."""


class ExportHTML(PaperworkExporter):
    """Class for HTML paperwork exports."""

    file_extension = ".html"

    def generate_html(self) -> list[str]:
        """Generate HTML of each paperwork.

        Includes HTML DOCTYPE and <html> tags so a valid webpage would be
            simply all entries concatenated together.
        """
        html = ["<!DOCTYPE html>\n<html>\n"]
        html.extend(p.make_html() for p in self.paperwork)
        html.append(["</html>"])

        return html

    def make(self) -> Path:
        """Make an HTML file with the provided paperwork."""
        html = self.generate_html()
        with self.filename.open("w") as f:
            for h in html:
                # Get rid of border-collapse for HTML (why?)
                clean_html = h.replace("border-collapse: collapse", "border-collapse: initial")
                f.write(clean_html)

        return self.filename


class ExportPDF(ExportHTML):
    """Class for PDF paperwork exports.

    This uses `weasyprint` to generate PDFs, which is a HTML to PDF package.
    There's a lot of print-specific CSS attributes but they're pretty poorly supported
        by most webbrowsers; `weasyprint` is pretty good at them though.
    """

    file_extension = ".pdf"

    def make(self) -> Path:
        """Make a PDF with the provided paperwork."""
        html = self.generate_html()
        documents = []
        documents.extend(weasyprint.HTML(string=h).render() for h in html)

        # This method generates each report individually and collates them
        # Means that page numbers reset per report
        all_pages = [page for document in documents for page in document.pages]

        documents[0].copy(all_pages).write_pdf(self.filename)
        return self.filename


class ExportExcel(PaperworkExporter):
    """Class for Excel paperwork exports."""

    file_extension = ".xlsx"

    def make(self) -> None:
        """Make an Excel workbook with provided paperwork.

        Note that the read/write buffer is the file, so each additional write/edit
            will re-dump to the file. This is a consequence of how openpyxl is structured.
        """
        wb = Workbook()
        wb.save(self.filename)

        for p in self.paperwork:
            p.make_excel(self.filename)

        # Get rid of default first sheet
        wb = openpyxl.load_workbook(self.filename)
        del wb["Sheet"]
        wb.save(self.filename)

        return self.filename
