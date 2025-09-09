import argparse
from importlib.metadata import version
import logging
import os

import pandas as pd
from weasyprint import HTML
import openpyxl
from openpyxl.workbook import Workbook
from rich.logging import RichHandler

from lighting_paperwork.channel_hookup import ChannelHookup
from lighting_paperwork.color_cut_list import ColorCutList
from lighting_paperwork.gobo_pull import GoboPullList
from lighting_paperwork.helpers import ShowData
from lighting_paperwork.instrument_schedule import InstrumentSchedule
from lighting_paperwork.vectorworks_xml import VWExport

logger = logging.getLogger(__name__)


def is_file(path: str) -> str:
    """Determines if a path is a file or not"""
    if not os.path.isfile(path):
        raise argparse.ArgumentTypeError("Path is not a valid file")

    return path


def main() -> None:
    """Main CLI function"""

    parser = argparse.ArgumentParser()
    # TODO add dtale support for editing
    # TODO add options for only outputting certain types of reports
    parser.add_argument("file", help="CSV or XML from Vectorworks", type=is_file)
    parser.add_argument("--show", help="Show name")
    parser.add_argument("--ld", help="Lighting designer initials")
    parser.add_argument("--rev", help="Revision string (ex. 'Rev. A')")
    parser.add_argument('--version', action='version', version=version('lighting-paperwork'))
    parser.add_argument("-log", "--loglevel", default="info", help="Change to the log level. One of debug, info (default), warning, error, critical")
    output_group = parser.add_argument_group(
        "Output style", "Select what type of output should be generated (default PDF)"
    )
    exclusive_output_group = output_group.add_mutually_exclusive_group()
    exclusive_output_group.add_argument(
        "--html",
        action="store_const",
        help="Exports reports into a HTML file (primarily for PDF layout debugging).",
        const="html",
        dest="output_type",
    )
    exclusive_output_group.add_argument(
        "--excel",
        action="store_const",
        help="Export reports into an Excel (xlsx) file.",
        const="excel",
        dest="output_type",
    )
    exclusive_output_group.add_argument(
        "--pdf",
        action="store_const",
        help="Export reports into a PDF.",
        const="pdf",
        dest="output_type",
    )

    parser.set_defaults(output_type="pdf")

    args = parser.parse_args()
    logging.basicConfig(level=args.loglevel.upper(), format="%(message)s", datefmt="[%X]", handlers=[RichHandler()])

    show_info = ShowData(args.show, args.ld, args.rev)

    if "csv" in args.file:
        # Converter is to supress the warning when I set addr=0 to empty string
        vw_export = pd.read_csv(
            args.file, sep="\t", header=0, converters={"Absolute Address": str}
        )

        # Clear VW's default "None" character
        vw_export = vw_export.replace("-", "")

    elif "xml" in args.file:
        vw_export = VWExport(args.file).export_df()

    if args.output_type in ("html", "pdf"):
        # Generate all paperwork HTML
        html = []
        html.append(ChannelHookup(vw_export, show_info).make_html())
        html.append(InstrumentSchedule(vw_export, show_info).make_html())
        html.append(ColorCutList(vw_export, show_info).make_html())
        html.append(GoboPullList(vw_export, show_info).make_html())

        # TODO add metadata to top of file
        # TODO add ToC for multi-report documents

        if args.output_type == "html":
            filename = show_info.generate_slug() + ".html"
            with open(filename, "w") as f:
                f.write("<!DOCTYPE html>")
                f.write("<html>")

                for h in html:
                    f.write(h)

                f.write("</html>")

                logger.info("HTML published to %s", filename)
        elif args.output_type == "pdf":
            # Generate paperwork PDF
            logger.info("Generating PDF...")
            filename = show_info.generate_slug() + ".pdf"
            documents = []
            for h in html:
                documents.append(HTML(string=h).render())

            # this method generates each report individually and collates them
            #   -> page numbers reset per report
            all_pages = [page for document in documents for page in document.pages]

            documents[0].copy(all_pages).write_pdf(filename)

            logger.info("PDF published to %s", filename)
    elif args.output_type == "excel":
        # Looks like our read/write buffer is the file, not anything internal to Python.
        # So each additional write/edit will dump to the file.

        excel_path = show_info.generate_slug() + ".xlsx"

        # Create blank wb
        wb = Workbook()
        wb.save(excel_path)

        ChannelHookup(vw_export, show_info).make_excel(excel_path)
        InstrumentSchedule(vw_export, show_info).make_excel(excel_path)
        ColorCutList(vw_export, show_info).make_excel(excel_path)
        GoboPullList(vw_export, show_info).make_excel(excel_path)

        # Get rid of the default first sheet
        wb = openpyxl.load_workbook(excel_path)
        del wb["Sheet"]
        wb.save(excel_path)

        logger.info("Excel workbook published to %s", excel_path)


if __name__ == "__main__":
    main()
