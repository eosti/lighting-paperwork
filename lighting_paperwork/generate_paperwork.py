"""CLI tool for generating lighting paperwork."""

import argparse
import logging
from importlib.metadata import version
from pathlib import Path

import pandas as pd
from rich.logging import RichHandler

from lighting_paperwork.channel_hookup import ChannelHookup
from lighting_paperwork.color_cut_list import ColorCutList
from lighting_paperwork.gobo_pull import GoboPullList
from lighting_paperwork.helpers import ShowData
from lighting_paperwork.instrument_schedule import InstrumentSchedule
from lighting_paperwork.paperwork_exporters import ExportExcel, ExportHTML, ExportPDF
from lighting_paperwork.vectorworks_xml import VWExport

logger = logging.getLogger(__name__)


def is_file(path: str) -> str:
    """Determine if a path is a file or not."""
    if not Path(path).is_file():
        raise argparse.ArgumentTypeError("Path is not a valid file")

    return path


def main(argv: list[str] | None = None) -> None:
    """Run main CLI function."""
    parser = argparse.ArgumentParser()
    # TODO(eosti): add dtale support for editing
    # https://github.com/eosti/lighting-paperwork/issues/12
    parser.add_argument("file", help="CSV or XML from Vectorworks", type=is_file)
    parser.add_argument("--show", help="Show name")
    parser.add_argument("--ld", help="Lighting designer initials")
    parser.add_argument("--rev", help="Revision string (ex. 'Rev. A')")
    parser.add_argument("--version", action="version", version=version("lighting-paperwork"))
    parser.add_argument(
        "-log",
        "--loglevel",
        default="info",
        help="Change to the log level. One of debug, info (default), warning, error, critical",
    )
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

    args = parser.parse_args(argv)
    logging.basicConfig(
        level=args.loglevel.upper(),
        format="%(message)s",
        datefmt="[%X]",
        handlers=[RichHandler()],
    )

    show_info = ShowData(args.show, args.ld, args.rev)

    if "csv" in args.file:
        # Converter is to supress the warning when I set addr=0 to empty string
        vw_export = pd.read_csv(args.file, sep="\t", header=0, converters={"Absolute Address": str})

        # Clear VW's default "None" character
        vw_export = vw_export.replace("-", "")

    elif "xml" in args.file:
        vw_export = VWExport(args.file).export_df()

    else:
        raise RuntimeError("Only supports csv and xml")

    paperwork_list = [
        ChannelHookup(vw_export, show_info),
        InstrumentSchedule(vw_export, show_info),
        ColorCutList(vw_export, show_info),
        GoboPullList(vw_export, show_info),
    ]

    if args.output_type == "html":
        output_path = ExportHTML(show_info.generate_slug(), paperwork_list).make()
        logger.info("HTML published to %s", output_path)
    elif args.output_type == "pdf":
        output_path = ExportPDF(show_info.generate_slug(), paperwork_list).make()
        logger.info("PDF published to %s", output_path)
    elif args.output_type == "excel":
        output_path = ExportExcel(show_info.generate_slug(), paperwork_list).make()
        logger.info("Excel workbook published to %s", output_path)
    else:
        raise ValueError(f"Unknown output type {args.output_type}")


if __name__ == "__main__":
    main()
