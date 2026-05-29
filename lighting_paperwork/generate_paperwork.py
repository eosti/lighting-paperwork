"""CLI tool for generating lighting paperwork."""

import logging

import pandas as pd
from rich.logging import RichHandler

from lighting_paperwork.channel_hookup import ChannelHookup
from lighting_paperwork.color_cut_list import ColorCutList
from lighting_paperwork.gobo_pull import GoboPullList
from lighting_paperwork.instrument_schedule import InstrumentSchedule
from lighting_paperwork.paperwork_exporters import ExportExcel, ExportHTML, ExportPDF
from lighting_paperwork.paperwork_settings import CLISettings
from lighting_paperwork.vectorworks_xml import VWExport

logger = logging.getLogger(__name__)


def main() -> None:
    """Run main CLI function."""
    # TODO(eosti): add dtale support for editing
    # https://github.com/eosti/lighting-paperwork/issues/12

    settings = CLISettings()

    logging.basicConfig(
        level=settings.log_level.upper(),
        format="%(message)s",
        datefmt="[%X]",
        handlers=[RichHandler()],
    )

    # add default PDF logic

    if settings.input_file.suffix == ".csv":
        # Converter is to suppress the warning when I set addr=0 to empty string
        vw_export = pd.read_csv(
            settings.input_file, sep="\t", header=0, converters={"Absolute Address": str}
        )

        # Clear VW's default "None" character
        vw_export = vw_export.replace("-", "")

    elif settings.input_file.suffix == ".xml":
        vw_export = VWExport(settings.input_file).export_df()

    else:
        raise RuntimeError("Only supports csv and xml")

    paperwork_list = [
        ChannelHookup(vw_export, settings.paperwork),
        InstrumentSchedule(vw_export, settings.paperwork),
        ColorCutList(vw_export, settings.paperwork),
        GoboPullList(vw_export, settings.paperwork),
    ]

    if settings.output_type == "html":
        output_path = ExportHTML(
            settings.paperwork.show_info.generate_slug(), paperwork_list
        ).make()
        logger.info("HTML published to %s", output_path)
    elif settings.output_type == "pdf":
        output_path = ExportPDF(settings.paperwork.show_info.generate_slug(), paperwork_list).make()
        logger.info("PDF published to %s", output_path)
    elif settings.output_type == "excel":
        output_path = ExportExcel(
            settings.paperwork.show_info.generate_slug(), paperwork_list
        ).make()
        logger.info("Excel workbook published to %s", output_path)
    else:
        raise AssertionError


if __name__ == "__main__":
    main()
