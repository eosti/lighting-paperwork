"""CLI tool for generating lighting paperwork."""

import logging
import sys
from pathlib import Path

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
    logger.debug(settings.model_dump_json())

    logging.basicConfig(
        level=settings.log_level.upper(),
        format="%(message)s",
        datefmt="[%X]",
        handlers=[RichHandler()],
    )

    if settings.input_file is not None and settings.input_file.suffix.lower() == ".yaml":
        logger.info("Using settings file %s", settings.input_file)
        settings.model_config["yaml_file"] = str(settings.input_file)
        settings.__init__()

    if settings.data_file is None:
        logger.critical("Must provide an input data file.")
        sys.exit(1)

    if settings.data_file.suffix.lower() == ".csv":
        vw_export = pd.read_csv(
            settings.data_file, sep="\t", header=0, dtype=str, keep_default_na=False
        )

        # Clear VW's default "None" character
        vw_export = vw_export.replace("-", "")

    elif settings.data_file.suffix.lower() == ".xml":
        vw_export = VWExport(settings.data_file).export_df()

    else:
        raise RuntimeError("Only supports csv and xml")

    paperwork_list = [
        ChannelHookup(vw_export, settings.paperwork),
        InstrumentSchedule(vw_export, settings.paperwork),
        ColorCutList(vw_export, settings.paperwork),
        GoboPullList(vw_export, settings.paperwork),
    ]

    output_dir = Path.cwd() if settings.output_dir is None else Path(settings.output_dir)

    if settings.output_type == "html":
        output_path = ExportHTML(
            output_dir, settings.paperwork.show_info.generate_slug(), paperwork_list
        ).make()
        logger.info("HTML published to %s", output_path)
    elif settings.output_type == "pdf":
        output_path = ExportPDF(
            output_dir, settings.paperwork.show_info.generate_slug(), paperwork_list
        ).make()
        logger.info("PDF published to %s", output_path)
    elif settings.output_type == "excel":
        output_path = ExportExcel(
            output_dir, settings.paperwork.show_info.generate_slug(), paperwork_list
        ).make()
        logger.info("Excel workbook published to %s", output_path)
    else:
        raise AssertionError

    sys.exit(0)


if __name__ == "__main__":
    main()
