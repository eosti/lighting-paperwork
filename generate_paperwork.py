import argparse
import logging
import os
import re

import pandas as pd
from weasyprint import HTML

from channel_hookup import ChannelHookup
from color_cut_list import ColorCutList
from gobo_pull import GoboPullList
from helpers import ShowData
from instrument_schedule import InstrumentSchedule
from vectorworks_xml import VWExport


def is_file(path: str) -> str:
    if not os.path.isfile(path):
        raise argparse.ArgumentTypeError("Path is not a valid file")

    return path


def main() -> None:
    logging.basicConfig(level=logging.DEBUG)

    parser = argparse.ArgumentParser()
    parser.add_argument("raw", help="Raw instrument from Vectorworks", type=is_file)
    parser.add_argument("--show", help="Show name")
    parser.add_argument("--ld", help="Lighting designer initials")
    parser.add_argument("--rev", help="Revision string (ex. 'Rev. A')")

    args = parser.parse_args()

    show_info = ShowData(args.show, args.ld, args.rev)

    if "csv" in args.raw:
        # Converter is to supress the warning when I set addr=0 to empty string
        vw_export = pd.read_csv(
            args.raw, sep="\t", header=0, converters={"Absolute Address": str}
        )

        # Clear VW's default "None" character
        vw_export = vw_export.replace("-", "")

    elif "xml" in args.raw:
        vw_export = VWExport(args.raw).export_df()

    # Generate all paperwork HTML
    html = []
    html.append(ChannelHookup(vw_export, show_info).make_html())
    html.append(InstrumentSchedule(vw_export, show_info).make_html())
    html.append(ColorCutList(vw_export, show_info).make_html())
    html.append(GoboPullList(vw_export, show_info).make_html())

    # Generate paperwork PDF
    documents = []
    for h in html:
        documents.append(HTML(string=h).render())

    all_pages = [page for document in documents for page in document.pages]
    output_filename = (
        f"{show_info.show_name.replace(' ', '')}_Paperwork_"
        + re.sub(r"\W+", "", show_info.revision)
        + ".pdf"
    )

    documents[0].copy(all_pages).write_pdf(output_filename)


if __name__ == "__main__":
    main()
