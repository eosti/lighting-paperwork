import argparse
import logging

import pandas as pd

from channel_hookup import ChannelHookup
from helpers import ShowData
from vectorworks_xml import VWExport


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("raw", help="Raw instrument from Vectorworks")
    parser.add_argument("--show", help="Show name")
    parser.add_argument("--ld", help="Lighting designer initials")
    parser.add_argument("--rev", help="Revision string (ex. 'Rev. A')")

    args = parser.parse_args()
    logging.basicConfig(level=logging.DEBUG)

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

    chanhook = ChannelHookup(vw_export, show_data=show_info)

    with open("chans.html", "w") as f:
        f.write(chanhook.make())


if __name__ == "__main__":
    main()
