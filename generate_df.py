import re
from typing import List, Optional, Tuple

import numpy as np
import pandas as pd
from natsort import natsort_keygen, natsorted

from helpers import parse_gel


def instrument_schedule(vw_export: pd.DataFrame) -> pd.DataFrame:
    filter_fields = [
        "Position",
        "Unit Number",
        "Purpose",
        "Instrument Type",
        "Wattage",
        "Color",
        "Gobo 1",
        "Channel",
        "Absolute Address",
    ]

    chan_fields = pd.DataFrame(vw_export[filter_fields], columns=filter_fields)
    # Need to have a position to show up in the instrument schedule
    chan_fields["Position"] = chan_fields["Position"].replace("", np.nan)
    chan_fields = chan_fields.dropna(subset=["Position"])

    chan_fields = combine_instrtype(chan_fields)
    chan_fields = format_address_slash(chan_fields)
    chan_fields = combine_gelgobo(chan_fields)
    chan_fields = chan_fields[
        [
            "Position",
            "Unit Number",
            "Purpose",
            "Instr Type & Load",
            "Color & Gobo",
            "Channel",
            "Addr",
        ]
    ]

    return chan_fields


def colorcut(vw_export: pd.DataFrame) -> pd.DataFrame:
    filter_fields = ["Color", "Frame Size"]

    chan_fields = pd.DataFrame(vw_export[filter_fields], columns=filter_fields)
    # Seperate colors and diffusion into dict list
    color_dict = []
    for index, row in chan_fields.iterrows():
        if row["Frame Size"] != "" and not pd.isnull(row["Frame Size"]):
            framesize = row["Frame Size"]
        else:
            framesize = "Unknown"
        if (
            row["Color"].strip() != ""
            and row["Color"] != "N/C"
            and not pd.isnull(row["Color"])
        ):
            for i in row["Color"].strip().split("+"):
                # Works for single gels too
                if len(i.split("x")) > 1:
                    # Repeat gel situation (i.e. L201x2)
                    for j in range(0, int(i.split("x")[1])):
                        gel = parse_gel(i.split("x")[0])
                        color_dict.append(
                            {
                                "Color": gel.name,
                                "Frame Size": framesize,
                                "Company": gel.company,
                                "Sort": gel.name_sort,
                            }
                        )
                else:
                    # Normal single gel
                    gel = parse_gel(i)
                    color_dict.append(
                        {
                            "Color": gel.name,
                            "Frame Size": framesize,
                            "Company": gel.company,
                            "Sort": gel.name_sort,
                        }
                    )

    colors = pd.DataFrame.from_dict(color_dict)
    colors = (
        colors.groupby(["Color", "Frame Size", "Sort"])["Color"]
        .count()
        .reset_index(name="Count")
    )
    # Hack for that silly Rosco company: 3xx values become xx.3
    colors = colors.sort_values(by=["Sort", "Frame Size"], key=natsort_keygen())
    colors = colors.drop(["Sort"], axis=1)

    return colors


def gobo_pull(vw_export: pd.DataFrame) -> pd.DataFrame:
    filter_fields = ["Gobo 1", "Gobo 2"]
    chan_fields = pd.DataFrame(vw_export[filter_fields], columns=filter_fields)
    gobo_list = []

    for index, row in chan_fields.iterrows():
        if row["Gobo 1"].strip() != "":
            gobo_list.append(row["Gobo 1"])
        if row["Gobo 2"].strip() != "":
            gobo_list.append(row["Gobo 2"])

    gobo_name, gobo_count = np.unique(gobo_list, return_counts=True)

    return pd.DataFrame(zip(gobo_name, gobo_count), columns=["Gobo Name", "Quantity"])
