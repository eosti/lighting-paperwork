import pandas as pd
import numpy as np
import re
from natsort import natsort_keygen, natsorted
from typing import List, Optional, Tuple

from helpers import parse_gel


def combine_instrtype(df: pd.DataFrame) -> pd.DataFrame:
    """
    Combines the Instrument Type and Power fields into one
    """
    instload = []
    for index, row in df.iterrows():
        # Consistently format power
        if row["Wattage"] != "":
            # we want it to be [number]W
            power = re.sub(r"[^\d\.]", "", row["Wattage"])
            powerstr = power + "W"
        else:
            power = None
            powerstr = None

        # Make sure power shows up once, after the instrument type
        if powerstr is None:
            tmp = row["Instrument Type"]
        else:
            # Remove from instrument type (if existing)
            instrtype = re.sub(rf"\s*{power}\w+\s*", "", row["Instrument Type"])
            tmp = instrtype + " " + powerstr

        # If accessory, add that here
        # if row["Accessory Inventory"] != "":
        #    tmp += ", " + row["Accessory Inventory"]

        instload.append(tmp)

    # Clean up by replacing old cols with new one
    # TODO: Get accessories in here
    # df.drop(["Instrument Type", "Wattage", "Accessory Inventory"], axis=1, inplace=True)
    new_df = df.drop(["Instrument Type", "Wattage"], axis=1)
    new_df["Instr Type & Load"] = instload
    return new_df


# Only operates on Gobo 1
def combine_gelgobo(df: pd.DataFrame) -> pd.DataFrame:
    gelgobo = []
    for index, row in df.iterrows():
        # If no gel replace with N/C
        if row["Color"] == "":
            tmp = "N/C"
        else:
            tmp = row["Color"]

        # Append gobo if exists
        if row["Gobo 1"] != "":
            tmp += ", T: " + row["Gobo 1"]

        gelgobo.append(tmp)

    # Clean up by replacing old cols with new one
    new_df = df.drop(["Color", "Gobo 1"], axis=1)
    new_df["Color & Gobo"] = gelgobo
    return new_df


def format_address_slash(df: pd.DataFrame) -> pd.DataFrame:
    for row in df.itertuples():
        absaddr = int(df.at[row.Index, "Absolute Address"])
        if absaddr == 0:
            # If no address set, replace it with a blank
            df.at[row.Index, "Absolute Address"] = ""
        else:
            universe = int((absaddr - 1) / 512) + 1

            if universe == 1:
                address = absaddr
                df.at[row.Index, "Absolute Address"] = f"{address}"
            else:
                address = ((absaddr - 1) % 512) + 1
                df.at[row.Index, "Absolute Address"] = f"{universe}/{address}"

    slashed_df = df.rename(columns={"Absolute Address": "Addr"})
    return slashed_df


def channel_hookup(vw_export: pd.DataFrame) -> pd.DataFrame:
    # Format data
    filter_fields = [
        "Channel",
        "Absolute Address",
        "Position",
        "Unit Number",
        "Purpose",
        "Instrument Type",
        "Wattage",
        "Color",
        "Gobo 1",
    ]
    chan_fields = pd.DataFrame(vw_export[filter_fields], columns=filter_fields)
    # Need to have a channel to show up in the channel hookup
    chan_fields["Channel"] = chan_fields["Channel"].replace("", np.nan)
    chan_fields = chan_fields.dropna(subset=["Channel"])

    chan_fields = combine_instrtype(chan_fields)
    chan_fields = format_address_slash(chan_fields)
    chan_fields = combine_gelgobo(chan_fields)
    chan_fields = chan_fields.rename(columns={"Channel": "Chan", "Unit Number": "U#"})
    chan_fields = chan_fields.sort_values(
        by=["Chan", "Addr", "Position", "U#"], key=natsort_keygen()
    )

    chan_fields = chan_fields[["Chan", "Addr", "Position", "U#", "Purpose", "Instr Type & Load", "Color & Gobo"]]
    return chan_fields


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
    chan_fields = chan_fields[["Position", "Unit Number", "Purpose", "Instr Type & Load", "Color & Gobo", "Channel", "Addr"]]

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
