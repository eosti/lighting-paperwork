"""
Generate lighting export by going to File > Export > Export Lighting Device Data...
Select all fields and then export. Output should be a txt file.
"""

# TODO: fix same color merge column, remove empty channel
# TODO: format numbers as numbers in excel
# TODO: accessories???
# TODO: embed UUID into a hidden cell somehow
# TODO: Add a field in VW as a "rep" indicator, then use that on color hookup to determine existing vs new
# TODO: add title page with links
# TODO: make setting widths relative to 100% page
import argparse
import os
import re
from typing import List, Optional, Tuple
import logging

import pandas as pd
from pandas.io.formats.style import Styler
from natsort import natsort_keygen, natsorted
from vectorworks_xml import VWExport
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.workbook import Workbook
from dataclasses import dataclass

import generate_df
import excel_format
from helpers import ShowData, FontStyle


def is_file(path: str) -> str:
    if not os.path.isfile(path):
        raise argparse.ArgumentTypeError("Path is not a valid file")

    return path


def split_by_position(df: pd.DataFrame) -> Tuple[pd.DataFrame, List[str]]:
    # Step one: sort position names
    unique_vals = df["Position"].unique()

    # Sort by Cat (descending), Elec (ascending), other
    elec_list = natsorted([x for x in unique_vals if re.match(r"^Elec\s\d", x)])
    lx_list = natsorted([x for x in unique_vals if re.match(r"^LX\d", x)])
    cat_list = natsorted(
        [x for x in unique_vals if re.match(r"^Cat\s\d", x)], reverse=True
    )
    foh_list = natsorted([x for x in unique_vals if re.match(r"^FOH\s\d", x)])

    # This will need to be tweaked per venue
    box_booms_ds = natsorted(
        [x for x in unique_vals if re.match(r"^DS[RL]\sBox Boom", x)]
    )
    box_booms_us = natsorted(
        [x for x in unique_vals if re.match(r"^US[RL]\sBox Boom", x)]
    )

    booms = natsorted([x for x in unique_vals if re.match(r"^S[RL]\sBoom\s\d", x)])
    ladders = natsorted([x for x in unique_vals if re.match(r"^S[RL]\sLadder", x)])

    # The order of this is what'll make the order in the schedule!
    special_positions = (
        cat_list
        + foh_list
        + elec_list
        + lx_list
        + booms
        + box_booms_ds
        + box_booms_us
        + ladders
    )

    other_list = natsorted([x for x in unique_vals if x not in (special_positions)])

    position_names = special_positions + other_list

    # Step two: create unique dataframes by position name
    sorted_df = []
    for i in position_names:
        pos_df = df.loc[df["Position"] == i].copy()
        pos_df = pos_df.drop(["Position"], axis=1)
        pos_df = pos_df.rename(columns={"Channel": "Chan", "Unit Number": "U#"})
        pos_df = pos_df.sort_values(by=["U#"], key=natsort_keygen())
        sorted_df.append(pos_df)

    return sorted_df, position_names


def add_channel_hookup(
    wb: Workbook, vw_export: pd.DataFrame, show_info: ShowData
) -> None:
    chan_fields = generate_df.channel_hookup(vw_export)

    # Dump to worksheet
    ws = wb.create_sheet("Channel Hookup", -1)
    for r in dataframe_to_rows(chan_fields, index=False, header=True):
        ws.append(r)

    # Format
    excel_format.add_title(ws, "Channel Hookup", show_info)
    excel_format.page_setup(ws, 1)
    excel_format.set_col_widths(ws, [60, 40, 80, 30, 80, 190, 130])
    excel_format.channelhookup_cols(ws)
    excel_format.repeated_channel(ws)


def style_chan_hookup(
    hookup: pd.DataFrame, chan_style: FontStyle, body_style: FontStyle
):
    style_df = hookup.copy()
    # Set borders based on channel data
    prev_row = (None, None)
    for index, data in hookup.iterrows():
        style_df.loc[index, :] = ""
        if prev_row == (None, None):
            prev_row = (index, data)
            continue

        if data["Chan"] == "&nbsp;":
            style_df.loc[
                prev_row[0], :
            ] += "border-bottom: 0px solid black; page-break-after: avoid; "
        else:
            style_df.loc[prev_row[0], :] += "border-bottom: 1px solid black; "

        prev_row = (index, data)

    # Set font based on column
    for col_name, col in style_df.items():
        if col_name == "Chan":
            style_df[col_name] += f"{chan_style.to_css()}; vertical-align: middle; "
        else:
            style_df[col_name] += f"{body_style.to_css()}; vertical-align: middle; "

        if col_name in ["Chan", "U#", "Addr"]:
            style_df[col_name] += "text-align: center; "
        else:
            style_df[col_name] += "text-align: left; "

    return style_df


def verify_width(width: List[int]):
    total_count = 0
    for i in width:
        total_count += i

    if total_count > 100:
        print(f"Too long! Used {total_count}%!")


def style_chan_header(index: pd.Series, header_style: FontStyle, col_width: List[int]):
    verify_width(col_width)
    style = [
        f"{header_style.to_css()}; border-top: 1px solid black; border-bottom: 1px solid black; "
        for _ in index
    ]

    for idx, val in enumerate(index):
        if val in ["Chan", "U#", "Addr"]:
            style[idx] += "text-align: center; "
        else:
            style[idx] += "text-align: left; "

    for idx, _ in enumerate(index):
        style[idx] += f"width: {col_width[idx]}%; "

    return style


def add_channel_hookup_html(vw_export: pd.DataFrame, show_info: ShowData) -> None:
    chan_fields = generate_df.channel_hookup(vw_export)

    header_style = FontStyle("Calibri", "bold", 12)
    body_style = FontStyle("Calibri", "normal", 11)
    chan_style = FontStyle("Calibri", "bold", 18)

    styled = Styler.from_custom_template(".", "header_footer.tpl")(chan_fields)
    styled = styled.apply(
        style_chan_hookup, axis=None, chan_style=chan_style, body_style=body_style
    )
    styled = styled.hide()
    styled = styled.apply_index(
        style_chan_header,
        header_style=header_style,
        col_width=[10, 5, 13, 5, 13, 32, 22],
        axis=1,
    )
    styled = styled.set_table_styles(
        [{"selector": "", "props": [("border-spacing", "0"), ("width", "100%")]}],
        overwrite=False,
    )

    head_style = FontStyle("Calibri", "normal", 12)
    title_style = FontStyle("Calibri", "bold", 22)

    with open("chans.html", "w") as f:
        f.write(
            styled.to_html(
                style_footer_right=head_style.to_css(),
                style_footer_left=head_style.to_css(),
                style_header_right=head_style.to_css(),
                style_header_left=head_style.to_css(),
                style_header_center=f"{title_style.to_css()}",
                content_footer_left="Channel Hookup",
                content_header_right="Show Name<br>LD Names",
                content_header_left="DATE HERE<br>Rev. A",
                content_header_center="Channel Hookup",
                pagenum_bottom_right=True,
            )
        )


def add_instrument_schedule(
    wb: Workbook, vw_export: pd.DataFrame, show_info: ShowData
) -> None:
    positions, position_names = split_by_position(
        generate_df.instrument_schedule(vw_export)
    )
    ws = wb.create_sheet("Instrument Schedule", -1)
    for idx, pos in enumerate(positions):
        ws.append([position_names[idx]])
        start_row = ws.max_row
        for r in dataframe_to_rows(pos, index=False, header=True):
            ws.append(r)
        ws.append([])

        excel_format.instrschedule_cols(ws, start_row)

    # Format
    excel_format.add_title(ws, "Instrument Schedule", show_info)
    excel_format.page_setup(ws)
    excel_format.set_col_widths(ws, [30, 110, 210, 160, 40, 40])
    excel_format.instr_schedule_pagebreaks(ws)


def add_colorcuts(wb: Workbook, vw_export: pd.DataFrame, show_info: ShowData) -> None:
    colors = generate_df.colorcut(vw_export)

    # Dump to worksheet
    ws = wb.create_sheet("Color Cut List", -1)
    for r in dataframe_to_rows(colors, index=False, header=True):
        ws.append(r)

    # Format
    excel_format.add_title(ws, "Color Cut List", show_info)
    excel_format.page_setup(ws, 1)
    excel_format.set_col_widths(ws, [80, 100, 50])
    excel_format.repeated_gels(ws)
    excel_format.colorcut_cols(ws)


def add_gobos(wb: Workbook, vw_export: pd.DataFrame, show_info: ShowData) -> None:
    gobo_list = generate_df.gobo_pull(vw_export)

    # Dump to Worksheet
    ws = wb.create_sheet("Gobo Pull List", -1)
    ws.append(["Gobo", "Qty"])
    for r in dataframe_to_rows(gobo_list, index=False, header=False):
        ws.append(r)

    # Format
    excel_format.add_title(ws, "Gobo Pull List", show_info)
    excel_format.page_setup(ws, 1)
    excel_format.set_col_widths(ws, [200, 50])
    excel_format.gobos(ws)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("raw", help="Raw instrument from Vectorworks", type=is_file)
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

    wb = Workbook()

    add_channel_hookup_html(vw_export, show_info)
    return
    add_channel_hookup(wb, vw_export, show_info)
    add_instrument_schedule(wb, vw_export, show_info)
    add_colorcuts(wb, vw_export, show_info)
    add_gobos(wb, vw_export, show_info)

    del wb["Sheet"]
    output_filename = (
        f"{show_info.show_name.replace(' ', '')}_Paperwork_"
        + re.sub(r"\W+", "", show_info.revision)
        + ".xlsx"
    )
    wb.save(output_filename)


if __name__ == "__main__":
    main()
