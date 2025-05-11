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
import argparse
import os
import re
from typing import List, Optional, Tuple

import numpy as np
import pandas as pd
from natsort import natsort_keygen, natsorted
from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.workbook import Workbook
from openpyxl.worksheet.page import PageMargins
from openpyxl.worksheet.pagebreak import Break
from openpyxl.worksheet.worksheet import Worksheet
from vectorworks_xml import VWExport

DIFFUSION_FIELD = "User Field 1"

X_PADDING = 0.7
Y_PADDING = 0.7
Y_PADDING_HEADER = 0.4
HEAD_FOOT_PAD = 0.2

PAGE_HEIGHT_INCHES = 11


def is_file(path: str) -> str:
    if not os.path.isfile(path):
        raise argparse.ArgumentTypeError("Path is not a valid file")

    return path


def page_setup(ws: Worksheet, rows_to_repeat: int = 0) -> None:
    # Set page size, margins, default view
    ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
    ws.page_setup.paperSize = ws.PAPERSIZE_LETTER
    ws.print_options.horizontalCentered = True
    ws.sheet_view.view = "pageLayout"

    ws.page_margins = PageMargins(
        bottom=Y_PADDING,
        top=Y_PADDING + Y_PADDING_HEADER,
        left=X_PADDING,
        right=X_PADDING,
        header=HEAD_FOOT_PAD,
        footer=HEAD_FOOT_PAD,
    )
    ws.sheet_view.showGridLines = False

    # Force the title + column names to repeat
    if rows_to_repeat > 0:
        ws.print_title_rows = f"1:{rows_to_repeat}"


def add_title(ws: Worksheet, name: str) -> None:
    # Header
    ws.oddHeader.left.text = f"&[Date]\n {REVISION_STRING}"
    ws.oddHeader.left.size = 12
    ws.oddHeader.right.text = f"{SHOW_NAME}\nLD: {LD_NAME}"
    ws.oddHeader.right.size = 12

    # Title
    ws.oddHeader.center.text = name
    ws.oddHeader.center.size = 22
    ws.oddHeader.center.font = "Calibri,Bold"

    # Footer
    ws.oddFooter.left.text = name
    ws.oddFooter.left.size = 12
    ws.oddFooter.right.text = "Page &[Page] of &[Pages]"
    ws.oddFooter.right.size = 12


def set_col_widths(ws: Worksheet, width: List[int]) -> None:
    # Validate width (assuming 96ppi)
    # TODO: This check doesn't work
    max_width = (8.5 - 2 * X_PADDING) * 96
    for i in width:
        max_width -= i
    if max_width < 0:
        raise ValueError("Widths are too wide for page")

    for i in range(1, ws.max_column + 1):
        ws.column_dimensions[get_column_letter(i)].width = int(width[i - 1] / 7)


def format_channelhookup_cols(ws: Worksheet) -> None:
    ws.row_dimensions[1].height = 20
    side = Side(border_style="thin", color="FF000000")
    for i, cell in enumerate(ws["1:1"]):
        if i == 0 or i == 1 or i == 3:
            # Center channel, address, and U#
            cell.alignment = Alignment(horizontal="center", vertical="center")
        else:
            cell.alignment = Alignment(horizontal="left", vertical="center")
        cell.font = Font(size=12, bold=True)
        cell.border = Border(bottom=side, top=side)

    # Data formatting
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for idx, cell in enumerate(row):
            if idx == 0:
                # Channel formatting
                cell.font = Font(size=18, bold=True)
                cell.number_format = "(0)"
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif idx == 1 or idx == 3:
                # Center U#, address
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                # Normal cell
                cell.font = Font(size=11)
                cell.alignment = Alignment(
                    horizontal="left", vertical="center", wrap_text=True
                )


def format_instrschedule_cols(ws: Worksheet, start_row: int) -> None:
    # Format the header
    ws.merge_cells(
        start_row=start_row, start_column=1, end_row=start_row, end_column=ws.max_column
    )
    header_cell = ws.cell(start_row, 1)
    header_cell.font = Font(size=18, bold=True)
    header_cell.alignment = Alignment(horizontal="left", vertical="center")

    # Format the column labels
    side = Side(border_style="thin", color="FF000000")
    for i, cell in enumerate(ws[f"{start_row + 1}:{start_row + 1}"]):
        if i == 0:
            # Center the U# column
            cell.alignment = Alignment(horizontal="center", vertical="center")
        else:
            cell.alignment = Alignment(horizontal="left", vertical="center")

        cell.font = Font(size=12, bold=True)
        cell.border = Border(bottom=side, top=side)

    # Format data
    for row in ws.iter_rows(min_row=start_row + 2, max_row=ws.max_row):
        # Add dashed line if next row isn't the same U#
        side = Side(border_style="dashed", color="FF000000")
        if row[0].value != ws.cell(row[0].row + 1, 1).value:
            add_line = True
        else:
            add_line = False

        for idx, cell in enumerate(row):
            if idx == 0:
                # Center the U# column
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif idx == 4:
                # Channel formatting
                cell.number_format = "(0)"
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif idx == 5:
                # Address formatting
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(
                    horizontal="left", vertical="center", wrap_text=True
                )

            cell.font = Font(size=11)
            if add_line:
                cell.border = Border(bottom=side)


def format_colorcut_cols(ws: Worksheet) -> None:
    side = Side(border_style="thin", color="FF000000")
    for i, cell in enumerate(ws["1:1"]):
        if i == 0:
            # Center color
            cell.alignment = Alignment(horizontal="center", vertical="bottom")
        else:
            cell.alignment = Alignment(horizontal="left", vertical="bottom")
        cell.font = Font(size=14, bold=True)
        cell.border = Border(bottom=side)

    ws.row_dimensions[1].height = 30

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for idx, cell in enumerate(row):
            if idx == 0:
                # Gel formatting
                cell.font = Font(size=16)
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.font = Font(size=14)
                cell.alignment = Alignment(
                    horizontal="left", vertical="center", wrap_text=True
                )

    # TODO: add bar on last row


def format_gobos(ws: Worksheet) -> None:
    side = Side(border_style="thin", color="FF000000")
    for i, cell in enumerate(ws["1:1"]):
        if i == 0:
            # Center color
            cell.alignment = Alignment(horizontal="center", vertical="bottom")
        else:
            cell.alignment = Alignment(horizontal="left", vertical="bottom")
        cell.font = Font(size=14, bold=True)
        cell.border = Border(bottom=side)

    ws.row_dimensions[1].height = 30

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for idx, cell in enumerate(row):
            if idx == 0:
                # Gobo formatting
                cell.font = Font(size=14)
                cell.alignment = Alignment(horizontal="left", vertical="center")
            else:
                cell.font = Font(size=14)
                cell.alignment = Alignment(
                    horizontal="left", vertical="center", wrap_text=True
                )

    # TODO: add bar on last row


def format_repeated_gels(ws: Worksheet) -> None:
    side = Side(border_style="thin", color="FF000000")

    for row in range(2, ws.max_row):
        # Search forward until non-matching row found
        end_row = row
        max_row = ws.max_row
        while (
            end_row <= max_row
            and ws.cell(end_row + 1, 1).value == ws.cell(row, 1).value
        ):
            end_row += 1

        if row != end_row:
            # Merge cells
            ws.merge_cells(start_row=row, end_row=end_row, start_column=1, end_column=1)

        for col in range(1, ws.max_column + 1):
            ws.cell(row, col).border = Border(top=side)

        row = end_row - 1


def format_repeated_channel(ws: Worksheet) -> None:
    prev_channel = 0
    prev_row = None
    side = Side(border_style="thin", color="FF000000")

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        if row[0].value == prev_channel:
            # This is a repeated channel!
            row[0].value = ""
            for idx, cell in enumerate(row):
                # if value the same as the last, don't repeat the data
                # exception: do repeat U# since this could get confusing
                if cell.value == prev_row[idx].value and idx != 3:
                    if prev_row[idx].value != "":
                        # Don't "-ify empty fields
                        cell.value = '"'
                    cell.alignment = Alignment(horizontal="left", vertical="center")

        else:
            prev_channel = row[0].value
            prev_row = row
            for idx, cell in enumerate(row):
                cell.border = Border(top=side)


def combine_instrtype(df: pd.DataFrame, add_after: str) -> pd.DataFrame:
    instloadacc = []
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

        instloadacc.append(tmp)

    # Clean up by replacing old cols with new one
    # TODO: Get accessories in here
    # df.drop(["Instrument Type", "Wattage", "Accessory Inventory"], axis=1, inplace=True)
    new_df = df.drop(["Instrument Type", "Wattage"], axis=1)
    new_df.insert(
        new_df.columns.get_loc(add_after) + 1, "Instr Type & Load", instloadacc
    )
    return new_df


# Only operates on Gobo 1
def combine_gelgobo(df: pd.DataFrame, add_after: str) -> pd.DataFrame:
    gelgobo = []
    for index, row in df.iterrows():
        # If no gel replace with N/C
        if row["Color"] == "":
            tmp = "N/C"
        else:
            tmp = row["Color"]

        # If diffusion, add that too
        if row[DIFFUSION_FIELD] != "":
            if tmp == "N/C":
                tmp = row[DIFFUSION_FIELD]
            else:
                tmp += " + " + row[DIFFUSION_FIELD]

        # Append gobo if exists
        if row["Gobo 1"] != "":
            tmp += ", T: " + row["Gobo 1"]

        gelgobo.append(tmp)

    # Clean up by replacing old cols with new one
    new_df = df.drop(["Color", DIFFUSION_FIELD, "Gobo 1"], axis=1)
    new_df.insert(new_df.columns.get_loc(add_after) + 1, "Color & Gobo", gelgobo)
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
    ladders = natsorted([x for x in unique_vals if re.match(r"^S[RL]\sLadder", x)])

    # The order of this is what'll make the order in the schedule!
    special_positions = (
        cat_list
        + foh_list
        + elec_list
        + lx_list
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


# Good heavens this is hacky
def instr_schedule_pagebreaks(ws: Worksheet) -> None:
    # Goal: each position should fit on a page (or at least take up a full page otherwise)
    pos_start_index = 0
    last_height = 0
    cur_height = 0.0

    # Note: all math here done in inches. it's hacky but also excel sucks so
    PAGE_FUDGE = 0.7  # Adjust this if you have weird overflow issues.
    PAGE_HEIGHT = (PAGE_HEIGHT_INCHES - (Y_PADDING * 2 + Y_PADDING_HEADER)) - PAGE_FUDGE

    TYPE_LINEBREAK_LEN = 30
    COLOR_LINEBREAK_LEN = 25

    for row in range(2, ws.max_row):
        # calculate how long this position is
        if ws.cell(row, 5).value is None and ws.cell(row, 1).value is None:
            # end of position
            if last_height + cur_height > PAGE_HEIGHT:
                # we don't want to add this position to the same page, add pagebreak
                ws.row_breaks.append(Break(id=pos_start_index - 1))
                last_height = cur_height
            else:
                last_height = last_height + cur_height + 0.22

            cur_height = 0
        else:
            # Predict height based on value
            if (not ws.cell(row, 1).value.isdigit()) and ws.cell(row, 2).value is None:
                # Position names have heights of 0.33"
                cur_height += 0.33
                pos_start_index = row
            elif ws.cell(row, 1).value.isdigit() or ws.cell(row, 1).value == "U#":
                # Probably a U#, which has default height 0.22"
                if (
                    len(ws.cell(row, 3).value) > TYPE_LINEBREAK_LEN
                    or len(ws.cell(row, 4).value) > COLOR_LINEBREAK_LEN
                ):
                    # Double height
                    cur_height += 0.44
                else:
                    cur_height += 0.22
            else:
                # dunno, assume it's a standard row
                cur_height += 0.22


def add_channel_hookup(wb: Workbook, vw_export: pd.DataFrame) -> None:
    # Format data
    fields = [
        "Channel",
        "Absolute Address",
        "Position",
        "Unit Number",
        "Purpose",
        "Instrument Type",
        "Wattage",
        "Color",
        DIFFUSION_FIELD,
        "Gobo 1",
    ]
    chan_fields = pd.DataFrame(vw_export[fields], columns=fields)
    # Need to have a channel to show up in the channel hookup
    chan_fields["Channel"] = chan_fields["Channel"].replace("", np.nan)
    chan_fields = chan_fields.dropna(subset=["Channel"])

    chan_fields = combine_instrtype(chan_fields, "Purpose")
    chan_fields = format_address_slash(chan_fields)
    chan_fields = combine_gelgobo(chan_fields, "Instr Type & Load")
    chan_fields = chan_fields.rename(columns={"Channel": "Chan", "Unit Number": "U#"})
    chan_fields = chan_fields.sort_values(
        by=["Chan", "Addr", "Position", "U#"], key=natsort_keygen()
    )

    # Dump to worksheet
    ws = wb.create_sheet("Channel Hookup", -1)
    for r in dataframe_to_rows(chan_fields, index=False, header=True):
        ws.append(r)

    # Format
    add_title(ws, "Channel Hookup")
    page_setup(ws, 1)
    set_col_widths(ws, [60, 40, 80, 30, 80, 190, 130])
    format_channelhookup_cols(ws)
    format_repeated_channel(ws)


def add_instrument_schedule(wb: Workbook, vw_export: pd.DataFrame) -> None:
    # Format data
    fields = [
        "Position",
        "Unit Number",
        "Purpose",
        "Instrument Type",
        "Wattage",
        "Color",
        DIFFUSION_FIELD,
        "Gobo 1",
        "Channel",
        "Absolute Address",
    ]

    chan_fields = pd.DataFrame(vw_export[fields], columns=fields)
    # Need to have a position to show up in the instrument schedule
    chan_fields["Position"] = chan_fields["Position"].replace("", np.nan)
    chan_fields = chan_fields.dropna(subset=["Position"])

    chan_fields = combine_instrtype(chan_fields, "Purpose")
    chan_fields = format_address_slash(chan_fields)
    chan_fields = combine_gelgobo(chan_fields, "Instr Type & Load")
    positions, position_names = split_by_position(chan_fields)

    # Dump to worksheet
    ws = wb.create_sheet("Instrument Schedule", -1)
    for idx, pos in enumerate(positions):
        ws.append([position_names[idx]])
        start_row = ws.max_row
        for r in dataframe_to_rows(pos, index=False, header=True):
            ws.append(r)
        ws.append([])

        format_instrschedule_cols(ws, start_row)

    # Format
    add_title(ws, "Instrument Schedule")
    page_setup(ws)
    set_col_widths(ws, [30, 110, 210, 160, 40, 40])
    instr_schedule_pagebreaks(ws)


def parse_gel(gel: str) -> Tuple[str, str, str]:
    if gel.startswith("AP"):
        company = "Apollo"
    elif gel.startswith("G"):
        company = "GAM"
    elif gel.startswith("L"):
        company = "Lee"
    elif gel.startswith("R"):
        company = "Rosco"

    gelsort = gel
    if company == "Rosco":
        if re.match(r"^R3\d\d$", gel):
            # Rosco extended gel, this is basically a .5 gel
            gelsort = "R" + gel[2:] + ".3"

    return gel, company, gelsort


def add_colorcuts(wb: Workbook, vw_export: pd.DataFrame) -> None:
    fields = ["Color", DIFFUSION_FIELD, "Frame Size"]

    chan_fields = pd.DataFrame(vw_export[fields], columns=fields)
    # Seperate colors and diffusion into dict list
    color_dict = []
    for index, row in chan_fields.iterrows():
        if row["Frame Size"] != "" and not pd.isnull(row["Frame Size"]):
            framesize = row["Frame Size"]
        else:
            framesize = "Unknown"

        if row[DIFFUSION_FIELD] != "":
            disp, company, gelsort = parse_gel(row[DIFFUSION_FIELD])
            color_dict.append(
                {
                    "Color": disp,
                    "Frame Size": framesize,
                    "Company": company,
                    "Sort": gelsort,
                }
            )

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
                        disp, company, gelsort = parse_gel(i.split("x")[0])
                        color_dict.append(
                            {
                                "Color": disp,
                                "Frame Size": framesize,
                                "Company": company,
                                "Sort": gelsort,
                            }
                        )
                else:
                    # Normal single gel
                    disp, company, gelsort = parse_gel(i)
                    color_dict.append(
                        {
                            "Color": disp,
                            "Frame Size": framesize,
                            "Company": company,
                            "Sort": gelsort,
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

    # Dump to worksheet
    ws = wb.create_sheet("Color Cut List", -1)
    for r in dataframe_to_rows(colors, index=False, header=True):
        ws.append(r)

    # Format
    add_title(ws, "Color Cut List")
    page_setup(ws, 1)
    set_col_widths(ws, [80, 100, 50])
    format_repeated_gels(ws)
    format_colorcut_cols(ws)


def add_gobos(wb: Workbook, vw_export: pd.DataFrame) -> None:
    fields = ["Gobo 1", "Gobo 2"]
    chan_fields = pd.DataFrame(vw_export[fields], columns=fields)
    gobo_list = []

    for index, row in chan_fields.iterrows():
        if row["Gobo 1"].strip() != "":
            gobo_list.append(row["Gobo 1"])
        if row["Gobo 2"].strip() != "":
            gobo_list.append(row["Gobo 2"])

    gobo_name, gobo_count = np.unique(gobo_list, return_counts=True)

    # Dump to Worksheet
    ws = wb.create_sheet("Gobo Pull List", -1)
    ws.append(["Gobo", "Qty"])
    for i in range(0, len(gobo_name)):
        ws.append([gobo_name[i], gobo_count[i]])

    # Format
    add_title(ws, "Gobo Pull List")
    page_setup(ws, 1)
    set_col_widths(ws, [200, 50])
    format_gobos(ws)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("raw", help="Raw instrument from Vectorworks", type=is_file)
    parser.add_argument("--show", help="Show name")
    parser.add_argument("--ld", help="Lighting designer initials")
    parser.add_argument("--rev", help="Revision letter")

    args = parser.parse_args()

    global SHOW_NAME, LD_NAME, REVISION_STRING
    SHOW_NAME = args.show
    LD_NAME = args.ld
    REVISION_STRING = args.rev

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

    add_channel_hookup(wb, vw_export)
    add_instrument_schedule(wb, vw_export)
    add_colorcuts(wb, vw_export)
    add_gobos(wb, vw_export)

    del wb["Sheet"]
    output_filename = (
        f"{SHOW_NAME.replace(' ', '')}_Paperwork_"
        + re.sub(r"\W+", "", REVISION_STRING)
        + ".xlsx"
    )
    wb.save(output_filename)


if __name__ == "__main__":
    main()
