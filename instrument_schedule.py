"""Generator for an instrument hookup"""

import logging
import re
from typing import List, Self, Tuple

import numpy as np
import pandas as pd
from natsort import natsort_keygen, natsorted
from pandas.io.formats.style import Styler

from helpers import FontStyle
from paperwork import PaperworkGenerator
from style import default_position_style

logger = logging.getLogger(__name__)


class InstrumentSchedule(PaperworkGenerator):
    """
    Generates an instrument schedules with U#, instrument type, color,
        channel, address, etc per position.
    TODO: Accessories don't show up.
    """

    def __init__(
        self, *args, position_style: FontStyle = default_position_style, **kwargs
    ):
        super().__init__(*args, **kwargs)
        self.position_style = position_style

    col_widths = [5, 17, 36, 28, 7, 7]

    def generate_df(self) -> Self:
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

        self.df = pd.DataFrame(self.vw_export[filter_fields], columns=filter_fields)
        # Need to have a position to show up in the instrument schedule
        self.df["Position"] = self.df["Position"].replace("", np.nan)
        self.df = self.df.dropna(subset=["Position"])

        self.combine_instrtype().format_address_slash().combine_gelgobo().abbreviate_col_names()
        self.df = self.df.sort_values(
            by=["Position", "U#", "Purpose"], key=natsort_keygen()
        )

        self.df = self.df[
            [
                "Position",
                "U#",
                "Purpose",
                "Instr Type & Load",
                "Color & Gobo",
                "Chan",
                "Addr",
            ]
        ]

        return self

    def split_by_position(self) -> List[Tuple[pd.DataFrame, str]]:
        """
        Split dataframe into multiple dataframes, one per position.
        """
        # Step one: sort position names
        unique_vals = self.df["Position"].unique()

        # Sort by Cat (descending), Elec (ascending), other
        elec_list = natsorted([x for x in unique_vals if re.match(r"^Elec\s\d", x)])
        lx_list = natsorted([x for x in unique_vals if re.match(r"^LX\d", x)])
        cat_list = natsorted(
            [x for x in unique_vals if re.match(r"^Cat\s\d", x)], reverse=True
        )
        # TODO: might be nice to force a linebreak between categories
        foh_list = natsorted(
            [
                x
                for x in unique_vals
                if re.match(r"^FOH\s\d", x) or re.match(r"^FOH$", x)
            ]
        )

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
        sorted_dfs = []
        for i in position_names:
            pos_df = self.df.loc[self.df["Position"] == i].copy()
            pos_df = pos_df.drop(["Position"], axis=1)
            pos_df = pos_df.rename(columns={"Channel": "Chan", "Unit Number": "U#"})
            pos_df = pos_df.sort_values(by=["U#"], key=natsort_keygen())
            sorted_dfs.append((pos_df, i))

        return sorted_dfs

    @staticmethod
    def style_data(
        df: pd.DataFrame,
        body_style: FontStyle,
        col_width: List[int],
        border_weight: float,
    ):
        chan_border_style = f"{border_weight}px dashed black"
        style_df = df.copy()
        # Set borders based on channel data
        prev_row = (None, None)
        for index, data in df.iterrows():
            style_df.loc[index, :] = ""
            if prev_row == (None, None):
                style_df.loc[index, :] += f"border-bottom: {chan_border_style}; "
                prev_row = (index, data)
                continue

            if data["U#"] == prev_row[1]["U#"]:
                # Same U#, remove dashed line
                style_df.loc[prev_row[0], :] += "border-bottom: none; "
                style_df.loc[index, :] += f"border-bottom: {chan_border_style}; "
            else:
                style_df.loc[index, :] += f"border-bottom: {chan_border_style}; "

            prev_row = (index, data)

        # Last row gets a solid bottom border
        style_df.loc[prev_row[0], :] += (
            f"border-bottom: {border_weight}px solid black; "
        )

        # Set font based on column
        for col_name, _ in style_df.items():
            style_df[col_name] += (
                f"{body_style.to_css()}; vertical-align: middle; "
                f"width: {col_width[style_df.columns.get_loc(col_name)]}%; "
            )

            if col_name in ["Chan", "U#", "Addr"]:
                style_df[col_name] += "text-align: center; "
            else:
                style_df[col_name] += "text-align: left; "

        return style_df

    @staticmethod
    def style_fields(
        index: pd.Series,
        header_style: FontStyle,
        col_width: List[int],
        border_weight: float,
    ) -> List[str]:
        PaperworkGenerator.verify_width(col_width)
        style = [
            f"{header_style.to_css()}; border-top: {border_weight}px solid black; "
            f"border-bottom: {border_weight}px solid black; "
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

    def make_html(self) -> str:
        self.generate_df()
        positions = self.split_by_position()

        output_html = "<div id='paperwork-container' style='width: 670px;'>\n"
        for df, position in positions:
            styled = Styler.from_custom_template(".", "header_footer.tpl")(df)
            styled = styled.apply(
                type(self).style_data,
                axis=None,
                body_style=self.style.body,
                col_width=self.col_widths,
                border_weight=self.border_weight,
            )
            styled = styled.hide()
            styled = styled.apply_index(
                type(self).style_fields,
                header_style=self.style.field,
                col_width=self.col_widths,
                border_weight=self.border_weight,
                axis=1,
            )
            styled = styled.set_table_attributes('class="paperwork-table"')
            styled = styled.set_table_styles(
                self.default_table_style(), overwrite=False
            )
            styled = styled.set_table_styles(
                [{"selector": "", "props": "break-inside: avoid; margin-bottom: 5mm;"}],
                overwrite=False,
            )

            header_html = self.generate_header(
                styled.uuid,
                content_left=position,
                style_left=self.position_style.to_css(),
            )

            output_html += styled.to_html(
                generated_header=header_html,
            )

        output_html += "\n</div>"
        return output_html
