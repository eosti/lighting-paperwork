"""Generator for a channel hookup."""

import logging
from pathlib import Path
from typing import Self, override

import numpy as np
import pandas as pd
from natsort import natsort_keygen
from pandas.io.formats.style import Styler

from lighting_paperwork.helpers import FontStyle, FormattingQuirks
from lighting_paperwork.paperwork import PaperworkGenerator
from lighting_paperwork.style import default_chan_style

logger = logging.getLogger(__name__)


class ChannelHookup(PaperworkGenerator):
    """Paperwork generator for a channel hookup.

    Generates a channel hookup with channel, address,
    position, type, gobo, color, etc, sorted by channel.
    """

    @override
    def __init__(self, *args, chan_style: FontStyle = default_chan_style, **kwargs) -> None:  # noqa: ANN002, ANN003
        super().__init__(*args, **kwargs)
        self.chan_style = chan_style

    col_widths = (10, 6, 13, 5, 13, 32, 21)
    display_name = "Channel Hookup"

    @override
    def generate_df(self) -> Self:
        # Format data
        filter_fields = [
            "Channel",
            "Absolute Address",
            "Position",
            "Unit Number",
            "Purpose",
            "Instrument Type",
            "Wattage",
            "Accessory String",
            "Accessory Flag",
            "Color",
            "Gobo 1",
            "Gobo 2",
        ]
        self.verify_filter_fields(filter_fields)
        self.df = pd.DataFrame(self.vw_export[filter_fields], columns=filter_fields)

        # Need to have a channel to show up in the channel hookup
        self.df["Channel"] = self.df["Channel"].replace("", np.nan)
        self.df = self.df.dropna(subset=["Channel"])

        self.combine_instrtype().format_address_slash().combine_gelgobo()

        self.df = self.df.rename(columns={"Channel": "Chan", "Unit Number": "U#"})
        self.df = self.df.sort_values(
            by=["Chan", "Position", "U#", "Accessory Flag", "Addr"],
            key=natsort_keygen(),
        )
        self.df = self.df.reset_index(drop=True)
        self.df = self.df[
            [
                "Chan",
                "Addr",
                "Position",
                "U#",
                "Purpose",
                "Instr Type & Load & Acc",
                "Color & Gobo",
            ]
        ]

        self.repeated_index_val("Chan")
        return self

    @override
    @staticmethod
    def style_data(
        df: pd.DataFrame,
        body_style: FontStyle,
        col_width: list[int],
        border_weight: float,
        quirks: FormattingQuirks,
        chan_style: FontStyle = default_chan_style,
    ) -> pd.DataFrame:
        border_style = f"{border_weight}px solid black"
        style_df = df.copy()
        # Set borders based on channel data
        prev_row = (None, None)
        for index, data in df.iterrows():
            style_df.loc[index, :] = ""
            if prev_row == (None, None):
                style_df.loc[index, :] += f"border-bottom: {border_style}; "
                prev_row = (index, data)
                continue

            if data["Chan"] == quirks.empty_str:
                style_df.loc[prev_row[0], :] += "border-bottom: none; "
                style_df.loc[index, :] += f"border-bottom: {border_style}; "
            else:
                style_df.loc[index, :] += f"border-bottom: {border_style}; "

            prev_row = (index, data)

        # Set font based on column
        for col_name in style_df:
            style_df[col_name] += (
                f"vertical-align: middle; width: {col_width[style_df.columns.get_loc(col_name)]}%;"
            )
            if col_name == "Chan":
                style_df[col_name] += f"{chan_style.to_css()}; "
            else:
                style_df[col_name] += f"{body_style.to_css()}; "

            if col_name in ["Chan", "U#", "Addr"]:
                style_df[col_name] += "text-align: center; "
            else:
                style_df[col_name] += "text-align: left; "

        return style_df

    @override
    @staticmethod
    def style_fields(
        index: pd.Series,
        header_style: FontStyle,
        col_width: list[int],
        border_weight: float,
    ) -> list[str]:
        PaperworkGenerator.verify_width(col_width)
        style = [
            f"{header_style.to_css()}; border-top: {border_weight}px solid black;"
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

    def pagebreak_style(self) -> list[dict]:
        """Disallow pagebreaks between channels with the same number.

        TODO: this doesn't actually do much
        https://stackoverflow.com/questions/20481039/applying-page-break-before-to-a-table-row-tr
        """
        idxs = np.where(self.df["Chan"] == self.formatting_quirks.empty_str)

        selector_list = []
        # We want to select the row before a repeated channel,
        # but CSS selectors index from 1. These cancel out.
        selector_list.extend(f"tr:nth-child({i - 1 + 1})" for i in np.transpose(*idxs))

        style_list = []
        style_list.extend(
            {"selector": i, "props": [("break-after", "avoid-page")]} for i in selector_list
        )

        return style_list

    @override
    def _make_common(self) -> pd.io.formats.style.Styler:
        self.generate_df()

        styled = Styler.from_custom_template(
            Path(__file__).parent / "templates", "header_footer.tpl"
        )(self.df)
        styled = styled.apply(
            type(self).style_data,
            axis=None,
            chan_style=self.chan_style,
            body_style=self.style.body,
            col_width=self.col_widths,
            quirks=self.formatting_quirks,
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

        return styled  # noqa: RET504
