"""Generator for a channel hookup."""

import logging
from pathlib import Path
from typing import Self, Unpack, override

import numpy as np
import pandas as pd
from natsort import natsort_keygen
from pandas.io.formats.style import Styler

from lighting_paperwork.helpers import FontStyle
from lighting_paperwork.paperwork import PaperworkGenerator, StyleDataParams, StyleFieldParams
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
            key=natsort_keygen(),  # type: ignore[reportArgumentType]
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
    def style_data(df: pd.DataFrame, /, **kwargs: Unpack[StyleDataParams]) -> pd.DataFrame:
        if "chan_style" not in kwargs:
            kwargs["chan_style"] = default_chan_style

        border_style = f"{kwargs['border_weight']}px solid black"
        style_df = pd.DataFrame().reindex_like(df).astype(str)
        # Set borders based on channel data
        prev_row = (None, None)
        for index, data in df.iterrows():
            style_df.loc[index, :] = ""  # type: ignore[reportCallIssue, reportArgumentType]
            if prev_row == (None, None):
                style_df.loc[index, :] += f"border-bottom: {border_style}; "  # type: ignore[reportCallIssue, reportArgumentType]
                prev_row = (index, data)
                continue

            if data["Chan"] == kwargs["quirks"].empty_str:
                style_df.loc[prev_row[0], :] += "border-bottom: none; "  # type: ignore[reportCallIssue, reportArgumentType]
                style_df.loc[index, :] += f"border-bottom: {border_style}; "  # type: ignore[reportCallIssue, reportArgumentType]
            else:
                style_df.loc[index, :] += f"border-bottom: {border_style}; "  # type: ignore[reportCallIssue, reportArgumentType]

            prev_row = (index, data)

        # Set font based on column
        for col_name in style_df:
            style_df[col_name] += (
                "vertical-align: middle; width: "
                f"{kwargs['col_width'][style_df.columns.get_loc(col_name)]}%;"  # type: ignore[reportCallIssue, reportArgumentType]
            )
            if col_name == "Chan":
                style_df[col_name] += f"{kwargs['chan_style'].to_css()}; "
            else:
                style_df[col_name] += f"{kwargs['body_style'].to_css()}; "

            if col_name in ["Chan", "U#", "Addr"]:
                style_df[col_name] += "text-align: center; "
            else:
                style_df[col_name] += "text-align: left; "

        return style_df

    @override
    @staticmethod
    def style_fields(index: pd.Series, /, **kwargs: Unpack[StyleFieldParams]) -> list[str]:
        PaperworkGenerator.verify_width(kwargs["col_width"])
        style = [
            f"{kwargs['header_style'].to_css()}; "
            f"border-top: {kwargs['border_weight']}px solid black;"
            f"border-bottom: {kwargs['border_weight']}px solid black;"
            for _ in index
        ]

        for idx, val in enumerate(index):
            if val in ["Chan", "U#", "Addr"]:
                style[idx] += "text-align: center; "
            else:
                style[idx] += "text-align: left; "

        for idx, _ in enumerate(index):
            style[idx] += f"width: {kwargs['col_width'][idx]}%; "

        return style

    @override
    def _make_common(self) -> pd.io.formats.style.Styler:
        self.generate_df()

        styled = Styler.from_custom_template(
            str(Path(__file__).parent / "templates"), "header_footer.tpl"
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
            type(self).style_fields,  # type: ignore[reportArgumentType]
            header_style=self.style.field,
            col_width=self.col_widths,
            border_weight=self.border_weight,
            axis=1,
        )

        styled = styled.set_table_styles(self.pagebreak_repeated_index("Chan"), overwrite=False)

        return styled  # noqa: RET504
