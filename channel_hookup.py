import logging
from typing import List, Self

import numpy as np
import pandas as pd
from natsort import natsort_keygen
from pandas.io.formats.style import Styler

from helpers import FontStyle
from paperwork import PaperworkGenerator
from style import default_chan_style

logger = logging.getLogger(__name__)


class ChannelHookup(PaperworkGenerator):
    def __init__(self, *args, chan_style: FontStyle = default_chan_style, **kwargs):
        super().__init__(*args, **kwargs)
        self.chan_style = chan_style

    col_widths = [10, 5, 13, 5, 13, 32, 22]

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
            "Color",
            "Gobo 1",
        ]
        self.df = pd.DataFrame(self.vw_export[filter_fields], columns=filter_fields)
        # Need to have a channel to show up in the channel hookup
        self.df["Channel"] = self.df["Channel"].replace("", np.nan)
        self.df = self.df.dropna(subset=["Channel"])

        self.combine_instrtype().format_address_slash().combine_gelgobo()

        self.df = self.df.rename(columns={"Channel": "Chan", "Unit Number": "U#"})
        self.df = self.df.sort_values(
            by=["Chan", "Addr", "Position", "U#"], key=natsort_keygen()
        )
        self.df = self.df[
            [
                "Chan",
                "Addr",
                "Position",
                "U#",
                "Purpose",
                "Instr Type & Load",
                "Color & Gobo",
            ]
        ]

        self.repeated_channels()
        return self

    @staticmethod
    def style_data(
        df: pd.DataFrame,
        chan_style: FontStyle,
        body_style: FontStyle,
        col_width: List[int],
        border_weight: float,
    ):
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

            if data["Chan"] == "&nbsp;":
                style_df.loc[prev_row[0], :] += "border-bottom: none; "
                style_df.loc[index, :] += f"border-bottom: {border_style}; "
            else:
                style_df.loc[index, :] += f"border-bottom: {border_style}; "

            prev_row = (index, data)

        # Set font based on column
        for col_name, col in style_df.items():
            style_df[col_name] += f"vertical-align: middle; width: {col_width[style_df.columns.get_loc(col_name)]}%;"
            if col_name == "Chan":
                style_df[col_name] += f"{chan_style.to_css()}; "
            else:
                style_df[col_name] += f"{body_style.to_css()}; "

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
            f"{header_style.to_css()}; border-top: {border_weight}px solid black; border-bottom: {border_weight}px solid black; "
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

    def pagebreak_style(self) -> List[dict]:
        # TODO: this doesn't actually do much
        # https://stackoverflow.com/questions/20481039/applying-page-break-before-to-a-table-row-tr
        idxs = np.where(self.df["Chan"] == "&nbsp;")

        selector_list = []

        for i in np.transpose(*idxs):
            # We want to select the row before a repeated channel,
            # but CSS selectors index from 1. These cancel out.
            selector_list.append(f"tr:nth-child({i - 1 + 1})")

        style_list = []

        for i in selector_list:
            style_list.append({"selector": i, "props": [("break-after", "avoid-page")]})

        return style_list

    def make_html(self) -> str:
        self.generate_df()

        styled = Styler.from_custom_template(".", "header_footer.tpl")(self.df)
        styled = styled.apply(
            type(self).style_data,
            axis=None,
            chan_style=self.chan_style,
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
        styled = styled.set_table_styles(self.default_table_style(), overwrite=False)
        styled = styled.set_table_styles(self.pagebreak_style(), overwrite=False)

        header_html = self.generate_header(
            styled.uuid,
            content_right=f"{self.show_data.show_name}<br>{self.show_data.ld_name}",
            content_left=f"{self.show_data.print_date()}<br>{self.show_data.revision}",
            content_center="Channel Hookup",
            style_right=self.style.marginals.to_css() + "margin-bottom: 5%; ",
            style_left=self.style.marginals.to_css() + "margin-bottom: 5%; ",
            style_center=f"{self.style.title.to_css()}",
        )
        footer_html = self.generate_footer(
            styled.uuid,
            style_left=self.style.marginals.to_css(),
            content_left="Channel Hookup",
        )
        page_style = self.generate_page_style(
            "bottom-right", self.style.marginals.to_css()
        )

        return styled.to_html(
            generated_header=header_html,
            generated_footer=footer_html,
            generated_page_style=page_style,
        )
