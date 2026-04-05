"""Generator for a gobo pull list."""

import logging
from typing import Self, override

import numpy as np
import pandas as pd

from lighting_paperwork.helpers import FontStyle, FormattingQuirks
from lighting_paperwork.paperwork import PaperworkGenerator

logger = logging.getLogger(__name__)


class GoboPullList(PaperworkGenerator):
    """Generate a gobo pull list with gobo name and quantity."""

    col_widths = (80, 20)
    page_width = 40
    display_name = "Gobo Pull List"

    @override
    def generate_df(self) -> Self:
        filter_fields = ["Gobo 1", "Gobo 2"]
        self.verify_filter_fields(filter_fields)
        chan_fields = pd.DataFrame(self.vw_export[filter_fields], columns=filter_fields)
        gobo_list = []

        for _, row in chan_fields.iterrows():
            if row["Gobo 1"].strip() != "":
                gobo_list.append(row["Gobo 1"])
            if row["Gobo 2"].strip() != "":
                gobo_list.append(row["Gobo 2"])

        gobo_name, gobo_count = np.unique(gobo_list, return_counts=True)

        self.df = pd.DataFrame(
            zip(gobo_name, gobo_count, strict=True), columns=["Gobo Name", "Count"]
        )

        return self

    @override
    @staticmethod
    def style_data(
        df: pd.DataFrame,
        body_style: FontStyle,
        col_width: list[int],
        border_weight: float,
        quirks: FormattingQuirks,
    ) -> pd.DataFrame:
        border_style = f"{border_weight}px solid black"
        style_df = df.copy()

        style_df = style_df.astype(str)
        for index, _ in df.iterrows():
            style_df.loc[index, :] = ""
            style_df.loc[index, :] += f"border-bottom: {border_style}; "

        # Set font based on column
        for col_name in style_df:
            width_idx = style_df.columns.get_loc(col_name)
            style_df[col_name] += (
                f"{body_style.to_css()}; vertical-align: middle; width: {col_width[width_idx]}%; "
            )

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
            f"{header_style.to_css()}; border-bottom: {border_weight}px solid black; "
            for _ in index
        ]

        for idx, val in enumerate(index):
            if val == ["Gobo Name"]:
                style[idx] += "text-align: center; "
            else:
                style[idx] += "text-align: left; "

        for idx, _ in enumerate(index):
            style[idx] += f"width: {col_width[idx]}%; "

        return style
