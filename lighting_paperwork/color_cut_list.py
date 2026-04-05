"""Generator for a color cut list."""

import logging
from typing import Self, override

import pandas as pd
from natsort import natsort_keygen

from lighting_paperwork.helpers import (
    FontStyle,
    FormattingQuirks,
    Gel,
    parse_frame_size,
)
from lighting_paperwork.paperwork import PaperworkGenerator

logger = logging.getLogger(__name__)


class ColorCutList(PaperworkGenerator):
    """Generate a color pull list with color, size, and quantity."""

    col_widths = (34, 43, 23)
    page_width = 30
    display_name = "Color Cut List"

    @override
    def generate_df(self) -> Self:
        filter_fields = ["Color", "Frame Size"]
        self.verify_filter_fields(filter_fields)
        self.df = pd.DataFrame(self.vw_export[filter_fields], columns=filter_fields)

        # Seperate colors and diffusion into dict list
        color_dict = []
        for _, row in self.df.iterrows():
            framesize = parse_frame_size(row["Frame Size"])
            all_gels = Gel.parse_gel_string(row["Color"])
            for i in all_gels:
                if i.name in ("", "N/C"):
                    continue
                color_dict.append(
                    {
                        "Color": i.name,
                        "Frame Size": framesize,
                        "Company": i.company,
                        "Sort": i.name_sort,
                    }
                )

        colors = pd.DataFrame.from_dict(color_dict)
        colors = (
            colors.groupby(["Color", "Frame Size", "Sort"])["Color"]
            .count()
            .reset_index(name="Count")
        )
        colors = colors.sort_values(by=["Sort", "Frame Size"], key=natsort_keygen())
        colors = colors.drop(["Sort"], axis=1)

        self.df = colors
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
        # Set borders based on color data
        prev_row = (None, None)
        style_df = style_df.astype(str)
        for index, data in df.iterrows():
            style_df.loc[index, :] = ""
            if prev_row == (None, None):
                style_df.loc[index, :] += f"border-bottom: {border_style}; "
                prev_row = (index, data)
                continue

            if data["Color"] == prev_row[1]["Color"]:
                style_df.loc[prev_row[0], "Color"] += "border-bottom: none; "
                style_df.loc[index, "Color"] += f"color: {quirks.hidden_fmt}; "
                style_df.loc[index, :] += f"border-bottom: {border_style}; "
            else:
                style_df.loc[index, :] += f"border-bottom: {border_style}; "

            prev_row = (index, data)

        # Set font based on column
        for col_name in style_df:
            width_idx = style_df.columns.get_loc(col_name)
            style_df[col_name] += (
                f"{body_style.to_css()}; vertical-align: middle; width: {col_width[width_idx]}%; "
            )

            if col_name == ["Color"]:
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
            f"{header_style.to_css()}; border-bottom: {border_weight}px solid black; "
            for _ in index
        ]

        for idx, val in enumerate(index):
            if val == ["Color"]:
                style[idx] += "text-align: center; "
            else:
                style[idx] += "text-align: left; "

        for idx, _ in enumerate(index):
            style[idx] += f"width: {col_width[idx]}%; "

        return style
