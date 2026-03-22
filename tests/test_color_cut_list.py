import pytest

from lighting_paperwork.color_cut_list import ColorCutList


def test_parse_color_cut_list(vwx_export):
    paperwork = ColorCutList(vwx_export).generate_df()

    assert paperwork.df is not None

    gels = paperwork.df["Color"].unique()

    assert "N/C" not in gels
    assert "" not in gels
    # Checks for order and presence
    assert (
        gels
        == [
            "AP2200",
            "G440",
            "L201",
            "R01",
            "R02",
            "R42",
            "R342",
            "R43",
            "R79",
            "R83",
            "R383",
            "R119",
            "R124",
            "R125",
            "R126",
            "Red",
            "TBD",
            "TK499",
        ]
    ).all()

    # Verify "L201x2" parses correctly
    l201 = paperwork.df.loc[paperwork.df["Color"] == "L201"]
    assert len(l201) == 1
    assert l201.iloc[0]["Frame Size"] == '6.25"'
    assert l201.iloc[0]["Count"] == 2

    # Check for frames w/ multiple sizes
    r119 = paperwork.df.loc[paperwork.df["Color"] == "R119"]
    assert len(r119) == 2
    assert r119.iloc[0]["Frame Size"] == '6.25"'
    assert r119.iloc[1]["Frame Size"] == '7.5"'
    assert r119.iloc[0]["Count"] == 1
    assert r119.iloc[1]["Count"] == 6

    # Check for frames with no size
    r02 = paperwork.df.loc[paperwork.df["Color"] == "R02"]
    assert len(r02) == 2
    assert r02.iloc[0]["Frame Size"] == ""
    assert r02.iloc[1]["Frame Size"] == '3"'
    assert r02.iloc[0]["Count"] == 2
    assert r02.iloc[1]["Count"] == 1

    # Check for frames that are rectangular
    r124 = paperwork.df.loc[paperwork.df["Color"] == "R124"]
    assert len(r124) == 1
    assert r124.iloc[0]["Frame Size"] == '16"x12.2"'
    assert r124.iloc[0]["Count"] == 1
