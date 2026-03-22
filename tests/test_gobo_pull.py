import pytest

from lighting_paperwork.gobo_pull import GoboPullList

def test_parse_gobo_pull_list(vwx_export):
    paperwork = GoboPullList(vwx_export).generate_df()

    assert paperwork.df is not None

    gobos = paperwork.df["Gobo Name"].unique()

    assert (
        gobos == [
            "G635-Construction A",
            "GAM 222-Small Breakup",
            "GAM 636-Construction B",
            "GAM 673-Linear Breakup 2",
            "R77405"
        ]
    ).all()

    # Verify Gobo 1 will be counted properly
    r77405 = paperwork.df.loc[paperwork.df["Gobo Name"] == "R77405"]
    assert len(r77405) == 1
    assert r77405.iloc[0]["Count"] == 2

    # Verify Gobo 2 will be counted properly
    g222 = paperwork.df.loc[paperwork.df["Gobo Name"] == "GAM 222-Small Breakup"]
    assert len(g222) == 1
    assert g222.iloc[0]["Count"] == 1

    # Verify Gobo 1 and Gobo 2 sum
    g626 = paperwork.df.loc[paperwork.df["Gobo Name"] == "GAM 636-Construction B"]
    assert len(g626) == 1
    assert g626.iloc[0]["Count"] == 2
