"""Tests for the instrument schedule generator."""

import logging
import re

import pytest

from lighting_paperwork.instrument_schedule import InstrumentSchedule


def test_parse_instrument_schedule(caplog, vwx_export):
    caplog.set_level(logging.WARNING)
    paperwork = InstrumentSchedule(vwx_export).generate_df()

    assert paperwork.df is not None

    # Verify U#s increment within positions
    prev_u_val = -1
    this_position = ""
    for _, row in paperwork.df.iterrows():
        if row["Position"] != this_position:
            this_position = row["Position"]
            prev_u_val = -1
            continue
        assert row["U#"] != ""
        u_num = re.sub("[^0-9]", "", row["U#"])
        assert int(u_num) >= prev_u_val
        prev_u_val = int(u_num)


def test_position_sorting():
    input_positions = [
        "AP7",
        "AP38",
        "AP9",
        "Elec 3",
        "2 Elec",
        "LX 4",
        "31 LX",
        "LX5",
        "Cat 7",
        "11 Cat",
        "Pipe 6",
        "61 Pipe",
        "68 Pipe",
        "Pipe8",
        "The Attic",
        "Spot Booth",
        "E6",
        "E 8",
        "8E",
        "1 E",
        "SR Ladder 4",
        "SL Ladder 4",
        "SR Boom 6",
        "SL Boom 33",
        "DS Boom 7",
        "US Boom 66",
        "DSR Boom 879",
        "SL Box Boom",
        "SR Box Boom 6",
    ]

    output = InstrumentSchedule.sort_positions(input_positions, InstrumentSchedule.position_regexes)

    assert output == [
        "61 Pipe",
        "68 Pipe",
        "Pipe8",
        "Pipe 6",
        "1 E",
        "8E",
        "E6",
        "E 8",
        "2 Elec",
        "Elec 3",
        "31 LX",
        "LX5",
        "LX 4",
        "AP7",
        "AP9",
        "AP38",
        "11 Cat",
        "Cat 7",
        "SL Box Boom",
        "SR Box Boom 6",
        "DS Boom 7",
        "DSR Boom 879",
        "SL Boom 33",
        "SR Boom 6",
        "US Boom 66",
        "SL Ladder 4",
        "SR Ladder 4",
        "Spot Booth",
        "The Attic",
    ]


def test_split_by_position(caplog, vwx_export):
    caplog.set_level(logging.WARNING)
    dfs = InstrumentSchedule(vwx_export).generate_df().split_by_position()

    assert dfs is not None
    # Elec 2 is empty
    assert len(dfs) == pytest.NUM_POSITIONS - 1

    assert dfs[0][0] == "Pipe 8"
    pipe_eight = dfs[0][1]
    assert len(pipe_eight) == 7
    assert pipe_eight.iloc[0]["U#"] == "1"
    assert pipe_eight.iloc[2]["U#"] == "3"
    assert pipe_eight.iloc[2]["Addr"] == "&nbsp;"
    assert pipe_eight.iloc[3]["U#"] == "&nbsp;"
    assert pipe_eight.iloc[3]["Instr Type & Load & Acc"] == '"'
    assert pipe_eight.iloc[3]["Color & Gobo"] == '"'
    assert pipe_eight.iloc[3]["Chan"] == "204"
    assert pipe_eight.iloc[3]["Addr"] == "&nbsp;"

    assert dfs[2][0] == "3 Elec"
    third_elec = dfs[2][1]
    assert len(third_elec) == 8
    assert third_elec.iloc[1]["Chan"] == "&nbsp;"
    assert third_elec.iloc[5]["U#"] == "&nbsp;"
    assert third_elec.iloc[5]["Chan"] == '"'
    assert third_elec.iloc[5]["Addr"] == '"'
