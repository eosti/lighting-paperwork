import logging

import pytest

from lighting_paperwork.channel_hookup import ChannelHookup


def test_parse_channel_hookup(caplog, vwx_export):
    caplog.set_level(logging.WARNING)
    paperwork = ChannelHookup(vwx_export).generate_df()

    assert paperwork.df is not None

    # Verify channel order strictly increments
    prev_chan_val = -1
    for idx, row in paperwork.df.iterrows():
        channel_number = row["Chan"]
        if channel_number == "&nbsp;":
            continue
        assert int(channel_number) > prev_chan_val
        prev_chan_val = int(channel_number)

    # Verify number of channel rows
    # -2 is for the channel and accessory w/o channel number
    assert len(paperwork.df.index) == pytest.NUM_INSTRUMENTS + pytest.NUM_SMART_ACCS - 2

    # Check power handling
    # LEDBeam has power listed as '0 W'
    ledbeam = paperwork.df.loc[paperwork.df["Chan"] == "201"]
    assert len(ledbeam) == 1
    assert ledbeam.iloc[0]["Instr Type & Load & Acc"] == "Robe Lighting LEDBeam 150"
    assert (
        "Channel 201 is infinitely efficient (Robe Lighting LEDBeam 150 is 0W)"
        in caplog.text
    )

    # Other LEDBeam has power listed as '0W'
    ledbeam = paperwork.df.loc[paperwork.df["Chan"] == "202"]
    assert len(ledbeam) == 1
    assert ledbeam.iloc[0]["Instr Type & Load & Acc"] == "Robe Lighting LEDBeam 150"
    assert (
        "Channel 202 is infinitely efficient (Robe Lighting LEDBeam 150 is 0W)"
        in caplog.text
    )

    # Other LEDBeam has power listed as '0'
    ledbeam = paperwork.df.loc[paperwork.df["Chan"] == "203"]
    assert len(ledbeam) == 1
    assert ledbeam.iloc[0]["Instr Type & Load & Acc"] == "Robe Lighting LEDBeam 150"
    assert (
        "Channel 203 is infinitely efficient (Robe Lighting LEDBeam 150 is 0W)"
        in caplog.text
    )

    # Other LEDBeam has power listed as ''
    ledbeam = paperwork.df.loc[paperwork.df["Chan"] == "204"]
    assert len(ledbeam) == 1
    assert ledbeam.iloc[0]["Instr Type & Load & Acc"] == "Robe Lighting LEDBeam 150"
    assert (
        "Channel 204 is infinitely efficient (Robe Lighting LEDBeam 150 is 0W)"
        in caplog.text
    )

    # Other LEDBeam has power listed as '    '
    ledbeam = paperwork.df.loc[paperwork.df["Chan"] == "205"]
    assert len(ledbeam) == 1
    assert ledbeam.iloc[0]["Instr Type & Load & Acc"] == "Robe Lighting LEDBeam 150"
    assert (
        "Channel 205 is infinitely efficient (Robe Lighting LEDBeam 150 is 0W)"
        in caplog.text
    )

    # Light has power in power field and instrument type as '575 W'
    light = paperwork.df.loc[paperwork.df["Chan"] == "51"]
    assert len(light) == 1
    assert light.iloc[0]["Instr Type & Load & Acc"] == "ETC Source 4 PAR MFL 575W"

    # Light has power in power field as '' and instrument type as '575 W'
    light = paperwork.df.loc[paperwork.df["Chan"] == "52"]
    assert len(light) == 1
    assert light.iloc[0]["Instr Type & Load & Acc"] == "ETC Source 4 PAR NSP 575W"

    # Light has power in power field as '575 W' and instrument type as ''
    light = paperwork.df.loc[paperwork.df["Chan"] == "53"]
    assert len(light) == 1
    assert light.iloc[0]["Instr Type & Load & Acc"] == "ETC Source 4 PAR VNSP 575W"

    # Light has power in power field as '575 W' and instrument type as '750 W'
    light = paperwork.df.loc[paperwork.df["Chan"] == "54"]
    assert len(light) == 1
    assert light.iloc[0]["Instr Type & Load & Acc"] == "ETC Source 4 PAR NSP 575W"

    # Light has power in power field as '575' and instrument type as '750W'
    light = paperwork.df.loc[paperwork.df["Chan"] == "55"]
    assert len(light) == 1
    assert light.iloc[0]["Instr Type & Load & Acc"] == "ETC Source 4 PAR NSP 575W"
    assert (
        "Channel 54 has conflicting power values (575W, 750W). Using 575W."
        in caplog.text
    )

    # Scoop has power listed as '1.5 kW'
    scoop = paperwork.df.loc[paperwork.df["Chan"] == "41"]
    assert len(scoop) == 1
    assert scoop.iloc[0]["Instr Type & Load & Acc"] == "Altman 18in Scoop 1.5kW"

    # Other scoop has power listed as '1.5kW'
    scoop = paperwork.df.loc[paperwork.df["Chan"] == "42"]
    assert len(scoop) == 1
    assert scoop.iloc[0]["Instr Type & Load & Acc"] == "Altman 18in Scoop 1.5kW"

    # Instrument type is `Strand 6x16     `, verify spaces are stripped
    strand = paperwork.df.loc[paperwork.df["Chan"] == "38"]
    assert len(strand) == 1
    assert "Strand 6x16 1kW" in strand.iloc[0]["Instr Type & Load & Acc"]

    # Verify accessories exist
    acc = paperwork.df.loc[paperwork.df["Chan"] == "312"]
    assert len(acc) == 1
    assert "Rosco I-Cue" in acc.iloc[0]["Instr Type & Load & Acc"]
    assert acc.iloc[0]["Color & Gobo"] == ""
    assert acc.iloc[0]["Purpose"] == "Moving"

    acc = paperwork.df.loc[paperwork.df["Chan"] == "313"]
    assert len(acc) == 1
    assert "ChromaQ Chroma-Q" in acc.iloc[0]["Instr Type & Load & Acc"]
    assert acc.iloc[0]["Color & Gobo"] == ""
    assert acc.iloc[0]["Purpose"] == "Coloring"

    acc = paperwork.df.loc[paperwork.df["Chan"] == "311"]
    assert len(acc) == 1
    assert "ChromaQ Chroma-Q" in acc.iloc[0]["Instr Type & Load & Acc"]
    assert "Rosco I-Cue" in acc.iloc[0]["Instr Type & Load & Acc"]
    assert "Iris" in acc.iloc[0]["Instr Type & Load & Acc"]

    acc_idx = paperwork.df.loc[paperwork.df["Chan"] == "74"].index
    acc = paperwork.df.loc[
        [e for lst in [range(idx, idx + 4) for idx in acc_idx] for e in lst]
    ].copy()
    assert len(acc) == 4
    assert "Rosco I-Cue" in acc.iloc[1]["Instr Type & Load & Acc"]
    assert acc.iloc[1]["U#"] == acc.iloc[0]["U#"]
    assert "Rosco I-Cue" in acc.iloc[3]["Instr Type & Load & Acc"]
    assert acc.iloc[3]["U#"] == acc.iloc[2]["U#"]
    # Check repeated channel formatting
    for i in range(1, 4):
        assert acc.iloc[i]["Addr"] == '"'
        assert acc.iloc[i]["Position"] == '"'
        assert acc.iloc[i]["Chan"] == "&nbsp;"

    # Verify repeated address w/ same channel don't "-ify
    lx = paperwork.df.loc[paperwork.df["Chan"] == "79"]
    assert len(lx) == 1
    assert lx.iloc[0]["Addr"] == "74"

    # Check gobos
    lx = paperwork.df.loc[paperwork.df["Chan"] == "74"]
    assert len(lx) == 1
    assert lx.iloc[0]["Color & Gobo"] == "R42, T: GAM 673-Linear Breakup 2"

    lx = paperwork.df.loc[paperwork.df["Chan"] == "37"]
    assert len(lx) == 1
    assert lx.iloc[0]["Color & Gobo"] == "L201x2, T: GAM 636-Construction B, R77405"
