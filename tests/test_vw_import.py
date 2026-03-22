import pytest
import logging

from lighting_paperwork.vectorworks_xml import VWExport

def test_opening_files(vwx_export_file):
    with pytest.raises(ValueError):
        vwx_export = VWExport("notafile.asdf")

    with pytest.raises(FileNotFoundError):
        vwx_export = VWExport("notafile.xml")

    vwx_export = VWExport(vwx_export_file)


def test_parsing_file(caplog, vwx_export_file):
    caplog.set_level(logging.WARNING)
    vwx_export = VWExport(vwx_export_file)

    assert len(vwx_export.instruments) == pytest.NUM_INSTRUMENTS

    uid_list = []
    for i in vwx_export.instruments:
        uid_list.append(i.props["UID"])
    # All unique UIDs?
    assert len(uid_list) == len(set(uid_list))

    accessory_count = 0
    for i in vwx_export.instruments:
        accessory_count += len(i.accs)
    assert accessory_count == pytest.NUM_SMART_ACCS + pytest.NUM_DUMB_ACCS

    assert caplog.text == ""


def test_export_df(vwx_export_file):
    vwx_export = VWExport(vwx_export_file)
    df = vwx_export.export_df()

    # number of instrument + number of smart accessories
    assert len(df) == (pytest.NUM_INSTRUMENTS + pytest.NUM_SMART_ACCS)
    # Additional col for Node UID
    assert len(df.columns) == (len(vwx_export.field_mapping)) + 1
