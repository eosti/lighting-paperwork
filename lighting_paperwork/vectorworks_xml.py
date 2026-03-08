"""Tools for importing data from a Vectorworks Data Exchange XML file"""

import logging
from xml.etree import ElementTree as ET

import pandas as pd

logger = logging.getLogger(__name__)


class VWAccessory:
    """
    Definition of a lighting accessory.
    """

    def __init__(self, node):
        self.node_uid = node.tag
        self.props = {}
        for element in node:
            if element.text:
                # If value, store this
                self.props[element.tag] = element.text
            else:
                # discard if no data in element
                pass


class VWInstrument:
    """
    Definition of a lighting instrument.
    """

    def __init__(self, node):
        self.props = {}
        self.node_uid = node.tag
        self.accs = []
        for element in node:
            if element.text and element.text.strip():
                # If value, store this
                self.props[element.tag] = element.text
            elif element.tag == "Accessories":
                # Parse accessories if any
                for acc in element:
                    self.accs.append(VWAccessory(acc))
            else:
                # discard if no data in element
                pass


class VWExport:
    """
    Vectorworks XML ingester
    """

    def __init__(self, filename: str):
        self.instruments: list[VWInstrument] = []
        self.field_mapping = {}
        tree = ET.parse(filename)
        root = tree.getroot()

        mapping_data = root.find("ExportFieldList")
        if mapping_data is None:
            raise RuntimeError("Unable to find ExportFieldList")
        for field in mapping_data:
            if field.tag in ("AppStamp", "TimeStamp"):
                continue
            self.field_mapping[field.tag] = field.text

        instrument_data = root.find("InstrumentData")
        if instrument_data is None:
            raise RuntimeError("Unable to find InstrumentData")
        for instr in instrument_data:
            if "VWVersion" in instr.tag:
                self.vw_version = instr.text
            if "UID" in instr.tag:
                new_instrument = VWInstrument(instr)

                if new_instrument.props["Device_Type"] == "Accessory":
                    # ok so typically the parent is right behind it, so we check
                    # more on UIDs to do this cleanly later:
                    # https://forum.vectorworks.net/index.php?/topic/
                    #   24673-why-do-my-uids-keep-changing/&do=findComment&comment=117429
                    if (
                        new_instrument.node_uid.split("_")[1]
                        in self.instruments[-1].node_uid
                    ):
                        # If major UID numbers match, then that's good enough lol
                        self.instruments[-1].accs.append(VWAccessory(instr))
                    else:
                        logger.info("UID %s is an orphaned accessory", instr.tag)
                        self.instruments.append(new_instrument)
                else:
                    self.instruments.append(new_instrument)

        self.handle_accessories()
        logger.info("Imported %s instruments", len(self.instruments))

    def handle_accessories(self, filterlist=[], fuzzyfilterlist=["C-Clamp"]):
        """
        Adds accessories to a AccessoryString, and if the accessory is smart, make
            a seperate "special" instrument.

        filterlist is a list of exact matches for accessories that should be omitted.
        fuzzyfilterlist is a list of strings that if found in an accessory name,
            will ommit that accessory.
        """
        self.field_mapping["AccessoryString"] = "Accessory String"
        self.field_mapping["AccessoryFlag"] = "Accessory Flag"
        additional_accs = []
        for instr in self.instruments:
            if instr.accs == []:
                instr.props["AccessoryString"] = ""
                instr.props["AccessoryFlag"] = "0"
                continue
            acc_names = []
            for acc in instr.accs:
                name = acc.props["Symbol_Name"]
                name = name.replace("Light Acc", "").strip()
                if name in filterlist:
                    continue
                if [1 for f in fuzzyfilterlist if (f in name)]:
                    continue
                acc_names.append(name)

                if acc.props["Device_Type"] == "Accessory":
                    # Accessories are typically smarter and require their own entry
                    # Some data needs to be copied from the parent to make a new entry
                    acc.props["AccessoryFlag"] = "1"
                    acc.props["Unit_Number"] = instr.props["Unit_Number"]
                    acc.props["Inst_Type"] = name
                    # No recursive accessories
                    acc.props["AccessoryString"] = ""
                    additional_accs.append(acc)

            namestr = ""
            for idx, name in enumerate(acc_names):
                if idx == len(acc_names) - 1:
                    # Final name
                    namestr += name
                else:
                    namestr += f"{name}, "

            instr.props["AccessoryString"] = namestr

        # do this last to prevent recursive conversion
        self.instruments += additional_accs

    def export_df(self):
        """Converts ingested data into a DataFrame compatible with CSV imports"""
        header = ["__UID"]
        for _, v in self.field_mapping.items():
            header.append(str(v))

        all_instr = []
        for instr in self.instruments:
            if instr.props["Action"] == "Delete":
                # Don't export deleted instruments
                continue
            row = [instr.node_uid]
            for field in self.field_mapping:
                row.append(instr.props.get(field, ""))

            all_instr.append(row)

        return pd.DataFrame(all_instr, columns=header)


def main():
    """Test XML ingest by printing exported DF"""
    logging.basicConfig(level=logging.DEBUG)
    export = VWExport("vw.xml")

    print(export.export_df())


if __name__ == "__main__":
    main()
