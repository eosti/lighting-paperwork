"""Tools for importing data from a Vectorworks Data Exchange XML file."""

import logging
from xml.etree import ElementTree as ET

import pandas as pd
from defusedxml.ElementTree import parse as defusedxml_parse

logger = logging.getLogger(__name__)


class VWAccessory:
    """Definition of a lighting accessory.

    Attributes:
        node_uid: The UID of the accessory
        props: Dict of properties related to the accessory

    """

    def __init__(self, node: ET.Element) -> None:
        """Create an accessory from an Accessory-type XML node."""
        self.node_uid: str = node.tag
        self.props: dict[str, str] = {}
        for element in node:
            if element.text:
                # If value, store this
                self.props[element.tag] = element.text
            else:
                # discard if no data in element
                pass


class VWInstrument:
    """Definition of a lighting instrument.

    Attributes:
        node_uid: The UID of the instrument
        props: Dict of properties related to the instrument
        accs: List of associated VWAccessories

    """

    def __init__(self, node: ET.Element) -> None:
        """Create an instrument from an XML node."""
        self.props: dict[str, str] = {}
        self.node_uid: str = node.tag
        self.accs: list[VWAccessory] = []
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
    """Vectorworks XML ingester.

    Attributes:
        instruments: List of VWInstruments from the XML file
        field_mapping: Mapping of XML tags to "pretty" descriptions
        export_time: Timestamp of the XML export
        vw_version: VW version that generated the XML export
        vw_build: VW build that generated the XML export

    """

    def __init__(self, filename: str) -> None:
        """Parse an VW XML export into Python objects.

        Args:
            filename: A valid filename to a XML export

        """
        self.instruments: list[VWInstrument] = []
        self.field_mapping = {}

        if ".xml" not in filename:
            raise ValueError(f"Invalid filetype for VW import (got {filename}, expected *.xml)")

        tree = defusedxml_parse(filename)
        root = tree.getroot()

        if root is None:
            raise RuntimeError("Failed to parse XML file")

        mapping_data = root.find("ExportFieldList")
        if mapping_data is None:
            raise RuntimeError("Unable to find ExportFieldList")
        for field in mapping_data:
            if field.tag == "TimeStamp":
                self.export_time = field.text
            elif field.tag == "AppStamp":
                continue
            else:
                self.field_mapping[field.tag] = field.text

        instrument_data = root.find("InstrumentData")
        if instrument_data is None:
            raise RuntimeError("Unable to find InstrumentData")
        for instr in instrument_data:
            self.parse_instrument(instr)

        logger.debug("VW export generated at %s", self.export_time)
        logger.debug("Importing from VW version %s build %s", self.vw_version, self.vw_build)
        logger.info("Imported %s instruments", len(self.instruments))

    def parse_instrument(self, instr: ET.Element) -> None:
        """Parse an XML node into an instrument or metadata."""
        if "VWVersion" in instr.tag:
            self.vw_version = instr.text
        if "VWBuild" in instr.tag:
            self.vw_build = instr.text
        if "Action" in instr.tag and instr.text != "Entire Plot":
            logger.warning(
                "Export is not of entire plot, results may be incorrect (%s)",
                instr.text,
            )
        if "UID" in instr.tag:
            new_instrument = VWInstrument(instr)

            if new_instrument.props["Device_Type"] == "Accessory":
                # ok so typically the parent is right behind it, so we check
                # more on UIDs to do this cleanly later:
                # https://forum.vectorworks.net/index.php?/topic/
                #   24673-why-do-my-uids-keep-changing/&do=findComment&comment=117429
                if new_instrument.node_uid.split("_")[1] in self.instruments[-1].node_uid:
                    # If major UID numbers match, then that's good enough lol
                    self.instruments[-1].accs.append(VWAccessory(instr))
                else:
                    logger.info("%s is an orphaned accessory", instr.tag)
                    self.instruments.append(new_instrument)
            else:
                self.instruments.append(new_instrument)

    def handle_accessories(
        self, filterlist: tuple[str, ...] = (), fuzzyfilterlist: tuple[str, ...] = ("C-Clamp",)
    ) -> None:
        """Convert VW representation of accessories to one suited for paperwork.

        Adds accessories to a AccessoryString, and if the accessory is smart, make
            a seperate "special" instrument with AccessoryFlag prop set.

        Mutates the existing instruments list to add these accessories

        Args:
            filterlist: a list of exact matches for accessories that should be omitted.
            fuzzyfilterlist: a list of strings that if found in an accessory name,
                will ommit that accessory.

        Returns:
            None

        """
        # TODO(eosti): self.instruments really shouldn't be mutable
        # https://github.com/eosti/lighting-paperwork/issues/11
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
                if any(f in name for f in fuzzyfilterlist):
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

    def export_df(self) -> pd.DataFrame:
        """Convert ingested data into a DataFrame.

        Returns:
            DataFrame with all props listed as rows with their "pretty" name

        """
        self.handle_accessories()
        header = ["Node Tag"]
        header.extend(str(v) for v in self.field_mapping.values())

        all_instr = []
        for instr in self.instruments:
            if instr.props["Action"] == "Delete":
                # Don't export deleted instruments
                continue
            row = [instr.node_uid]
            row.extend(instr.props.get(field, "") for field in self.field_mapping)

            all_instr.append(row)

        return pd.DataFrame(all_instr, columns=header)
