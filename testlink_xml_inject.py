# Script usage:
# - export ONE testcase from testlink as xml file
# - create excel file (xls, xlsx) where first column is "Actions"
#   second column is "Expected Results". File should NOT contains
#   any headers or other information, data should be on first sheet
# - run this file with first parameter path to excel and second path
#   to xml. Actual command see down below.
# - Export xml file back to the testlink. You have options for already
#   existed cases: update latest version (prefered), create new version
#   and so on.
#
# Requirements:
# pip install pandas openpyxl
#
# ATTENTION: Next command will change xml file, if you want to preserve
# original file, please copy it manually.
#
# python testlink_xml_inject.py {path to excel file} {path to xml file}

# %%
import pandas as pd
import xml.etree.ElementTree as ET
import sys
import os

# %%
def _tag_text_to_CDATA(element: ET.Element) -> None:
    """Convert tag element text to CDATA xml type, and also add multiline text
    ----------
    Parameters:
    element: tag element
    ----------
    Returns: it is void fucntion
    """

    def _escape_cdata(text):
        # escape character data
        try:
            # it's worth avoiding do-nothing calls for strings that are
            # shorter than 500 characters, or so.  assume that's, by far,
            # the most common case in most applications.
            if "&" in text:
                text = text.replace("&", "&amp;")
            if ("<" in text) and (text.startswith("<![CDATA") is False):
                text = text.replace("<", "&lt;")
            if (">" in text) and (text.startswith("<![CDATA") is False):
                text = text.replace(">", "&gt;")
            return text
        except (TypeError, AttributeError):
            _raise_serialization_error(text)

    ET._escape_cdata = _escape_cdata

    for child in element:
        if child.text is not None:
            child.text = f"<![CDATA[{child.text}]]>"


def create_steps_from_excel(path_to_excel: str) -> ET.Element:
    """Create xml tag steps from excel file
    ----------
    Parameters:
    path_to_excel: path to excel file with steps
    ----------
    Returns: return xml tag steps in ET.Element format
    """
    path_to_excel = os.path.abspath(path_to_excel)
    steps_xlsx = pd.ExcelFile(path_to_excel)
    steps_sheet = pd.read_excel(steps_xlsx, header=None)
    steps_sheet.set_index((i + 1 for i in steps_sheet.index), inplace=True)

    steps = ET.Element("steps")
    for row in steps_sheet.index:
        step = ET.Element("step")
        step_number = ET.Element("step_number")
        step_number.text = str(row)
        actions = ET.Element("actions")
        actions.text = str(steps_sheet.at[row, 0])
        actions.text = actions.text.replace("\n", r"<br>")
        expectedresults = ET.Element("expectedresults")
        expectedresults.text = str(steps_sheet.at[row, 1])
        expectedresults.text = expectedresults.text.replace("\n", r"<br>")
        step.append(step_number)
        step.append(actions)
        step.append(expectedresults)
        _tag_text_to_CDATA(step)
        steps.append(step)
    return steps


def prepare_xml(path_to_xml: str) -> ET.Element:
    """Load testcase xml, and delete steps tag
    ----------
    Parameters:
    path_to_xml: path to xml file with testcase
    ----------
    Returns: return whole xml tree in ET.ElementTree and xml tag testcase in ET.Element formats
    """
    path_to_xml = os.path.abspath(path_to_xml)
    tree = ET.parse(path_to_xml)
    root = tree.getroot()
    testcase = root.find("./testcase")
    original_steps = testcase.find("./steps")
    if original_steps is not None:
        testcase.remove(original_steps)
    _tag_text_to_CDATA(testcase)
    return tree, testcase


def main(*args):
    path_to_excel = sys.argv[1]
    path_to_xml = sys.argv[2]
    steps = create_steps_from_excel(path_to_excel)
    tree, testcase = prepare_xml(path_to_xml)
    # append new tag steps from our excel file
    testcase.append(steps)
    # write our changes back to the xml file
    tree.write(path_to_xml, encoding="utf-8")


# %%
if __name__ == "__main__":
    main(*sys.argv)
# %%
