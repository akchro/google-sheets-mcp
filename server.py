from mcp.server.fastmcp import FastMCP
from sheets_functions import get_user_spreadsheet_ids, create_new_spreadsheet, copy_to_spreadsheet, mass_edit_spreadsheet

mcp = FastMCP("Google Sheets MCP")

@mcp.tool()
def get_spreadsheets():
    """
    Gets all the spreadsheets of the authenticated user

    return:
        A seperated string of spreadsheet names and IDs
    """
    spreadsheets = get_user_spreadsheet_ids()
    if spreadsheets is None:
        return "Unable to get spreadsheets"

    if len(spreadsheets) == 0:
        return "No spreadsheets found"

    return "\n---\n".join(spreadsheets)


@mcp.tool()
def create_spreadsheet(title: str):
    """
    Creates a spreadsheet with the given title
    :param title: The title of the spreadsheet
    :return: A string that contains the spreadsheet title and ID
    """
    spreadsheet = create_new_spreadsheet(title)

    if spreadsheet is None:
        return "Unable to create spreadsheet"

    return f"Created new spreadsheet with title {spreadsheet[0]} and ID {spreadsheet[1]}"


@mcp.tool()
def copy_spreadsheet(sheet_id_to: str, sheet_id_from: str):
    """
    Copies a spreadsheet from sheet_id_from to sheet_id_to

    :param sheet_id_to: The ID of the sheet to copy
    :param sheet_id_from: The ID of the sheet that is being copied from
    :return:
    """
    speadsheets = copy_to_spreadsheet(sheet_id_to, sheet_id_from)

    if speadsheets is None:
        return "Unable to copy spreadsheet"

    return f"Copied spreadsheet {speadsheets[0]} to {speadsheets[1]} with sheet name {speadsheets[2]}"


@mcp.tool()
def edit_spreadsheet(spreadsheet_id, ranges_and_values):
    """
Edits a spreadsheet with the given ID.

:param spreadsheet_id: The ID of the spreadsheet to edit.
:param ranges_and_values:
    A list of tuples, where each tuple contains:
    1. A string representing the range to update (e.g., "A1:B2").
    2. A list of lists, where each inner list represents a row of values to be inserted into that range.

    Example:
    ranges_and_values = [
        ("A1:B2", [["Hello", "World"], ["Foo", "Bar"]]),
        ("C1:D2", [["Test", "Value"], ["Example", "Update"]]),
    ]
    In this example:
    - The range "A1:B2" will be updated with two rows:
      - Row 1: ["Hello", "World"]
      - Row 2: ["Foo", "Bar"]
    - The range "C1:D2" will be updated with:
      - Row 1: ["Test", "Value"]
      - Row 2: ["Example", "Update"]

    The values in each list will be placed in the corresponding cells of the given range.

    :param value_input_option: The input option for how the data should be entered. Default is "USER_ENTERED", which means the values will be entered as if a user typed them. Other possible options include "RAW" (values are entered as-is, without any formatting or interpretation).

    :return: A dictionary with the result of the batch update, or None if there was an error. The dictionary contains the number of updated cells.
    """


    result = mass_edit_spreadsheet(spreadsheet_id, ranges_and_values)
    if result is None:
        return "Unable to edit spreadsheet"
    return f"{result.get('totalUpdatedCells')} cells updated."


if __name__ == "__main__":
    mcp.run(transport='stdio')
