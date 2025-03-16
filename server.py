from mcp.server.fastmcp import FastMCP
from sheets_functions import get_user_spreadsheet_ids, create_new_spreadsheet, copy_to_spreadsheet, mass_edit_spreadsheet, fill_ranges_with_colors_in_spreadsheet

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
        Fills specified ranges in a spreadsheet with colors.

        :param spreadsheet_id: The ID of the spreadsheet to edit.

        :param ranges_and_colors:
            A list of tuples, where each tuple contains:
            1. A string representing the range to fill with a color (e.g., "A1:B2").
            2. A string representing the color to fill the range with, provided in hex format (e.g., "#FF5733").

            Example:
            ranges_and_colors = [
                ("A1:B2", "#FF5733"),
                ("C1:D2", "#4287f5"),
                ("E3:F3", "#FFC300"),
            ]

            In this example:
            - The range "A1:B2" will be filled with the color "#FF5733" (a shade of red).
            - The range "C1:D2" will be filled with the color "#4287f5" (a shade of blue).
            - The range "E3:F3" will be filled with the color "#FFC300" (a shade of yellow).

            The color is applied to the entire range of cells specified.

    """

    result = mass_edit_spreadsheet(spreadsheet_id, ranges_and_values)
    if result is None:
        return "Unable to edit spreadsheet"
    return f"{result.get('totalUpdatedCells')} cells updated."


@mcp.tool()
def fill_spreadsheet(spreadsheet_id, ranges_and_colors):
    """
    Fills specified ranges in a spreadsheet with colors.

    :param spreadsheet_id: The ID of the spreadsheet to edit.

    :param ranges_and_colors:
        A list of tuples, where each tuple contains:
        1. A string representing the range to fill with a color (e.g., "A1:B2").
        2. A string representing the color to fill the range with, provided in hex format (e.g., "#FF5733").

        Example:
        ranges_and_colors = [
            ("A1:B2", "#FF5733"),
            ("C1:D2", "#4287f5"),
            ("E3:F3", "#FFC300"),
        ]

        In this example:
        - The range "A1:B2" will be filled with the color "#FF5733" (a shade of red).
        - The range "C1:D2" will be filled with the color "#4287f5" (a shade of blue).
        - The range "E3:F3" will be filled with the color "#FFC300" (a shade of yellow).

        The color is applied to the entire range of cells specified.

    :param value_input_option:
        The input option for how the data should be entered. This parameter is not used in this function as it primarily deals with formatting and color filling, but it's included for consistency with Google Sheets API conventions.

    :return:
        A dictionary with the result of the batch update, or None if there was an error. The dictionary contains the number of updated cells.
"""

    result = fill_ranges_with_colors_in_spreadsheet(spreadsheet_id, ranges_and_colors)
    if result is None:
        return "Unable to fill spreadsheet"
    return f"Cells updated with colors."

if __name__ == "__main__":
    mcp.run(transport='stdio')
