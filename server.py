from mcp.server.fastmcp import FastMCP
from sheets_functions import get_user_spreadsheet_ids

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


if __name__ == "__main__":
    mcp.run(transport='stdio')