import os.path
from typing import List, Any

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# If modifying these scopes, delete the file token.json.
SCOPES = ["https://www.googleapis.com/auth/drive.readonly", "https://www.googleapis.com/auth/spreadsheets"]


def get_user_spreadsheet_ids() -> List[str] | None:
    creds = None
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
            creds = flow.run_local_server(port=0)
        with open("token.json", "w") as token:
            token.write(creds.to_json())

    try:
        service = build("drive", "v3", credentials=creds)
        results = service.files().list(q="mimeType='application/vnd.google-apps.spreadsheet'",
                                       fields="files(id, name)").execute()
        files = results.get("files", [])

        if not files:
            print("No spreadsheets found.")
            return []

        return [f"{file["name"]}: {file["id"]}" for file in files]

    except HttpError as err:
        print(f"An error occurred: {err}")
        return None


def create_new_spreadsheet(title: str) -> tuple[str, str] | None:
    creds = None
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
            creds = flow.run_local_server(port=0)
        with open("token.json", "w") as token:
            token.write(creds.to_json())

    try:
        service = build("sheets", "v4", credentials=creds)
        spreadsheet = {"properties": {"title": title}}
        spreadsheet = (
            service.spreadsheets()
            .create(body=spreadsheet, fields="spreadsheetId")
            .execute()
        )
        print(f"Spreadsheet ID: {(spreadsheet.get('spreadsheetId'))}")
        return (title, spreadsheet.get("spreadsheetId"))
    except HttpError as error:
        print(f"An error occurred: {error}")
        return None


def copy_to_spreadsheet(sheet_id_to: str, sheet_id_from: str) -> tuple[str, str, str] | None:
    creds = None
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
            creds = flow.run_local_server(port=0)
        with open("token.json", "w") as token:
            token.write(creds.to_json())

    try:
        service = build("sheets", "v4", credentials=creds)

        # Get the name of the source spreadsheet and the sheet to copy
        source_spreadsheet = service.spreadsheets().get(spreadsheetId=sheet_id_from).execute()
        source_spreadsheet_name = source_spreadsheet["properties"]["title"]
        sheet_id_from_actual = source_spreadsheet["sheets"][0]["properties"]["sheetId"]
        sheet_name_from = source_spreadsheet["sheets"][0]["properties"]["title"]

        # Get the name of the destination spreadsheet
        destination_spreadsheet = service.spreadsheets().get(spreadsheetId=sheet_id_to).execute()
        destination_spreadsheet_name = destination_spreadsheet["properties"]["title"]

        # Copy the sheet to the destination
        copied_sheet = service.spreadsheets().sheets().copyTo(
            spreadsheetId=sheet_id_from,
            sheetId=sheet_id_from_actual,
            body={"destinationSpreadsheetId": sheet_id_to}
        ).execute()

        # Return the relevant names
        return (
            source_spreadsheet_name,
            destination_spreadsheet_name,
            copied_sheet["title"]
        )

    except HttpError as error:
        print(f"An error occurred: {error}")
        return None


def mass_edit_spreadsheet(spreadsheet_id: str, ranges_and_values: List[tuple[str, List[List[Any]]]],
                          value_input_option: str = "USER_ENTERED") -> dict | None:
    creds = None
    # Check if the token.json file exists and load credentials from it
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)

    # If credentials are not valid, refresh or authenticate again
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
            creds = flow.run_local_server(port=0)
        with open("token.json", "w") as token:
            token.write(creds.to_json())

    try:
        service = build("sheets", "v4", credentials=creds)

        # Prepare the request body for batch update
        data = [
            {"range": range_name, "values": values}
            for range_name, values in ranges_and_values
        ]

        body = {
            "valueInputOption": value_input_option,
            "data": data
        }

        # Execute the batch update request
        result = service.spreadsheets().values().batchUpdate(
            spreadsheetId=spreadsheet_id, body=body
        ).execute()

        # Return the result which contains the number of updated cells
        print(f"{result.get('totalUpdatedCells')} cells updated.")
        return result

    except HttpError as error:
        print(f"An error occurred: {error}")
        return None


def hex_to_rgb(hex_color: str) -> dict:
    """
    Converts a hex color code (e.g., "#FF5733") to RGB format for Google Sheets API.

    :param hex_color: The hex color code (e.g., "#FF5733").
    :return: A dictionary with the RGB values for Google Sheets backgroundColor (e.g., {"red": 1.0, "green": 0.341, "blue": 0.2}).
    """
    hex_color = hex_color.lstrip('#')  # Remove leading '#' if present
    r, g, b = bytes.fromhex(hex_color)  # Convert hex to RGB
    return {
        "red": r / 255,
        "green": g / 255,
        "blue": b / 255
    }


def fill_ranges_with_colors_in_spreadsheet(spreadsheet_id: str, ranges_and_colors: List[tuple[str, str]],
                                           value_input_option: str = "USER_ENTERED") -> dict | None:

    creds = None
    # Check if the token.json file exists and load credentials from it
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)

    # If credentials are not valid, refresh or authenticate again
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
            creds = flow.run_local_server(port=0)
        with open("token.json", "w") as token:
            token.write(creds.to_json())

    try:
        service = build("sheets", "v4", credentials=creds)

        # Prepare the requests for filling colors
        requests = []

        for range_to_fill, hex_color in ranges_and_colors:
            # Convert hex color to RGB format
            rgb_color = hex_to_rgb(hex_color)

            # Parse the range (e.g., "A1:B2")
            range_parts = range_to_fill.split(":")
            start_cell = range_parts[0]
            end_cell = range_parts[1]

            # Calculate the row and column indices
            start_row = int(start_cell[1:]) - 1  # Convert to 0-indexed
            start_col = ord(start_cell[0].upper()) - ord('A')  # Convert column letter to index
            end_row = int(end_cell[1:])  # Exclusive, so we do not subtract 1 here
            end_col = ord(end_cell[0].upper()) - ord('A') + 1  # To include the last column

            # Prepare the cells to fill within the specified range
            rows = []
            for row_index in range(start_row, end_row):
                row_values = []
                for col_index in range(start_col, end_col):
                    row_values.append({
                        "userEnteredFormat": {
                            "backgroundColor": rgb_color
                        }
                    })
                rows.append({"values": row_values})

            # Create the updateCells request
            requests.append({
                "updateCells": {
                    "range": {
                        "sheetId": 0,  # Assuming sheetId 0, update with correct sheetId if needed
                        "startRowIndex": start_row,
                        "endRowIndex": end_row,
                        "startColumnIndex": start_col,
                        "endColumnIndex": end_col
                    },
                    "rows": rows,
                    "fields": "userEnteredFormat.backgroundColor"
                }
            })

        # Execute the batch update request
        body = {"requests": requests}
        result = service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()

        # Return the result which contains the number of updated cells
        print(f"Cells updated with colors.")
        return result

    except HttpError as error:
        print(f"An error occurred: {error}")
        return None


if __name__ == "__main__":
    ranges_and_colors = [
        ("A1:B2", "#FF5733"),  # Hex color for red-orange
        ("C3:D4", "#FFFF00"),  # Hex color for yellow
        ("E5:F6", "#00FF00"),  # Hex color for green
        ("G7:H8", "#800080")  # Hex color for purple
    ]
    fill_ranges_with_colors_in_spreadsheet(
        spreadsheet_id="1bsmwdHv5wcxPbHnex4RGkTrwUoR-vl10gdO-L5Ohlrg", ranges_and_colors=ranges_and_colors
    )

