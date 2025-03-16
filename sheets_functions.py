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


if __name__ == "__main__":
    # List of tuples: (range_name, values)
    ranges_and_values = [
        ("A1:B2", [["Hello", "World"], ["Foo", "Bar"]]),
        ("C1:D2", [["Test", "Value"], ["Example", "Update"]]),
    ]
    # Call mass_edit_spreadsheet function to batch update values
    result = mass_edit_spreadsheet(
        "1RDoRbpYKZflWJCRPT72dnnmrvfFK4stx1eaR38XKQag",  # Your spreadsheet ID
        ranges_and_values  # Ranges and their values to be updated
    )
    print(result)
