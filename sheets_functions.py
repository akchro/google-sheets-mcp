import os.path
from email.headerregistry import Group
from typing import List

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# If modifying these scopes, delete the file token.json.
GDRIVE_SCOPES = ["https://www.googleapis.com/auth/drive.readonly"]

def get_user_spreadsheet_ids() -> List[str]  | None:
  creds = None
  if os.path.exists("token.json"):
    creds = Credentials.from_authorized_user_file("token.json", GDRIVE_SCOPES)
  if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
      creds.refresh(Request())
    else:
      flow = InstalledAppFlow.from_client_secrets_file("credentials.json", GDRIVE_SCOPES)
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


if __name__ == "__main__":
  spreadsheets = get_user_spreadsheet_ids()
  print(spreadsheets)
