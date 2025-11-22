
import os.path

import google.auth.transport.requests
import google.oauth2.credentials
import google_auth_oauthlib.flow
import googleapiclient.discovery
from googleapiclient.errors import HttpError

# If modifying these scopes, delete the file token.json.
SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]

# The ID and range of a sample spreadsheet.
SAMPLE_SPREADSHEET_ID = "12BykWguB9piR9RCTkXYsRM2427dsYrZyA75bI978IBo"
SAMPLE_RANGE_NAME = "Sheet1!A2:E"

cwd = os.path.dirname(__file__)
secret_dir = os.path.join(cwd, "secret")

def main():

  try:
    spread_sheet = get_service()

    # Call the Sheets API
    result = (
        spread_sheet.values()
        .get(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=SAMPLE_RANGE_NAME)
        .execute()
    )
    values = result.get("values", [])

    if not values:
      print("No data found.")
      return

    print("Name, Major:")
    for row in values:
      # Print columns A and E, which correspond to indices 0 and 4.
      print(f"{row[0]}, {row[4]}")
  except HttpError as err:
    print(err)


def get_service():
    return googleapiclient.discovery.build("sheets", "v4", credentials=get_creds()).spreadsheets()


def get_creds() -> google.oauth2.credentials.Credentials:
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists(token_path()):
        creds = google.oauth2.credentials.Credentials.from_authorized_user_file(token_path(), SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(google.auth.transport.requests.Request())
        else:
            flow = google_auth_oauthlib.flow.InstalledAppFlow.from_client_secrets_file(
                creds_path(), SCOPES
            )
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open(token_path(), "w") as token:
            token.write(creds.to_json())
    return creds


def creds_path() -> str:
    return os.path.join(secret_dir, "credentials.json")


def token_path() -> str:
    return os.path.join(secret_dir, "token.json")


if __name__ == "__main__":
  main()