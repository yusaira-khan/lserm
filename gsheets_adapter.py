import functools
import os.path

import google.auth.transport.requests
import google.oauth2.credentials
import google_auth_oauthlib.flow
import googleapiclient.discovery
import googleapiclient.errors


class GsheetsAdapter:
    # If modifying these scopes, delete the file token.json.
    SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

    SPREADSHEET_ID = "12BykWguB9piR9RCTkXYsRM2427dsYrZyA75bI978IBo"
    SECRET_DIR = os.path.join(os.path.dirname(__file__), "secret")
    CREDENTIALS_PATH = os.path.join(SECRET_DIR, "credentials.json")
    TOKEN_PATH = os.path.join(SECRET_DIR, "token.json")

    def __init__(self):
        self._creds = None

    def get_values(self, range_name):
        return (
            self.spreadsheet_service.values()
            .get(spreadsheetId=self.SPREADSHEET_ID, range=range_name)
            .execute()
        )

    def create_sheet(self, title):
        request_body = {"requests": [{
            "updateSpreadsheetProperties": {
                "properties": {"title": title},
                "fields": "title",
            }
        }]}
        self.spreadsheet_service.batchUpdate(spreadsheetId=self.SPREADSHEET_ID,
                                             body=request_body).execute()

    @property
    def spreadsheet_service(self):
        return googleapiclient.discovery.build("sheets", "v4",
                                               credentials=self.credentials).spreadsheets()

    @property
    def credentials(self) -> google.oauth2.credentials.Credentials:
        # The file token.json stores the user's access and refresh tokens, and is
        # created automatically when the authorization flow completes for the first
        # time.
        if self._creds is None and os.path.exists(self.TOKEN_PATH):
            self._creds = google.oauth2.credentials.Credentials.from_authorized_user_file(
                self.TOKEN_PATH,
                self.SCOPES)
        # If there are no (valid) credentials available, let the user log in.
        if not self._creds or not self._creds.valid:
            if self._creds and self._creds.expired and self._creds.refresh_token:
                self._creds.refresh(google.auth.transport.requests.Request())
            else:
                flow = google_auth_oauthlib.flow.InstalledAppFlow.from_client_secrets_file(
                    self.CREDENTIALS_PATH, self.SCOPES
                )
                self._creds = flow.run_local_server(port=0)
            # Save the credentials for the next run
            with open(self.TOKEN_PATH, "w") as token:
                token.write(self._creds.to_json())
        return self._creds

    @staticmethod
    def column_number(letters: str)->int:
        if len(letters) == 0 or len(letters) > 3:
            raise ValueError(f"letter size must be between 1 to 3 letters, given='{letters}' size={len(letters)}")

        number = 0
        for l in letters.upper():
            n = ord(l) + 1 - ord("A")
            number = number * 26 + n
        return number

    @staticmethod
    def column_letter(number: int)->str:
        if number < 1 or number > 18278:
            raise ValueError(f"number must be between 1 and 18278 (inclusive), given={number}")

        letters = ""
        while number > 0:
            digit = (number - 1) % 26 + 1
            number = (number - 1) // 26
            letters = chr(ord("A") - 1 + digit) + letters
        return letters


def main():
    values = GsheetsAdapter().get_values("Sheet1!A1").get("values", [])
    if not values:
        print("No data found.")
        return

    print("Name, Major:")
    for row in values:
        print(f"{row[0]}, {row[4]}")


if __name__ == "__main__":
    main()
