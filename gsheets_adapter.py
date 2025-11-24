import os.path
import re

import google.auth.transport.requests
import google.oauth2.credentials
import google_auth_oauthlib.flow
import googleapiclient.discovery
import googleapiclient.errors




class GsheetsAdapter:
    SPREADSHEET_ID = "12BykWguB9piR9RCTkXYsRM2427dsYrZyA75bI978IBo"
    SECRET_DIR = os.path.join(os.path.dirname(__file__), "secret")
    CREDENTIALS_PATH = os.path.join(SECRET_DIR, "credentials.json")
    TOKEN_PATH = os.path.join(SECRET_DIR, "token.json")
    USER_ENTERED = "USER_ENTERED"

    def __init__(self):
        self._creds = None
        self._service = None

    def get_values(self, range_address):
        return (
            self.spreadsheet_service.values()
            .get(spreadsheetId=self.SPREADSHEET_ID, range=range_address)
            .execute()
        )

    def set_values(self, values, range_address):
        return (
            self.spreadsheet_service
            .values()
            .update(
                spreadsheetId=self.SPREADSHEET_ID,
                valueInputOption=self.USER_ENTERED,
                range=range_address,
                body={"values": values},
            )
            .execute()
        )

    def append_values(self, values, range_address):
        service = self.spreadsheet_service
        return (
            service.spreadsheets()
            .values()
            .append(
                spreadsheetId=self.SPREADSHEET_ID,
                valueInputOption=self.USER_ENTERED,
                range=range_address,
                body={"values": values},
            )
            .execute()
        )

    def create_sheet(self, title: str):
        request_body = {"requests": [{
            "addSheet": {
                "properties": {"title": title},
            }
        }]}
        return self.spreadsheet_service.batchUpdate(spreadsheetId=self.SPREADSHEET_ID,
                                                    body=request_body).execute()
    def find_or_create_sheet(self, title: str):
        try:
            return self.get_sheet_by_title(title)
        except LookupError:
            return self.create_sheet(title)

    def get_sheets(self):
        spreadsheet_response = self.spreadsheet_service.get(
            spreadsheetId=self.SPREADSHEET_ID).execute()
        return spreadsheet_response["sheets"]

    def get_sheet_by_title(self, title: str):
        sheets_list = self.get_sheets()
        with_title = [s for s in sheets_list if s["properties"]["title"] == title]
        if len(with_title) ==0:
            raise LookupError
        if len(with_title) != 1:
            raise ValueError(str(sheets_list))

    def delete_sheet(self, sheet_id: str):
        request_body = {"requests": [{
            "deleteSheetRequest": {
                "properties": {"sheetId": sheet_id},
            }
        }]}
        return self.spreadsheet_service.batchUpdate(spreadsheetId=self.SPREADSHEET_ID,
                                                    body=request_body).execute()

    @property
    def spreadsheet_service(self) -> googleapiclient.discovery.Resource:
        if self._service is None:
            self._service = googleapiclient.discovery.build("sheets", "v4",
                                               credentials=self.credentials).spreadsheets()
        return self._service

    @property
    def credentials(self) -> google.oauth2.credentials.Credentials:
        # If modifying these scopes, delete the file token.json.
        scopes = ["https://www.googleapis.com/auth/spreadsheets"]

        # The file token.json stores the user's access and refresh tokens, and is
        # created automatically when the authorization flow completes for the first
        # time
        if self._creds is None and os.path.exists(self.TOKEN_PATH):
            self._creds = google.oauth2.credentials.Credentials.from_authorized_user_file(
                self.TOKEN_PATH,
                scopes)
        # If there are no (valid) credentials available, let the user log in.
        if not self._creds or not self._creds.valid:
            if self._creds and self._creds.expired and self._creds.refresh_token:
                self._creds.refresh(google.auth.transport.requests.Request())
            else:
                flow = google_auth_oauthlib.flow.InstalledAppFlow.from_client_secrets_file(
                    self.CREDENTIALS_PATH, scopes
                )
                self._creds = flow.run_local_server(port=0)
            # Save the credentials for the next run
            with open(self.TOKEN_PATH, "w") as token:
                token.write(self._creds.to_json())
        return self._creds

    @staticmethod
    def column_number(letters: str) -> int:
        if len(letters) == 0 or len(letters) > 3:
            raise ValueError(
                f"letter size must be between 1 to 3 letters, given='{letters}' size={len(letters)}")

        number = 0
        for l in letters.upper():
            n = ord(l) + 1 - ord("A")
            number = number * 26 + n
        return number

    @staticmethod
    def column_letter(number: int) -> str:
        if number < 1 or number > 18278:
            raise ValueError(f"number must be between 1 and 18278 (inclusive), given={number}")

        letters = ""
        while number > 0:
            digit = (number - 1) % 26 + 1
            number = (number - 1) // 26
            letters = chr(ord("A") - 1 + digit) + letters
        return letters

    @staticmethod
    def color_hex_to_dict(color_hex: str) -> dict[str, float]:
        color_hex = color_hex.lower()
        keys = ("red", "green", "blue")
        color_dict = {}
        for i in range(3):
            key = keys[i]
            hex_idx = i * 2
            byte_hex = color_hex[hex_idx:hex_idx + 2]
            float_value = int(byte_hex, 16) / 255.0
            color_dict[key] = float_value
        return color_dict

    @staticmethod
    def color_dict_to_hex(color_dict: dict[str, float]) -> str:
        keys = ("red", "green", "blue")
        color_value = 0
        for k in keys:
            v = color_dict[k]
            h = int(v * 255.0)
            color_value = color_value * 256 + h
        hex_color = hex(color_value)[2:]
        if len(hex_color) < 6:
            hex_color = "0" * (6 - len(hex_color)) + hex_color
        return hex_color

    @classmethod
    def address(cls, row: int, column: int):
        return cls.column_letter(column) + str(row)

    @classmethod
    def range_address(cls, sheet_title: str, start_row:int, start_col:int, end_row:int, end_col:int):
        start_address = cls.address(start_row, start_col)
        end_address = cls.address(end_row, end_col)
        return f"{sheet_title}!{start_address}:{end_address}"


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
