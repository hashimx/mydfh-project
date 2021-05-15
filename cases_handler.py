# Script to update the G sheet using Google API's

import pandas as pd
import gspread
from tabulate import tabulate
from oauth2client.service_account import ServiceAccountCredentials
import traceback

# define the scope
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']


class Cases:
    def __init__(self):
        """
        Initialization/ Constructor
        """
        # Add credentials to the account
        self.creds = ServiceAccountCredentials. \
            from_json_keyfile_name(r'dfh-project.json', scope)
        # Authorize the clientsheet
        self.client = gspread.authorize(self.creds)
        # Get the instance of the Spreadsheet
        self.sheet = self.client.open('my_Sheet')

    def check_if_open_cases_in_closed_sheet(self) -> None:
        """
        Method to determine if the open cases are existing in closed sheet. Uses Google API's and Pandas for processing

        1) Read given google sheet
        2) Check the opened cases, get unique values
        3) Check the closed cases, get unique values
        4) Compare Values based on the retrieved keys
        5) Print if matching Values are found

        :return: None
        """
        # Sheet indexes, can vary based on sheet
        opened, closed = 1, 2
        open_index_string , close_index_string = "Please check UnAssigned sheet as well", "SRFID"

        # Get all open Cases
        open_cases = self.sheet.get_worksheet(opened)
        # Get all closed Cases
        closed_cases = self.sheet.get_worksheet(closed)

        # Convert to Data Frames
        open_records_df = pd.DataFrame.from_dict(open_cases.get_all_records())
        closed_records_df = pd.DataFrame.from_dict(closed_cases.get_all_records())
        srfid_open = open_records_df[open_index_string].unique()
        srfid_closed = closed_records_df[close_index_string].unique()
        for item in srfid_open:
            if isinstance(item, int):
                if item in srfid_closed:
                    print(":Duplicate Found:")
                    a = open_records_df.loc[open_records_df[open_index_string] == item]
                    print(tabulate(a))


if __name__ == '__main__':
    try:
        cases = Cases()
        cases.check_if_open_cases_in_closed_sheet()
    except Exception as e:
        error_msg = traceback.format_exc()
        raise Exception(f'Exception >>>>>>>>>> {error_msg}')
