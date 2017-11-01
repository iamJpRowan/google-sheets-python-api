import string

from .sheet import Sheet
from .api import init_service


class Spreadsheet:

    def __init__(self, spreadsheet_id=None):
        self.service = init_service()

        # TODO: Add a method to create sheet if it doesn't already exist
        if not spreadsheet_id:
            # return spreadsheet_id
            pass

        self.id_ = spreadsheet_id
        self.link = 'https://docs.google.com/spreadsheets/d/{}'.format(
            self.id_)
        # Get the full sheet object from the API
        self.properties = self.service.spreadsheets().get(
            spreadsheetId=self.id_).execute()

        # Create A1 dictionary
        self.a1 = {}
        for num1, ltr1 in enumerate(string.ascii_uppercase, 1):
            self.a1[num1] = ltr1
            self.a1[ltr1] = num1
            for num2, ltr2 in enumerate(string.ascii_uppercase, 1):
                self.a1['{}{}'.format(ltr1, ltr2)] = num2 + (num1 * 26)
                self.a1[num2 + (num1 * 26)] = '{}{}'.format(ltr1, ltr2)

        # Create a sheet object from any existing sheets
        self.sheets = {sheet['properties']['title']: Sheet(self, properties=sheet['properties'])
                       for sheet in self.properties['sheets']}

    def batch_update(self, requests):
        return self.service.spreadsheets().batchUpdate(spreadsheetId=self.id_,
                                                       body={"requests": requests}).execute()

    def sheet_ids(self):
        return [sheet.id_ for title, sheet in self.sheets.items()]

    def sheet_titles(self):
        return [title for title in self.sheets.keys()]

    def delete_all_sheets(self, protected_sheets=[]):
        if not protected_sheets:
            self.create_sheet('stub')
            protected_sheets=['stub']
        sheets = [
            sheet for title, sheet in self.sheets.items() if title not in protected_sheets]
        requests = [{"deleteSheet": {"sheetId": sheet.id_}} for sheet in sheets]
        self.sheets = {title: sheet for title, sheet in self.sheets.items() if title in protected_sheets}
        if requests:
            return self.batch_update(requests)

    def create_sheet(self, title, headers=[[]]):
        return Sheet(self, title=title, headers=headers)

    def delete_sheet(self, title):
        self.batch_update([{"deleteSheet": {"sheetId": self.sheets[title].id_}}])
        self.sheets.remove(title)
