import re


class Sheet:

    def __init__(self, spreadsheet, title=None, properties=None, headers=[[]]):
        self.spreadsheet = spreadsheet
        self.headers = headers
        self.a1 = self.spreadsheet.a1

        # Create the sheet if it doesn't already exist
        if properties:
            self.properties = properties
            self.title = self.properties['title']
        else:
            self.title = title
            self.properties = self.spreadsheet.batch_update({
                "addSheet": {
                    "properties": self.create_properties()
                }
            })['replies'][0]['addSheet']['properties']
            # Write the header rows
            self.update_values('A1', self.headers)
            # Add new sheet object to Spreadsheet
            self.spreadsheet.sheets[self.title] = self

        # Add more attributes to sheet object
        self.id_ = self.properties['sheetId']
        self.link = '{}/edit#gid={}'.format(self.spreadsheet.link, self.id_)

    def columnCount(self):
        return self.properties['gridProperties'].get('columnCount')

    def frozenColumnCount(self):
        return self.properties['gridProperties'].get('frozenColumnCount')

    def frozenRowCount(self):
        return self.properties['gridProperties'].get('frozenRowCount')

    def rowCount(self):
        return self.properties['gridProperties'].get('rowCount')

    def add_dimension(self, amount, dimension):
        request = {
            "appendDimension": {
                "sheetId": self.id_,
                "dimension": dimension,
                "length": amount
            }
        }
        self.spreadsheet .batch_update([request])

    def create_properties(self):
        return {
            "gridProperties": {
                "columnCount": max(len(l) for l in self.headers),
                "frozenRowCount": len(self.headers),
                "rowCount": len(self.headers) + 1
            },
            "index": 1,
            "sheetType": "GRID",
            "title": self.title
        }

    def update_values(self, start, values, dimension='ROWS', valueInputOption='USER_ENTERED'):
        """ Updates the sheet with new array of values

        PARAMETERS:
            start: the 'A1' style position to start from
            dimension: Should be either ROWS or COLUMNS
            values: An array of array to be used as either columns or rows"""
        range_ = self.determine_range(start, values)
        value_range_body = {
            "range": range_,
            "majorDimension": dimension,
            "values": values,
        }
        return self.spreadsheet.service.spreadsheets().values().update(
            spreadsheetId=self.spreadsheet.id_, range=range_, body=value_range_body,
            valueInputOption=valueInputOption).execute()

    def refresh_properties(self):
        response = self.spreadsheet.service.spreadsheets().get(
            spreadsheetId=self.spreadsheet.id_).execute()
        for sheet in response['sheets']:
            if sheet['properties']['sheetId'] == self.id_:
                self.properties = (sheet['properties'])

    def delete_values(self, start_index=None, end_index=None, dimension='ROWS'):
        value_range_body = {
            "sheetId": self.id_,
            "dimension": dimension,
            "startIndex": start_index,
            "endIndex": end_index,
        }
        requests = [{"deleteDimension": {"range": value_range_body}}]
        return self.spreadsheet.batch_update(requests)

    def determine_range(self, start, rows):
        start_column = self.a1[''.join(re.findall("[a-zA-Z]+", start))]
        end_column = self.a1[start_column + max(len(l) for l in rows)]
        end_row = int(re.findall('\d+', start)[0]) + len(rows)
        end = '{}{}'.format(end_column, end_row)
        return "'{}'!{}:{}".format(self.title, start, end)
