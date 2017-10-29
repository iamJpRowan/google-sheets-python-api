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
            self.update_rows('A1', self.headers)
            self.spreadsheet.sheets[self.title] = self
        self.id_ = self.properties['sheetId']

        # Attach the Sheet object to the SpreadSheet
        self.spreadsheet.sheet_ids.append(self.id_)
        self.spreadsheet.sheet_titles.append(self.title)
        self.link = '{}/edit#gid={}'.format(self.spreadsheet.link, self.id_)


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

    def update_rows(self, start, rows, valueInputOption='USER_ENTERED'):
        range_ = self.determine_range(start, rows)
        value_range_body = {
            "range": range_,
            "majorDimension": 'ROWS',
            "values": rows,
        }
        return self.spreadsheet.service.spreadsheets().values().update(
            spreadsheetId=self.spreadsheet.id_, range=range_, body=value_range_body,
            valueInputOption=valueInputOption).execute()

    def determine_range(self, start, rows):
        start_column = self.a1[''.join(re.findall("[a-zA-Z]+", start))]
        end_column = self.a1[start_column + max(len(l) for l in rows)]
        end_row = int(re.findall('\d+', start)[0]) + len(rows)
        end = '{}{}'.format(end_column, end_row)
        return "'{}'!{}:{}".format(self.title, start, end)
