

# get the instance of the Spreadsheet
sheet = client.open('Holiday Form (Responses)')
print(sheet.id)

# get the first sheet of the Spreadsheet
sheet_instance = sheet.get_worksheet(0)

raw = client.export(sheet.id, ExportFormat.CSV)
print(raw)


def get_last():
    timestamp = sheet_instance.cell(sheet_instance.row_count, 1).value
    entryid = sheet_instance.cell(sheet_instance.row_count, 2).value
    applicant = sheet_instance.cell(sheet_instance.row_count, 3).value
    return f'The last entry was added at {timestamp} by {applicant} (ID = {entryid})'


print(get_last())
