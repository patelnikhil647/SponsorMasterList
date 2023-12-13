import gspread
from oauth2client.service_account import ServiceAccountCredentials #1st and 2nd imports are for Google API
from pprint import pprint #for Pretty Print, but not necessary

#scopes to help creds find file
scope = ["https://spreadsheets.google.com/feeds",
         "https://www.googleapis.com/auth/spreadsheets",
         "https://www.googleapis.com/auth/drive.file",
         "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("Sponsors_and_Partners_Secret.json", scope)
client = gspread.authorize(creds)

#sheet object
sheet = client.open("Sponsorship/Partnership Master List 2020/2021").sheet1

#new_row = ["this", "is", "a", "test", "brother"]
#sheet.insert_row(new_row, 4)
#sheet.update_cell(4,5, "bruv")
#sheet.delete_rows(31,32)

data = sheet.get_all_records()
row = sheet.row_values(2)
col = sheet.col_values(1)
cell = sheet.cell(2,5).value
pprint(sheet.row_values(4))

numRows = sheet.row_count
print(len(data)) #prints all rows that have data

cell_list = sheet.range("A3:C4") #gets a list of all the cells in that range
