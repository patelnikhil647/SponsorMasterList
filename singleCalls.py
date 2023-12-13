import gspread
from oauth2client.service_account import ServiceAccountCredentials
from pprint import pprint

scope = ["https://spreadsheets.google.com/feeds",
         "https://www.googleapis.com/auth/spreadsheets",
         "https://www.googleapis.com/auth/drive.file",
         "https://www.googleapis.com/auth/drive"]

creds = ServiceAccountCredentials.from_json_keyfile_name("Sponsors_and_Partners_Secret.json", scope)
client = gspread.authorize(creds)

sheet = client.open("Sponsorship/Partnership Master List 2020/2021").sheet1

#THIS WILL ONLY SORT THE ORGANIZATION NAMES, NOT THE DATA ATTACHED
def orgSort():
    orgs = sheet.col_values(1)[2:]
    cell_list = sheet.range("A3:A37")
    pprint(cell_list)
    orgs.sort()
    i = 0
    for cell in cell_list:
        cell.value = orgs[i]
        i += 1
    pprint(cell_list)
    sheet.update_cells(cell_list)

#orgSort()