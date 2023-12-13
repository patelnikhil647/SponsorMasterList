import main
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
from pprint import pprint

scope = ["https://spreadsheets.google.com/feeds",
         "https://www.googleapis.com/auth/spreadsheets",
         "https://www.googleapis.com/auth/drive.file",
         "https://www.googleapis.com/auth/drive"]

creds = ServiceAccountCredentials.from_json_keyfile_name("Sponsors_and_Partners_Secret.json", scope)
client = gspread.authorize(creds)

wksheet = client.open("Sponsorship/Partnership Master List 2020/2021")
sheet1 = client.open("Sponsorship/Partnership Master List 2020/2021").sheet1
sheet2 = client.open("Sponsorship/Partnership Master List 2020/2021").worksheet("Last Save")

def addWhileSorting(letter = "p"):
    li = ["b", "c", "R", "f", "a", "y", "l"]
    print(li)
    li.sort(key=str.lower)
    print(li)
    li.append(letter)
    print(li)
    li.sort(key=str.lower)
    print(li)

    for i in range(len(li)):
        if li[i] == letter:
            index = i + 1
    print(index)

def typeOfBlankCells(r=3, c=4):
    string = sheet1.cell(r,c)
    print(string)
    print(type(string))
    print(str(string))

def sortingCells(): #CANT SORT CELLS
    cells = sheet2.range("A1:A12")
    print(cells)
    cells.sort()
    print(cells)

def usingFindFunction():
    row = sheet2.find("h", in_column=1)
    print(row)

def compareTwoSheets():
    sheet1_data = sheet1.get_all_values()
    pprint(sheet1_data)
    print(len(sheet1_data))
    sheet2_data = sheet2.get_all_values()
    pprint(sheet2_data)
    print(len(sheet2_data))
    sheet2_data = sheet2_data[:-2]
    pprint(sheet2_data)
    print(len(sheet2_data))
    if sheet1_data == sheet2_data:
        print(True)
    else:
        print(False)


def testSave():
    title = "Last Save"
    save_sheet_temp = client.open("Sponsorship/Partnership Master List 2020/2021").worksheet(title)
    wksheet.del_worksheet(save_sheet_temp)
    del save_sheet_temp
    info = sheet1.copy_to("1nfDBqNArdear20SrnIVD-2vIy0YPXdW7CrwG26p2iSY")
    save_sheet = client.open("Sponsorship/Partnership Master List 2020/2021").worksheet(info.get("title"))
    info["title"] = title
    save_sheet.update_title(info.get("title"))
    save_sheet.insert_row(["Last Updated: " + str(datetime.now())[:19]], 3 + len(save_sheet.get_all_records()))
    save_sheet.format("A" + str(1 + len(save_sheet.get_all_records())), {"textFormat": {"bold": True}})
    return info


"""
addWhileSorting()
typeOfBlankCells()
sortingCells()
main.addOrg("Senor Thickems", "wizardog@yahoo.com", "LA; Vancouver", "Silver", "$800", "Somewhat Publicity", "6")
main.addOrgs(["Apple", "Senor Thickems"], ["sjobs@gmail.com", "wizardog@yahoo.com"],
             ["Vancouver", "LA; Vancouver"], ["Platinum", "Silver"], ["A frick ton", "$800"],
             ["Massive Publicity", "Somewhat Publicity"], ["Steve Jobs", ""], ["CEO", ""], ["Kaden", ""],
             ["Lord", "Sr."])
main.addOrgs(["Apple", "Senor Thickems", "Orange"], ["sjobs@gmail.com", "wizardog@yahoo.com", "MJackson@gmail.com"],
             ["Vancouver", "LA; Vancouver", "LA"], ["Platinum", "Silver", "Gold"],
             ["A frick ton", "$800", "Xbox Consoles"], ["Massive Publicity", "Somewhat Publicity", "Decent Publicity"],
             ["Steve Jobs", "", "Michael Jackson"], ["CEO", "", "Marketing Head"], ["Kaden", "", "Nikhil"],
             ["Lord", "Sr.", "King of Pop"])
usingFindFunction()
"""

