import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import re
from pprint import pprint

scope = ["https://spreadsheets.google.com/feeds",
         "https://www.googleapis.com/auth/spreadsheets",
         "https://www.googleapis.com/auth/drive.file",
         "https://www.googleapis.com/auth/drive"]

creds = ServiceAccountCredentials.from_json_keyfile_name("Sponsors_and_Partners_Secret.json", scope)
client = gspread.authorize(creds)

wksheet = client.open("Sponsorship/Partnership Master List 2020/2021")
sheet = client.open("Sponsorship/Partnership Master List 2020/2021").sheet1
changes_sheet = client.open("Sponsorship/Partnership Master List 2020/2021").worksheet("Changes")
save_sheet = client.open("Sponsorship/Partnership Master List 2020/2021").worksheet("Last Save")


# # #

def addToChanges(isAdd, r, c, changed):
    """

    :type isAdd: boolean
    :type r: int
    :type c: str
    :type changed: str
    """
    aos = "+" if isAdd else "-"
    newUpdate = [aos, r, c, changed, str(datetime.now())[:19]]
    changes_sheet.insert_row(newUpdate, 2)

    if changes_sheet.row_values(102) is not None:
        changes_sheet.delete_rows(102)


def findOrgRow(orgName, sht=sheet):
    """

    :type orgName: str
    :type sht: object

    :returns: row number given organization name
    :rtype: int
    """
    allOrgs = sht.col_values(1)
    i = 1
    for aorg in allOrgs:
        if orgName.lower() == aorg.lower():
            return i
        i += 1


def save():
    sheet_values = sheet.get_all_values()
    save_sheet.delete_rows(1, len(save_sheet.get_all_values()))
    save_sheet.insert_rows(sheet_values)
    save_sheet.format("A1:J1", {"textFormat": {"bold": True}})
    now = str(datetime.now())[:19]
    save_sheet.insert_rows([["Last Saved: " + now], ["Last Compared: " + now, True]], 2 + len(save_sheet.get_all_values()))
    save_sheet.format("A{0}:A{1}".format(len(save_sheet.get_all_values()) - 1, len(save_sheet.get_all_values())),
                      {"textFormat": {"bold": True}})


def timeLastSaved():
    """

    :returns: date and time masterlist was last saved
    :rtype: str
    """
    date = save_sheet.get_all_values()[-2][0]
    return date


def compareLastSaved():
    """

    :returns: if Masterlist and Last Saved are the same, return True, else return a dictionary with information on
    what is different
    :rtype: bool or dict
    """
    sheet_values = sheet.get_all_values()
    save_sheet_values = save_sheet.get_all_values()[:-3]
    save_sheet.update_cell(len(save_sheet.get_all_values()), 1, "Last Compared: " + str(datetime.now())[:19])
    if sheet_values == save_sheet_values:
        save_sheet.update_cell(len(save_sheet.get_all_values()), 2, True)
        return True
    else:
        different = {"Masterlist": [], "Save List": []}
        diff = [i for i in sheet_values + save_sheet_values if i not in sheet_values or i not in save_sheet_values]
        for i in diff:
            if i in sheet_values:
                row = findOrgRow(i[0])
                individual_different = (row, i[0])  # tuple: (row, orgName)
                different["Masterlist"].append(individual_different)
            else:
                row = findOrgRow(i[0], save_sheet)
                individual_different = (row, i[0])
                different["Save List"].append(individual_different)
        save_sheet.update_cell(len(save_sheet.get_all_values()), 2, str(different))
        for i in different.get("Masterlist"):
            print(sheet.row_values(i[0]))
        print("-")
        for i in different.get("Save List"):
            print(save_sheet.row_values(i[0]))
        return different


def comparedData():
    """

    :returns: timestamp and what's different in last compared
    :rtype: list
    """
    data = save_sheet.get_all_values()[-1][:2]
    return data


def restore():
    rows_to_restore = save_sheet.get("A3:J" + str(len(save_sheet.get_all_values()) - 3))
    sheet.delete_rows(3, len(sheet.get_all_values()))
    sheet.insert_rows(rows_to_restore, 3)
    save_sheet.update_cell(len(save_sheet.get_all_values()), 2, "RESTORED")


def addOrg(orgName, email, rounds, level, given, gaveBack, contact='', position='', connection='', honorific=''):
    """

    :type orgName: str
    :type email: str
    :type rounds: str
    :type level: str
    :type given: str
    :type gaveBack: str
    :type contact: str
    :type position: str
    :type connection: str
    :type honorific: str
    """
    newRow = [orgName, contact, email, position, connection, rounds, honorific, level, given, gaveBack]
    orgs = sheet.col_values(1)[2:]
    orgs.append(orgName)
    orgs.sort(key=str.lower)

    for i in range(len(orgs)):
        if orgs[i] == orgName:
            index = i
            rowNum = index + 1 + 2
            # the +1 is to align index (which starts at 0) to rows (which starts at 1)
            # the +2 is to add two rows because the first two rows aren't supposed to be messed with
            break

    sheet.insert_row(newRow, rowNum)
    addToChanges(True, rowNum, "A-J", "; ".join(newRow))


def addOrgs(orgNames, emails, roundss, levels, givens, gaveBacks, contacts, positions, connections, honorifics):
    """

    :type orgNames: list of strings
    :type emails: list of strings
    :type roundss: list of strings
    :type levels: list of strings
    :type givens: list of strings
    :type gaveBacks: list of strings
    :type contacts: list of strings
    :type positions: list of strings
    :type connections: list of strings
    :type honorifics: list of strings

    :raises: :class: "AssertionError": All args (type: list) must be of equal size
    """
    assert len(orgNames) == len(emails) == len(roundss) == len(levels) == len(givens) == len(gaveBacks) == len(
        contacts) == len(positions) == len(connections) == len(
        honorifics), "All args (type: list) must be of equal size"
    for i in range(len(orgNames)):
        addOrg(orgNames[i], emails[i], roundss[i], levels[i], givens[i], gaveBacks[i], contacts[i], positions[i],
               connections[i], honorifics[i])


def delOrgs(rows=None, orgs=None):
    """

    :type rows: list of ints
    :type orgs: list of strings

    :raises: :class: "TypeError": "orgs" argument must be of type list
    :raises: :class: "TypeError": "rows" argument must be of type list
    :raises: :class: "ValueError": Both args cannot be of type None
    """
    if orgs is not None:
        if type(orgs) is not list:
            raise TypeError("\"orgs\" argument must be of type list")
        rowsToDelete = []
        for org in orgs:
            rowsToDelete.append(findOrgRow(org))
        for c in range(rowsToDelete.count(None)):
            rowsToDelete.remove(None)
        rowsToDelete.sort(reverse=True)
        for rowToDelete in rowsToDelete:
            addToChanges(False, rowToDelete, "A-J", "; ".join(sheet.row_values(rowToDelete)))
            sheet.delete_rows(rowToDelete)
        if len(rowsToDelete) < len(orgs):
            print("One or more organizations originally named has not been deleted, check spelling")
    elif rows is not None:
        if type(rows) is not list:
            raise TypeError("\"rows\" argument must be of type list")
        rows.sort(reverse=True)
        for row in rows:
            addToChanges(False, row, "A-J", "; ".join(sheet.row_values(row)))
            sheet.delete_rows(row)
    else:
        raise ValueError("Both args cannot be of type None")


def updateRounds(orgName, updatedString, isReset=False):
    """

    :type orgName: str
    :type updatedString: str
    :type isReset: bool
    """
    COLUMN_TO_UPDATE = 6
    rowToUpdate = findOrgRow(orgName)
    rounds = sheet.col_values(COLUMN_TO_UPDATE)
    noChange = False
    try:
        round_ = rounds[rowToUpdate - 1]
    except IndexError:
        isReset = True
        noChange = True
    if isReset:
        string = updatedString
        if not noChange:
            addToChanges(False, rowToUpdate, "F", round_)
    else:
        string = round_ + "; " + updatedString
    sheet.update_cell(rowToUpdate, COLUMN_TO_UPDATE, string)
    addToChanges(True, rowToUpdate, "F", updatedString)


def updateContactInfo(org, name="", pos="", connection="", honorific=""):
    """

    :type org: str
    :type name: str
    :type pos: str
    :type connection: str
    :type honorific: str

    :raises: :class: "ValueError": At least one argument must have value
    """
    if name == pos == connection == honorific == "":
        raise ValueError("At least one argument must have value")
    columns = ["B", "D", "E", "G"]
    currentColumns = []
    columnsDeleted = []
    updated = []
    columnsUpdated = []
    currentValues = []
    rowNum = findOrgRow(org)
    rowValues = sheet.row_values(rowNum)
    for i in range(len(rowValues)):
        if i == 1 or i == 3 or i == 4 or i == 6:
            currentValues.append(rowValues[i])
    print(currentValues)
    for i in range(len(currentValues)):
        if currentValues[i] != "":
            currentColumns.append(columns[i])
    print(currentColumns)

    if name != "":
        sheet.update_cell(rowNum, 2, name)
        updated.append(name)
        columnsUpdated.append("B")
    if pos != "":
        sheet.update_cell(rowNum, 4, pos)
        updated.append(pos)
        columnsUpdated.append("D")
    if connection != "":
        sheet.update_cell(rowNum, 5, connection)
        updated.append(connection)
        columnsUpdated.append("E")
    if honorific != "":
        sheet.update_cell(rowNum, 7, honorific)
        updated.append(honorific)
        columnsUpdated.append("G")
    if len(currentColumns) > 0:
        for i in columnsUpdated:
            for j in currentColumns:
                if i == j:
                    columnsDeleted.append(j)
                    break
        addToChanges(False, rowNum, "; ".join(columnsDeleted), "; ".join(currentValues))
    addToChanges(True, rowNum, "; ".join(columnsUpdated), "; ".join(updated))


# updateRounds("1password", "LA")
# addToChanges(1, 2434, "jds", "test")
# updateContactInfo("senor thickems", pos="CFO", honorific="Sr.")
# delOrgs(orgs=["Senor thickems", "Apple", "Orange"])
# save()

"""addOrg("K1 Packaging", "jennie.chang@k1packaging.com", "LA", "Silver", "Provided Stickers for the Event", "On Website", contact="", position="", connection="", honorific="")
r = findOrgRow("K1 Packaging")
print(r)
sheet.delete_rows(r+1)"""