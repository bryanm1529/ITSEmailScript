import sys
# from colorama import Fore, init, Back, Style
from openpyxl import load_workbook
import re
import quickstart
import openpyxl


"""Reading excel file using openpyxl"""
#load excel file
workbook = load_workbook(filename="/Users/bryan/copyBanner.xlsx")
workbook.active = workbook["Ellucian"]
sheet = workbook.active

#Colors declarations
my_red = openpyxl.styles.colors.Color(rgb='00FF0000')
red_fill = openpyxl.styles.fills.PatternFill(
    patternType='solid', fgColor=my_red)



#function for getting last row in sheet so we can add a new release in the sheet
#gets the last empty row

def get_last_row(ws):
    last_row = ws.max_row

    while ws.cell(column=1, row=last_row).value is None and last_row > 0:
        last_row -= 1
    last_col_a_value = ws.cell(column=1, row=last_row).value
    return last_row + 1

def not_rcnj_related(release,ws):
    current_row = str(get_last_row(ws))
    insert_release(release,ws)
    for cell in ws[current_row+":"+current_row]:
        cell.fill = red_fill

    rcnj_cell = ws["B"+current_row]

    rcnj_cell.value = "No"

def insert_release(release,ws):
    last_row = get_last_row(ws)
    new_cell = ws.cell(column=1, row=last_row)
    if new_cell.value is None:
        new_cell.value = release
    else:
        print("Cell is already populated")


release = "Student 6.9"
not_rcnj_related(release,sheet)
#modify the desired cell
# sheet["A1"] = "Full Name"


# Enumerate the cells in the  row we want to update to then color them if the banner is ramappo related
# for cell in ws["2:2"]:
#     cell.style = 'red_italic'

#trying to iterate over banner releases column values
# for row in sheet.iter_rows(min_row=1, max_col=1, max_row=167, values_only=True):
#     for release in banner_releases:
#         #Want to update sheet if release is in row identically
#         if release == row:

#         #want to regex for release without version number
#         # elif release == revised_release:

#         #Update sheet with new releases if they are not there at all
#         else:
#             #call function to get last row

# save the file
workbook.save(filename="/Users/bryan/copyBanner.xlsx")




