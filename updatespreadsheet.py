import sys
# from colorama import Fore, init, Back, Style
from openpyxl import load_workbook
import re
import quickstart
import openpyxl
from string import digits


"""Reading excel file using openpyxl"""
# load excel file
workbook = load_workbook(filename="/Users/bryan/copyBanner.xlsx")
workbook.active = workbook["Ellucian"]
# ws
sheet = workbook.active

# Colors declarations
my_red = openpyxl.styles.colors.Color(rgb='00FF0000')
red_fill = openpyxl.styles.fills.PatternFill(
    patternType='solid', fgColor=my_red)


def update_sheet(banner_releases, ws):
    for i, (release, link) in enumerate(banner_releases):
        cell_tracker = 0
        for cell in ws['A']:
            if cell.value == None:
                break
            cell_tracker += 1
            repost_date_cell = ws['F'+str(cell_tracker)]
            link_cell = ws["I"+str(cell_tracker)]
            banner_releases_cell = cell
            # Case where we only need to update repost date
            if str(banner_releases_cell.value) == release:
                repost_date_cell.value = str(quickstart.final_date)
                link_cell.value = link
                break

            # Case where we need to update the version number
            elif str(banner_releases_cell.value.translate(digits)) == release.translate(digits):
                banner_releases_cell.value = release
                repost_date_cell.value = (quickstart.final_date)
                link_cell.value = link
                break
        not_rcnj_related(release, link, ws)


# function for getting last row in sheet so we can add a new release in the sheet
# gets the last empty row
def get_last_row(ws):
    last_row = ws.max_row

    while ws.cell(column=1, row=last_row).value is None and last_row > 0:
        last_row -= 1
    last_col_a_value = ws.cell(column=1, row=last_row).value
    return last_row + 1


def not_rcnj_related(release, link, ws):
    current_row = str(get_last_row(ws))
    insert_release(release, ws)
    for cell in ws[current_row+":"+current_row]:
        cell.fill = red_fill

    rcnj_cell = ws["B"+current_row]
    rcnj_cell.value = "No"

    email_date_cell = ws["D"+current_row]
    email_date_cell.value = str(quickstart.final_date)
    repost_date_cell = ws["E"+current_row]
    repost_date_cell.value = str(quickstart.final_date)

    link_cell = ws["I"+current_row]
    link_cell.value = link


def insert_release(release, ws):
    last_row = get_last_row(ws)
    new_cell = ws.cell(column=1, row=last_row)
    if new_cell.value is None:
        new_cell.value = release
    else:
        print("Cell is already populated")


release_test = [('Human Resources 8.19.2', 'https://ellucian.force.com/clients/s/releases/a111M00000RP3OY/ba-hr-8192'), ('Finance 8.13.1.3',
                                                                                                                         'https://ellucian.force.com/clients/s/releases/a111M00000RP3OY/ba-hr-8192'), ('Testing 8.13.1.3', 'https://ellucian.force.com/clients/s/releases/a111M00000RP3OY/ba-hr-8192')]
update_sheet(release_test, sheet)

# modify the desired cell
# sheet["A1"] = "Full Name"


# Enumerate the cells in the  row we want to update to then color them if the banner is ramappo related
# for cell in ws["2:2"]:
#     cell.style = 'red_italic'

# trying to iterate over banner releases column values
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
