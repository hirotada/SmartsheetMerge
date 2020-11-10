#
# merge-from-excel.py
# Merge the Smartsheet sheet, which has been imported from an Excel sheet, with the updated Excel sheet.
#
#  Prerequisites:
#    Environment:
#    - Smartsheet Python SDK needs to be installed to the python environment
#        See https://smartsheet-platform.github.io/api-docs/#sdks-and-sample-code for detail
#    - Python 3.4, 3.5, 3.6 (required by the above SDK)
#        tested with Python 3.6.12 (built locally from Python-3.6.12.tgz) running on CentOS 7.8.2003
#    - Set Smartsheet access token to SMARTSHEET_ACCESS_TOKEN environment variable
#        See https://smartsheet-platform.github.io/api-docs/#authentication-and-access-tokens
#    Sheet:
#    - Mandatory columns (column title string) of Smartsheet sheets:
#        + Uniq keys: if below two columns are the same, this script treats the row in the existing Smartsheet sheet
#                     is the same row in the updated Excel (and the row is updated if there is any differences)
#            "Opp No"       Currently this is hardcoded as one of uniq key
#            "Detail Key"   Currently this is hardcoded as one of uniq key
#        + Comparison result column: This script sets NOT_EXIST, NO_UPDATE, UPDATED or NEWLY_ADDED accordingly
#            "Check Update" Dropdown(Single Select) type and NOT_EXIST, NO_UPDATE, UPDATED and NEWLY_ADDED are the values
#    - Current limitations:
#        + Smartsheet, which will be merged, is hardcoded.  Please change MASTER_SHEET_ID to the Sheet ID
#            You can check the Sheet ID from File --> Properties...
#        + Excel sheet is imported as its sheet name "Japan_Deal_intake_New" in the workspce(ID:232723978708868).
#            Currently, the sheet name and the workspce is hardcoded.
#        + The imported sheet is not removed (need to remove manually)
#        + Many debug messages are printed in STDOUT
#        + There is no backup option.  It is recommended to make a backup of Smartsheet sheet before executing this script
#
#  References:
#    https://smartsheet-platform.github.io/api-docs/
#    https://github.com/smartsheet-platform/smartsheet-python-sdk
#    https://github.com/smartsheet-samples/python-read-write-sheet
#    http://smartsheet-platform.github.io/smartsheet-python-sdk/smartsheet.models.html
#
# Version 0.1 (early beta)  November 10, 2020  by Hirotada Sasaki
#


# Install the smartsheet sdk with the command: pip install smartsheet-python-sdk
import smartsheet
import logging
import os
import io,sys
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')    # This is for treating UTF-8 charactors, e.g. Japanese

_dir = os.path.dirname(os.path.abspath(__file__))

# Master Seet ID
MASTER_SHEET_ID = '5980351874000772'

# The API identifies columns by Id, but it's more convenient to refer to column names. Store a map here
new_column_map = {}
master_column_map = {}

# Helper function to find cell in a row (for new sheet imported from Excel)
def get_cell_by_column_name_new(row, column_name):
    try:
        column_id = new_column_map[column_name]
        return row.get_column(column_id)
    except KeyError:
        return None

# Helper function to find cell in a row (for existing master sheet)
def get_cell_by_column_name_master(row, column_name):
    try:
        column_id = master_column_map[column_name]
        return row.get_column(column_id)
    except KeyError:
        return None

# Compare each cell value of the row, and return the row object including update data
def compare_and_update_row(new_row, master_row, new_sheet, master_sheet):
    update_row = smart.models.Row()
    updates = 0
    for new_column in new_sheet.columns:
        master_cell = get_cell_by_column_name_master(master_row, new_column.title)
        if master_cell is not None:
            if new_row.get_column(new_column.id).value != master_cell.value:
                update_cell = smart.models.Cell()
                update_cell.column_id = master_column_map[new_column.title]
                update_cell.value = new_row.get_column(new_column.id).value
                update_row.id = master_row.id
                update_row.cells.append(update_cell)
                updates += 1
                master_column = master_sheet.get_column_by_title(new_column.title)
                print("Updating(not saved yet) Row:" + str(new_row.row_number) + " Column:" + str(master_column.index + 1) + " --> " + str(new_row.get_column(new_column.id).value))
    update_cell = smart.models.Cell()
    update_cell.column_id = master_column_map["Check Update"]    
    if updates > 0:
        update_cell.value = "UPDATED"
    else:
        update_cell.value = "NO_UPDATE"
    update_row.id = master_row.id
    update_row.cells.append(update_cell)
    
    return update_row

# Return the row object including data to be added
def add_row(new_row, master_row, new_sheet, master_sheet):
    update_row = smart.models.Row()
    for new_column in new_sheet.columns:
        master_cell = get_cell_by_column_name_master(master_row, new_column.title)
        if master_cell is not None:
            if new_row.get_column(new_column.id).value: # skip empty cell
                update_cell = smart.models.Cell()
                update_cell.column_id = master_column_map[new_column.title]
                update_cell.value = new_row.get_column(new_column.id).value
                update_cell.strict = False
                update_row.cells.append(update_cell)
                master_column = master_sheet.get_column_by_title(new_column.title)
                print("Updating(not saved yet) Row:" + str(new_row.row_number) + " Column:" + str(master_column.index + 1) + " --> " + str(new_row.get_column(new_column.id).value))
    update_cell = smart.models.Cell()
    update_cell.column_id = master_column_map["Check Update"]    
    update_cell.value = "NEWLY_ADDED"
    update_row.to_top = True
    update_row.cells.append(update_cell)
    
    return update_row
    

# TODO: Replace the body of this function with your code
# This *example* looks for rows with a "Status" column marked "Complete" and sets the "Remaining" column to zero
#
# Return a new Row with updated cell values, else None to leave unchanged
def evaluate_row_and_build_updates(source_row):
    # Find the cell and value we want to evaluate
    status_cell = get_cell_by_column_name(source_row, "Status")
    status_value = status_cell.display_value
    if status_value == "Complete":
        remaining_cell = get_cell_by_column_name(source_row, "Remaining")
        if remaining_cell.display_value != "0":  # Skip if already 0
            print("Need to update row #" + str(source_row.row_number))

            # Build new cell value
            new_cell = smart.models.Cell()
            new_cell.column_id = column_map["Remaining"]
            new_cell.value = 0

            # Build the row to update
            new_row = smart.models.Row()
            new_row.id = source_row.id
            new_row.cells.append(new_cell)

            return new_row

    return None


print("Starting ...")

# Initialize client. Uses the API token in the environment variable "SMARTSHEET_ACCESS_TOKEN"
smart = smartsheet.Smartsheet()
# Make sure we don't miss any error
smart.errors_as_exceptions(True)

# Log all calls
logging.basicConfig(filename='rwsheet.log', level=logging.INFO)

# Import the sheet
# Workspace: My Snadbox (ID:232723978708868)
# result = smart.Sheets.import_xlsx_sheet(_dir + '/Japan_Deal_intake_New.xlsx', header_row_index=0)
result = smart.Workspaces.import_xlsx_sheet(232723978708868, _dir + '/Japan_Deal_intake_New.xlsx', 'Japan_Deal_intake_New', header_row_index=0)

# Load entire sheet
new_sheet = smart.Sheets.get_sheet(result.data.id)
print("Loaded " + str(len(new_sheet.rows)) + " rows from sheet: " + new_sheet.name)

# Master Sheet: Japan_Deal_Intake (ID:5391510515541892)
master_sheet = smart.Sheets.get_sheet(MASTER_SHEET_ID)

print("Loaded " + str(len(master_sheet.rows)) + " rows from sheet: " + master_sheet.name)

# Build column map for later reference - translates column names to column id
for new_column in new_sheet.columns:
    new_column_map[new_column.title] = new_column.id

for master_column in master_sheet.columns:
    master_column_map[master_column.title] = master_column.id

# Initialize "Check Update" column
# (Set "NOT_EXIST" to all cells of "Check Update" column)
rowsToUpdate = []
for master_row in master_sheet.rows:
    new_cell = smartsheet.models.Cell()
    new_cell.column_id = master_column_map["Check Update"]
    new_cell.value = "NOT_EXIST"
    new_row = smartsheet.models.Row()
    new_row.id = master_row.id
    new_row.cells.append(new_cell)
    rowsToUpdate.append(new_row)
update_row = smart.Sheets.update_rows(MASTER_SHEET_ID, rowsToUpdate)

# Note: Uniq Key: "Opp No" + "Detail Key"
rowsToUpdate = []
rowsToAdd = []
for new_row in new_sheet.rows:
    need_to_add = True
    new_opp_no = get_cell_by_column_name_new(new_row, "Opp No")
    for master_row in master_sheet.rows:
        master_opp_no = get_cell_by_column_name_master(master_row, "Opp No")
        if new_opp_no.value == master_opp_no.value:
            new_detail_key = get_cell_by_column_name_new(new_row, "Detail Key")
            master_detail_key = get_cell_by_column_name_master(master_row, "Detail Key")
            if new_detail_key.value == master_detail_key.value:
#                print("Opp No:" + new_opp_no.value + " " + new_detail_key.value + " found in the master sheet.")
                need_to_add = False
                rowToUpdate = compare_and_update_row(new_row, master_row, new_sheet, master_sheet)
#                print(rowToUpdate)
                if rowToUpdate is not None:
                    rowsToUpdate.append(rowToUpdate)

    if need_to_add: # meaning this line exists in the new sheet, but not exists in the master sheet
        rowToAdd = add_row(new_row, master_row, new_sheet, master_sheet)
        print(rowToAdd)
        rowsToAdd.append(rowToAdd)



# Write updated cells back to Smartsheet
if rowsToUpdate:
    print("Writing " + str(len(rowsToUpdate)) + " rows back to sheet id " + str(master_sheet.id))
    result = smart.Sheets.update_rows(MASTER_SHEET_ID, rowsToUpdate)

# Write added rows to Smartsheet
if rowsToAdd:
    print("Adding " + str(len(rowsToAdd)) + " rows to sheet id " + str(master_sheet.id))
    result = smart.Sheets.add_rows(MASTER_SHEET_ID, rowsToAdd)

print("Done")
