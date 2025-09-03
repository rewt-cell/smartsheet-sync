import smartsheet
import os

# Get the token from environment variable
ACCESS_TOKEN = os.getenv("SMARTSHEET_TOKEN")

if not ACCESS_TOKEN:
    print("Error: SMARTSHEET_TOKEN not found in environment variables.")
    exit(1)

# Initialize the Smartsheet client
smartsheet_client = smartsheet.Smartsheet(ACCESS_TOKEN)

# Replace with your actual sheet IDs
SOURCE_SHEET_ID = 'hP678C6c64PRfjJrQHpmq92HQWMFWqVhcJJXF3J1'
TARGET_SHEET_ID = '2pgPfFrWJwCxM386gQgr68HvXwPC36fr3QXFj2Q1'

# Get both sheets
source_sheet = smartsheet_client.Sheets.get_sheet(SOURCE_SHEET_ID)
target_sheet = smartsheet_client.Sheets.get_sheet(TARGET_SHEET_ID)

# Map column names to IDs for both sheets
def get_column_map(sheet):
    return {col.title: col.id for col in sheet.columns}

source_columns = get_column_map(source_sheet)
target_columns = get_column_map(target_sheet)

# Build a lookup for NSI Number in target sheet
target_lookup = {}
for row in target_sheet.rows:
    for cell in row.cells:
        if cell.column_id == target_columns['NSI Number']:
            target_lookup[cell.value] = row
            break

# Columns to update
columns_to_update = [
    "Customer Name",
    "Expected no. of units in the next 6 months",
    "Software/firmware Version",
    "Transfer Region",
    "Transfer Region Owners",
    "Order/Forecast in place",
    "Expected delivery date of first unit/batch/pilot",
    "Comments from requestor",
    "Comments from Sergio",
    "Last Amendment Date"
]

# Prepare rows to update
rows_to_update = []

for source_row in source_sheet.rows:
    nsi_number = None
    for cell in source_row.cells:
        if cell.column_id == source_columns['NSI Number']:
            nsi_number = cell.value
            break

    if nsi_number and nsi_number in target_lookup:
        target_row = target_lookup[nsi_number]
        updated_cells = []

        for col_name in columns_to_update:
            source_value = next((c.value for c in source_row.cells if c.column_id == source_columns[col_name]), None)
            updated_cells.append({
                'column_id': target_columns[col_name],
                'value': source_value
            })

        for cell_data in updated_cells:
            cell = smartsheet.models.Cell()
            cell.column_id = cell_data['column_id']
            cell.value = cell_data['value']
            target_row.cells.append(cell)

        rows_to_update.append(target_row)

# Push updates to Smartsheet
if rows_to_update:
    response = smartsheet_client.Sheets.update_rows(TARGET_SHEET_ID, rows_to_update)
    print(f"Updated {len(response.data)} rows.")
else:
    print("No matching rows found to update.")

# Allow script to run directly
if __name__ == "__main__":
    pass  # Already runs when executed

