import openpyxl
from openpyxl.utils import get_column_letter
import re

# Load the Excel file
file_path = 'YW1 Inbound PL_with_unmapped.xlsx'
wb = openpyxl.load_workbook(file_path)

# Source sheets
yw1_sheet = wb['YW1 Inbound PL']
container_sheet = wb['Container']

# Template sheets
sci_sheet = wb['SCI Template - Single Batch']
pl_sheet = wb['PL Template - Single Batch']

print("="*80)
print("ADDING FORMULAS TO TEMPLATES")
print("="*80)

# YW1 Inbound PL column mapping
yw1_cols = {
    'Batch': 'A', 'PO': 'B', 'ASIN': 'C', 'Item': 'D', 'IC PO_': 'E', 'IC SKU': 'F',
    'SM (Buyer)': 'G', 'CM (Purchasing Buyer)': 'H', 'Brand': 'I', 'Description': 'J',
    'Chinese Description': 'K', 'Material': 'L', 'Vendor Id': 'M', 'Vendor Name': 'N',
    'CN Vendor Name': 'O', 'Price': 'P', 'Currency': 'Q', 'UPC/EAN': 'R',
    'Tax Compliance': 'S', 'INB%23': 'T', 'INB%23 Later': 'U', 'Qty': 'V',
    'Remaining Qty': 'W', 'Qty/Ctn': 'X', 'Cartons': 'Y', 'Dimensions': 'Z',
    'Ctn Weight': 'AA', 'CBM': 'AB', 'Total Weight': 'AC', 'WH ETA': 'AD',
    'WH ATA': 'AE', 'Adjusted Qty': 'AF', 'Length': 'AG', 'Width': 'AH', 'Height': 'AI'
}

# ============================
# SCI TEMPLATE FORMULAS
# ============================

print("\n--- Processing SCI Template ---")

# SCI Template uses K9 as the Batch ID reference cell
# Row 22 is headers, data starts at row 23

sci_header_row = 22
sci_data_start = 23
sci_data_end = 122  # Allow for up to 100 rows of data

# Column mappings for SCI Template (based on header row 22)
sci_col_map = {
    'A': 'PO',           # PO# -> YW1:B
    'B': 'Batch',        # Batch ID -> YW1:A
    'C': 'Photo',        # Product Photo -> UNMAPPED
    'D': 'ASIN',         # Razin -> YW1:C (ASIN)
    'E': 'Description',  # Product name -> YW1:J
    'F': 'Material',     # Composition -> YW1:L
    'G': 'HS',           # HS code -> UNMAPPED
    'H': 'Qty',          # Quantity (Units) -> YW1:V
    'I': 'Cartons',      # Carton quantity -> YW1:Y
    'J': 'Qty/Ctn',      # Parcels per MC -> YW1:X
    'K': 'CBM',          # Carton Vol. (CBM) -> YW1:AB
    'L': 'Price',        # Unit Price -> YW1:P
    'M': 'TotalPrice'    # Total Price -> Calculated (Price * Qty)
}

print(f"Adding formulas to SCI Template rows {sci_data_start} to {sci_data_end}")

# Helper function to safely set cell value
def safe_set_cell(sheet, cell_ref, value):
    try:
        cell = sheet[cell_ref]
        # Check if cell is part of a merged cell
        if isinstance(cell, openpyxl.cell.cell.MergedCell):
            return False
        cell.value = value
        return True
    except Exception as e:
        print(f"  Warning: Could not set {cell_ref}: {e}")
        return False

# Add formulas for each data row
for row in range(sci_data_start, sci_data_end + 1):
    # Column A: PO# - use INDEX/MATCH to find PO based on Batch in column B
    safe_set_cell(sci_sheet, f'A{row}', f'=IFERROR(INDEX(\'YW1 Inbound PL\'!B:B,MATCH(B{row},\'YW1 Inbound PL\'!A:A,0)),"")')

    # Column B: Batch ID - use array formula or reference to K9 for batch lookup
    # For now, we'll extract unique batches - but typically this should be filled from a list
    if row == sci_data_start:
        safe_set_cell(sci_sheet, f'B{row}', f'=IFERROR(INDEX(\'YW1 Inbound PL\'!A:A,{row - sci_data_start + 2}),"")')
    else:
        safe_set_cell(sci_sheet, f'B{row}', f'=IFERROR(INDEX(\'YW1 Inbound PL\'!A:A,{row - sci_data_start + 2}),"")')

    # Column D: Razin (ASIN)
    safe_set_cell(sci_sheet, f'D{row}', f'=IFERROR(INDEX(\'YW1 Inbound PL\'!C:C,MATCH(B{row},\'YW1 Inbound PL\'!A:A,0)),"")')

    # Column E: Product name (Description)
    safe_set_cell(sci_sheet, f'E{row}', f'=IFERROR(INDEX(\'YW1 Inbound PL\'!J:J,MATCH(B{row},\'YW1 Inbound PL\'!A:A,0)),"")')

    # Column F: Composition (Material)
    safe_set_cell(sci_sheet, f'F{row}', f'=IFERROR(INDEX(\'YW1 Inbound PL\'!L:L,MATCH(B{row},\'YW1 Inbound PL\'!A:A,0)),"")')

    # Column H: Quantity (Units)
    safe_set_cell(sci_sheet, f'H{row}', f'=IFERROR(INDEX(\'YW1 Inbound PL\'!V:V,MATCH(B{row},\'YW1 Inbound PL\'!A:A,0)),"")')

    # Column I: Carton quantity
    safe_set_cell(sci_sheet, f'I{row}', f'=IFERROR(INDEX(\'YW1 Inbound PL\'!Y:Y,MATCH(B{row},\'YW1 Inbound PL\'!A:A,0)),"")')

    # Column J: Parcels per MC (Qty/Ctn)
    safe_set_cell(sci_sheet, f'J{row}', f'=IFERROR(INDEX(\'YW1 Inbound PL\'!X:X,MATCH(B{row},\'YW1 Inbound PL\'!A:A,0)),"")')

    # Column K: Carton Vol. (CBM)
    safe_set_cell(sci_sheet, f'K{row}', f'=IFERROR(INDEX(\'YW1 Inbound PL\'!AB:AB,MATCH(B{row},\'YW1 Inbound PL\'!A:A,0)),"")')

    # Column L: Unit Price
    safe_set_cell(sci_sheet, f'L{row}', f'=IFERROR(INDEX(\'YW1 Inbound PL\'!P:P,MATCH(B{row},\'YW1 Inbound PL\'!A:A,0)),"")')

    # Column M: Total Price (Price * Qty)
    safe_set_cell(sci_sheet, f'M{row}', f'=IF(AND(L{row}<>"",H{row}<>""),L{row}*H{row},"")')

print(f"✓ Added formulas to {sci_data_end - sci_data_start + 1} rows in SCI Template")

# ============================
# PL TEMPLATE FORMULAS
# ============================

print("\n--- Processing PL Template ---")

# PL Template uses M10 as the Batch ID reference cell
# Row 22 is headers, data starts at row 23

pl_header_row = 22
pl_data_start = 23
pl_data_end = 122

# Column mappings for PL Template
pl_col_map = {
    'A': 'PO',            # PO# -> YW1:B
    'B': 'Batch',         # Batch ID -> YW1:A
    'C': 'Photo',         # Product Photo -> UNMAPPED
    'D': 'ASIN',          # Razin -> YW1:C
    'E': 'Description',   # Product name -> YW1:J
    'F': 'Dimensions',    # Product Package Dim -> Calculated from Length x Height x Width
    'G': 'Weight',        # Product Package Weight -> YW1:AA (Ctn Weight)
    'H': 'HS',            # HS code -> UNMAPPED
    'I': 'Cartons',       # Number of Cartons -> YW1:Y
    'J': 'Qty/Ctn',       # Parcels per MC -> YW1:X
    'K': 'Qty',           # Quantity (Units) -> YW1:V
    'L': 'NetWeight',     # Net Weight -> UNMAPPED
    'M': 'GrossWeight',   # Gross Weight -> YW1:AA
    'N': 'MasterDim',     # Master Car. Dim -> Calculated from Length x Height x Width
    'O': 'MasterVol'      # Master Carton Vol. -> UNMAPPED
}

print(f"Adding formulas to PL Template rows {pl_data_start} to {pl_data_end}")

for row in range(pl_data_start, pl_data_end + 1):
    # Column A: PO#
    safe_set_cell(pl_sheet, f'A{row}', f'=IFERROR(INDEX(\'YW1 Inbound PL\'!B:B,MATCH(B{row},\'YW1 Inbound PL\'!A:A,0)),"")')

    # Column B: Batch ID
    if row == pl_data_start:
        safe_set_cell(pl_sheet, f'B{row}', f'=IFERROR(INDEX(\'YW1 Inbound PL\'!A:A,{row - pl_data_start + 2}),"")')
    else:
        safe_set_cell(pl_sheet, f'B{row}', f'=IFERROR(INDEX(\'YW1 Inbound PL\'!A:A,{row - pl_data_start + 2}),"")')

    # Column D: Razin (ASIN)
    safe_set_cell(pl_sheet, f'D{row}', f'=IFERROR(INDEX(\'YW1 Inbound PL\'!C:C,MATCH(B{row},\'YW1 Inbound PL\'!A:A,0)),"")')

    # Column E: Product name
    safe_set_cell(pl_sheet, f'E{row}', f'=IFERROR(INDEX(\'YW1 Inbound PL\'!J:J,MATCH(B{row},\'YW1 Inbound PL\'!A:A,0)),"")')

    # Column F: Product Package Dim (LxHxW) - Calculated
    safe_set_cell(pl_sheet, f'F{row}', f'=IFERROR(INDEX(\'YW1 Inbound PL\'!AG:AG,MATCH(B{row},\'YW1 Inbound PL\'!A:A,0))&"x"&INDEX(\'YW1 Inbound PL\'!AI:AI,MATCH(B{row},\'YW1 Inbound PL\'!A:A,0))&"x"&INDEX(\'YW1 Inbound PL\'!AH:AH,MATCH(B{row},\'YW1 Inbound PL\'!A:A,0)),"")')

    # Column G: Product Package Weight (Ctn Weight / Qty per Ctn)
    safe_set_cell(pl_sheet, f'G{row}', f'=IFERROR(INDEX(\'YW1 Inbound PL\'!AA:AA,MATCH(B{row},\'YW1 Inbound PL\'!A:A,0))/INDEX(\'YW1 Inbound PL\'!X:X,MATCH(B{row},\'YW1 Inbound PL\'!A:A,0)),"")')

    # Column I: Number of Cartons
    safe_set_cell(pl_sheet, f'I{row}', f'=IFERROR(INDEX(\'YW1 Inbound PL\'!Y:Y,MATCH(B{row},\'YW1 Inbound PL\'!A:A,0)),"")')

    # Column J: Parcels per MC
    safe_set_cell(pl_sheet, f'J{row}', f'=IFERROR(INDEX(\'YW1 Inbound PL\'!X:X,MATCH(B{row},\'YW1 Inbound PL\'!A:A,0)),"")')

    # Column K: Quantity (Units)
    safe_set_cell(pl_sheet, f'K{row}', f'=IFERROR(INDEX(\'YW1 Inbound PL\'!V:V,MATCH(B{row},\'YW1 Inbound PL\'!A:A,0)),"")')

    # Column M: Gross Weight (Ctn Weight * Cartons)
    safe_set_cell(pl_sheet, f'M{row}', f'=IFERROR(INDEX(\'YW1 Inbound PL\'!AA:AA,MATCH(B{row},\'YW1 Inbound PL\'!A:A,0))*INDEX(\'YW1 Inbound PL\'!Y:Y,MATCH(B{row},\'YW1 Inbound PL\'!A:A,0)),"")')

    # Column N: Master Car. Dim (LxHxW)
    safe_set_cell(pl_sheet, f'N{row}', f'=IFERROR(INDEX(\'YW1 Inbound PL\'!AG:AG,MATCH(B{row},\'YW1 Inbound PL\'!A:A,0))&"x"&INDEX(\'YW1 Inbound PL\'!AI:AI,MATCH(B{row},\'YW1 Inbound PL\'!A:A,0))&"x"&INDEX(\'YW1 Inbound PL\'!AH:AH,MATCH(B{row},\'YW1 Inbound PL\'!A:A,0)),"")')

print(f"✓ Added formulas to {pl_data_end - pl_data_start + 1} rows in PL Template")

# Save the workbook
output_file = 'YW1 Inbound PL_with_formulas.xlsx'
wb.save(output_file)

print("\n" + "="*80)
print(f"✓ Successfully added formulas to both templates!")
print(f"✓ Saved to: {output_file}")
print("="*80)
print("\nSummary:")
print("  - SCI Template: Added formulas for 10/13 fields (3 unmappable)")
print("  - PL Template: Added formulas for 11/15 fields (4 unmappable)")
print("\nUnmappable fields are documented in the 'Unmappable Fields' tab")
