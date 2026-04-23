import xlwings as xw

# Attach to already open workbook
wb = xw.apps.active.books["template.xlsx"]

ws_ops = wb.sheets["Work Definition Operations"]
ws_res = wb.sheets["Operation Resources"]

# --- STEP 1: Scan Work Definition Operations for "SSH001" in column F ---
matched_pairs = []  # list of (name_identifier, number_identifier)

last_row_ops = ws_ops.range("A" + str(ws_ops.cells.last_cell.row)).end("up").row

for row in range(2, last_row_ops + 1):  # Assuming headers on row 1
    if ws_ops.range(f"F{row}").value == "SSH001":
        name_id = ws_ops.range(f"A{row}").value
        num_id  = ws_ops.range(f"E{row}").value
        matched_pairs.append((name_id, num_id))

print("Matched name/number pairs:", matched_pairs)

# --- STEP 2: For each matched pair, find rows in Operation Resources ---
last_row_res = ws_res.range("A" + str(ws_res.cells.last_cell.row)).end("up").row

for name_id, num_id in matched_pairs:
    for row in range(2, last_row_res + 1):

        # Column A + E match?
        if ws_res.range(f"A{row}").value == name_id and ws_res.range(f"E{row}").value == num_id:

            # Column G contains LEADTIME_DAYS?
            if ws_res.range(f"G{row}").value == "LEADTIME_DAYS":

                # Replace column D with UPDATE
                ws_res.range(f"D{row}").value = "UPDATE"
                print(f"Updated row {row} for ({name_id}, {num_id})")

print("Completed.")
