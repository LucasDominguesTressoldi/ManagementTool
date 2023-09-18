def get_row(sheet):
    try:
        row = 4
        keep_going = True
        while keep_going:
            cell_1 = sheet[f"B{row}"].value
            cell_2 = sheet[f"B{row + 1}"].value
            if not cell_1 and not cell_2:
                keep_going = False
            else:
                row += 1

        return row
    except Exception:
        return False
