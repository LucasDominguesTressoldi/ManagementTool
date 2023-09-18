def create_cell(sheet, row, col, value, black, center):
    cell = sheet.cell(row=row, column=col, value=value)
    cell.border = black
    cell.alignment = center
    return cell
