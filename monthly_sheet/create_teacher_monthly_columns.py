from datetime import datetime as dt
from tools.create_sheet_cell import create_cell


def create_teacher_monthly_columns(
    teacher_db_sheet, teacher_wb, months, orange_bg, bold, center, black
):
    teacher_db_maxrow = teacher_db_sheet.max_row + 1
    for row in range(2, teacher_db_maxrow):
        teacher_db_data = teacher_db_sheet[f"A{row}"].value
        teacher_sheet = teacher_wb[f"{teacher_db_data}"]
        teacher_month_row = teacher_sheet.max_row + 2
        teacher_sheet.merge_cells(
            start_row=teacher_month_row,
            end_row=teacher_month_row,
            start_column=2,
            end_column=9,
        )
        sheet_cell = teacher_sheet.cell(
            row=teacher_month_row,
            column=2,
            value=f"{months[dt.now().month]}".upper(),
        )
        sheet_cell.fill = orange_bg
        sheet_cell.font = bold
        sheet_cell.alignment = center
        for col in range(2, 10):
            teacher_sheet.cell(row=teacher_month_row, column=col).border = black

        main_data = {
            "column": [
                "Date",
                "Hour",
                "Student",
                "Subject",
                "Plan",
                "Attendance",
                "Price",
                "Observation",
            ]
        }

        for num, col_value in enumerate(main_data["column"], start=2):
            cell = create_cell(
                teacher_sheet,
                teacher_month_row + 1,
                num,
                col_value,
                black,
                center,
            )
            cell.font = bold
