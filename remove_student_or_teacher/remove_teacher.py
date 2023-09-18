import openpyxl
from gui.create_gui import make_msgbox
from files_paths.files import teacher_file


def delete_teacher(teacher):
    if teacher != "":
        file_name = teacher_file

        try:
            wb = openpyxl.load_workbook(file_name)
            wb.remove(wb[f"{teacher}"])

            database_sheet = wb["PythonDatabase"]

            row = 1
            while row <= database_sheet.max_row:
                if database_sheet[f"A{row}"].value == teacher:
                    database_sheet.delete_rows(row, 1)
                    break
                else:
                    row += 1

            wb.save(file_name)
            wb.close()
            make_msgbox(
                "SUCCESS", f"The teacher {teacher} was removed succesfully!", "check"
            )
        except Exception as e:
            make_msgbox(
                "ERROR",
                f"An error occured when trying to remove the teacher {teacher}!\n\nERROR: {e}.\n\nTry again.",
                "cancel",
            )
    else:
        make_msgbox("WARNING", "Type the teacher name!", "warning")
