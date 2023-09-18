import os
import shutil
import openpyxl
from datetime import datetime as dt
from gui.create_gui import make_msgbox
from files_paths.files import (
    backup_file,
    management_student_file,
    teacher_file,
    register_student_file,
)


def create_backup():
    try:
        file_name = backup_file
        curr_year = dt.now().year
        curr_month = dt.now().strftime("%B")
        curr_day = dt.now().strftime("%d-%m-%Y")

        wb = openpyxl.load_workbook(management_student_file)
        sheet = wb["PythonDatabase"]

        if os.path.exists(f"{file_name}/{curr_month}/{curr_day}"):
            try:
                sheet["B1"].value = "do not change month"
                wb.save(management_student_file)
                wb.close()
            except:
                make_msgbox(
                    "ERROR",
                    "Close all Excel files before open the application!",
                    "cancel",
                )
                return False
            return True

        if not os.path.exists(file_name):
            os.makedirs(file_name)

        dir_month = os.path.join(file_name, curr_month)

        if not os.path.exists(dir_month):
            os.makedirs(dir_month)
            sheet["B1"].value = "change month"
        else:
            sheet["B1"].value = "do not change month"

        try:
            wb.save(management_student_file)
            wb.close()
        except Exception:
            make_msgbox(
                "ERROR",
                "Close all Excel files before open the application!",
                "cancel",
            )
            return False

        dir_day = os.path.join(dir_month, curr_day)

        if not os.path.exists(dir_day):
            list_dir_month = os.listdir(dir_month)
            if len(list_dir_month) == 0:
                os.makedirs(dir_day)
            else:
                for folder in list_dir_month:
                    os.rename(f"{dir_month}/{folder}", dir_day)

            old_file_registered_students = os.path.join(
                dir_day, f"RegisteredStudents.xlsx"
            )
            old_file_student = os.path.join(dir_day, f"StudentsManagement.xlsx.xlsx")
            old_file_teacher = os.path.join(dir_day, f"TeachersManagement.xlsx")
            shutil.copy(register_student_file, old_file_registered_students)
            shutil.copy(management_student_file, old_file_student)
            shutil.copy(teacher_file, old_file_teacher)
            make_msgbox(
                "SUCCESS",
                "Backup file created successfully!",
                "check",
            )
            return True
    except Exception as e:
        make_msgbox(
            "ERROR",
            f"Backup file was not created!\n\nERROr: {e}\n\nRestart the application and try again.",
            "cancel",
        )
        return False
