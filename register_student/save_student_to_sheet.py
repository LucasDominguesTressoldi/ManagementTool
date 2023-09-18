import openpyxl
from gui.create_gui import make_msgbox
from openpyxl.styles import Border, Side, Alignment
from files_paths.files import register_student_file


def insert_new_student_to_sheet(
    name,
    responsible_name,
    responsible_rg,
    responsible_cpf,
    address,
    responsible_email,
    email,
    responsible_cel,
    cel,
    birth,
    initial_date,
    expiration_date,
    classes,
    school,
    grade,
    nf_number,
):
    warning_msg = make_msgbox(
        "ERROR",
        "For any information to be added to an Excel file, it must be closed to prevent possible failures and/or errors. "
        "So make sure Excel is closed!\n\nDo you want to continue?",
        "question",
        "Continue",
        "Cancel",
    )

    if warning_msg.get() == "Continue":
        if classes != {}:
            try:
                file_name = register_student_file
                wb = openpyxl.load_workbook(file_name)
                sheet = wb["Cadastro alunos FIXOS"]

                row = 1
                row_number = 0
                selected_row = 0
                keep_going = True

                while keep_going:
                    cell_1 = sheet[f"B{row}"].value
                    cell_2 = sheet[f"B{row + 1}"].value
                    if not cell_1 and not cell_2:
                        selected_row = row
                        row_number = int(sheet[f"B{selected_row - 1}"].value) + 1
                        row_number = (
                            str(row_number)
                            if row_number >= 10
                            else "0" + str(row_number)
                        )
                        keep_going = False
                    row += 1

                center = Alignment(horizontal="center", vertical="center")
                border_style = Side(border_style="thin", color="000000")
                black = Border(
                    left=border_style,
                    right=border_style,
                    top=border_style,
                    bottom=border_style,
                )

                for key in classes:
                    data = [
                        row_number,
                        name,
                        responsible_name,
                        responsible_rg,
                        responsible_cpf,
                        address,
                        responsible_email,
                        email,
                        responsible_cel,
                        cel,
                        birth,
                        initial_date,
                        f"$ {classes[key]['price']}",
                        expiration_date,
                        classes[key]["day"],
                        classes[key]["class_hours"],
                        classes[key]["class_quantity"],
                        classes[key]["subject"],
                        classes[key]["teacher"],
                        school,
                        grade,
                        nf_number,
                    ]

                    for col, value in enumerate(data, start=2):
                        cell = sheet.cell(selected_row, col, value)
                        cell.border = black
                        cell.alignment = center

                    selected_row += 1

                try:
                    wb.save(file_name)
                    wb.close()
                    make_msgbox(
                        "SUCCESS",
                        f"The student {name} was created succesfully!",
                        "check",
                    )
                except Exception as e:
                    make_msgbox(
                        "ERROR",
                        "The Excel file to be modified is opened!\nClose it and try again.",
                        "cancel",
                    )
            except Exception as e:
                make_msgbox(
                    "ERROR",
                    f"The student was not created!\n\nERROR: {str(e)}\n\nTry again.",
                    "cancel",
                )
        else:
            make_msgbox(
                "ERROR",
                "The student sheet was NOT created because there is not registered class!\nMinimum required: 1 registered class.\n\nTry again.",
                "cancel",
            )
    else:
        make_msgbox("WARNING", "Execution canceled!", "info")
