from gui.create_gui import make_msgbox
from tools.format_to_str import to_string
from register_student.create_student_sheet import create_sheet
from register_student.create_student_contract import create_contract
from register_student.save_student_to_sheet import insert_new_student_to_sheet


def student_functions_manager(
    check_register_student,
    check_student_contract,
    check_student_sheet,
    name,
    birth,
    cpf,
    email,
    cel,
    school,
    grade,
    id_number,
    address,
    chosen_plan,
    classes,
    responsible_name,
    responsible_email,
    responsible_cpf,
    responsible_rg,
    responsible_cel,
    initial_date,
    expiration_date,
    nf_number,
):
    classes_to_str = to_string(classes)

    if check_register_student == 1:
        insert_new_student_to_sheet(
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
        )
    if check_student_sheet == 1:
        create_sheet(
            name, classes, classes_to_str, chosen_plan, initial_date, expiration_date
        )
    if check_student_contract == 1:
        create_contract(
            name,
            classes_to_str,
            responsible_name,
            responsible_cpf,
            id_number,
            address,
            chosen_plan,
            initial_date,
            expiration_date,
        )
    if all(
        condition == 0
        for condition in [
            check_register_student,
            check_student_sheet,
            check_student_contract,
        ]
    ):
        make_msgbox(
            "WARNING",
            "None option selected!\nSelect at least one of the three options.",
            "warning",
        )
