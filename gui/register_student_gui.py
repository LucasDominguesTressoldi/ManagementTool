from gui.create_classes_gui import make_class
from register_student.student_management import student_functions_manager
from gui.create_gui import make_gui, make_entry, make_button, make_checkbox


def student_gui(root):
    window = make_gui("New Student", root)

    name = make_entry(window, "Full Name", 0, 0)
    birth = make_entry(window, "Birth", 1, 0)
    cpf = make_entry(window, "ID", 2, 0)
    email = make_entry(window, "Email", 3, 0)
    cel = make_entry(window, "Cellphone Number", 4, 0)
    school = make_entry(window, "School Name", 5, 0)
    grade = make_entry(window, "Grade", 0, 1)
    id_number = make_entry(window, "Guardian ID", 1, 1)
    address = make_entry(window, "Address", 2, 1)
    chosen_plan = make_entry(window, "Chosen Study Plan", 3, 1)
    responsible_name = make_entry(window, "Guardian Name", 4, 1)
    responsible_email = make_entry(window, "Guardian Email", 5, 1)
    responsible_cpf = make_entry(window, "Guardian ID", 0, 2)
    responsible_rg = make_entry(window, "Guardian SSN", 1, 2)
    responsible_cel = make_entry(window, "Guardian Cellphone", 2, 2)
    initial_date = make_entry(window, "Start Date", 3, 2)
    expiration_date = make_entry(window, "Due Date", 4, 2)
    nf_number = make_entry(window, "Invoice", 5, 2)

    classes = {}
    make_button(window, "Set Classes", 10, 1, 1, lambda: make_class(window, classes))

    check_register_student = make_checkbox(window, "Add Student to Registration", 11, 1)
    check_student_contract = make_checkbox(window, "Create Student Contract", 10, 0)
    check_student_sheet = make_checkbox(window, "Create Student Sheet", 11, 0)

    make_button(
        window,
        "Confirm",
        10,
        2,
        1,
        lambda: student_functions_manager(
            check_register_student.get(),
            check_student_contract.get(),
            check_student_sheet.get(),
            name.get(),
            birth.get(),
            cpf.get(),
            email.get(),
            cel.get(),
            school.get(),
            grade.get(),
            id_number.get(),
            address.get(),
            chosen_plan.get(),
            classes,
            responsible_name.get(),
            responsible_email.get(),
            responsible_cpf.get(),
            responsible_rg.get(),
            responsible_cel.get(),
            initial_date.get(),
            expiration_date.get(),
            nf_number.get(),
        ),
    )

    make_button(window, "Exit", 11, 2, 1, window.destroy)
