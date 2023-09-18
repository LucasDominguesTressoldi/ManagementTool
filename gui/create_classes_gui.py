import openpyxl
from gui.create_gui import (
    make_gui,
    make_entry,
    make_button,
    make_combobox,
    make_label,
    make_msgbox,
)
from files_paths.files import teacher_file


def make_class(root, classes):
    window = make_gui("Set Classes", root)

    teacher_wb = openpyxl.load_workbook(teacher_file)
    teacher_sheet = teacher_wb["PythonDatabase"]

    teacher_row = teacher_sheet.max_row + 1
    teacher_set = set()
    for row in range(2, teacher_row):
        teacher_set.add(teacher_sheet[f"A{row}"].value)

    make_label(window, "Classes added:", 0, 0)
    number_class_label = make_label(window, f"{len(classes)}", 0, 1, 2)
    subject = make_combobox(
        window,
        ["Math", "English", "Chemistry", "History", "Philosophy", "Physics"],
        1,
        0,
    )
    if len(teacher_set) > 0:
        teacher = make_combobox(window, list(teacher_set), 2, 0)
    else:
        teacher = make_combobox(window, ["No Teacher"], 2, 0)
    day = make_combobox(
        window,
        ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"],
        3,
        0,
    )
    class_duration = make_entry(window, "Class Duration", 4, 0)
    class_quantity = make_entry(window, "Class Quantity/Week", 1, 1)
    price = make_entry(window, "R$ Price", 2, 1)
    class_hours = make_entry(window, "Class Hour", 3, 1)

    make_button(
        window,
        "Add Class",
        4,
        1,
        1,
        lambda: new_class(
            subject,
            teacher,
            class_hours,
            class_quantity,
            class_duration,
            price,
            day,
            classes,
            number_class_label,
        ),
    )
    make_button(
        window,
        "Remove Class",
        5,
        1,
        1,
        lambda: remove_classes(classes, number_class_label),
    )
    make_button(
        window, "Confirm and Exit", 5, 0, 1, lambda: confirm_exit(window, classes)
    )


def new_class(
    subject,
    teacher,
    class_hours,
    class_quantity,
    class_duration,
    price,
    day,
    classes,
    number_class_label,
):
    if all(
        items != ""
        for items in [
            subject.get(),
            teacher.get(),
            class_hours.get(),
            class_quantity.get(),
            class_duration.get(),
            price.get(),
            day.get(),
        ]
    ):
        try:
            int_price = int(price.get())
            int_class_quantity = int(class_quantity.get())
            key = len(classes)
            classes[key] = {}
            classes[key]["subject"] = subject.get()
            classes[key]["teacher"] = teacher.get()
            classes[key]["class_hours"] = class_hours.get()
            classes[key]["class_quantity"] = int_class_quantity
            classes[key]["class_duration"] = class_duration.get()
            classes[key]["price"] = int_price
            classes[key]["day"] = day.get()
            number_class_label.configure(text=f"{len(classes)}")
        except Exception:
            make_msgbox(
                "WARNING",
                "The price and class quantity/week must contain only numbers.",
                "warning",
            )
    else:
        make_msgbox("WARNING", "Fill all the fields to continue!", "warning")


def confirm_exit(window, classes):
    window.destroy()
    if len(classes) > 0:
        make_msgbox("SUCCESS", f"{len(classes)} class added!", "check")
    else:
        make_msgbox(
            "WARNING",
            "No class added!",
            "warning",
        )


def remove_classes(classes, number_class_label):
    if len(classes) > 0:
        classes.popitem()
        number_class_label.configure(text=f"{len(classes)}")
    else:
        make_msgbox("INFO", "No class added!", "info")
