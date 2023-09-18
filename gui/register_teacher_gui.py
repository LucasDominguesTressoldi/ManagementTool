from register_teacher.create_teacher_sheet import create_teacher
from gui.create_classes_gui import make_gui, make_entry, make_button


def teacher_gui(root):
    window = make_gui("New Teacher", root)

    name = make_entry(window, "Teacher Name", 0, 0)

    make_button(window, "Confirm", 0, 1, 1, lambda: create_teacher(name.get()))
