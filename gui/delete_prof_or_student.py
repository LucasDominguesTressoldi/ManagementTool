from gui.create_gui import make_gui, make_button, make_entry
from remove_student_or_teacher.remove_student import delete_student
from remove_student_or_teacher.remove_teacher import delete_teacher


def delete_prof_student(root):
    window = make_gui("Delete Student or Teacher", root)

    student = make_entry(window, "Student Name", 0, 0)
    teacher = make_entry(window, "Teacer Name", 0, 1)

    make_button(
        window, "Delete Student", 1, 0, 1, lambda: delete_student(student.get())
    )
    make_button(
        window, "Delete Teacher", 1, 1, 1, lambda: delete_teacher(teacher.get())
    )

    make_button(window, "Exit", 2, 0, 2, window.destroy)
