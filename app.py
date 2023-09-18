from backup.create_backup import create_backup
from gui.register_student_gui import student_gui
from gui.register_teacher_gui import teacher_gui
from gui.student_attendance_gui import attendance_gui
from gui.delete_prof_or_student import delete_prof_student
from monthly_sheet.create_monthly_sheet import add_new_sheet
from gui.create_gui import make_gui, make_button, make_label, make_msgbox

root = make_gui("School Management Tool")

bck_done = create_backup()

new_sheets_added = add_new_sheet()

if bck_done == False or new_sheets_added == False:
    make_msgbox("ERROR", "Something went wrong!\n\nRestart the program.", "cancel")

make_label(root, "Choose one of the options below:", 0, 0, 2)

make_button(root, "New Student", 1, 0, 1, lambda: student_gui(root))

make_button(root, "Check Attendance", 1, 1, 1, lambda: attendance_gui(root))

make_button(root, "New Teacher", 2, 0, 1, lambda: teacher_gui(root))

make_button(root, "Delete Teacher/Student", 2, 1, 1, lambda: delete_prof_student(root))

make_button(root, "Exit", 3, 0, 2)

root.mainloop()
