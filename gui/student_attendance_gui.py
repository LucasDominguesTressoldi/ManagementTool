from datetime import datetime as dt
from verify_attendance.insert_attendace_to_student import update_attendance
from gui.create_gui import make_gui, make_entry, make_button, make_combobox


def attendance_gui(root):
    window = make_gui("Check Attendance", root)

    student_name = make_entry(window, "Student Name", 0, 0)
    subject = make_combobox(
        window,
        ["Math", "English", "Chemistry", "History", "Philosophy", "Physics"],
        0,
        1,
    )
    hour = make_entry(window, "Class Hour", 1, 0)
    chosen_attendance_opt = make_combobox(window, ["Done", "Pending", "Missed"], 1, 1)
    day = make_combobox(
        window,
        [
            str(dt.now().day),
            "1",
            "2",
            "3",
            "4",
            "5",
            "6",
            "7",
            "8",
            "9",
            "10",
            "11",
            "12",
            "13",
            "14",
            "15",
            "16",
            "17",
            "18",
            "19",
            "20",
            "21",
            "22",
            "23",
            "24",
            "25",
            "26",
            "27",
            "28",
            "29",
            "30",
            "31",
        ],
        2,
        0,
    )
    month = make_combobox(
        window,
        [
            str(dt.now().month),
            "1",
            "2",
            "3",
            "4",
            "5",
            "6",
            "7",
            "8",
            "9",
            "10",
            "11",
            "12",
        ],
        2,
        1,
    )
    make_button(
        window,
        "Confirm",
        3,
        1,
        1,
        lambda: update_attendance(
            student_name.get(),
            chosen_attendance_opt.get(),
            day.get(),
            month.get(),
            subject.get(),
            hour.get(),
        ),
    )
    make_button(window, "Exit", 3, 0, 1, window.destroy)
