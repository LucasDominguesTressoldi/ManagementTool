import sys
import customtkinter as ctk
from CTkMessagebox import CTkMessagebox


def make_gui(title, root="root"):
    ctk.set_appearance_mode("light")
    ctk.set_default_color_theme(r"assets\orange-theme.json")

    if root == "root":
        gui = ctk.CTk()
    else:
        gui = ctk.CTkToplevel(root)
    gui.title(title)
    gui.wm_iconbitmap(r"assets\app-icon.ico")
    gui.resizable(False, False)
    return gui


def make_button(root, text, row, col, col_span=1, command=sys.exit):
    ctk.CTkButton(
        root,
        text=f"{text}",
        width=200,
        font=("Montserrat", 15, "bold"),
        command=command,
    ).grid(row=row, column=col, columnspan=col_span, padx=10, pady=10)


def make_label(root, text, row, col, col_span=1):
    label = ctk.CTkLabel(root, text=text, font=("Montserrat", 20, "bold"))
    label.grid(row=row, column=col, padx=10, pady=10, columnspan=col_span)
    return label


def make_entry(root, text, row, col):
    entry = ctk.CTkEntry(
        root, placeholder_text=f"{text}", width=200, font=("Montserrat", 13, "bold")
    )
    entry.grid(row=row, column=col, padx=10, pady=10)
    return entry


def make_msgbox(title, text, icon, opt_1="", opt_2="", opt_3=""):
    if opt_3 != "":
        return CTkMessagebox(
            title=title,
            message=text,
            icon=icon,
            option_1=opt_1,
            option_2=opt_2,
            option_3=opt_3,
        )
    elif opt_2 != "":
        return CTkMessagebox(
            title=title,
            message=text,
            icon=icon,
            option_1=opt_1,
            option_2=opt_2,
        )
    elif opt_1 != "":
        return CTkMessagebox(title=title, message=text, icon=icon, option_1=opt_1)
    else:
        return CTkMessagebox(title=title, message=text, icon=icon)


def make_checkbox(root, text, row, col, col_span=1):
    checkbox = ctk.CTkCheckBox(root, text=text)
    checkbox.grid(row=row, column=col, padx=10, pady=10, columnspan=col_span)
    return checkbox


def make_combobox(root, values, row, col):
    combobox = ctk.CTkComboBox(root, values=values)
    combobox.grid(row=row, column=col, padx=10, pady=10)
    return combobox
