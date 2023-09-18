from docx import Document
from datetime import datetime as dt
from gui.create_gui import make_msgbox
from files_paths.files import main_contract_file, new_contract_file


def create_contract(
    name,
    classes_to_str,
    responsible_name,
    responsible_cpf,
    id_number,
    address,
    chosen_plan,
    initial_date,
    expiration_date,
):
    try:
        doc = Document(main_contract_file)

        text_to_change = {
            "SUBJECT_HERE": classes_to_str[0],
            "NAME_HERE": name,
            "CLASS_QUANTITY_HERE": str(classes_to_str[3]),
            "CLASS_DURATION_HERE": classes_to_str[5],
            "CLASS_DAYS_HERE": classes_to_str[4],
            "CLASS_HOURS_HERE": classes_to_str[2],
            "RESPONSIBLE_HERE": responsible_name,
            "RESPONSIBLE_CPF": responsible_cpf,
            "ID_NUMBER_HERE": id_number,
            "ADDRESS_HERE": address,
            "CHOSEN_PLAN_HERE": chosen_plan,
            "DAY_HERE": str(dt.now().day),
            "MONTH_HERE": str(dt.now().strftime("%B")).lower(),
            "YEAR_HERE": str(dt.now().year),
            "INITIAL_DATE_HERE": initial_date,
            "EXPIRATION_DATE_HERE": expiration_date,
            "DAY_DATE_HERE": str(dt.now().day),
            "MONTH_DATE_HERE": str(dt.now().strftime("%B")).lower(),
            "YEAR_DATE_HERE": str(dt.now().year),
        }

        for paragraph in doc.paragraphs:
            for text, new_text in text_to_change.items():
                if text in paragraph.text:
                    paragraph.text = paragraph.text.replace(text, new_text)

        try:
            doc.save(rf"{new_contract_file}\Contract_{name}.docx")
            make_msgbox(
                "SUCCESS",
                f"The student contract {name} was created succesfully!",
                "check",
            )
        except Exception as e:
            make_msgbox(
                "ERROR",
                f"ERROR: The student contract was not created because there was an error trying to save it!\nTry again.",
                "cancel",
            )
    except Exception as e:
        make_msgbox(
            "ERROR",
            f"The student contract has not been created!\n\nERROR: {str(e)}\n\nTry again.",
            "cancel",
        )
