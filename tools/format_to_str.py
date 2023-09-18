def to_string(classes):
    price = []
    subject = []
    teacher = []
    class_days = []
    class_hours = []
    class_quantity = []
    class_duration = []

    for key in classes:
        price.append(classes[key]["price"])
        class_days.append(classes[key]["day"])
        subject.append(classes[key]["subject"])
        teacher.append(classes[key]["teacher"])
        class_hours.append(classes[key]["class_hours"])
        class_quantity.append(classes[key]["class_quantity"])
        class_duration.append(classes[key]["class_duration"])

    return [
        make_str(subject),
        make_str(teacher),
        make_str(class_hours),
        len(class_quantity),
        make_str(class_days),
        make_str(class_duration),
        sum(price),
    ]


def make_str(arr):
    item = []
    str_value = ""
    [item.append(element) for element in arr if element not in item]
    item_length = len(item)

    if item_length == 1:
        str_value = item[0]
    elif item_length == 2:
        str_value = f"{item[0]} e {item[1]}"
    else:
        for i, value in enumerate(item, start=1):
            if i + 1 > item_length:
                str_value += f" e {value}"
            else:
                if i + 2 > item_length:
                    str_value += f"{value}"
                else:
                    str_value += f"{value}, "
    return str_value
