from datetime import date, timedelta, datetime as dt


def calc_date(target_days):
    days = []
    weekday_names = [
        "Monday",
        "Tuesday",
        "Wednesday",
        "Thursday",
        "Friday",
        "Saturday",
        "Sunday",
    ]

    month = dt.now().month

    for target_day_name in target_days:
        target_day_number = weekday_names.index(target_day_name)
        days_date = date(dt.now().year, month, dt.now().day)

        while days_date.weekday() != target_day_number:
            days_date += timedelta(days=1)

        while days_date.month == month:
            formatted_date = days_date.strftime("%d/%m/%Y")
            weekday = weekday_names[days_date.weekday()]
            days.append(f"{formatted_date}, {weekday}")
            days_date += timedelta(days=7)

    days.sort(key=lambda x: dt.strptime(x.split(", ")[0], "%d/%m/%Y"))
    return days
