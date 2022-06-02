from tkcalendar import Calendar
from datetime import date


def get_current_date():
    curr = date.today()
    return {"year": curr.year, "month": curr.month, "day": curr.day}


current_date = get_current_date()


def CalenderUI(app):
    return Calendar(
        app,
        selectmode="day",
        year=current_date["year"],
        month=current_date["month"],
        day=current_date["day"],
        locale="ko_KR",
    )
