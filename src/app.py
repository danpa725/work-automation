from tkinter import *
from calendar_ui import CalenderUI
from automation import automation
from load import load_dir
from system import win_set_time
from time import sleep

AUTO_CONFIG = {
    "slected_date": "NOT",
    "exceldir": "NOT",
}


def init_automation():
    # config check
    if AUTO_CONFIG["slected_date"] == "NOT":
        show_slected_value(Label["app"], "날짜를 먼저 선택하세요")
        return

    if AUTO_CONFIG["exceldir"] == "NOT":
        show_slected_value(Label["app"], "파일을 먼저 선택하세요")

        return

    # set window date
    date = list(AUTO_CONFIG["slected_date"].replace(
        " ", "").strip().split("."))
    date_tuple = (int(f'20{date[0]}'), int(date[1]), int(date[2]), 0, 0, 0, 0)
    win_set_time(date_tuple)

    sleep(2.5)

    # automation process
    automation(exceldir=AUTO_CONFIG["exceldir"],
               id_location="E", iter_time=1.75)

    return


def set_automation_config(config, slected_date="NOT", exceldir="NOT"):
    if slected_date == "NOT":
        config["exceldir"] = exceldir
        return
    if exceldir == "NOT":
        config["slected_date"] = slected_date
        return


def show_slected_value(label, value):
    label.config(text=value)


def set_slected_date():
    slected_date = calendar.get_date()
    show_slected_value(Label["calendar"], f"window 날짜: {slected_date} 설정")

    set_automation_config(AUTO_CONFIG, slected_date=slected_date)

    return


def set_excel_dir():
    filedir = load_dir()
    show_slected_value(Label["filedir"], f"파일 위치: {filedir}")

    set_automation_config(AUTO_CONFIG, exceldir=filedir)

    return


# -------------------------------- UI --------------------------------

# App
App = Tk()
app_config = {"title": "식권 입력", "size": "500x500"}
App.title(app_config["title"])
App.geometry(app_config["size"])
App.resizable(True, True)

# Label
Label = {
    "app": Label(App, text="식권 입력", padx=12, pady=10, relief="solid", borderwidth=2),
    "filedir": Label(App, text="A. 파일을 선택하세요", padx=5, pady=5),
    "calendar": Label(App, text="B. 날짜를 선택하세요", padx=5, pady=5),
}

# Button
Button = {
    "date": Button(App, command=set_slected_date, text="날짜 선택", padx=5, pady=5),
    "file": Button(App, command=set_excel_dir, text="액셀 파일 선택", padx=5, pady=5),
    "automation": Button(App, command=init_automation, text="자동화 시작", padx=5, pady=5),
}

# Calendar
calendar = CalenderUI(App)

# Add Component
Label["app"].pack()

Label["filedir"].pack()
Button["file"].pack()

Label["calendar"].pack()
Button["date"].pack()

Button["automation"].pack()
calendar.pack(pady=10)

# App initialize
App.mainloop()
