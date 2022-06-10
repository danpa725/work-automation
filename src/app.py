from tkinter import *
from calendar import Calendar
from automation import automation
from load import load_dir
from excel import get_column
from system import _win_set_time
from time import sleep

AUTO_CONFIG = {
    "slected_date": "NOT",
    "exceldir": "NOT",
}

is_config_setup_finishied = (
    AUTO_CONFIG["slected_date"] != "NOT" and AUTO_CONFIG["exceldir"] != "NOT"
)


def init_automation(config):
    # window date
    date = config["slected_date"].split("-")
    date_tuple = (date[0], date[1], date[2], 0, 0, 0)
    _win_set_time(date_tuple)

    sleep(2.5)
    # automation process
    id_list = get_column(config["exceldir"], "E")
    automation(id_list, 1.75)

    return


def set_automation_config(config, slected_date="NOT", exceldir="NOT"):
    config["slected_date"] = slected_date
    config["exceldir"] = exceldir
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
app_config = {"title": "식권 입력", "size": "640x400"}
App.title(app_config["title"])
App.geometry(app_config["size"])
App.resizable(False, False)

# Label
Label = {
    "app": Label(App, text="식권 입력", padx=5, pady=5),
    "calendar": Label(App, text="날짜를 선택하세요", padx=5, pady=5),
    "filedir": Label(App, text="파일을 선택하세요", padx=5, pady=5),
}
Label["app"].pack()
Label["calendar"].pack()
Label["filedir"].pack()

# Button
Button = {
    "date": Button(App, command=set_slected_date, text="날짜 선택", padx=5, pady=5),
    "file": Button(App, command=load_dir, text="액셀 파일 선택", padx=5, pady=5),
    "automation": Button(App, command=init_automation, text="자동화 시작", padx=5, pady=5),
}
Button["date"].pack()
Button["file"].pack()
if is_config_setup_finishied == True:
    Button["automation"].pack()

# Calendar
calendar = Calendar(App)

# App initialize
App.mainloop()
