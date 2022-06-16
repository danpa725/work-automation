from tkinter import *
from calendar_ui import CalenderUI
from automation import automation
from load import load_dir
from system import win_set_time
from time import sleep

CONFIG_DEFAULT = "DEFAULT"
AUTO_CONFIG = {
    "slected_date": CONFIG_DEFAULT,
    "exceldir": CONFIG_DEFAULT,
    "extract_column": CONFIG_DEFAULT
}


# def get_mouse_pos():
#     print(win32gui.GetCursorPos())
# get_mouse_pos()


def init_automation():
    # config check
    if AUTO_CONFIG["slected_date"] == CONFIG_DEFAULT:
        show_slected_value(Label["app"], "날짜를 먼저 선택하세요")
        return

    if AUTO_CONFIG["exceldir"] == CONFIG_DEFAULT:
        show_slected_value(Label["app"], "파일을 먼저 선택하세요")
        return

    if AUTO_CONFIG["extract_column"] == CONFIG_DEFAULT:
        show_slected_value(Label["app"], "추출할 열을 입력하고 Enter 입력, (ex: A)")

    # set window date
    date = list(AUTO_CONFIG["slected_date"].replace(
        " ", "").strip().split("."))
    date_tuple = (int(f'20{date[0]}'), int(date[1]), int(date[2]), 0, 0, 0, 0)
    win_set_time(date_tuple)

    sleep(2.5)

    # automation process
    automation(
        exceldir=AUTO_CONFIG["exceldir"],
        id_location=AUTO_CONFIG["extract_column"],
        iter_time=1.75)

    return


def set_automation_config(config: dict[str, str], slected_date: str = CONFIG_DEFAULT, exceldir: str = CONFIG_DEFAULT, extract_column: str = CONFIG_DEFAULT):
    if slected_date != CONFIG_DEFAULT:
        config["slected_date"] = slected_date
        return
    if exceldir != CONFIG_DEFAULT:
        config["exceldir"] = exceldir
        return
    if extract_column != CONFIG_DEFAULT:
        config["extract_column"] = extract_column
        return


def show_slected_value(label: Label, value: str):
    label.config(text=value)
    return


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


def set_extract_column(event):
    slected_colum = Input["extract_column"].get()[:1].capitalize()

    if slected_colum == "":
        show_slected_value(Label["extract_column"], f"적어도 한글자 이상 입력하세요")
        return
    if slected_colum.isalpha() == False:
        show_slected_value(Label["extract_column"], f"알파벳 한글자를 입력하고 Enter 입력")
        return

    show_slected_value(Label["extract_column"], f"선택한 열: {slected_colum}")
    set_automation_config(AUTO_CONFIG, extract_column=slected_colum)

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
    "extract_column": Label(App, text="B. 추출 열을 선택하세요", padx=5, pady=5),
    "calendar": Label(App, text="C. 날짜를 선택하세요", padx=5, pady=5),


}

# Button
Button = {
    "date": Button(App, command=set_slected_date, text="날짜 선택", padx=5, pady=5),
    "file": Button(App, command=set_excel_dir, text="액셀 파일 선택", padx=5, pady=5),
    "automation": Button(App, command=init_automation, text="자동화 시작", padx=5, pady=5),
}

# Input
Input = {
    "extract_column": Entry(App)
}

# Calendar
calendar = CalenderUI(App)

# Add Component
Label["app"].pack()

# file
Label["filedir"].pack()
Button["file"].pack()

# column
Label["extract_column"].pack()
Input["extract_column"].bind("<Return>", set_extract_column)
Input["extract_column"].pack()

# calendar
Label["calendar"].pack()
Button["date"].pack()

# start btn
Button["automation"].pack()

# slect calendar
calendar.pack(pady=10)

# App initialize
App.mainloop()
