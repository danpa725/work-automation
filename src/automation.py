import pyautogui as pg
from tkinter import messagebox
from time import sleep
from excel import get_column

BUTTON_POS = {"x": 174, "y": 204}

WINDOW_TITLE = "노인복지종합시스템PLUS"


def automation(exceldir, id_location, iter_time):
    try:
        focus_window = pg.getWindowsWithTitle(WINDOW_TITLE)[0]
        focus_window.maximize()

        if focus_window.isActive == False:
            focus_window.activate()

        # extract id list
        id_list = get_column(exceldir, id_location)

        for id in id_list:
            pg.moveTo(
                focus_window.left +
                BUTTON_POS["x"], focus_window.top + BUTTON_POS["y"]
            )
            pg.click(focus_window.top +
                     BUTTON_POS["x"], focus_window.top + BUTTON_POS["y"])
            pg.write(id)
            pg.press("enter")

            sleep(iter_time)

    except:
        messagebox.showerror("error", "회원 회비입금관리가 최대로 열려 있는지 확인하세요")
        return
