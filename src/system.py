from datetime import datetime


def win_set_time(time_tuple: tuple):
    import win32api

    dayOfWeek = datetime(*time_tuple).isocalendar()[2]
    time = time_tuple[:2] + (dayOfWeek,) + time_tuple[2:]
    win32api.SetSystemTime(*time)

    return
