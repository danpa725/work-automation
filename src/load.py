from tkinter import filedialog


def load_dir():
    filename = filedialog.askopenfilename(
        initialdir="/",
        title="액셀 파일 선택",
        filetypes=(("*.xlsx")),
    )

    return filename
