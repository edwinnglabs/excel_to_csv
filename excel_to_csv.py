import tkinter as tk
import pandas as pd
import os
# need import cmath to work around pyinstaller bug
import cmath
from tkinter.filedialog import askopenfilename
from openpyxl import load_workbook

FILE_PATH = ""
FILE_NAME = ""
FILE_EXTENSION = ""
DIR_PATH = ""

# Understanding Widget Naming Conventions
# (https://realpython.com/python-gui-tkinter/#assigning-widgets-to-frames-with-frame-widgets)
# Widget Class	Variable Name Prefix	Example
# Label	lbl	lbl_name
# Button	btn	btn_submit
# Entry	ent	ent_age
# Text	txt	txt_notes
# Frame	frm	frm_address


def browse_files():
    global FILE_PATH, FILE_NAME, FILE_EXTENSION, DIR_PATH
    FILE_PATH = askopenfilename(
        initialdir="/",
        title="Select a File",
        filetypes=(("Excel Files", ".xlsx"),
                   ("all files", "*.*"),
                   )
    )
    if not FILE_PATH:
        return
    else:
        # meta data
        _, FILE_EXTENSION = os.path.splitext(FILE_PATH)
        DIR_PATH, FILE_NAME = os.path.split(FILE_PATH)
        # update status
        lbl_file_browser.configure(text="File Opened: {}\nFile Name: {}\nFile Ext.: {}".format(
            FILE_PATH, FILE_NAME, FILE_EXTENSION)
        )


def convert_file_to_csv():
    # read
    wb = load_workbook(filename=FILE_PATH)
    ws = wb.active
    df = pd.DataFrame(ws.values)
    # write
    # os.mkdir(dir_path)
    new_file_name = FILE_NAME.replace(FILE_EXTENSION, '.csv')
    df.to_csv('{}/{}'.format(DIR_PATH, new_file_name), index=False, header=True)
    # update status
    lbl_status.configure(text="Status: Done!")


# create frames and window
app = tk.Tk()
frm_upper = tk.Frame(master=app, width=700, height=150, bg="black")
frm_lower = tk.Frame(master=app, width=700, height=100, bg="black")
app.title('Excel Files Convertor')

# positions
frm_upper.pack(fill=tk.BOTH, expand=True)
frm_lower.pack(fill=tk.BOTH, expand=True)

# buttons, labels and widgets
lbl_file_browser = tk.Label(master=frm_upper,
                            text="Select a File",
                            wraplength=500,
                            # width=20, height=4,
                            bg="black",
                            fg="white")

btn_file_browser = tk.Button(master=frm_upper,
                             text="Browse Files",
                             width=20, height=4,
                             bg="black",
                             # this fixes mac problem
                             highlightbackground="black",
                             fg="white",
                             command=browse_files)

lbl_status = tk.Label(master=frm_lower,
                      text="Status: ",
                      # width=20, height=4,
                      bg="black",
                      fg="white")

btn_export = tk.Button(master=frm_lower,
                       text="Run",
                       width=20, height=4,
                       bg="black",
                       # this fixes mac problem
                       highlightbackground="black",
                       fg="white",
                       command=convert_file_to_csv)

# positions
lbl_file_browser.place(x=0, y=0)
btn_file_browser.place(x=500, y=0)
lbl_status.place(x=0, y=0)
btn_export.place(x=500, y=0)

app.mainloop()
