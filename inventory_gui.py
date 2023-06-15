import customtkinter
import tkinter
from tkinter import filedialog
from fileinput import filename
import os
from cProfile import run
from inventory import *
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from openpyxl import styles
import numpy as np
from datetime import datetime, date

customtkinter.set_appearance_mode("System")
customtkinter.set_default_color_theme("blue")

app = customtkinter.CTk()

app.geometry("300x300")
app.title("ConnectWise/QuickBooks Inventory")
app.configure(fg_color="#2D1723", corner_radius=15)

load_cw = customtkinter.CTkButton(master=app, text="Select ConnectWise Report", width=175, corner_radius=15, height=50, command=load_cw_report)
load_cw.place(relx = 0.5, rely=0.18, anchor=tkinter.CENTER)

load_qb = customtkinter.CTkButton(master=app, text="Select QuickBooks Report", width=175, corner_radius=15, height=50, command=load_qb_report)
load_qb.place(relx = 0.5, rely=0.36, anchor=tkinter.CENTER)

button_run = customtkinter.CTkButton(master=app, text="Compare reports", width=175, corner_radius=15, height=50, command=run_report)
button_run.place(relx = 0.5, rely = 0.54, anchor=tkinter.CENTER)

button_open = customtkinter.CTkButton(master=app, text="Open Inventory Results", width=175, corner_radius=15, height=50, fg_color="#469F41", hover_color="#33742F", command=saved_report)
button_open.place(relx = 0.5, rely = 0.72, anchor=tkinter.CENTER)

button_exit = customtkinter.CTkButton(master=app, text="Exit", width=75, corner_radius=15, height=35, fg_color="red", hover_color="#C70039", command=app.destroy)
button_exit.place(relx=.5, rely=.90, anchor=tkinter.CENTER)

app.mainloop()