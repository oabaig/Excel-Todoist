import openpyxl
import os
import subprocess
import tkinter as tk
from tkinter import filedialog

root = tk.Tk()
root.withdraw()

file_path = filedialog.askopenfilename(filetypes=(("Excel files", "*.xlsx"), ("All files", "*.")))

wb = openpyxl.load_workbook(file_path)