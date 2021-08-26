import openpyxl
import os
import tkinter as tk
from tkinter import filedialog
import todoist
from dotenv import load_dotenv

root = tk.Tk()
root.withdraw()

load_dotenv('---.env')
TOKEN = os.getenv('TOKEN')

todoist_api = todoist.TodoistAPI(TOKEN)
todoist_api.sync()

"""
example adding item
item = todoist_api.items.add('first test task', due={
    "date": "2021-08-25"
})
print(item)
todoist_api.commit()
"""

file_path = filedialog.askopenfilename(filetypes=(("Excel files", "*.xlsx"), ("All files", "*.")))

if file_path == '':
    exit()

wb = openpyxl.load_workbook(file_path)
 
timeline_sheet = wb['Sheet1']

class_names = timeline_sheet['A']
dates = timeline_sheet['B']
times = timeline_sheet['C']
assignment_names = timeline_sheet['D']
assignment_types = timeline_sheet['E']

print(type(dates[0].value))