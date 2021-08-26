import openpyxl
import os
import tkinter as tk
from tkinter import filedialog
import todoist
from dotenv import load_dotenv

load_dotenv('---.env')
TOKEN = os.getenv('TOKEN')

print(TOKEN)

todoist_api = todoist.TodoistAPI(TOKEN)
todoist_api.sync()
print(todoist_api['user']['full_name']) 

item = todoist_api.items.add('first test task')
todoist_api.commit()


root = tk.Tk()
root.withdraw()

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