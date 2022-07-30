import openpyxl as op
import tkinter as tk
import tkinter.filedialog as fd
import os
import string

d = dict(enumerate(string.ascii_uppercase, 1))

base = tk.Tk()
base.withdraw()
# base.geometry('640x480')
base.title("Grade Viwer")

currdir = os.getcwd()
filePath = fd.askopenfilename(parent=base, initialdir=currdir,
                              title='Please select an excel file.',
                              filetypes=[
                                ("Excel Files", ".xlsx")
                              ])

print(filePath)
path = filePath
# path = input("Copy and paste excel file name here.") + ".xlsx"
# path = 'Test' + ".xlsx"

subNum = int(input("Enter Subject Code : "))
wb = op.load_workbook(path)
sheet_obj = wb.active

maxRow = sheet_obj.max_row
maxCol = sheet_obj.max_column

cell_obj = sheet_obj['A2': str(d[maxCol])+str(maxRow)]

for cell1, cell2, cell3 in cell_obj:
    if cell2.value == subNum:
        cell2 = sheet_obj.cell(row=cell2.row + 1, column=cell2.column)
        cell3 = sheet_obj.cell(row=cell3.row + 1, column=cell3.column)
        print(cell1.value, cell2.value, cell3.value)
