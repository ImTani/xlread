import openpyxl as op
import tkinter as tk
import tkinter.filedialog as fd
import os
import sys
import string

d = dict(enumerate(string.ascii_uppercase, 1))
again = True

base = tk.Tk()
# base.withdraw()
base.geometry('640x480')
base.title("Grade Viwer")

currdir = os.getcwd()

path = ''


def OpenFile():
    base.lift()
    base.attributes('-topmost', True)
    base.after_idle(base.attributes, '-topmost', False)

    filePath = fd.askopenfilename(parent=base, initialdir=currdir,
                                  title='Please select an excel file.',
                                  filetypes=[
                                     ("Excel Files", ".xlsx")
                                  ])

    global path
    path = filePath


browseButton = tk.Button(text="Browse Files", padx=10, pady=10,
                         command=lambda: OpenFile())
browseButton.pack()

while again is True:

    filePath = path

    if filePath != '':
        print("The chosen file is :", filePath)
        again = False
        break
    else:
        a = input("No file selected. Do you want to try again? (y/n) : ")

    if a == "n":
        sys.exit()


os.system("pause")

subNum = int(input("Enter Subject Code : "))
wb = op.load_workbook(path)
sheet_obj = wb.active

firstRow = sheet_obj.min_row
firstCol = sheet_obj.min_column
maxRow = sheet_obj.max_row
maxCol = sheet_obj.max_column

cell_obj = sheet_obj[str(d[firstCol]) + str(firstRow):
                     str(d[maxCol]) + str(maxRow)]

newWB = op.Workbook()
ws = newWB.active
ws.title = "Marks for Sub. Code " + str(subNum)

ws['A1'] = "Name"
ws['B1'] = "Marks"
ws['C1'] = "Grade"

h = 3
for cell1, cell2, cell3 in cell_obj:
    if cell2.value == subNum:
        cell2 = sheet_obj.cell(row=cell2.row + 1, column=cell2.column)
        cell3 = sheet_obj.cell(row=cell3.row + 1, column=cell3.column)
        ws.cell(row=h, column=1).value = cell1.value
        ws.cell(row=h, column=2).value = cell2.value
        ws.cell(row=h, column=3).value = cell3.value

        h += 2

savePath = fd.askdirectory(initialdir=currdir, title="Select location to save")

newWB.save(str(savePath) + "/Student Marks.xlsx")

print("New Excel file created! \nFile Located in :", savePath)

os.system("pause")

base.mainloop()
