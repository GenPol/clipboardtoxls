import pyperclip


# creating_spreadsheet.py
from openpyxl import Workbook
def create_workbook(path):
    workbook = Workbook()
    workbook.save(path)
if __name__ == "__main__":
    create_workbook("hello.xlsx")

lst = []
l = pyperclip.paste()

while True:
    s = pyperclip.paste()
    if s == l:
        continue
    if s == 'exit':
        break
    else:
        lst.append(s)
        print(lst)
        l = s


# adding_data.py
from openpyxl import Workbook
def create_workbook(path):
    workbook = Workbook()
    sheet = workbook.active

# create background
    sheet["B1"] = 'Original word'
    sheet["C1"] = 'Translate'

# Filling Excel File
    num = 2
    while num != len(lst)+2:
        sheet["A"+str(num)] = num - 1
        sheet["B" + str(num)] = lst[num - 2]
        num = num + 1

    workbook.save(path)
if __name__ == "__main__":
    create_workbook("hello.xlsx")