import tkinter as tk
from tkinter import filedialog as fd
import re
import os

import xlrd
import win32com.client


class MainHandler:
    def __init__(self):
        self.source_file = None
        self.target_file = None
        self.data_dict = {}

    def create_dict(self):
        rb = xlrd.open_workbook(self.source_file, formatting_info=True)
        sheet = rb.sheet_by_index(0)
        vals = (sheet.row_values(rownum) for rownum in range(sheet.nrows))
        for i in vals:
            if re.search('обои виниловые на', i[4]):
                format_str = re.findall('\d+', i[4])[3]
                self.data_dict[format_str] = i[7]

    def write_data(self):
        file_name = os.path.basename(self.target_file)
        xl = win32com.client.Dispatch("Excel.Application")
        wb = xl.Workbooks.Open(Filename=self.target_file)
        xl.Application.Run("{}!Module1.Unprotect".format(file_name))

        sheet = wb.ActiveSheet
        last_line = 0
        while True:
            last_line += 1
            if sheet.Cells(last_line, 1).value is None:
                break

        for item, value in self.data_dict.items():
            sheet.Cells(last_line, 1).value = int(item)
            sheet.Cells(last_line, 2).value = value
            last_line += 1

        xl.Application.Run("{}!Module1.Protect".format(file_name))
        xl.Application.Save()
        xl.Application.Quit()

    def preview(self):
        if self.source_file and self.target_file:
            self.create_dict()
            return True
        else:
            return False

    def start(self):
        if self.source_file and self.target_file:
            if not self.data_dict:
                self.preview()
            self.write_data()
            return True
        else:
            return False

########################################################################################################################

app = MainHandler()
root = tk.Tk()

btn1_open = tk.Button(root, text='open source file', width=20)
label1 = tk.Label(root, text='')
btn2_open = tk.Button(root, text='open file target', width=20)
label2 = tk.Label(root, text='')
btn_start = tk.Button(root, text='start', width=20)
btn_preview = tk.Button(root, text='preview', width=20)
listbox = tk.Listbox(root, height=20, width=80)
state_lable = tk.Label(root, text='')

def display_data():
    listbox.delete(0, listbox.size())
    for item, value in app.data_dict.items():
        listbox.insert(tk.END, '{}: {}'.format(item, value))

def get_path1(event):
    file1_path = fd.askopenfilename()
    label1['text'] = file1_path
    app.source_file = file1_path

def get_path2(event):
    file2_path = fd.askopenfilename()
    label2['text'] = file2_path
    app.target_file = file2_path

def preview(event):
    if state_lable['text'] != 'Data recorded':
        if app.preview():
            state_lable['text'] = 'File read'
            display_data()
        else:
            state_lable['text'] = 'No source or target file specified'

def start(evant):
    if app.start():
        display_data()
        state_lable['text'] = 'Data recorded'
    else:
        state_lable['text'] = 'No source or target file specified'

btn1_open.bind('<Button-1>', get_path1)
btn2_open.bind('<Button-1>', get_path2)
btn_preview.bind('<Button-1>', preview)
btn_start.bind('<Button-1>', start)

btn1_open.grid(row=0, column=0, sticky=tk.W)
label1.grid(row=0, column=1, sticky=tk.W)
btn2_open.grid(row=1, column=0, sticky=tk.W)
label2.grid(row=1, column=1, sticky=tk.W)
listbox.grid(row=2, columnspan=5)
state_lable.grid(row=3, columnspan=5, sticky=tk.W)
btn_preview.grid(row=4, column=0, sticky=tk.W, ipady=10)
btn_start.grid(row=4, column=1, sticky=tk.W, ipady=10)

if __name__ == '__main__':
    root.title('GS1_aggregate')
    root.geometry('500x400')
    root.mainloop()