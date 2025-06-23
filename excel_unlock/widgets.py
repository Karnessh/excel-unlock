import tkinter as tk
from tkinter import ttk
from . import models as m
from typing import List

class listview(ttk.Treeview):
    listviewSetup = [
        {'column' : 'col1', 'name' : 'col1', 'minsize' : 50, 'size' : 100 }
    ]

    def __init__(self, *args, setup=listviewSetup, **kwargs):

        columnList = []
        for list in setup:
            columnList.append(list['column'])

        super().__init__( *args, columns=columnList, **kwargs)
        index = 0
        for list in setup:
            self.addHeading(index, list)
            index += 1

    def addHeading(self, index, headings):
        super().column( index, minwidth=headings['minsize'], 
                       width=headings['size'])
        super().heading( index, text=headings['name'],anchor=tk.W)

    def addData(self, excelfile_list: List[m.ExcelSheet]):
        for data in excelfile_list:
            if data.get_file_type() != 'none':
                super().insert('', 'end',text=str(data.getIndex()), values=data.get_file_info() )
        