import tkinter as tk
import zipfile
import os.path 
from tkinter import ttk
from . import widgets as w
from . import models as m
from tkinter import filedialog
from tkinter import messagebox
from typing import List
from ttkbootstrap import Style



class PasswordRemovalWindows(tk.Frame):
    """Main windows of the app, used to remove the passwords from excel 
 sheets and workbooks"""
    listviewSetup = [
        {'column' : 'hType', 'name' : 'Type', 'minsize' : 20, 'size' : 100 },
        {'column' : 'hName', 'name' : 'Name', 'minsize' : 20, 'size' : 100 },
        {'column' : 'hFilename', 'name' : 'File name', 'minsize' : 20,
         'size' : 200 },
        {'column' : 'hProtection', 'name' : 'Protection', 'minsize' : 20,
         'size' : 100 },
        {'column' : 'hRemoveProtection', 'name' : 'Remove Protection',
         'minsize' : 20, 'size' : 160},
    ]
    excelfile_list: List[m.ExcelSheet] = []
    file:str = ""
    path:str = ""
    sheetsNameList = {}
    
  
    def __init__(self, parent, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)
        self.s = ttk.Style()
        self.configure(background='white')
        style = Style(theme='sandstone')
        self.loadedZipfileName = tk.StringVar(value='No file selected, please\
 push the open file button to select a file')
        self.saveZipFileName = tk.StringVar()

        self.columnconfigure(0, weight=1)

    #Initializing top frame and controls
        topFrame = ttk.LabelFrame(self, text='Protected file to Unprotect')
        topFrame.columnconfigure(0, weight=0)
        topFrame.columnconfigure(1, weight=1)
        topFrame.grid(sticky=tk.W + tk.E, padx=10, pady=10, row=0)
        #Initializing controls from top frame
        btnOpenFile =  ttk.Button(topFrame, command=self.loadFile,
                                 text="Open file")
        btnOpenFile.grid(row=0, column=0, sticky=tk.W + tk.N + tk.S, padx=10, pady=10,
                         rowspan=2)
        
        labelOpenFileTitle = ttk.Label(topFrame,text='Path and filename :')
        labelOpenFileTitle.grid(row=0, column=1, sticky=tk.W + tk.E,
                                padx=10)
        labelOpenFile = ttk.Entry(topFrame, 
                                textvariable=self.loadedZipfileName)
        labelOpenFile.configure(state=tk.DISABLED)
        labelOpenFile.grid(row=1, column=1, sticky=tk.W + tk.E, 
                           padx=10,
                           pady=(0,10))

    #Initializing middle frame and controls
        middleFrame = ttk.LabelFrame(self, 
                                     text='Workbook and sheets management')
        middleFrame.columnconfigure(0, weight=1)
        middleFrame.columnconfigure(1, weight=2)
        middleFrame.columnconfigure(2, weight=0)
        middleFrame.grid(sticky=tk.W + tk.E, padx=10, pady=10, row=1)

        self.sheetList = w.listview(middleFrame, setup=self.listviewSetup,
                                show='headings')
        self.sheetList.configure(selectmode='extended')
        self.sheetList.grid(row=1, column=0, columnspan=2, 
                            padx=(10,0), pady=(10,0))
        test = ttk.Checkbutton(self.sheetList)
        
        self.sheetListScroll = ttk.Scrollbar(middleFrame,
                                            orient=tk.VERTICAL,
                                            command=self.sheetList.yview,
                                            style='test.Vertical.TScrollbar'
                                            )
        
        self.sheetListScroll.grid(row=1, column=2,
                                  sticky=tk.N + tk.S +tk.E,
                                  padx=(0,10), pady=(10,0))
        self.sheetList.configure(yscrollcommand = self.sheetListScroll.set)
        self.buttonSelect = ttk.Button(middleFrame, 
                                       text='Remove protection on selected',
                                       command=self.removedProtectionSelected
                                       )
        self.buttonSelect.grid(row=2, column=0, 
                               sticky=tk.W + tk.E, 
                               padx=10, 
                               pady=10)
        self.buttonAll = ttk.Button(middleFrame, 
                                    text='Remove protection on all',
                                    command=lambda: 
                                    self.removedProtectionSelected(True))
        self.buttonAll.grid(row=2,column=1, columnspan=2, 
                            sticky=tk.W + tk.E, padx=10, pady=10)
        
        
    #Initializing bottom frame and controls

        bottomFrame = ttk.LabelFrame(self, text='Save file management')
        bottomFrame.columnconfigure(0, weight=0)
        bottomFrame.columnconfigure(1, weight=1)
        bottomFrame.columnconfigure(2, weight=0)
        bottomFrame.grid(sticky=tk.W + tk.E, padx=10, pady=(0,10), row=2)

        #Initializing controls from bottom frame
        btnChangePath = ttk.Button(bottomFrame, text="Change Path",
                                    command=self.changeSaveFilename)
        btnChangePath.grid(row=0, column=0, 
                           rowspan=2, sticky=tk.W +tk.N + tk.S,
                           padx=10, pady=10)
        
        labelSaveFile = ttk.Label(bottomFrame, text='Path and savefile name')
        labelSaveFile.grid(row=0, column=1, sticky=tk.W + tk.E,
                            padx=10, pady=(10,0))
        textBoxSaveFile = ttk.Entry(bottomFrame, 
                                    textvariable=self.saveZipFileName)
        textBoxSaveFile.grid(row=1, column=1, sticky=tk.W + tk.E,
                             padx=10, pady=(0,10))

        btnSave = ttk.Button(bottomFrame,text="Execute and Save", 
                             command=self.saveExcelFile)
        btnSave.grid(row=0, column=2, rowspan=2,
                     sticky=tk.N + tk.S + tk.E, padx=10, pady=10)
        
    def saveExcelFile(self):
        filename = self.saveZipFileName.get()
        m.update_file(filename, self.excelfile_list)
        if zipfile.is_zipfile(filename):
            messagebox.showinfo('Info', 'File was saved succesfully')
        else:
            messagebox.showinfo('Info', 'There was an issue with the save')
        
        self.path = os.path.dirname(filename)
        self.file = os.path.basename(filename)
        self.setSaveFile()
        self.refreshFile(filename)
    
    def changeSaveFilename(self):
        path = os.path.dirname(self.saveZipFileName.get())
        file = os.path.basename(self.saveZipFileName.get())
        filename = filedialog.asksaveasfilename(defaultextension='.xlsx',
                    filetypes=[('.Excel file', ('.xlsx','.xlsm'))],
                    title='Select an excel file to save to',
                    initialdir=path,
                    initialfile=file)
        if not (filename == ''):
            self.saveZipFileName.set(filename)

    def loadFile(self):
        filename = filedialog.askopenfilename(defaultextension='.xlsx',
                    filetypes=[('.Excel file', ('.xlsx','.xlsm'))],
                    title='Select an excel file')
        self.path = os.path.dirname(filename)
        self.file = os.path.basename(filename)
        self.setSaveFile()
        self.refreshFile(filename)
    
    def setSaveFile(self):
        self.saveZipFileName.set(self.path + self.getNewFileName(self.file))

    def getNewFileName(self, filename:str)-> str:
        '''Take a filename and add _unprotected at the end of the filename
        and before the extention
        
        Args:
            filename: The filename to modify
            
        Returns:
            The string of the modified filename'''
        splitfile = filename.split('.')
        newFilename:str = '/'
        for split in splitfile:
            # we split the string and add the _unprotected right before the
            # extention
            if split == 'xlsx' or split == 'xlsm':
                split = '_unprotected.' + split
            else:
                # Make sure the dot is added after each split, but not for the 
                # last one before the extention
                if not (split == splitfile[len(splitfile)-2]):
                    split = split + '.'
            newFilename = newFilename + split
        return newFilename

    def removedProtectionSelected(self, selectall=False):
        if selectall == False:
            selections = self.sheetList.selection()
        else:
            selections = self.sheetList.get_children()
        message = ''
        
        for selection in selections:
            sheetListCurrentSelection = self.sheetList.item(selection)
            
            currentsheet = self.excelfile_list[
                int(sheetListCurrentSelection['text'])]
            result = currentsheet.set_for_removal()
            currentsheet.findSheetName()
            message = message + '\n' + result[1]
        messagebox.showinfo('Protection status', message=message)
        self.refreshSheetList()

    def refreshSheetList(self):
        self.sheetList.delete(*self.sheetList.get_children())
        self.sheetList.addData(self.excelfile_list)

    def refreshFile(self, file):
        if file:
            
            if zipfile.is_zipfile(file):
                self.loadedZipfileName.set(file)

                with zipfile.ZipFile(file, mode='r') as excel_zipfile:
                    self.excelfile_list = []
                    index = 0
                    for list in excel_zipfile.namelist():
                        if not (list[-3:] == 'xml'):
                            print(list[-3:])
                            continue
                        filedata = excel_zipfile.read(list).decode("utf-8")
                        self.excelfile_list.append(
                                            m.ExcelSheet(
                                                        list,
                                                        filedata, 
                                                        fromIndex=index
                                                        )
                                                    )
                        if self.excelfile_list[
                                        index].check_file_type() == 'workbook':
                            self.sheetsNameList = self.excelfile_list[
                                        index].findSheetName()
                        index += 1
                    for sheet in self.excelfile_list:
                        if sheet.get_file_type() == 'worksheet':
                            #print(sheet.getSheetNumber())
                            #print(self.sheetsNameList[sheet.getSheetNumber()])
                            sheet.setName(
                                self.sheetsNameList[sheet.getSheetNumber()])

                        

                self.refreshSheetList()

            else:
                messagebox.showerror('Wrong filetype',
                                        file + ' is not a valid file')
                
        
                
    
