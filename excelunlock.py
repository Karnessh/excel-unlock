import zipfile
from filemanagement.functions import *
from filemanagement.objects import *
from typing import List

excelfile_list: List[ExcelSheet] = []
version_number:str = '0.1'

print(f'Welcome to the Excel protection remover script ver {version_number}')
print('Enter the name of the file you want to unlock :', end=None)

#filename = input()
# Not asking the filename for testing purposes
filename = './excel_sheet/excel1.xlsx'


#make sure the file entered is a valid zip file
# is_excelfile = check_excelfile()

if zipfile.is_zipfile(filename):
    try:
        #file = open(filename,"r")
        zfile = zipfile.ZipFile(filename, mode='r')
    except():
        print('Error opening the file, it may not exist')
        exit()

    print('File opened')
    for list in zfile.namelist():
        if (list.find("xl/worksheets/") != -1) and not (list.find("xl/worksheets/_rels/") != -1):
            #open_file_in_zipfile(zfile, list)
            filedata = zfile.read(list).decode("utf-8")
            excelSheet = ExcelSheet(list,filedata)
            excelfile_list.append(excelSheet)
    zfile.close()
else:
    print(filename + ' is not a zipfile')

for excel_sheet in excelfile_list:
    print(excel_sheet.get_filename_path())
    print(excel_sheet.get_is_protected())

