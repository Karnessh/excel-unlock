import zipfile
import os
import sys
from filemanagement.functions import *
from filemanagement.objects import *
from typing import List

excelfile_list: List[ExcelSheet] = []
menu: List[ExcelSheet] = []
version_number:str = '0.2'
os_clear = ''

# Checking which os is running the script to know which command to use to clear
# the screen
if os.name == 'nt':
    os_clear = 'cls'
else:
    os_clear = 'clear'

os.system(os_clear)

print(f'Welcome to the Excel protection remover script ver {version_number}')
print('Enter the name of the file you want to unlock : ', end='')

#filename = input()
# Not asking the filename for testing purposes
#zip_filename = './excel_sheet/excel1.xlsx'
zip_filename = input()

# make sure the file entered is a valid zip file
if zipfile.is_zipfile(zip_filename):
    with zipfile.ZipFile(zip_filename, mode='r') as excel_zipfile:

        print('File opened')
        for list in excel_zipfile.namelist():
           
            filedata = excel_zipfile.read(list).decode("utf-8")
            excelfile_list.append(ExcelSheet(list,filedata))
           
else:
    print(zip_filename + ' is not a zipfile\n')
    sys.exit()
# Creating interface
choice = ''
while (choice != "Q") and (choice != "q"):
    os.system(os_clear)

        # Create menu to work with the files

    print(f'\nList of excel sheets currently protected in {zip_filename}\n')
    print(f'  -Type-      | -Protection-      |   -Filename-')
    x = 1
    for excel_sheet in excelfile_list:
        if excel_sheet.print_menuitem(x):
            menu.append(excel_sheet)
            x += 1

    print('\nX. Apply the removal of the protection to archive')
    print('Q. Quit the script')
    print('\nPlease enter the number of the file you wish to remove the \
protection : ', end='')
            
        # Validating choices and processing it

    choice = input()
    message = ''
    if validate_choice(choice, len(menu)):
        message = menu[int(choice) - 1].set_for_removal()
    elif (choice == "X") or (choice == "x"):
        update_file(zip_filename, excelfile_list)
        message = 'File written successfully. '
    elif (choice == "Q") or (choice == "q"):
        continue
    else:
        message = 'Please enter a valid choice. '
    print(message + 'Press any key to continue. \n')
    input()




