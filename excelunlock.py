import zipfile
import os
from filemanagement.functions import *
from filemanagement.objects import *
from typing import List

excelfile_list: List[ExcelSheet] = []
version_number:str = '0.1'

# To be change later depending on the os running the app
os.system('cls')
#print(os.name)

print(f'Welcome to the Excel protection remover script ver {version_number}')
print('Enter the name of the file you want to unlock :', end=None)

#filename = input()
# Not asking the filename for testing purposes
filename = './excel_sheet/excel1.xlsx'
to_be_removed = input()

#make sure the file entered is a valid zip file
# is_excelfile = check_excelfile()

if zipfile.is_zipfile(filename):
    with zipfile.ZipFile(filename, mode='r') as excel_zip_file:

        print('File opened')
        for list in excel_zip_file.namelist():
           
           # Making sure I get the right files
           if (list.find("xl/worksheets/") != -1) and not \
              (list.find("xl/worksheets/_rels/") != -1):

                filedata = excel_zip_file.read(list).decode("utf-8")
                excelfile_list.append(ExcelSheet(list,filedata))
        choice = ""
        while (choice != "Q") and (choice != "q"):
            os.system("cls")
            print(f'\nList of excel sheets currently protected in {filename}\n')

            x = 1
            for excel_sheet in excelfile_list:
                print(str(x) + '. ' + excel_sheet.get_filename_path() + \
                    ' | ' + excel_sheet.get_file_state())
                x += 1

            print('\nPlease enter the number of the file you wish to remove the \
protection : ', end="")
            choice = input()
            message = ''
            if validate_choice(choice, len(excelfile_list)):
                message = excelfile_list[int(choice) - 1].remove_protection()
            elif (choice != "Q") and (choice != "q"):
                message = 'Please enter a valid choice. '
            else:
                continue
            print(message + 'Press any key to continue. \n')
            input()
            
                    




    


else:
    print(filename + ' is not a zipfile')



