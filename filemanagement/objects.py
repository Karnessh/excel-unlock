from filemanagement.functions import *
from enum import Enum

#class ResultState(Enum):
#    WORKED = 1
#    NOT_PROTECTED = 2

class ExcelSheet():

    file_type:str = ""
    is_protected:bool = False
    is_for_protection_removal:bool = False
    filename_path:str = ''
    file_data:str = ''
    protection_string:str = ''
    

    def __init__(self, filename_path:str, file_data:str) -> None:
        """ Initializing class and checking sheet protection """
        self.filename_path = filename_path
        self.file_data = file_data
        self.file_type = self.check_file_type()
        self.is_protected = self.check_protection()

    def get_is_protected(self)-> bool:
        """ Return if the sheet is protected """
        return self.is_protected
    
    def get_file_state(self) -> str:
        if self.is_protected:
            return "PROTECTED"
        return "UNPROTECTED" 
    
    def get_filename_path(self)-> str:
        """ Return the file path in the zipfile """
        return self.filename_path
    
    def get_file_data(self)-> str:
        """ return the sheet data """
        return self.file_data
    
    def get_file_type(self)-> str:
        """ return the sheet data """
        return self.file_type
    
    def get_file_data_bytes(self) -> bytes:
        return self.file_data.encode('utf-8')
    
    def set_for_removal(self) -> str:
        if (self.is_protected):
            self.is_for_protection_removal = True
            return 'Protection is set to be removed '
        
        return 'File is already unprotected'
        
    def execute_change(self)-> bytes:
        """ Remove the sheet protection and return a status code """
        if (self.is_for_protection_removal == True):
            self.file_data = remove_range_substring(self.file_data,\
                            self.protection_string, '/>')
            self.is_protected = False

        return self.get_file_data().encode('utf-8')
    
    def check_file_type(self) -> str:
        result = self.file_data.find('<worksheet')
        if result != -1:
            self.protection_string = '<sheetProtection'
            return 'worksheet'
        
        result = self.file_data.find('<workbook')
        if result != -1:
            self.protection_string = '<workbookProtection'
            return 'workbook'

        return 'none'
    
    def check_protection(self) -> bool:
        """ Check if the cheet is protected and change the object \n
        is_protected variable to reflect that """
        start_index = self.file_data.find(self.protection_string)
        if start_index != -1:
            return True  
        
        return False # Substring not found
    
    def print_menuitem(self, menuNumber) -> bool:
        if self.file_type != "none":
            print(f'{str(menuNumber)}. {self.get_file_type():9} \
 |  {self.get_file_state():12} |  {self.get_filename_path()}')
            
            return True
        else:
            return False
