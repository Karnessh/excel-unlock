import zipfile
from filemanagement.functions import *
from enum import Enum

class ResultState(Enum):
    WORKED = 1
    NOT_PROTECTED = 2

class ExcelSheet():

    is_protected:bool = False
    filename_path:str = ''
    file_data:str = ''
    

    def __init__(self, filename_path:str, file_data:str) -> None:
        """ Initializing class and checking sheet protection """
        self.filename_path = filename_path
        self.file_data = file_data
        self.is_protected = self.check_protection()

    def get_is_protected(self)-> bool:
        """ Return if the sheet is protected """
        return self.is_protected
    
    def get_filename_path(self)-> str:
        """ Return the file path in the zipfile """
        return self.filename_path
    
    def get_file_data(self)-> str:
        """ return the sheet data """
        return self.file_data
    
    def remove_protection(self)-> ResultState:
        """ Remove the sheet protection and return a status code """
        if (self.is_protected):
            self.file_data = remove_range_substring(self.file_data, '<sheetProtection', '/>')
            return ResultState.WORKED

        return ResultState.NOT_PROTECTED
    
    def check_protection(self) -> bool:
        """ Check if the cheet is protected and change the object is_protected variable to reflect that """
        start_index = self.file_data.find('<sheetProtection')
        if start_index == -1:
            return False  # Substring not found

        return True
    
