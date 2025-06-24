import zipfile

class ExcelSheet():

    file_type:str = ""
    filename_path:str = ''
    sheetName:str = ''
    sheetNumber:int = 0
    file_data:str = ''
    fileDataBytes:bytes = b''
    index:int = 0
    is_protected:bool = False
    is_for_protection_removal:bool = False
    protection_string:str = ''
    removal:str = ''
    isXmlFile = False
    

    def __init__(self, filename_path:str, file_data:bytes , 
                 fromIndex = 0) -> None:
        """ Initializing class and checking sheet protection """
        self.filename_path = filename_path
        self.fileDataBytes = file_data
        if self.filename_path[-3:] == 'xml':
            self.file_data = self.fileDataBytes.decode("utf-8")
            self.isXmlFile = True
        self.file_type = self.check_file_type()
        self.is_protected = self.check_protection()
        self.index = fromIndex
        if self.file_type == 'worksheet':
            self.findSheetNumber()

    def findSheetNumber(self):
        result = self.filename_path.split('.')
        self.sheetNumber = int(result[0][19:])
        #print(self.sheetNumber)
    
    def get_is_protected(self)-> bool:
        """ Return if the sheet is protected """
        return self.is_protected
    
    def get_file_state(self) -> str:
        """ convert the boolean protected to a string for better readability """
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
    
    def getSheetNumber(self)->str:
        return str(self.sheetNumber)
    
    def get_is_for_protection_removal(self):
        """ return if the sheet is set for protection removal """
        return self.is_for_protection_removal
    
    def get_file_info(self)-> tuple:
        setForRemoval = "No"
        if self.get_is_for_protection_removal():
            setForRemoval = "Yes"
        if not self.get_is_protected():
            setForRemoval = "N/A"
        value = (self.get_file_type(),
                self.sheetName, 
                self.get_filename_path(), 
                self.get_file_state(),
                setForRemoval)
        return value
    
    def getIndex(self):
        return self.index
    
    def set_for_removal(self) -> tuple:
        """ Flag the sheet for protection removal"""
        filePathName = self.get_filename_path()
        if (self.is_protected):
            self.is_for_protection_removal = True
            return (True, f'{filePathName} is flagged to be unprotected')
        
        return (False, f'{filePathName} is already unprotected')
    
    def setIndex(self, index):
        self.index = index

    def setName(self, name:str):
        self.sheetName = name
        
    def execute_change(self)-> bytes:
        """ Remove the sheet protection and return a status code """
        if (self.is_for_protection_removal == True):
            self.file_data = remove_range_substring(self.file_data,\
                            self.protection_string, '/>')
            self.is_protected = False
            self.is_for_protection_removal = False
            self.fileDataBytes = self.get_file_data().encode('utf-8')

        return self.fileDataBytes
    
    def check_file_type(self) -> str:
        """ Define the type of file """
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
        if self.get_file_type != 'none':
            start_index = self.file_data.find(self.protection_string)
            if start_index != -1:
                return True  
        
        return False # Substring not found
    
    def is_flagged(self) -> str:
        """ Return * if the sheet is flag for protection removal"""
        string = ''
        if self.get_is_for_protection_removal():
            return '*'
        else:
            return ''


    def print_menuitem(self, menuNumber) -> bool:
        """ Print a menu item depending on the file type"""
        if self.file_type != "none":
            print(f'{str(menuNumber)}. {self.get_file_type():9} \
 |  {self.get_file_state():12} {self.is_flagged():2}  |\
    {self.get_filename_path()}')
            
            return True
        else:
            return False
        
    def findSheetName(self)-> dict:
        startSheets = self.file_data.find('<sheets>') + 8
        endSheets = self.file_data.find('</sheets>')
        sheets = self.file_data[startSheets:endSheets]
        sheetDict = {}
        process = True
        while process == True:
            startSheet = sheets.find('<sheet name=') + 13
            endsheet = sheets.find('sheetId=') - 2
            sheetName = sheets[startSheet:endsheet]
            idStart = sheets.find('r:id=')
            idEnd = sheets.find('"', idStart + 9)
            sheetID = sheets[idStart + 9: idEnd]
            print(sheets)
            print(f'idstart={idStart}, idend={idEnd}, sheetID={sheetID}')
            sheets = sheets[idEnd:]
            if startSheet < 13:
                process = False
            else:
                sheetDict[sheetID] = sheetName

        return sheetDict

        
        

def remove_range_substring(string_to_strip :str, start_substring:str,\
                            end_substring:str) -> str:
    """
    Finds and removes the string between the start_substring and the \n
    end_substring and the the substring themself

    Args:
      string_to_strip: The original string.
      start_substring: The first substring to occurence to find.
      end_substring: The next substring to find

    Returns:
      A new string with the substring removed, or the original string if the
      substring is not found.
    """
    start_index = string_to_strip.find(start_substring)
    if start_index == -1:
        return string_to_strip  # Substring not found
    end_index = string_to_strip.find(end_substring, start_index)
    if end_index == -1:
        return string_to_strip  # Substring not found
    

    return string_to_strip[:start_index] + \
          string_to_strip[end_index + len(end_substring):]

def validate_choice(choice:str, range_of:int) -> bool:
    if choice.isdigit():
        if (int(choice) > 0) and (int(choice) < range_of + 1):
          return True

    return False 

def update_file(zip_filename:str, excel_list):
    with zipfile.ZipFile(zip_filename, mode='w', 
                         compression=zipfile.ZIP_DEFLATED) as excel_zipfile:
        for excel_file in excel_list:
            with excel_zipfile.open(excel_file.get_filename_path(), 'w') \
              as sheet:
                sheet.write(excel_file.execute_change())

