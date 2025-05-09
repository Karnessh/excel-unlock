import zipfile

def open_file_in_zipfile(current_zipfile:zipfile.ZipFile, file:str)-> None:
    """
    


    """
    filedata = current_zipfile.read(file).decode("utf-8")
    filedata = remove_range_substring(filedata, '<sheetProtection', '/>')
    for line in filedata.split('\n') :
        print(line)

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
    
