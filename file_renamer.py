import os
from pathlib import Path
import glob

def renamer(folder_path, client_name):
    """Renames all .xlsx files in the given folder by appending the client's name to their names."""
    if not folder_path.endswith(os.path.sep):
        folder_path += os.path.sep
    
    excel_files = glob.glob(f"{folder_path}*.xlsx")
    
    for old_file_path in excel_files:
        directory, old_file_name = os.path.split(old_file_path)
        base_name, extension = os.path.splitext(old_file_name)
        new_file_name = f"{base_name}_{client_name}{extension}"
        new_file_path = os.path.join(directory, new_file_name)
        
        try:
            os.rename(old_file_path, new_file_path)
            print(f"Renamed '{old_file_name}' to '{new_file_name}'")
        except Exception as e:
            print(f"Error renaming '{old_file_name}': {e}")
    print("\nFile Rename Successfully")