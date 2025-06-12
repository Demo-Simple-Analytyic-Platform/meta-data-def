## This module provides functions to check if a folder exists and to create a folder if it does not exist.
# meta-def-example/4-processing-python/modules/fso.py

# This module references libraries from the Python Standard Library.
import os

# This module provides functions to check if a folder exists
def folder_exists(folder_path):
    if os.path.isdir(folder_path):
        print(f"The folder '{folder_path}' exists.")
        return True
    else:
        print(f"The folder '{folder_path}' does not exist.")
        return False

# This module provides functions to create a folder if it does not exist
def create_folder(folder_path):
    try:
        os.makedirs(folder_path, exist_ok=True)  # `exist_ok=True` prevents errors if the folder already exists
        print(f"Folder '{folder_path}' created successfully.")
    except Exception as e:
        print(f"Error creating folder '{folder_path}': {e}")