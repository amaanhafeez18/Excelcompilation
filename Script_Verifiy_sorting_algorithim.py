import os
from glob import glob
import re

def extract_range_from_filename(filename):
    # Use regex to find the numeric range in the filename
    match = re.search(r'_(\d+)_(\d+)\.xlsx$', filename)
    if match:
        return (int(match.group(1)), int(match.group(2)))  # Return tuple of (start, end)
    return (0, 0)  # Default if no match found

def sort_excel_files(folder_path):
    # Get a list of all Excel files in the folder
    excel_files = glob(os.path.join(folder_path, '*.xlsx'))
    
    # Sort the files based on the numeric range in their filenames
    excel_files.sort(key=extract_range_from_filename)

    return excel_files

# Test the sorting function
folder_path = 'Retail'  # Replace with the path to your folder with Excel files
sorted_files = sort_excel_files(folder_path)

# Print the sorted file names
print("Sorted Excel Files:")
for file in sorted_files:
    print(os.path.basename(file))
