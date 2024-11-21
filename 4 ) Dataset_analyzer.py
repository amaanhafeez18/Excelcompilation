import os
import re
import pandas as pd

def extract_range_from_filename(filename):
    # Using regex to find the numeric range in the filename, so the records are analyzed in correct order
    match = re.search(r'_(\d+)_(\d+)\.xlsx$', filename)
    if match:
        return (int(match.group(1)), int(match.group(2)))  # Return tuple of (start, end)
    return (0, 0)  # Default if no match found

def count_rows_in_company_details(folder_path,sheet_to_analyze):
    total_rows = 0
    counter_2000 = 0
    counter_1999 = 0
    counter_less_than_1999 = 0

    # Get a list of all Excel files in the folder and sort them based on numeric ranges in their filenames
    excel_files = [f for f in os.listdir(folder_path) if f.endswith('.xlsx')]
    excel_files.sort(key=extract_range_from_filename)  # Sort files by numeric range
    print(f"Found {len(excel_files)} Excel files in the folder: {folder_path}")

    for filename in excel_files:
        file_path = os.path.join(folder_path, filename)
        try:
            # Load the workbook
            workbook = pd.ExcelFile(file_path)
            if sheet_to_analyze in workbook.sheet_names:
                # Read the 'Company Details' sheet into a DataFrame
                sheet_data = pd.read_excel(workbook, sheet_to_analyze)
                # Count the number of rows (excluding the header)
                num_rows = sheet_data.shape[0]  # Count rows
                print(f"{filename}: {num_rows} records")
                total_rows += num_rows
                if (num_rows == 2000):
                    counter_2000 += 1
                elif (num_rows == 1999):
                    counter_1999 += 1
                elif (num_rows < 1999):
                    counter_less_than_1999 += 1
                
            else:
                print(f"{filename}: {sheet_to_analyze} sheet not found.")
        except Exception as e:
            print(f"Error reading {filename}: {e}")

    print("The number of records is calculated by subtracting 1 (Header value) from the total number of rows in a given file")
    print(f"Total records for {sheet_to_analyze} across all files is: {total_rows}\n")
    print(f"The total amount of files with records equal to 2000 are: {counter_2000}")
    print(f"The total amount of files with records equal to 1999 are: {counter_1999}")
    print(f"The total amount of files with records less than 1999 are: {counter_less_than_1999}")


# Use the function
# available sheets to analyse 'Company Details' 'Financial Info' 
sheet_to_analyze = 'Executive'
folder_path = 'Retail'  # Replace with the path to your folder with Excel files
count_rows_in_company_details(folder_path,sheet_to_analyze)
