import os

# Specify the path to the folder containing the files
folder_path = r"Retail"
counter_num_files = 0
# Iterate over all files in the folder
for filename in os.listdir(folder_path):
    # Check if the file is an Excel file
    if filename.endswith('.xlsx'):
        # Find the position of '_advancesearch'
        index = filename.find('_advancesearch')
        
        # If '_advancesearch' is found, rename the file
        if index != -1:
            # Create the new filename by slicing the original filename
            new_filename = filename[:index] + '.xlsx'
            # Create the full path for the old and new filenames
            old_file_path = os.path.join(folder_path, filename)
            new_file_path = os.path.join(folder_path, new_filename)
            
            # Rename the file
            os.rename(old_file_path, new_file_path)
            print(f'Renamed: {filename} to {new_filename}')
            counter_num_files =+ 1
print(f"\nThe number of files renamed is: {counter_num_files}\n")
