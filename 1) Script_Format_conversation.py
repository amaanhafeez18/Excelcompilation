import os

# Specify the directory containing the .csv files
folder_path = 'Retail'

# Iterate through all files in the specified folder
for filename in os.listdir(folder_path):
    if filename.endswith('.csv'):
        # Construct full file path
        csv_file = os.path.join(folder_path, filename)
        
        # Create the new file name with .xlsx extension
        xlsx_file = os.path.join(folder_path, filename.replace('.csv', '.xlsx'))
        
        # Rename the file
        os.rename(csv_file, xlsx_file)

print("Renaming complete!")
