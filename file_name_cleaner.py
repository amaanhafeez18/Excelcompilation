import os

def file_name_clean(folder_path):
    counter_num_files = 0
    # Iterate through all files in the specified folder
    for filename in os.listdir(folder_path):
        if filename.endswith('.csv'):
            # Construct full file path
            index = filename.find('_advancesearch')
            if index != -1:
                # Create the new filename by slicing the original filename
                new_filename = filename[:index]+ '.xlsx'
            else : 
                new_filename = filename + '.xlsx'
            csv_file = os.path.join(folder_path, filename)

            
            # Create the new file name with .xlsx extension
            xlsx_file = os.path.join(folder_path, new_filename)
            
            # Rename the file
            os.rename(csv_file, xlsx_file)
            print(f'Renamed: {filename} to {new_filename}')
            counter_num_files += 1
    print(f"\nThe number of files renamed is: {counter_num_files}\n")

def main():
    folder_path = 'Retail'
    file_name_clean(folder_path)

if __name__ == "__main__":
    main()
