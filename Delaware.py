# IMPORTANT REMINDER:

# BEFORE YOU BEGIN PROCESSING, READ THE INSTRUCTIONS CAREFULLY AND DOUBLE-CHECK THAT EVERYTHING IS SET UP CORRECTLY. 
# IF THE SETUP IS NOT DONE PROPERLY, IT WILL LEAD TO SIGNIFICANT ISSUES OR ERRORS. 

import pandas as pd
import glob
import os
import re

# Prompt user for the base directory
directory_path = input("Enter the base directory path: ")
# C:\...\...\The University of Southern Mississippi\...\W Drive\Userfiles\AKale\Resource Allocation Rubric\AY_*_*
# This takes the base file 'ET_DELAWARE_STUDY_BASE*.xlsx' (* is a wildcard)
# This script does the following:
# Finds Files: It looks for Excel files starting with ET_DELAWARE_STUDY_BASE (the asterisk * means it can match any characters after that) in the folder you specified.
# Calculates Data: It calculates the number of courses taught by dividing SCH Load by Enrl Load and stores this in a new column called # OF COURSES TAUGHT.
# Creates Pivot Table: It generates a summary (pivot table) that shows:
# The count of classes for each ID.
# The total number of courses taught (with a maximum limit of 3).
# Saves Results: The modified data and pivot table are saved in two locations:
# In the OUTPUT folder.
# In the same folder where the original files were located.

# Validate if the directory exists
if not os.path.isdir(directory_path):
    print(f"The directory '{directory_path}' does not exist. Please check the path and try again.")
    exit()

# Function to load and process ET_DELAWARE_STUDY_BASE* files
def load_and_process_files(directory_path):
    file_pattern = os.path.join(directory_path, 'ET_DELAWARE_STUDY_BASE*.xlsx')
    excel_files = glob.glob(file_pattern)

    if not excel_files:
        print(f"No files found matching pattern 'ET_DELAWARE_STUDY_BASE*' in the directory: {directory_path}")
        return

    for file_path in excel_files:
        print(f"Processing file: {file_path}")
        
        # Extract the year or number from the file name using regex
        match = re.search(r'(\d{4})', os.path.basename(file_path))
        if match:
            year = match.group(1)
            output_file_name = f'DELAWARE_{year}.xlsx'
        else:
            print(f"Could not extract year from file name: {os.path.basename(file_path)}")
            continue
        
        # Save to the original directory
        output_file_path = os.path.join(directory_path, output_file_name)
        print(f"Output file will be: {output_file_path}")
        
        # Create the OUTPUT folder directly inside the base directory if it doesn't exist
        output_folder = os.path.join(directory_path, 'OUTPUT')
        
        if not os.path.exists(output_folder):
            print(f"The OUTPUT folder '{output_folder}' does not exist. Creating it now.")
            os.makedirs(output_folder)
        
        # Save to the OUTPUT folder
        output_file_path_in_output = os.path.join(output_folder, output_file_name)
        print(f"Also writing the file to the OUTPUT folder: {output_file_path_in_output}")
        
        # Sheet name
        sheet_name = 'sheet1'  # Update this if needed

        # Read the Excel sheet into a DataFrame
        try:
            df = pd.read_excel(file_path, sheet_name=sheet_name)
        except ValueError as e:
            print(e)
            # List all available sheets
            xls = pd.ExcelFile(file_path)
            print("Available sheet names:", xls.sheet_names)
            continue

        # Check the columns in the DataFrame
        print("Columns in the DataFrame:", df.columns)

        # Calculate the number of courses taught
        if 'df' in locals():
            df['# OF COURSES TAUGHT'] = df['SCH Load'] / df['Enrl Load']
            
            # Cap the # OF COURSES TAUGHT to 3 for values greater than 3
            df['# OF COURSES TAUGHT'] = df['# OF COURSES TAUGHT'].clip(upper=3)

            # Reorder the columns in the specified order
            column_order = [
                'ID', '# OF COURSES TAUGHT', 'Class Nbr', 'Course ID', 'Section', 
                'Catalog', 'Subject', 'Career', 'Load Factor', 'Tot Enrl', 
                'Tot Hrs C', 'Tot Ghrs', 'Title', 'Min Units', 'Max Units', 
                'Instructor', 'Cls Load', 'Enrl Load', 'SCH Load', 
                'AVG_SCH', 'USM SCH Fr', 'USM SCH So', 'USM SCH Jr', 
                'USM SCH Sr', 'USM SCH Ms', 'USM SCH Sp', 'USM SCH Do', 
                'DEPT_CIP_Code', 'DEPT_CHAIR_EMPLID', 'DEPT_HEAD', 'INSTR_DEPT'
            ]

            # Ensure all columns are present in the DataFrame
            for col in column_order:
                if col not in df.columns:
                    print(f"Column '{col}' is missing from the DataFrame.")
            
            # Reorder the DataFrame
            df = df[column_order]

            # Create a pivot table
            pivot_table = df.pivot_table(
                index=['ID'],  # Replace 'ID' with the actual column name for ID if different
                values=['Class Nbr', '# OF COURSES TAUGHT'],
                aggfunc={
                    'Class Nbr': 'count',  # Count of Class Nbr
                    '# OF COURSES TAUGHT': 'sum'  # Sum of capped hours taught
                }
            ).reset_index()

            # Rename columns for clarity
            pivot_table.rename(columns={'Class Nbr': 'Count of Class Nbr', '# OF COURSES TAUGHT': 'Sum of # OF COURSES TAUGHT'}, inplace=True)

            # Reorder the pivot table columns
            pivot_table = pivot_table[['ID', 'Count of Class Nbr', 'Sum of # OF COURSES TAUGHT']]

            # Save the updated DataFrame and pivot table to both locations
            with pd.ExcelWriter(output_file_path) as writer:
                df.to_excel(writer, sheet_name='Updated Data', index=False)
                pivot_table.to_excel(writer, sheet_name='Pivot Table', index=False)

            with pd.ExcelWriter(output_file_path_in_output) as writer:
                df.to_excel(writer, sheet_name='Updated Data', index=False)
                pivot_table.to_excel(writer, sheet_name='Pivot Table', index=False)

            print("New output files created successfully in both locations.")

# Load and process the files
load_and_process_files(directory_path)