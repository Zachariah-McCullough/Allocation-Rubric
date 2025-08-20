# YOU NEED THE CHANGE THE GRANTS FILE HERE TO BE .............. GRANTS_AY_**_** (** THESE ARE YEARS)
# THIS ONE CAN BE FINNICKY AT BEST, HENCE WHY IT IS ITERATED TO MAKE SURE IT SAVES

import pandas as pd
import os
import glob

# Prompt user for the base directory
directory_path = input("Enter the base directory path: ")
# C:\...\...\The University of Southern Mississippi\IR Office - Documents (1)\W Drive\Userfiles\AKale\Resource Allocation Rubric\AY_**_**\FACULTY SUCCESS
# The script checks for the existence of various files in the specified directory using patterns, such as Applied_Research_AY_*, Awards_AY_*, and others.
# If files matching the patterns are found, it confirms the number of files detected. If no files are found, it prints an appropriate message.
# The script looks for the IPEDS HR file in the base directory and loads the MASTER_IPEDS_HR sheet, converting the ID column to a string and creating a new column, ID_String.
# For each file type, the script attempts to load the corresponding file. If an error occurs while loading a file, the script tries to repair it, particularly for the "High Impact Practices Scheduled Learning" files.
# Replacements are made in the Home Campus/Teaching Site (Most Recent) column (e.g., replacing 'Hattiesburg' with 'HBG').
# It removes the leading 'W' from the USERNAME column if present.
# Existing ID_String and HEGIS Code columns are dropped to avoid duplication.
# A VLOOKUP-like merge is performed with the MASTER_IPEDS_HR sheet using the USERNAME column to add the HEGIS Code.
# After processing and modifying the files, the updated data is saved back into the original file.
# Adding the MASTER_IPEDS_HR Sheet to Grants Files:
# For the Grants files, the MASTER_IPEDS_HR sheet is added to the file without further modifications.
# At the end, the script provides a summary of all the processed files, showing which files were found and processed successfully.

# Define the base directories for FACULTY SUCCESS and MASTER_IPEDS_HR_Component_Survey
base_dir_ipeds = os.path.dirname(directory_path)
base_dir_awards = directory_path

# Locate the Applied Research file
applied_research_pattern = os.path.join(base_dir_awards, 'Applied_Research_AY_*')
applied_research_files = glob.glob(applied_research_pattern)

# Ensure we have found Applied Research file
if not applied_research_files:
    print(f"No Applied Research files found matching pattern: {applied_research_pattern}")
else:
    print(f"Found {len(applied_research_files)} Applied Research files.")

# Locate the Awards file
awards_pattern = os.path.join(base_dir_awards, 'Awards_AY_*')
awards_files = glob.glob(awards_pattern)

# Ensure we have found Awards file
if not awards_files:
    print(f"No Awards files found matching pattern: {awards_pattern}")
else:
    print(f"Found {len(awards_files)} Awards files.")

# Locate the Creative Works file
creative_works_pattern = os.path.join(base_dir_awards, 'Creative_Works_AY_*')
creative_works_files = glob.glob(creative_works_pattern)

# Ensure we have found Creative Works file
if not creative_works_files:
    print(f"No Creative Works files found matching pattern: {creative_works_pattern}")
else:
    print(f"Found {len(creative_works_files)} Creative Works files.")

# Locate the High Impact Practices file
hip_pattern = os.path.join(base_dir_awards, 'High_Impact_Practices_Directed_Service_Learning_AY_*')
hip_files = glob.glob(hip_pattern)

# Ensure we have found High Impact Practices file
if not hip_files:
    print(f"No High Impact Practices files found matching pattern: {hip_pattern}")
else:
    print(f"Found {len(hip_files)} High Impact Practices files.")

# Locate the High Impact Practices Scheduled Learning file
hip_scheduled_pattern = os.path.join(base_dir_awards, 'High_Impact_Practices_Scheduled_Learning_AY_*')
hip_scheduled_files = glob.glob(hip_scheduled_pattern)

# Ensure we have found High Impact Practices Scheduled Learning file
if not hip_scheduled_files:
    print(f"No High Impact Practices Scheduled Learning files found matching pattern: {hip_scheduled_pattern}")
else:
    print(f"Found {len(hip_scheduled_files)} High Impact Practices Scheduled Learning files.")

# Locate the Presentations file
presentation_pattern = os.path.join(base_dir_awards, 'Presentations_AY_*')
presentation_pattern_files = glob.glob(presentation_pattern)

# Ensure we have found Presentations file
if not presentation_pattern_files:
    print(f"No Presentation files found matching pattern: {presentation_pattern}")
else:
    print(f"Found {len(presentation_pattern_files)} Presentations files.")

# Locate the IP file
ip_pattern = os.path.join(base_dir_awards, 'IP_AY_*')
ip_pattern_files = glob.glob(ip_pattern)

# Ensure we have found IP file
if not ip_pattern_files:
    print(f"No IP files found matching pattern: {ip_pattern}")
else:
    print(f"Found {len(ip_pattern_files)} IP files.")

# Locate the Grants file
grant_pattern = os.path.join(base_dir_awards, 'Grants_AY_*')
grant_pattern_files = glob.glob(grant_pattern)

# Ensure we have found Grants file
if not grant_pattern_files:
    print(f"No Grants files found matching pattern: {grant_pattern}")
else:
    print(f"Found {len(grant_pattern_files)} Grants files.")

# Locate the Publications file
publications_pattern = os.path.join(base_dir_awards, 'Publications*.xlsx')
publications_files = glob.glob(publications_pattern)

# Ensure we have found Publications file
if not publications_files:
    print(f"No Publications files found matching pattern: {publications_pattern}")
else:
    print(f"Found {len(publications_files)} Publications files.")

# Locate the Fall IPEDS HR file
ipeds_file_pattern = os.path.join(base_dir_ipeds, 'Fall_*_IPEDS_HR_Component_Survey*.xlsx')
ipeds_files = glob.glob(ipeds_file_pattern)

# Ensure we have found the correct IPEDS HR file
if not ipeds_files:
    print(f"No IPEDS HR files found matching pattern: {ipeds_file_pattern}")
    exit()
else:
    ipeds_file = ipeds_files[0]  # Take the first match
    print(f"Found IPEDS HR file: {ipeds_file}")

# Load the MASTER_IPEDS_HR sheet from the IPEDS HR file
try:
    master_ipeds_df = pd.read_excel(ipeds_file, sheet_name='MASTER_IPEDS_HR')
    print("Successfully loaded the MASTER_IPEDS_HR sheet.")

    # Convert the ID column to string and create a new column next to it
    master_ipeds_df['ID_String'] = master_ipeds_df['ID'].astype(str)
    print("Successfully converted ID to string and created a new column 'ID_String'.")

except Exception as e:
    print(f"An error occurred while loading the MASTER_IPEDS_HR sheet: {e}")
    exit()

# Function to repair and save an Excel file
def repair_and_save_excel(file_path, output_path):
    """
    Try to read a corrupted Excel file and save it to a new file to fix potential corruption.
    """
    try:
        # Load the Excel file
        df = pd.read_excel(file_path)
        print(f"Successfully loaded {file_path}")

        # Save it as a new Excel file
        df.to_excel(output_path, index=False)
        print(f"Successfully saved to {output_path}")

    except Exception as e:
        print(f"Failed to load or save the file: {e}")

# Function to process files
def process_file(file_name, file_type):
    print(f"\nProcessing {file_type} file: {file_name}")
    
    # Load the respective file
    try:
        df = pd.read_excel(file_name)
        print(f"Successfully loaded the {file_type} file.")

    except Exception as e:
        print(f"An error occurred while processing the {file_type} file: {e}")

        # Attempt to repair the file if it's the High Impact Practices Scheduled Learning file
        if 'High Impact Practices Scheduled Learning' in file_type:
            repaired_file_name = file_name.replace('.xlsx', '_repaired.xlsx')
            repair_and_save_excel(file_name, repaired_file_name)
            return  # Skip further processing

        return  # Skip to the next file if there's another error

    # Create a dictionary for replacements in 'Home Campus/Teaching Site (Most Recent)'
    replacements = {
        'Hattiesburg': 'HBG',
        'Gulf Park': 'USMGC',
        'GCRL': 'USMGC',
        'Stennis': 'USMGC'
    }

    # Perform the replacements in the specified column
    if 'Home Campus/Teaching Site (Most Recent)' in df.columns:
        df['Home Campus/Teaching Site (Most Recent)'] = df['Home Campus/Teaching Site (Most Recent)'].replace(replacements)
        print(f"Successfully performed replacements in 'Home Campus/Teaching Site (Most Recent)' column.")

    # Remove leading 'W' from the 'USERNAME' column
    if 'USERNAME' in df.columns:
        df['USERNAME'] = df['USERNAME'].astype(str)  # Convert to string
        df['USERNAME'] = df['USERNAME'].str.lstrip('W')  # Now safe to remove leading 'W'
        print("Successfully removed leading 'W' from 'USERNAME' column.")

    # Drop existing ID_String and HEGIS Code columns to avoid duplication
    if 'ID_String' in df.columns:
        df.drop(columns=['ID_String'], inplace=True)
        print("Dropped existing 'ID_String' column from the DataFrame.")

    if 'HEGIS Code' in df.columns:
        df.drop(columns=['HEGIS Code'], inplace=True)
        print("Dropped existing 'HEGIS Code' column from the DataFrame.")

    # Perform VLOOKUP-like merge with MASTER_IPEDS_HR based on USERNAME
    try:
        df = df.merge(
            master_ipeds_df[['ID_String', 'HEGIS Code']],
            left_on='USERNAME', right_on='ID_String', how='left'
        )
        print("Successfully performed VLOOKUP-like merge and added 'HEGIS Code'.")

    except Exception as e:
        print(f"An error occurred while performing the merge: {e}")
        return  # Skip to the next file if there's an error

    # Save the modified DataFrame back to the original file
    try:
        with pd.ExcelWriter(file_name, engine='openpyxl', mode='w') as writer:
            df.to_excel(writer, sheet_name=file_type, index=False)
            master_ipeds_df.to_excel(writer, sheet_name='MASTER_IPEDS_HR', index=False)

        print(f"Successfully saved changes to the {file_type} file: {file_name}")

    except Exception as e:
        print(f"An error occurred while saving to the {file_type} file: {e}")

# Process Applied Research file
for applied_research_file in applied_research_files:
    process_file(applied_research_file, "Applied Research")

# Process Awards file
for awards_file in awards_files:
    process_file(awards_file, "Awards")

# Process Creative Works file
for creative_works_file in creative_works_files:
    process_file(creative_works_file, "Creative Works")

# Process High Impact Practices file
for hip_file in hip_files:
    process_file(hip_file, "High Impact Practices")

# Process High Impact Practices Scheduled Learning file
for hip_scheduled_file in hip_scheduled_files:
    process_file(hip_scheduled_file, "Scheduled Learning")

# Process Presentations file
for presentation_file in presentation_pattern_files:
    process_file(presentation_file, "Presentations")

# Process IP file
for ip_file in ip_pattern_files:
    process_file(ip_file, "IP")

# Process Grants file
for grant_file in grant_pattern_files:
    try:
        # Open the existing Grants file in write mode and add MASTER_IPEDS_HR sheet
        with pd.ExcelWriter(grant_file, engine='openpyxl', mode='a') as writer:
            master_ipeds_df.to_excel(writer, sheet_name='MASTER_IPEDS_HR', index=False)
        print(f"Successfully added MASTER_IPEDS_HR sheet to the Grants file: {grant_file}")
    except Exception as e:
        print(f"An error occurred while adding MASTER_IPEDS_HR to the Grants file {grant_file}: {e}")

# Process Publications file
for publications_file in publications_files:
    process_file(publications_file, "Publications")

# Additional representation for processed files
print("\nProcessing summary of found files:")
for file_type, files in [ ("Applied Research", applied_research_files), 
                          ("Awards", awards_files), 
                          ("Creative Works", creative_works_files),
                          ("High Impact Practices", hip_files),
                          ("High Impact Practices Scheduled Learning", hip_scheduled_files),
                          ("Presentations", presentation_pattern_files),
                          ("IP", ip_pattern_files), 
                          ("Grants", grant_pattern_files),
                          ("Publications", publications_pattern)]:

    if files:
        print(f"{file_type}:")
        for file in files:
            print(f"- {file}")

print("PROCESSING COMPLETE --- ALL FILES SUCCESSFUL.")