import pandas as pd
import os
import glob
import openpyxl
from openpyxl import load_workbook
import re

# Prompt user for the base directory
directory_path = input("Enter the base directory path: ")
# C:\...\...\The University of Southern Mississippi\IR Office - Documents (1)\W Drive\Userfiles\AKale\Resource Allocation Rubric\AY_**_**\FACULTY SUCCESS
# This code is doing a lot of individual things with a lot of calculations, it does it one at a time, so there shouldn't be too many issues here

###########################################################
# PART 1: Process Applied Research
###########################################################
print("STAGE 1: Processing Applied Research files...")

# Extract the academic year (AY) from the directory path
match = re.search(r'AY_(\d{2})_(\d{2})', directory_path)
if not match:
    raise ValueError("Academic year (AY_XX_XX) not found in the directory path.")
start_year = int(f"20{match.group(1)}")  # e.g., AY_23_24 -> 2023
end_year = int(f"20{match.group(2)}")    # e.g., AY_23_24 -> 2024
print(f"Filtering for years: {start_year} and {end_year}")

# Use glob to find the file in the specified directory
file_path_directed = glob.glob(os.path.join(directory_path, "Applied_Research_AY_*.xlsx"))

# Ensure at least one file is found
if not file_path_directed:
    print("No file found with the specified pattern for Applied Research.")
else:
    # Load the first matching file, specifically the 'Applied Research' sheet
    file_path = file_path_directed[0]
    df_directed = pd.read_excel(file_path, sheet_name='Applied Research')

    # Step 1: Filter rows where TYPE = "Applied"
    df_directed = df_directed[df_directed['TYPE'] == 'Applied']

    # Step 2: Filter for START_START and START_END in the dynamic years
    df_directed['START_START'] = pd.to_datetime(df_directed['START_START'], errors='coerce')  # Ensure datetime
    df_directed['START_END'] = pd.to_datetime(df_directed['START_END'], errors='coerce')      # Ensure datetime
    df_directed = df_directed[
        (df_directed['START_START'].dt.year.isin([start_year, end_year])) & 
        (df_directed['START_END'].dt.year.isin([start_year, end_year]))
    ]

    # Step 3: Standardize the 'Home Campus/Teaching Site (Most Recent)' column
    if 'Home Campus/Teaching Site (Most Recent)' in df_directed.columns:
        # Normalize the case of the column for consistent matching
        df_directed['Home Campus/Teaching Site (Most Recent)'] = df_directed['Home Campus/Teaching Site (Most Recent)'] \
            .str.strip() \
            .str.title()  # Standardize to title case (e.g., "New York" instead of "new york")
        
        # Debug: print out unique values before mapping
        #print("Before mapping 'Home Campus/Teaching Site (Most Recent)':")
        #print(df_directed['Home Campus/Teaching Site (Most Recent)'].unique())

        # Manual mappings for known variations
        campus_name_map = {
            'Hattiesburg': 'HBG',
            'Online': 'HBG',
            'Gcrl': 'USMGC',
            'Stennis': 'USMGC',
            'Mrc': 'USMGC',
            'Gulf Park': 'USMGC'
        }

        # Apply the mappings to standardize the campus names (case-insensitive matching)
        df_directed['Home Campus/Teaching Site (Most Recent)'] = df_directed['Home Campus/Teaching Site (Most Recent)'] \
            .map(lambda x: campus_name_map.get(x, x))  # Default to the original if no mapping is found

        # Debug: print out unique values after mapping
        #print("After mapping 'Home Campus/Teaching Site (Most Recent)':")
        #print(df_directed['Home Campus/Teaching Site (Most Recent)'].unique())
    else:
        print("Column 'Home Campus/Teaching Site (Most Recent)' not found. Skipping standardization.")  

    # Step 4: Create the 'score' column by multiplying ID_String by 1.1
    df_directed['score'] = df_directed['ID_String'] * 1.1

    # Verify the 'score' column has been created correctly
    #print(df_directed[['ID_String', 'score']].head())  # Check if 'score' exists and is correct

    # Step 5: Create the pivot table without 'sum' and with 'score' directly
    pivot_table = df_directed.pivot_table(
        index=['HEGIS Code', 'Home Campus/Teaching Site (Most Recent)'],  # Rows of the pivot table
        values='ID_String',                                            # Only count the 'ID_String' occurrences
        aggfunc='count'                                                # Use count aggregation for 'ID_String'
    )

    # Add the 'score' column to the pivot table (this will be calculated from the count)
    pivot_table['score'] = pivot_table['ID_String'] * 1.1  # Using the count to multiply by 1.1

    # Step 6: Rename the 'ID_String' column to 'count'
    pivot_table = pivot_table.rename(columns={'ID_String': 'count'})

    # Step 7: Save the pivot table back to the Excel file, overriding 'AR Pivot' if it exists
    with pd.ExcelWriter(file_path, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
        pivot_table.to_excel(writer, sheet_name='AR Pivot')

    print(f"Pivot table with 'score' saved to the Excel file in sheet 'AR Pivot'.")

###########################################################
# PART 2: Process Creative Works
###########################################################
print("STAGE 2: Processing Creative Works files...")

# Extract the academic year (AY) from the directory path
match = re.search(r'AY_(\d{2})_(\d{2})', directory_path)
if not match:
    raise ValueError("Academic year (AY_XX_XX) not found in the directory path.")
start_year = int(f"20{match.group(1)}")  # e.g., AY_23_24 -> 2023
end_year = int(f"20{match.group(2)}")    # e.g., AY_23_24 -> 2024
print(f"Filtering for years: {start_year} and {end_year}")

# Use glob to find the file in the specified directory
file_path_creative = glob.glob(os.path.join(directory_path, "Creative_Works_AY_*.xlsx"))

# Ensure at least one file is found
if not file_path_creative:
    print("No file found with the specified pattern for Creative Works.")
else:
    # Load the first matching file, specifically the 'Creative Works' sheet
    file_path = file_path_creative[0]
    df_creative = pd.read_excel(file_path, sheet_name='Creative Works')

    # Step 1: Remove rows with blanks (for all types)
    df_creative = df_creative.dropna(subset=['TYPE'])

    # Step 2: Filter for specific statuses (Presented, Performance, Exhibited, or Published)
    valid_statuses = ['Presented', 'Performed', 'Exhibited', 'Published']
    df_creative = df_creative[df_creative['STATUS'].isin(valid_statuses)]

    # Step 3: Remove rows with blanks in the 'STATUS' column
    df_creative = df_creative.dropna(subset=['STATUS'])

    # Step 4: Filter for 'Academic' in the 'ACADEMIC' column
    df_creative = df_creative[df_creative['ACADEMIC'] == 'Academic']

    # Step 5: Standardize the 'Home Campus/Teaching Site (Most Recent)' column
    if 'Home Campus/Teaching Site (Most Recent)' in df_creative.columns:
        # Normalize the case of the column for consistent matching
        df_creative['Home Campus/Teaching Site (Most Recent)'] = df_creative['Home Campus/Teaching Site (Most Recent)'] \
            .str.strip() \
            .str.title()  
        
        # Debug: print out unique values before mapping
        #print("Before mapping 'Home Campus/Teaching Site (Most Recent)':")
        #print(df_creative['Home Campus/Teaching Site (Most Recent)'].unique())

        # Manual mappings for known variations
        campus_name_map = {
            'Hattiesburg': 'HBG',
            'Online': 'HBG',
            'Gcrl': 'USMGC',
            'Stennis': 'USMGC',
            'Mrc': 'USMGC',
            'Gulf Park': 'USMGC'
        }

        # Apply the mappings to standardize the campus names (case-insensitive matching)
        df_creative['Home Campus/Teaching Site (Most Recent)'] = df_creative['Home Campus/Teaching Site (Most Recent)'] \
            .map(lambda x: campus_name_map.get(x, x))  # Default to the original if no mapping is found

        # Debug: print out unique values after mapping
        #print("After mapping 'Home Campus/Teaching Site (Most Recent)':")
        #print(df_creative['Home Campus/Teaching Site (Most Recent)'].unique())
    else:
        print("Column 'Home Campus/Teaching Site (Most Recent)' not found. Skipping standardization.")

    # Step 6: Filter rows based on the 'START_START' year range (same as in Part 1)
    df_creative['START_START'] = pd.to_datetime(df_creative['START_START'], errors='coerce')  # Ensure datetime
    df_creative = df_creative[
        (df_creative['START_START'].dt.year.isin([start_year, end_year]))
    ]

    # Step 7: Create the pivot table
    pivot_table_creative = df_creative.pivot_table(
        index=['HEGIS Code', 'Home Campus/Teaching Site (Most Recent)'],  # Rows of the pivot table
        values='ID_String',                                       
        aggfunc='count'  # Use count aggregation
    )

    # Add the 'score' column to the pivot table (calculated from the count)
    pivot_table_creative['score'] = pivot_table_creative['ID_String'] * 1.25  # Using count to multiply by 1.25

    # Step 8: Rename the 'ID_String' column to 'count'
    pivot_table_creative = pivot_table_creative.rename(columns={'ID_String': 'count'})

    # Step 9: Save the pivot table back to the Excel file, overriding 'CW Pivot' if it exists
    with pd.ExcelWriter(file_path, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
        pivot_table_creative.to_excel(writer, sheet_name='CW Pivot')

    print(f"Pivot table with 'score' saved to the Excel file in sheet 'CW Pivot'.")
    
###########################################################
# PART 3: Process Presentations
###########################################################
print("STAGE 3: Processing Presentations files...")

# Extract the academic year (AY) from the directory path
match = re.search(r'AY_(\d{2})_(\d{2})', directory_path)
if not match:
    raise ValueError("Academic year (AY_XX_XX) not found in the directory path.")
start_year = int(f"20{match.group(1)}")  # e.g., AY_23_24 -> 2023
end_year = int(f"20{match.group(2)}")    # e.g., AY_23_24 -> 2024
print(f"Filtering for years: {start_year} and {end_year}")

# Use glob to find the file in the specified directory
file_path_presentations = glob.glob(os.path.join(directory_path, "Presentations_AY_*.xlsx"))

if not file_path_presentations:
    print("No file found with the specified pattern for Presentations.")
else:
    # Load the first matching file
    file_path = file_path_presentations[0]
    
    # Load the Presentations sheet
    df_presentations = pd.read_excel(file_path, sheet_name='Presentations')

    # Ensure that the column contains consistent capitalization
    df_presentations['INVACC'] = df_presentations['INVACC'].str.strip().str.capitalize()

    # Map 'Accepted' to 1.0 and 'Invited' to 1.5
    invacc_map = {'Accepted': 1.0, 'Invited': 1.5}
    df_presentations['INVACC'] = df_presentations['INVACC'].map(invacc_map)

    # Handle unmapped or missing values by setting them to 0
    df_presentations['INVACC'] = df_presentations['INVACC'].fillna(0)

    # Debug: Print a summary of the INVACC column
    #print("Mapped INVACC values:")
    #print(df_presentations['INVACC'].value_counts())

    # Step 2: Remove rows with NaN in the 'SCOPE' column
    if 'SCOPE' in df_presentations.columns:
        df_presentations = df_presentations[df_presentations['SCOPE'].notna()]
    else:
        print("Column 'SCOPE' not found. Skipping SCOPE filtering.")

    # Step 3: Filter rows where ACADEMIC column equals "academic"
    if 'ACADEMIC' in df_presentations.columns:
        df_presentations['ACADEMIC'] = df_presentations['ACADEMIC'].str.strip().str.lower()
        df_presentations = df_presentations[df_presentations['ACADEMIC'] == 'academic']
    else:
        print("Column 'ACADEMIC' not found. Skipping academic filtering.")

    # Step 4: Ensure date columns are valid for filtering
    if 'DATE_START' in df_presentations.columns and 'DATE_END' in df_presentations.columns:
        df_presentations['DATE_START'] = pd.to_datetime(df_presentations['DATE_START'], errors='coerce')
        df_presentations['DATE_END'] = pd.to_datetime(df_presentations['DATE_END'], errors='coerce')

        # Filter rows where dates fall within the start and end years
        df_presentations = df_presentations[
            (df_presentations['DATE_START'].dt.year.isin([start_year, end_year])) & 
            (df_presentations['DATE_END'].dt.year.isin([start_year, end_year]))
        ]
    else:
        print("Date columns 'DATE_START' and 'DATE_END' not found. Skipping date filtering.")

    # Step 5: Standardize the 'Home Campus/Teaching Site (Most Recent)' column
    if 'Home Campus/Teaching Site (Most Recent)' in df_presentations.columns:
        # Normalize the case of the column for consistent matching
        df_presentations['Home Campus/Teaching Site (Most Recent)'] = df_presentations['Home Campus/Teaching Site (Most Recent)'] \
            .str.strip() \
            .str.title() 

        # Debug: print out unique values before mapping
        #print("Before mapping 'Home Campus/Teaching Site (Most Recent)':")
        #print(df_presentations['Home Campus/Teaching Site (Most Recent)'].unique())

        # Manual mappings for known variations
        campus_name_map = {
            'Hattiesburg': 'HBG',
            'Online': 'HBG',
            'Gcrl': 'USMGC',
            'Stennis': 'USMGC',
            'Mrc': 'USMGC',
            'Gulf Park': 'USMGC'
        }
        
        # Apply the mappings to standardize the campus names (case-insensitive matching)
        df_presentations['Home Campus/Teaching Site (Most Recent)'] = df_presentations['Home Campus/Teaching Site (Most Recent)'] \
            .map(lambda x: campus_name_map.get(x, x))  # Default to the original if no mapping is found

        # Debug: print out unique values after mapping
        #print("After mapping 'Home Campus/Teaching Site (Most Recent)':")
        #print(df_presentations['Home Campus/Teaching Site (Most Recent)'].unique())
    else:
        print("Column 'Home Campus/Teaching Site (Most Recent)' not found. Skipping standardization.")

    # Step 6: Create the pivot table
    pivot_table_presentations = df_presentations.pivot_table(
        index=['HEGIS Code', 'Home Campus/Teaching Site (Most Recent)'],  # Rows
        values='INVACC',  # Aggregate column
        aggfunc='sum',  # Summing INVACC values
        fill_value=0  # Replace NaN with 0 in the pivot table
    )

    # Step 7: Add a new column for updated INVACC
    pivot_table_presentations['INVACC_Updated'] = pivot_table_presentations['INVACC'] * 1.1

    # Debug: View pivot table
    #print("Pivot Table with INVACC_Updated:")
    #print(pivot_table_presentations)

    # Step 8: Write the pivot table to the Excel file
    with pd.ExcelWriter(file_path, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
        pivot_table_presentations.to_excel(writer, sheet_name='Presentations Pivot')

    print("Presentations Pivot sheet created successfully.")

###########################################################
# PART 4: Process Grants Files
###########################################################
print("STAGE 4: Processing Grants files...")

# Step 1: Load the Grants file and sheets
file_path_grants = glob.glob(os.path.join(directory_path, "Grants_AY_*.xlsx"))
if not file_path_grants:
    print("No file found with the specified pattern for Grants.")
else:
    # Load the first matching file, including both sheets
    file_path = file_path_grants[0]
    excel_data = pd.ExcelFile(file_path)
    
    # Load the Grants sheet
    df_grants = excel_data.parse('Sheet1')
    print("Grants sheet loaded.")
    
    # Load the MASTER_IPEDS_HR sheet (second sheet in the file)
    df_master = excel_data.parse(sheet_name=1)  # Second sheet by position
    print("MASTER_IPEDS_HR sheet loaded.")

    # Step 2: Drop existing 'Location' and 'HEGIS_Code' columns to override them
    df_grants.drop(columns=['Location', 'HEGIS_Code'], errors='ignore', inplace=True)

    # Step 3: Perform the first VLOOKUP for Location
    df_grants = pd.merge(
        df_grants,
        df_master[['ID', 'Location']],  # Explicitly select only required columns
        on='ID',                        # Merge on the ID column
        how='left'                      # Left join to keep all rows in df_grants
    )

    # Step 4: Standardize the 'Location' column in df_grants
    if 'Location' in df_grants.columns:
        # Normalize the case of the column for consistent matching
        df_grants['Location'] = df_grants['Location'].str.strip().str.title() 

        # Debug: print out unique values before mapping
        #print("Before mapping 'Location':")
        #print(df_grants['Location'].unique())

        # Manual mappings for known variations
        location_name_map = {
            'Hattiesburg': 'HBG',
            'Online': 'HBG',
            'Gcrl': 'USMGC',
            'Stennis': 'USMGC',
            'Mrc': 'USMGC',
            'Gulf Park': 'USMGC'
        }
        
        # Apply the mappings to standardize the location names (case-insensitive matching)
        df_grants['Location'] = df_grants['Location'].map(lambda x: location_name_map.get(x, x))  # Default to the original if no mapping is found

        # Debug: print out unique values after mapping
        #print("After mapping 'Location':")
        #print(df_grants['Location'].unique())
    else:
        print("Column 'Location' not found. Skipping standardization.")

    # Step 5: Perform the second VLOOKUP for HEGIS Code
    df_grants = pd.merge(
        df_grants,
        df_master[['ID', 'HEGIS Code']],  # Explicitly select only required columns
        on='ID',                         # Merge on the ID column
        how='left'                       # Left join to keep all rows in df_grants
    )

    # Rename HEGIS Code column for consistency
    df_grants.rename(columns={'HEGIS Code': 'HEGIS_Code'}, inplace=True)

    # Debug: Verify HEGIS Code values
    #print("HEGIS Code values:")
    #print(df_grants['HEGIS_Code'].head())

    # Step 6: Create Pivot Table
    #print("Creating Pivot Table...")
    pivot_table = pd.pivot_table(
        df_grants,
        values='ID',         # The column to aggregate
        index=['HEGIS_Code', 'Location'],  # Rows: HEGIS_Code and Location
        aggfunc='count',     # Count the number of IDs
        fill_value=0         # Replace NaN with 0 in the result
    )

    # Convert Pivot Table to DataFrame for Writing
    pivot_table_df = pivot_table.reset_index()

    # Add a new column to the Pivot Table that multiplies the count of IDs by 1.1
    pivot_table_df['ID x 1.1'] = pivot_table_df['ID'] * 1.1

    # Step 7: Write the updated Grants sheet back to the same file and the Pivot Table to a new sheet
    with pd.ExcelWriter(file_path, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
        df_grants.to_excel(writer, sheet_name='Sheet1', index=False)  # Write the Grants sheet
        pivot_table_df.to_excel(writer, sheet_name='GN Pivot', index=False)  # Write the Pivot Table

    #print("Grants sheet updated with both VLOOKUPs and standardized Location names.")
    print("Pivot Table saved to 'GN Pivot' sheet with 'ID x 1.1' column.")

###########################################################
# PART 5: Process Awards Files
###########################################################
print("STAGE 5: Processing Awards files...")

# Step 1: Load the Awards file and sheets
file_path_awards = glob.glob(os.path.join(directory_path, "Awards_AY_*.xlsx"))
if not file_path_awards:
    print("No file found with the specified pattern for Awards.")
else:
    # Load the first matching file
    file_path = file_path_awards[0]
    excel_data = pd.ExcelFile(file_path)
    
    # Load the Awards sheet
    df_awards = excel_data.parse('Awards')
    print("Awards sheet loaded.")
    
    # Load the MASTER_IPEDS_HR sheet (second sheet in the file)
    df_master = excel_data.parse(sheet_name=1)  # Second sheet by position
    print("MASTER_IPEDS_HR sheet loaded.")

    # Step 2: Convert `ID_String` to numeric
    # If conversion fails, replace with NaN
    df_awards['ID_String'] = pd.to_numeric(df_awards['ID_String'], errors='coerce')

    # Debug: Check the conversion
    #print("Converted `ID_String` to numeric:")
    #print(df_awards['ID_String'].head())

    # Step 3: Perform the first VLOOKUP for Location
    df_awards = pd.merge(
        df_awards,
        df_master[['ID_String', 'Location']],  # Explicitly select only required columns
        on='ID_String',                        # Merge on the ID column
        how='left'                      # Left join to keep all rows in df_awards
    )

    # Step 4: Standardize the 'Home Campus/Teaching Site (Most Recent)' column in df_awards
    if 'Location' in df_awards.columns:
        # Normalize the case of the column for consistent matching
        df_awards['Location'] = df_awards['Location'] \
            .str.strip() \
            .str.title()  # Standardize to title case (e.g., "New York" instead of "new york")
        
        # Debug: print out unique values before mapping
        #print("Before mapping 'Location':")
        #print(df_awards['Location'].unique())

        # Manual mappings for known variations
        campus_name_map = {
            'Hattiesburg': 'HBG',
            'Online': 'HBG',
            'Gcrl': 'USMGC',
            'Stennis': 'USMGC',
            'Mrc': 'USMGC',
            'Gulf Park': 'USMGC'
        }
        
        # Apply the mappings to standardize the campus names (case-insensitive matching)
        df_awards['Location'] = df_awards['Location'] \
            .map(lambda x: campus_name_map.get(x, x))  # Default to the original if no mapping is found

        # Debug: print out unique values after mapping
        #print("After mapping 'Location':")
        #print(df_awards['Location'].unique())
    else:
        print("Column 'Location' not found. Skipping standardization.")

    # Step 5: Apply filters for NOMREC and SCOPE
    df_filtered = df_awards[ 
        (df_awards['NOMREC'] == 'Received') & 
        (df_awards['SCOPE'].isin(['Scholarship/Creative Works/Research']))
    ]

    # Debug: Verify the filtered data
    #print("Filtered data based on NOMREC and SCOPE:")
    #print(df_filtered[['NOMREC', 'SCOPE', 'ID_String']].head())

    # Step 6: Create a pivot table with filtered data
    pivot_table = pd.pivot_table(
        df_filtered,
        values='ID_String',             # Count of ID_String in values
        index=['HEGIS Code', 'Location'],  # Rows: HEGIS_Code first, then Location
        aggfunc='count'                # Aggregation: count
    )

    # Debug: View the pivot table
    #print("Pivot table created with filtered data.")
    #print(pivot_table)

    # Step 6.1: Add a new column with values multiplied by 1.1
    pivot_table['ID_String_Multiplied'] = pivot_table['ID_String'] * 1.1

    # Debug: View the updated pivot table with the new column
    #print("Updated pivot table with new column (multiplied by 1.1):")
    #print(pivot_table)

    # Step 7: Write both the filtered Awards sheet and the pivot table to the same file
    with pd.ExcelWriter(file_path, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
        # Write filtered Awards sheet
        df_filtered.to_excel(writer, sheet_name='Awards_Filtered', index=False)
        
        # Write pivot table
        pivot_table.to_excel(writer, sheet_name='Awards Pivot')

    print("Awards sheet and pivot table updated successfully.")

###########################################################
# PART 6: Process IP Files
###########################################################
print("STAGE 6: Processing IP files...")

# Step 1: Load the IP file and sheets
file_path_IP = glob.glob(os.path.join(directory_path, "IP_AY_*.xlsx"))
if not file_path_IP:
    print("No file found with the specified pattern for IP.")
else:
    # Load the first matching file
    file_path = file_path_IP[0]
    excel_data = pd.ExcelFile(file_path)
    
    # Load the IP sheet
    df_IP = excel_data.parse('IP')
    print("IP sheet loaded.")
    
    # Load the MASTER_IPEDS_HR sheet (second sheet in the file)
    df_master = excel_data.parse(sheet_name=1)  # Second sheet by position
    print("MASTER_IPEDS_HR sheet loaded.")

    # Step 2: Ensure 'ID_String' exists in both sheets and merge for 'Location'
    if 'ID_String' in df_IP.columns and 'ID_String' in df_master.columns:
        df_IP = pd.merge(
            df_IP,
            df_master[['ID_String', 'Location']],  # Explicitly select only required columns
            on='ID_String',                        # Merge on the ID column
            how='left'                             # Left join to keep all rows in df_IP
        )
        print("Merged 'Location' from MASTER_IPEDS_HR sheet.")
    else:
        print("Error: 'ID_String' column missing in one of the dataframes.")

    # Step 3: Standardize the 'Home Campus/Teaching Site (Most Recent)' column in df_IP
    if 'Home Campus/Teaching Site (Most Recent)' in df_IP.columns:
        # Normalize the case of the column for consistent matching
        df_IP['Home Campus/Teaching Site (Most Recent)'] = df_IP['Home Campus/Teaching Site (Most Recent)'] \
            .str.strip() \
            .str.title()  # Standardize to title case (e.g., "New York" instead of "new york")
        
        # Debug: print out unique values before mapping
        #print("Before mapping 'Home Campus/Teaching Site (Most Recent)':")
        #print(df_IP['Home Campus/Teaching Site (Most Recent)'].unique())

        # Manual mappings for known variations
        campus_name_map = {
            'Hattiesburg': 'HBG',
            'Online': 'HBG',
            'Gcrl': 'USMGC',
            'Stennis': 'USMGC',
            'Mrc': 'USMGC',
            'Gulf Park': 'USMGC'
        }
        
        # Apply the mappings to standardize the campus names (case-insensitive matching)
        df_IP['Location'] = df_IP['Location'] \
            .map(lambda x: campus_name_map.get(x, x))  # Default to the original if no mapping is found

        # Debug: print out unique values after mapping
        #print("After mapping 'Location)':")
        #print(df_IP['Location'].unique())
    else:
        print("Column 'Location' not found. Skipping standardization.")

    # Step 4: Check if 'APPROVE_START' exists in the IP sheet before creating the pivot table
    if 'APPROVE_START' in df_IP.columns:
        # Create the pivot table only if 'APPROVE_START' column exists
        pivot_table_ip = pd.pivot_table(
            df_IP,
            values='APPROVE_START', # Count of APPROVE_START in values
            index=['HEGIS Code', 'Location'],  # Rows: HEGIS Code
            aggfunc='count' # Aggregation: count
        )
        print("Pivot table for IP created with HEGIS Code on rows and APPROVE_START count as values.")
        #print(pivot_table_ip)
        
        # Step 4.1: Add a new column 'Score' with values multiplied by 0.1
        pivot_table_ip['Score'] = pivot_table_ip['APPROVE_START'] * 0.1
        print("New 'Score' column added to the pivot table.")
        #print(pivot_table_ip[['APPROVE_START', 'Score']].head())

        # Step 5: Write both the filtered IP sheet and the pivot table to the same file
        with pd.ExcelWriter(file_path, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
            # Write the IP sheet (or use df_IP if needed)
            df_IP.to_excel(writer, sheet_name='IP_Filtered', index=False)
            
            # Write the IP pivot table to a new sheet 'IP Pivot'
            pivot_table_ip.to_excel(writer, sheet_name='IP Pivot')

        print("IP sheet and IP pivot table updated successfully.")
    else:
        print("Error: 'APPROVE_START' column is missing in the IP sheet. Pivot table creation aborted.")

###########################################################
# PART 7: Process Publications Files
###########################################################
print("STAGE 7: Processing Publications files...")

# Step 1: Load the Publications file and sheets
file_path_publications = glob.glob(os.path.join(directory_path, "Publications_AY_*.xlsx"))
if not file_path_publications:
    print("No file found with the specified pattern for Publications.")
else:
    # Load the first matching file
    file_path = file_path_publications[0]
    excel_data = pd.ExcelFile(file_path)
    
    # Load the Publications sheet
    df_publications = excel_data.parse('Publications')
    print("Publications sheet loaded.")
    
    # Load the MASTER_IPEDS_HR sheet (second sheet in the file)
    df_master = excel_data.parse(sheet_name=1)  # Second sheet by position
    print("MASTER_IPEDS_HR sheet loaded.")

    # Step 2: Ensure 'ID_String' exists in both sheets and merge for 'Location'
    if 'ID_String' in df_publications.columns and 'ID_String' in df_master.columns:
        df_publications = pd.merge(
            df_publications,
            df_master[['ID_String', 'Location']],  # Explicitly select only required columns
            on='ID_String',                        # Merge on the ID column
            how='left'                             # Left join to keep all rows in df_publications
        )
        print("Merged 'Location' from MASTER_IPEDS_HR sheet.")
    else:
        print("Error: 'ID_String' column missing in one of the dataframes.")

    # Step 3: Calculate contype_score
    if 'CONTYPE' in df_publications.columns:
        df_publications['contype_score'] = df_publications['CONTYPE'].apply(lambda x: 2 if x == 'Book' else 1)
    else:
        print("Column 'CONTYPE' not found. Skipping contype score calculation.")
        df_publications['contype_score'] = 0  # Default value if missing

    # Step 4: Calculate student_level_score using the updated logic
    student_level_columns = [
        col for col in df_publications.columns
        if col.startswith('INTELLCONT_AUTH_') and col.endswith('STUDENT_LEVEL')
    ]
    # Ensure 'student_level_score' is a float column
    df_publications['student_level_score'] = 0.0

    for index, row in df_publications.iterrows():
        student_level_check = sum(
            (row[col] == 'Graduate') or (row[col] == 'Undergraduate') 
            for col in student_level_columns if pd.notnull(row[col])
        )
        if student_level_check > 0:
            df_publications.at[index, 'student_level_score'] = 1.5
    print(f"Student level scores calculated across {len(student_level_columns)} columns.")

    # Step 5: Calculate the total_score
    if 'contype_score' in df_publications.columns and 'student_level_score' in df_publications.columns:
        df_publications['total_score'] = df_publications['contype_score'] + df_publications['student_level_score']
        print("Total score calculated and added to the DataFrame.")

    # Save the updated DataFrame back to a new Excel file to avoid issues
    new_file_path = file_path.replace(".xlsx", "_updated.xlsx")
    with pd.ExcelWriter(new_file_path, engine='openpyxl', mode='w') as writer:
        df_publications.to_excel(writer, sheet_name='Updated_Publications', index=False)
    print(f"Updated data saved to new file: {new_file_path}")

    # Step 6: Create Pivot Table
    try:
        # Filtered DataFrame
        filtered_df = df_publications[ 
            (df_publications['STATUS'] == 'Published') & 
            (df_publications['REFEREED'].isin(['Refereed', 'Peer-Reviewed']))
        ]

        # Verify column names for pivot table
        if not all(col in filtered_df.columns for col in ['HEGIS Code', 'Location_y', 'total_score']):
            raise KeyError("Missing one or more columns required for the pivot table: ['HEGIS Code', 'Location_y', 'total_score']")

        # Pivot table creation
        pivot_table = pd.pivot_table(
            filtered_df,
            values='total_score',         
            index=['HEGIS Code', 'Location_y'],  # Updated column names
            aggfunc='sum'
        ).reset_index()
        print("Pivot table created:\n", pivot_table.head())

        # Ensure correct data types for the pivot table columns
        pivot_table['total_score'] = pivot_table['total_score'].astype(float)
        pivot_table['adjusted_total_score'] = pivot_table['total_score'] * 1.25

        # Save the pivot table to the new file
        with pd.ExcelWriter(new_file_path, engine='openpyxl', mode='a') as writer:
            pivot_table.to_excel(writer, sheet_name='Pivot_Table', index=False)
        print("Pivot table saved to new file.")

        # Force Excel to recalculate and save the workbook to ensure everything is updated
        workbook = openpyxl.load_workbook(new_file_path)
        workbook.save(new_file_path)
        print(f"Workbook saved and recalculated: {new_file_path}")

    except KeyError as ke:
        print(f"KeyError during Pivot Table creation: {ke}")
    except Exception as e:
        print(f"Error during Pivot Table creation: {e}")