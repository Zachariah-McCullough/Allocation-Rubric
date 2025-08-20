import pandas as pd
import os
import glob

# Prompt user for the base directory
directory_path = input("Enter the base directory path: ")
# C:\...\...\The University of Southern Mississippi\...\W Drive\Userfiles\AKale\Resource Allocation Rubric\AY_*_*
# Loads data from the pivot table and the HR survey file into two DataFrames.
# Filters HR data by specific faculty ranks and merges it with the pivot table based on a common ID.
# Adds two new columns: # OF HRS TAUGHT and NEW CALC FTE, where FTE is calculated by dividing hours taught by 12.
# Creates a new pivot table summarizing FTE data by HEGIS Code and Full/Part-Time status.
# Adds row and column totals.
# Saves the result to multiple predefined directories (INSTRUCTIONAL EFFORT PART 1 & The file path), create missing directories (OUTPUT) if needed.

# Validate if the directory exists
if not os.path.isdir(directory_path):
    print(f"The directory '{directory_path}' does not exist. Please check the path and try again.")
    exit()

# Use a wildcard to find the correct DELAWARE_* file
delaware_file_pattern = os.path.join(directory_path, 'DELAWARE_*.xlsx')
delaware_files = glob.glob(delaware_file_pattern)

if not delaware_files:
    print(f"No files found matching pattern 'DELAWARE_*.xlsx' in the directory: {directory_path}")
    exit()

# Select the first matched file (assuming there could be multiple files)
pivot_table_file_path = delaware_files[0]

# Extract the year from the file name
year = os.path.basename(pivot_table_file_path).split('_')[-1].split('.')[0]

# Use a wildcard to find the correct IPEDS HR Component Survey file
ipeds_hr_file_pattern = os.path.join(directory_path, '*_IPEDS_HR_Component_Survey.xlsx')
ipeds_hr_files = glob.glob(ipeds_hr_file_pattern)

if not ipeds_hr_files:
    print(f"No files found matching pattern '*_IPEDS_HR_Component_Survey.xlsx' in the directory: {directory_path}")
    exit()

# Select the first matched IPEDS HR file
ipeds_hr_file_path = ipeds_hr_files[0]

# Load the pivot table from the Output file
try:
    pivot_table_df = pd.read_excel(pivot_table_file_path, sheet_name='Pivot Table')  # Adjust the sheet name if necessary
    print("Pivot Table Columns:", pivot_table_df.columns.tolist())  # Print the actual column names
except Exception as e:
    print(f"Error loading pivot table: {e}")
    exit()

# Load the master_ipeds_hr data from the HR component survey file
try:
    master_ipeds_hr_df = pd.read_excel(ipeds_hr_file_path, sheet_name='MASTER_IPEDS_HR')  # Adjust the sheet name if necessary
    print("Master IPEDS HR Columns:", master_ipeds_hr_df.columns.tolist())  # Print the actual column names
except Exception as e:
    print(f"Error loading master_ipeds_hr data: {e}")
    exit()

# Define the rank values of interest
rank_values_of_interest = [
    'Associate Professor', 'Assistant Professor', 'Instructor', 
    'Lecturer', 'Professor', 'No Rank'
]

# Filter the master_ipeds_hr DataFrame based on the Rank column
master_ipeds_hr_filtered = master_ipeds_hr_df[master_ipeds_hr_df['Rank'].isin(rank_values_of_interest)]

# Merge the pivot table with the filtered master IPEDS HR based on 'ID'
merged_df = pd.merge(master_ipeds_hr_filtered, pivot_table_df[['ID', 'Sum of # OF COURSES TAUGHT']], on='ID', how='left')

# Add two new columns:
merged_df['# OF HRS TAUGHT'] = merged_df['Sum of # OF COURSES TAUGHT']  # Assuming this column contains the hours
merged_df['NEW CALC FTE'] = merged_df['# OF HRS TAUGHT'] / 12

# Drop the redundant 'Sum of # OF COURSES TAUGHT' column
merged_df = merged_df.drop(columns=['Sum of # OF COURSES TAUGHT'])

# Reorder columns to place '# OF HRS TAUGHT' and 'NEW CALC FTE' right after 'ID'
cols = merged_df.columns.tolist()
id_index = cols.index('ID')
cols = cols[:id_index + 1] + ['# OF HRS TAUGHT', 'NEW CALC FTE'] + cols[id_index + 1:]
merged_df = merged_df[cols]

# Create a Pivot Table based on the merged data
pivot_table_result = pd.pivot_table(
    merged_df,
    values='NEW CALC FTE',
    index='HEGIS Code',
    columns='Full/Part',
    aggfunc='sum'
)

# Drop duplicate columns if they exist
pivot_table_result = pivot_table_result.loc[:, ~pivot_table_result.columns.duplicated()]

# Rename the columns to make them clearer
pivot_table_result = pivot_table_result.rename(columns={'F': 'Full-Time', 'P': 'Part-Time'})

# Round the pivot table values to 2 decimal places
pivot_table_result = pivot_table_result.round(2)

# Manually calculate the row and column grand totals
row_totals = pivot_table_result.sum(axis=1).round(2)  # Sum rows for grand total
column_totals = pivot_table_result.sum(axis=0).round(2)  # Sum columns for grand total

# Add the row totals as a new column to the pivot table
pivot_table_result['Grand Total'] = row_totals

# Add the column totals as a new row to the pivot table
column_totals['Grand Total'] = column_totals.sum().round(2)  # Ensure the total for the Grand Total is calculated
pivot_table_result.loc['Grand Total'] = column_totals
# Define the output file paths
output_file_paths = {
    'original': os.path.join(directory_path, f'INSTRUCTIONAL_FTE_{year}.xlsx'),
    'part1': os.path.join(directory_path, 'INSTRUCTIONAL EFFORT PART 1', f'INSTRUCTIONAL_FTE_{year}.xlsx'),
    'part2': os.path.join(directory_path, 'INSTRUCTIONAL EFFORT PART 2', f'INSTRUCTIONAL_FTE_{year}.xlsx'),
    'success': os.path.join(directory_path, 'SUCCESS', f'INSTRUCTIONAL_FTE_{year}.xlsx'),
    'faculty success': os.path.join(directory_path, 'FACULTY SUCCESS', f'INSTRUCTIONAL_FTE_{year}.xlsx'),
    'output': os.path.join(directory_path, 'OUTPUT', f'INSTRUCTIONAL_FTE_{year}.xlsx'),
}

# Ensure all directories exist before saving files
for path in output_file_paths.values():
    dir_name = os.path.dirname(path)
    # Logging the directory status
    if not os.path.exists(dir_name):
        print(f"Directory does not exist: {dir_name}. Creating directory.")
        os.makedirs(dir_name)
    else:
        print(f"Directory exists: {dir_name}.")

# Save the pivot table to the defined file paths
for key, output_file_path in output_file_paths.items():
    try:
        # Saving the pivot table to each specified path
        with pd.ExcelWriter(output_file_path, engine='openpyxl', mode='w') as writer:
            pivot_table_result.to_excel(writer, sheet_name='Pivot Table NEW CALC FTE')
        print(f"Pivot table created successfully: {output_file_path}")
    except Exception as e:
        print(f"Error writing to the Excel file at {output_file_path}: {e}")