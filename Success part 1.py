import pandas as pd
import glob
import os

# Prompt user for the base directory
directory_path = input("Enter the base directory path: ")
# C:\...\...\The University of Southern Mississippi\...\W Drive\Userfiles\AKale\Resource Allocation Rubric\AY_**_**\SUCCESS
# The script uses glob to find all files matching the pattern ET_RAF_ENROLLMENT_*.xlsx in the specified directory.
# Each file is processed to extract specific columns (ID, HEGIS Code, Term, Acad Org, Org Descr, Acad Group, Pri Prog Camp) and any rows with missing values are dropped.
# The data from all files is concatenated into a single DataFrame.
# The script replaces the value 'ONLNE' with 'HBG' in the Pri Prog Camp column to standardize the values.
# The processed data is saved as INSTRUCTIONAL_EFFORT_PART_2.xlsx in both the base directory and an OUTPUT folder (if it doesn't already exist).
# A pivot table is created from the combined data, with Org Descr, HEGIS Code, and Pri Prog Camp as the index and ID counted as the value.
# The script attempts to find a file matching INSTRUCTIONAL_FTE_*.xlsx, loads the Grand Total data from its Pivot Table NEW CALC FTE sheet, and merges it with the pivot table.
# The SCH/FTE is calculated by dividing the count of ID by the Grand Total (if the grand total is greater than zero). The SCORE is calculated as SCH/FTE * 0.20.
# For each unique HEGIS Code, the script aggregates data and creates total rows that sum the Total ID and average the SCH/FTE and SCORE.
# The total rows are appended to the merged table, and the table is sorted by HEGIS Code and Pri Prog Camp.
# The final merged table is written to an Excel file.
# The script sets formatting for the header row, numbers (to two decimal places) in relevant columns, and applies the required styles.
# A summary pivot table is created, aggregating the Total ID, Grand Total, SCH/FTE, and SCORE by HEGIS Code and Pri Prog Camp. The columns are flattened for readability.

# Validate if the directory exists
if not os.path.isdir(directory_path):
    print(f"The directory '{directory_path}' does not exist. Please check the path and try again.")
    exit()

# Get the parent directory from the base path
parent_directory = os.path.dirname(directory_path)

# Check for the OUTPUT folder in the parent directory
output_folder = os.path.join(parent_directory, 'OUTPUT')
if not os.path.exists(output_folder):
    print(f"The OUTPUT folder '{output_folder}' does not exist. Please check the parent directory.")
    exit()
else:
    print(f"OUTPUT folder found at: {output_folder}")

# Define the file pattern for ET_RAF_COMPLETIONS with wildcard for .xlsx files
file_pattern = os.path.join(directory_path, "ET_RAF_COMPLETIONS_*.xlsx")

# List all matching files
files = glob.glob(file_pattern)

# Check if any files were found
if not files:
    print(f"No files matching '{file_pattern}' were found in the directory.")
    exit()

# Load and merge all files into a single DataFrame
df_list = []
for file in files:
    try:
        df = pd.read_excel(file, header=1)  # Read .xlsx files and use the second row as header
        df_list.append(df)
    except Exception as e:
        print(f"Error reading {file}: {e}")

# Proceed only if we have data to merge
if df_list:
    merged_data = pd.concat(df_list, ignore_index=True)
else:
    print("No data to merge.")
    exit()

# Replace 'Online' with 'HBG' in the 'HEGIS Code' column
merged_data['Campus'] = merged_data['Campus'].replace('ONLNE', 'HBG')

# Load the INSTRUCTIONAL_FTE data
instructional_fte_pattern = os.path.join(directory_path, "INSTRUCTIONAL_FTE_*.xlsx")
instructional_fte_files = glob.glob(instructional_fte_pattern)

if instructional_fte_files:
    try:
        instructional_fte_data = pd.read_excel(instructional_fte_files[0], header=0)  # Assuming there's only one file
    except Exception as e:
        print(f"Error reading INSTRUCTIONAL_FTE file: {e}")
        exit()
else:
    print("No INSTRUCTIONAL_FTE files found.")
    exit()

# Create a list of required columns to verify if they exist
required_columns = ['Org Descr', 'HEGIS Code', 'Campus']
missing_columns = [col for col in required_columns if col not in merged_data.columns]

if missing_columns:
    print(f"Missing columns in merged data: {missing_columns}")
else:
   # Create the pivot table from merged_data
    pivot_table = merged_data.groupby(['Org Descr', 'HEGIS Code', 'Campus']).size().reset_index(name='Count_ID')

# Calculate total counts for each HEGIS_Code
total_counts = pivot_table.groupby('HEGIS Code')['Count_ID'].sum().reset_index()
total_counts['Org Descr'] = 'TOTAL'  # Set Org_Descr as TOTAL for total rows
total_counts['Campus'] = ''  # Set Campus as empty for total rows

# Combine the pivot table with total counts
combined_output = pd.concat([pivot_table, total_counts], ignore_index=True)

# Merge with INSTRUCTIONAL_FTE to get 'Grand Total' next to 'Count_ID'
final_output = pd.merge(combined_output, instructional_fte_data[['HEGIS Code', 'Grand Total']],
                         on='HEGIS Code', how='left')

# Calculate the ratio of Count_ID to Grand Total
final_output['Ratio_Count_ID_to_Grand_Total'] = final_output['Count_ID'] / final_output['Grand Total']

# Multiply the ratio by 0.175
final_output['Adjusted_Ratio'] = final_output['Ratio_Count_ID_to_Grand_Total'] * 0.175

# Sort the final output for better presentation
final_output = final_output.sort_values(by=['HEGIS Code', 'Org Descr'], ascending=[True, True])

# Modify the current pivot table structure to have one line for each HEGIS Code
summary_pivot_table = final_output.pivot_table(
    index='HEGIS Code', 
    columns='Campus', 
    values=['Count_ID', 'Ratio_Count_ID_to_Grand_Total', 'Adjusted_Ratio'], 
    aggfunc='sum',
    fill_value=0.00  # This ensures that any missing values are filled with 0
)

# Flatten the multi-level columns
summary_pivot_table.columns = ['_'.join(filter(None, col)).strip() for col in summary_pivot_table.columns]

# Reset the index to bring 'HEGIS Code' back as a column
summary_pivot_table.reset_index(inplace=True)

# Replace any remaining NaN values with 0.00 (although the fill_value=0 should already handle most cases)
summary_pivot_table.fillna(0.00, inplace=True)

# Define the output file name
output_file_name = "SUCCESS_PART_1.xlsx"

# Define the output file paths
output_file_path = os.path.join(directory_path, output_file_name)  # Save in the base directory
output_file_path_output_folder = os.path.join(output_folder, output_file_name)  # Save in the OUTPUT folder

# Save the Excel files to both locations
print(f"Saving to: {output_file_path} and {output_file_path_output_folder}")

with pd.ExcelWriter(output_file_path) as writer:
    merged_data.to_excel(writer, sheet_name='Merged Data', index=False)  # Save merged data
    final_output.to_excel(writer, sheet_name='Pivot Table', index=False)  # Save final output
    summary_pivot_table.to_excel(writer, sheet_name='Summary', index=False)
 
with pd.ExcelWriter(output_file_path_output_folder) as writer:
    merged_data.to_excel(writer, sheet_name='Merged Data', index=False)  # Save merged data
    final_output.to_excel(writer, sheet_name='Pivot Table', index=False)  # Save final output
    summary_pivot_table.to_excel(writer, sheet_name='Summary', index=False)

# Print a confirmation message
print(f"Output successfully written to: {output_file_path}")
print(f"Output successfully written to: {output_file_path_output_folder}")