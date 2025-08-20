import pandas as pd
import glob
import os

# Prompt user for the base directory
directory_path = input("Enter the base directory path: ")
# C:\...\...\The University of Southern Mississippi\...\W Drive\Userfiles\AKale\Resource Allocation Rubric\AY_**_**\SUCCESS
# It loads data from two sources:
# JR Graduation Rate CSV file (JR Graduation Rate_Full Data_data.csv).
# Multiple Excel files matching the ET_RAF_COMPLETIONS_*.xlsx pattern.
# The script merges the data based on the Discipline Desc column between the JR Graduation Rate data and ET RAF Completions data.
# It processes the Campus column by replacing "Online" with "Hattiesburg" and ensures the column names are free of extra spaces for correct matching.
# The script filters the data based on a specific column (JR Retention Year shift - Split 1) that contains '4'.
# A pivot table is created counting Student ID values and focusing on whether the degree was completed (Completed Degree (Y/N)).
# It calculates:
# The total number of students (Total).
# Graduation rate (JR GRAD RATE).
# A final Score based on the graduation rate, multiplying it by 0.175.
# A row for the sum of each HEGIS Code is inserted into the pivot table.
# Missing columns are handled by filling them with default values (e.g., 0).
# The data is aggregated into a summary pivot table, which includes metrics like:
# Total student counts (Y, Total).
# Graduation rates (JR GRAD RATE).
# Scores (Score).
# The final summary table is renamed for clarity and presentation.
# The processed data is saved to an Excel file (SUCCESS_PART_2.xlsx) both in the OUTPUT folder and the base directory.
# Multiple sheets are written into the Excel file:
# JR Graduation Rate: Merged JR Graduation Rate data.
# ET RAF Completions: Merged ET RAF data.
# pivot JR rate: Pivot table with graduation rates and scores.
# Summary: Summary pivot table with aggregated metrics.

# Validate if the directory exists
if not os.path.isdir(directory_path):
    print("Directory does not exist. Exiting...")
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

# File name for the output Excel in the OUTPUT folder
output_file_name = 'SUCCESS_PART_2.xlsx'
output_file_path = os.path.join(output_folder, output_file_name)

# Also set the path for saving in the base directory
base_output_file_path = os.path.join(directory_path, output_file_name)

# Load the JR Graduation Rate Data
jr_grad_file = os.path.join(directory_path, 'JR Graduation Rate_Full Data_data.csv')

# Check if the JR Graduation Rate file exists
if not os.path.isfile(jr_grad_file):
    print("JR Graduation Rate file does not exist. Exiting...")
    exit()

# Load the JR Graduation Rate file
jr_data = pd.read_csv(jr_grad_file)

# Load ET_RAF_COMPLETIONS Data using a wildcard
et_files = glob.glob(os.path.join(directory_path, "ET_RAF_COMPLETIONS_*.xlsx"))
et_data_list = []

# Check if ET_RAF_COMPLETIONS files exist
if not et_files:
    print("No ET_RAF_COMPLETIONS files found. Exiting...")
    exit()

# Load each ET_RAF_COMPLETIONS file, skipping the first row
for file in et_files:
    et_data = pd.read_excel(file, skiprows=1)  # Skip the first row
    et_data_list.append(et_data[['HEGIS Code', 'Discipline Desc']])

# Concatenate all loaded ET data and drop duplicates
merged_et_data = pd.concat(et_data_list, ignore_index=True).drop_duplicates()

# Rename columns for clarity
jr_data.rename(columns={'Primary Discipline': 'Discipline Desc'}, inplace=True)

# Replace 'Online' with 'Hattiesburg' in the 'Campus' column of JR Graduation Rate data
jr_data['Campus'] = jr_data['Campus'].replace('Online', 'Hattiesburg')

# Merge the dataframes on Discipline Desc
merged_data = pd.merge(jr_data, merged_et_data, on='Discipline Desc', how='left')

# Strip whitespace from column names for clean matching
merged_data.columns = merged_data.columns.str.strip()

# Check for the specific column with spaces included
expected_column = 'JR Retention  Year shift - Split 1'  # Note the extra spaces
if expected_column not in merged_data.columns:
    print(f"Column matching '{expected_column}' not found in merged_data. Exiting...")
    exit()

# Create COMPLETED DEGREE (Y/N) column based on Degree Completion Term
merged_data['COMPLETED DEGREE (Y/N)'] = merged_data['Degree Completion Term'].notnull().replace({True: 'Y', False: 'N'})

# Filter the data based on the identified retention column containing '4'
filtered_data = merged_data[merged_data[expected_column].astype(str).str.contains('4', case=False, na=False)]

# Create the pivot table from the filtered data
pivot_table = filtered_data.pivot_table(
    index=['HEGIS Code', 'Campus'],  # Rows (removed 'Primary School')
    columns='COMPLETED DEGREE (Y/N)',  # Columns for Completed Degree
    values='Student ID',  # Values to count
    aggfunc='count',  # Count of Student ID
    fill_value=0  # Fill missing values with 0
)

# Focus only on 'Y' and create a total column for completed degrees
pivot_table['Total'] = pivot_table.get('Y', 0) + pivot_table.get('N', 0)  # Total counts Y and N

# Create the JR GRAD RATE column by calculating the ratio of Y to Total
pivot_table['JR GRAD RATE'] = pivot_table['Y'] / pivot_table['Total']

# Add a new column 'Score' based on the 'JR GRAD RATE' column
pivot_table['Score'] = pivot_table['JR GRAD RATE'] * 0.175

# Drop the 'N' column if it exists
pivot_table.drop(columns=['N'], inplace=True, errors='ignore')

# Add a sum row for each HEGIS Code by grouping by 'HEGIS Code' and summing
hegis_sums = pivot_table.groupby('HEGIS Code').sum()

# Insert 'Total' row under each HEGIS Code
def insert_sums(df, sums):
    result = pd.DataFrame()
    for hegis_code in df.index.get_level_values('HEGIS Code').unique():
        temp_df = df.xs(hegis_code, level='HEGIS Code', drop_level=False)
        result = pd.concat([result, temp_df])
        sum_row = pd.DataFrame(sums.loc[hegis_code]).T
        sum_row.index = pd.MultiIndex.from_tuples([(hegis_code, 'SUM')], names=['HEGIS Code', 'Campus'])
        result = pd.concat([result, sum_row])
    return result

# Call the function to insert sum rows under each HEGIS Code
pivot_table_with_sums = insert_sums(pivot_table, hegis_sums)

# Reset the index to convert MultiIndex to columns for saving
pivot_table_with_sums.reset_index(inplace=True)

# Create final_output based on pivot_table_with_sums
final_output = pivot_table_with_sums.copy()  # Copy the pivot table for further processing

# Check if required columns exist
required_columns = ['HEGIS Code', 'Campus', 'Y', 'Total', 'JR GRAD RATE', 'Score']
missing_columns = [col for col in required_columns if col not in final_output.columns]

if missing_columns:
    print(f"Missing columns in final_output: {missing_columns}")
    
    # Handle missing columns by defining them with default values
    for col in missing_columns:
        final_output[col] = 0  # Or any logic to fill missing columns

# Create a summary pivot table to hold the required metrics
summary_pivot_table = final_output.pivot_table(
    index='HEGIS Code',
    columns='Campus',
    values=['Y', 'Total', 'JR GRAD RATE', 'Score'],  # Ensure all columns are valid
    aggfunc='sum',
    fill_value=0.00
)

# Reset index for better readability
summary_pivot_table.reset_index(inplace=True)

# Optional: Rename columns for better clarity
summary_pivot_table.columns = ['HEGIS Code', 'Hattiesburg Count JR', ' Total JR', 
                                'USM Gulf Coast JR', 'Hattiesburg Score', 'Total Scores', 
                                'USM Gulf Coast Score', 'Total Hattiesburg', 'Sum Total', 
                                'USM Gulf Coast Total', 'Hattiesburg Y', 'Sum Y', 'USM Gulf Coast']

# Display the resulting summary pivot table
print("Summary Pivot Table:")
print(summary_pivot_table)

# Function to write DataFrames to Excel
def write_to_excel(writer, sheet_name, df):
    if sheet_name in writer.sheets:
        df.to_excel(writer, sheet_name=sheet_name, index=False, if_sheet_exists='replace')
    else:
        df.to_excel(writer, sheet_name=sheet_name, index=False)

# Save the Excel files to the OUTPUT folder and base directory
print(f"Saving to: {output_file_path}")
print(f"Also saving to: {base_output_file_path}")

# Write DataFrames to the same Excel file in the OUTPUT folder
with pd.ExcelWriter(output_file_path, engine='openpyxl', mode='w') as writer:
    write_to_excel(writer, 'JR Graduation Rate', merged_data)
    write_to_excel(writer, 'ET RAF Completions', merged_et_data)
    write_to_excel(writer, 'pivot JR rate', pivot_table_with_sums)
    write_to_excel(writer, 'Summary', summary_pivot_table)

# Write DataFrames to the same Excel file in the base directory
with pd.ExcelWriter(base_output_file_path, engine='openpyxl', mode='w') as writer:
    write_to_excel(writer, 'JR Graduation Rate', merged_data)
    write_to_excel(writer, 'ET RAF Completions', merged_et_data)
    write_to_excel(writer, 'pivot JR rate', pivot_table_with_sums)
    write_to_excel(writer, 'Summary', summary_pivot_table)

# Print confirmation message for the output path
print(f"Output successfully written to: {output_file_path}")
print(f"Output also successfully written to: {base_output_file_path}")