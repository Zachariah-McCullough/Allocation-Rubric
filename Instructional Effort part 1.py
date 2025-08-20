import pandas as pd
import os
import glob
import numpy as np

# Function to load and combine ET_RAF_COURSE_SCH_* files
def load_et_raf_files(directory_path):
    file_pattern = os.path.join(directory_path, 'ET_RAF_COURSE_SCH_*.xlsx')
    excel_files = glob.glob(file_pattern)

    if not excel_files:
        print(f"No files found matching pattern 'ET_RAF_COURSE_SCH_*' in the directory: {directory_path}")
        return None

    combined_data = pd.DataFrame()
    
    for file in excel_files:
        try:
            df = pd.read_excel(file, sheet_name=0, header=1)
            required_columns = ['ID', 'SCH Load', 'WithinDisc(1)/InterDisc(1.5)', 
                                'Instr HEGIS Code', 'Instr HEGIS Descr', 
                                'Instr School', 'Instr College', 
                                'Instr HEGIS AS OF Term', 'Class HEGIS Code', 
                                'Campus', 'Class Nbr']
            df = df[required_columns]
            df.dropna(inplace=True)
            combined_data = pd.concat([combined_data, df], ignore_index=True)
            print(f"Successfully processed file: {file}")
        except Exception as e:
            print(f"Error processing file {file}: {e}")
    
    return combined_data

# Load and combine ET_RAF_COURSE_SCH_* data
directory_path = input("Enter the base directory path: ")
combined_data = load_et_raf_files(directory_path)
# C:\...\...\The University of Southern Mississippi\...\W Drive\Userfiles\AKale\Resource Allocation Rubric\AY_*_*\INSTRUCTIONAL EFFORT PART 1
# The script loads and combines data from all files matching the pattern ET_RAF_COURSE_SCH_*.xlsx in a specified directory.
# It extracts the relevant columns, drops any rows with missing data, and concatenates the data into a single DataFrame.
# It calculates a new column, SCH Score, by multiplying SCH Load by WithinDisc(1)/InterDisc(1.5).
# The columns are reordered so that SCH Score appears after WithinDisc(1)/InterDisc(1.5).
# The processed data is saved to an OUTPUT folder. If this folder doesn't exist, it's created.
# The output is saved in two locations:
    # The base directory.
    # The OUTPUT subdirectory (in the parent directory).
# A pivot table is created with Instr School, Instr HEGIS Code, and Campus as indices, and SCH Score as the values (summed up).
# Total rows are added to the pivot table, showing the sum of SCH Score per school and HEGIS code.
# The script attempts to load data from a file named INSTRUCTIONAL_FTE_4241.xlsx, specifically its Pivot Table NEW CALC FTE sheet, to obtain a grand total for comparison.
# This data is merged with the pivot table to calculate the SCH/FTE ratio (i.e., dividing SCH Score by the Grand Total), and the Score is calculated by multiplying the SCH/FTE by 0.10.
# A new column, Combined Score, is created based on whether the row is a total row or not.
# The columns SCH/FTE, Score, and Combined Score are filled with 0 where necessary, and formatted to two decimal places.
# The script writes the final pivot table (with grand totals) to the Excel file.
# It formats the header and totals in the Excel sheet.
# A summary pivot table is created to show aggregated metrics like Combined Score, SCH/FTE, and Score for each campus, and written to a separate sheet.

if combined_data is not None:
    print("ET_RAF_COURSE_SCH_* data loaded and combined.")

    # Create the SCH Score column
    combined_data['SCH Score'] = combined_data['SCH Load'] * combined_data['WithinDisc(1)/InterDisc(1.5)']

    # Reorder columns to place 'SCH Score' after 'WithinDisc(1)/InterDisc(1.5)'
    ordered_columns = [
        'ID', 'SCH Load', 'WithinDisc(1)/InterDisc(1.5)', 'SCH Score',
        'Instr HEGIS Code', 'Instr HEGIS Descr', 
        'Instr School', 'Instr College', 
        'Instr HEGIS AS OF Term', 'Class HEGIS Code', 
        'Campus', 'Class Nbr'
    ]
    combined_data = combined_data[ordered_columns]

    # Define the output file name
    output_file_name = 'INSTRUCTIONAL_EFFORT_PART_1.xlsx'
    
    # Define the path for the output folder in the parent directory
    parent_directory = os.path.dirname(directory_path)
    output_folder = os.path.join(parent_directory, 'OUTPUT')

    # Ensure the OUTPUT folder exists; if not, create it
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
        print(f"Created directory: {output_folder}")
    else:
        print(f"Directory already exists: {output_folder}")

    # Save the file to the OUTPUT folder
    output_file_path_output_folder = os.path.join(output_folder, output_file_name)
    print(f"Saving to: {output_file_path_output_folder}")

    # Save the Excel file to both locations
    output_file_path_base = os.path.join(directory_path, output_file_name)

    for output_file_path in [output_file_path_base, output_file_path_output_folder]:
        with pd.ExcelWriter(output_file_path, engine='xlsxwriter') as writer:
            # Write the combined data to a new sheet
            combined_data.to_excel(writer, sheet_name='Combined Data', index=False)

            # Create a pivot table for the combined data
            pivot_table = combined_data.pivot_table(
                index=['Instr School', 'Instr HEGIS Code', 'Campus'],
                values='SCH Score',
                aggfunc='sum'
            ).reset_index()

            # Group by 'Instr HEGIS Code' and calculate totals
            total_rows = pivot_table.groupby(['Instr School', 'Instr HEGIS Code'], as_index=False).agg({'SCH Score': 'sum'})
            total_rows['Campus'] = ''  # Leave campus blank for the total rows
            total_rows = total_rows.rename(columns={'SCH Score': 'TOTAL'})

            # Append the total rows to the pivot table
            final_table = pd.concat([pivot_table, total_rows], ignore_index=True).sort_values(by=['Instr School', 'Instr HEGIS Code', 'Campus'])

            # Attempt to load the Grand Total from INSTRUCTIONAL_FTE_4241.xlsx
            try:
                ft_file_path = os.path.join(directory_path, 'INSTRUCTIONAL_FTE_4241.xlsx')
                grand_total_data = pd.read_excel(ft_file_path, sheet_name='Pivot Table NEW CALC FTE')

                # Adjust based on actual column names
                grand_total_data = grand_total_data[['HEGIS Code', 'Grand Total']]  # Using the column from grand total

                # Merging using the correct column names and retaining the original order
                if 'Instr HEGIS Code' in final_table.columns and 'HEGIS Code' in grand_total_data.columns:
                    # Perform the merge
                    merged_table = final_table.merge(grand_total_data, left_on='Instr HEGIS Code', right_on='HEGIS Code', how='left')

                    # Define the final column order: retain existing order and add 'Grand Total' at the end
                    final_column_order = list(final_table.columns) + ['Grand Total']
                    merged_table = merged_table[final_column_order]

                    # Calculate SCH/FTE for every row, including total rows
                    merged_table['SCH/FTE'] = merged_table.apply(lambda row: row['TOTAL'] / row['Grand Total'] if row['Campus'] == '' else row['SCH Score'] / row['Grand Total'], axis=1)

                    # Combine SCH Score and TOTAL into one column
                    # Move Combined Score next to Campus
                    merged_table['Combined Score'] = merged_table.apply(
                        lambda row: row['TOTAL'] if row['Campus'] == '' else row['SCH Score'], axis=1
                    )

                    # Calculate Score as SCH/FTE * 0.10 for every row, including total rows
                    merged_table['Score'] = merged_table['SCH/FTE'] * 0.10
                else:
                    print("One of the DataFrames does not contain the required HEGIS code column.")

            except Exception as e:
                print(f"Error processing INSTRUCTIONAL_FTE_4241.xlsx for grand total: {e}")

            # Replace NaN values with 0 in 'SCH/FTE', 'Score', and 'Combined Score' columns
            merged_table['SCH/FTE'] = merged_table['SCH/FTE'].fillna(0)
            merged_table['Score'] = merged_table['Score'].fillna(0)
            merged_table['Combined Score'] = merged_table['Combined Score'].fillna(0)

            # Format the SCH/FTE column to two decimal places
            merged_table['SCH/FTE'] = merged_table['SCH/FTE'].round(2)

            # Define the Final Column order excluding the 'TOTAL' column for output
            final_column_order = [
                'Instr School', 'Instr HEGIS Code', 'Campus', 'Combined Score', 'Grand Total',
                'SCH/FTE', 'Score'
            ]
            merged_table = merged_table[final_column_order]

            # Write the final table to a new sheet
            merged_table.to_excel(writer, sheet_name='Pivot Table with Grand Total', index=False)

            # Access the workbook and the worksheet for formatting
            workbook = writer.book
            pivot_worksheet = writer.sheets['Pivot Table with Grand Total']

            # Set the header format for the pivot table
            header_format = workbook.add_format({'bold': True, 'border': 1})
            for col_num, value in enumerate(merged_table.columns.values):
                pivot_worksheet.write(0, col_num, value, header_format)

            # Format the totals in bold for the pivot table
            total_format = workbook.add_format({'bold': True})

            for row_num in range(1, len(merged_table) + 1):
                if merged_table.iloc[row_num - 1]['Campus'] == '':  # Identify total rows
                    # Also write the calculated Score for total rows
                    pivot_worksheet.write(row_num, merged_table.columns.get_loc('Score'), merged_table.iloc[row_num - 1]['Score'], total_format)

            # Set number format for SCH/FTE to two decimal places
            number_format = workbook.add_format({'num_format': '0.00'})
            pivot_worksheet.set_column('D:D', 12, number_format)  # Adjust 'D:D' based on your actual column index for SCH/FTE

            # Create a summary pivot table to hold the required metrics
            summary_pivot_table = merged_table.pivot_table(
                index='Instr HEGIS Code',
                columns='Campus',
                values=['Combined Score', 'Grand Total', 'SCH/FTE', 'Score'],  # Ensure all columns are valid
                aggfunc='sum',
                fill_value=0.00
            )

            # Reset index for better readability
            summary_pivot_table.reset_index(inplace=True)

            summary_pivot_table.columns = [
            'HEGIS Code', 
            'Total Combined Score', 
            'HBG Combined Score', 
            'USMGC Combined Score', 
            'Grand Total',
            'Grand Total 2', 
            'Grand Total 3', 
            'SCH/FTE Total', 
            'SCH/FTE HBG',
            'SCH/FTE USMGC', 
            'Total Score', 
            'HBG Score', 
            'USMGC Score'
            ]

            # Write the summary pivot table to a new sheet
            summary_pivot_table.to_excel(writer, sheet_name='Summary Table', index=False)
            
            print(f"Data successfully saved to {output_file_path}.")

print("All operations completed.")