import pandas as pd
import os
import glob

# Function to load and combine ET_RAF_COURSE_SCH_* files
def load_et_raf_files(directory_path):
    file_pattern = os.path.join(directory_path, 'ET_RAF_ENROLLMENT_*.xlsx')
    excel_files = glob.glob(file_pattern)

    if not excel_files:
        print(f"No files found matching pattern 'ET_RAF_ENROLLMENT_*' in the directory: {directory_path}")
        return None

    combined_data = pd.DataFrame()
    
    for file in excel_files:
        try:
            df = pd.read_excel(file, sheet_name=0, header=1)
            required_columns = ['ID', 'HEGIS Code', 'Term', 'Acad Org', 'Org Descr', 'Acad Group', 'Pri Prog Camp']
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
# C:\...\...\The University of Southern Mississippi\...\W Drive\Userfiles\AKale\Resource Allocation Rubric\AY_*_*\INSTRUCTIONAL EFFORT PART 2
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

if combined_data is not None:
    print("ET_RAF_ENROLLMENT_* data loaded and combined.")

    # Combine 'HBG' and 'ONLNE' into 'HBG'
    combined_data['Pri Prog Camp'] = combined_data['Pri Prog Camp'].replace({'ONLNE': 'HBG'})

    # Define the output file name
    output_file_name = 'INSTRUCTIONAL_EFFORT_PART_2.xlsx'
    
    # Define the path for the output folder in the parent directory
    parent_directory = os.path.dirname(directory_path)
    output_folder = os.path.join(parent_directory, 'OUTPUT')

    # Ensure the OUTPUT folder exists; if not, create it
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # Save the Excel file to both locations
    for output_file_path in [os.path.join(directory_path, output_file_name), os.path.join(output_folder, output_file_name)]:
        with pd.ExcelWriter(output_file_path, engine='xlsxwriter') as writer:
            # Write the combined data to a new sheet
            combined_data.to_excel(writer, sheet_name='Combined Data', index=False)

            # Create a pivot table for the combined data
            pivot_table = combined_data.pivot_table(
                index=['Org Descr', 'HEGIS Code', 'Pri Prog Camp'],
                values='ID',
                aggfunc='count'  # Count to get the number of IDs
            ).reset_index()

            # Attempt to load the Grand Total from INSTRUCTIONAL_FTE_*.xlsx
            try:
                ft_file_path = glob.glob(os.path.join(directory_path, 'INSTRUCTIONAL_FTE_*.xlsx'))  # Use glob to find the file
                if ft_file_path:
                    grand_total_data = pd.read_excel(ft_file_path[0], sheet_name='Pivot Table NEW CALC FTE')
                    grand_total_data = grand_total_data[['HEGIS Code', 'Grand Total']]  # Using the column from grand total

                    # Merging using the correct column names and retaining the original order
                    if 'HEGIS Code' in pivot_table.columns and 'HEGIS Code' in grand_total_data.columns:
                        # Perform the merge
                        merged_table = pivot_table.merge(grand_total_data, on='HEGIS Code', how='left')

                        # Calculate SCH/FTE for every row
                        merged_table['SCH/FTE'] = merged_table.apply(lambda row: row['ID'] / row['Grand Total'] if row['Grand Total'] > 0 else 0, axis=1)

                        # Calculate SCORE as SCH/FTE * 0.20
                        merged_table['SCORE'] = merged_table['SCH/FTE'] * 0.20

                        # Rename columns for clarity
                        merged_table.rename(columns={'ID': 'Total ID'}, inplace=True)

                        # Create a new DataFrame for the flattened totals
                        total_rows = []
                        for hegis_code in merged_table['HEGIS Code'].unique():
                            # Filter rows for the current HEGIS Code
                            hegis_data = merged_table[merged_table['HEGIS Code'] == hegis_code]

                            # Prepare a new row with summed values
                            total_row = {
                                'Org Descr': '',  # or None
                                'HEGIS Code': hegis_code,
                                'Pri Prog Camp': '',  # or None
                                'Total ID': hegis_data['Total ID'].sum(),
                                'Grand Total': hegis_data['Grand Total'].iloc[0],  # Use the first row for Grand Total
                                'SCH/FTE': hegis_data['SCH/FTE'].mean() if not hegis_data['SCH/FTE'].isnull().all() else 0,  # Average or set to 0
                                'SCORE': hegis_data['SCORE'].mean() if not hegis_data['SCORE'].isnull().all() else 0  # Average or set to 0
                            }
                            total_rows.append(total_row)

                        # Convert the list of total rows into a DataFrame
                        total_rows_df = pd.DataFrame(total_rows)

                        # Concatenate the original merged_table with the total_rows_df
                        final_table = pd.concat([merged_table, total_rows_df], ignore_index=True)

                        # Optionally, sort by HEGIS Code to keep it organized
                        final_table.sort_values(by=['HEGIS Code', 'Pri Prog Camp'], inplace=True)

                    else:
                        print("One of the DataFrames does not contain the required HEGIS code column.")
                else:
                    print("No INSTRUCTIONAL_FTE_*.xlsx files found.")

            except Exception as e:
                print(f"Error processing INSTRUCTIONAL_FTE_*.xlsx for grand total: {e}")

            # Write the final pivot table to a new sheet
            final_table.to_excel(writer, sheet_name='Pivot Table with Grand Total', index=False)

            # Access the workbook and the worksheet for formatting
            workbook = writer.book
            pivot_worksheet = writer.sheets['Pivot Table with Grand Total']

            # Set the header format for the pivot table
            header_format = workbook.add_format({'bold': True, 'border': 1})
            for col_num, value in enumerate(final_table.columns.values):
                pivot_worksheet.write(0, col_num, value, header_format)

            # Set number format for Total ID, SCH/FTE, and SCORE to two decimal places
            number_format = workbook.add_format({'num_format': '0.00'})
            pivot_worksheet.set_column('D:D', 12, number_format)  # Adjust based on your actual column index for Total ID
            pivot_worksheet.set_column('E:E', 12, number_format)  # Adjust based on your actual column index for SCH/FTE
            pivot_worksheet.set_column('F:F', 12, number_format)  # Adjust based on your actual column index for SCORE
            pivot_worksheet.set_column('G:G', 12, number_format)  # Adjust based on your actual column index for Grand Total

            # Create a summary pivot table to hold the required metrics
            summary_pivot_table = final_table.pivot_table(
                index='HEGIS Code',
                columns='Pri Prog Camp',
                values=['Total ID', 'Grand Total', 'SCH/FTE', 'SCORE'],  # Ensure all columns are valid
                aggfunc='sum',
                fill_value=0.00
            )

            # Reset index for better readability
            summary_pivot_table.reset_index(inplace=True)

            # Flatten the MultiIndex columns
            summary_pivot_table.columns = [' '.join(col).strip() if isinstance(col, tuple) else col for col in summary_pivot_table.columns.values]

            # Write the summary pivot table to a new sheet
            summary_pivot_table.to_excel(writer, sheet_name='Summary Table', index=False)

            print(f"File saved to: {output_file_path}")

else:
    print("Merged table not created due to previous errors.")