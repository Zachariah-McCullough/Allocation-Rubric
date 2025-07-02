# Allocation-Rubric #
Overview

This script automates a complex, multi-step process involving the aggregation and analysis of Excel files related to HEGIS Codes (Higher Education General Information Survey). It filters, merges, and calculates various scores—including Instructional Effort, Student Success, and Student Engagement—for specified HEGIS Codes and outputs a clean, benchmarked Excel report.

#### This script eliminates up to two weeks of tedious manual work — transforming what used to take days of repetitive copying, filtering, and formatting across multiple files into a seamless, automated process completed in minutes. ####

How It Works
1. User Input for Directory: 

    - You’re prompted to enter the base directory where your Excel files are stored.
    - The script looks for files with specific naming patterns in that directory.

2. Filter HEGIS Codes from Instructional FTE Files: 

    - Looks for files like INSTRUCTIONAL_FTE_*.xlsx.
    - Reads the 'Pivot Table NEW CALC FTE' sheet from each.
    - Filters rows based on a predefined list of HEGIS Codes.
    - Adds a "STANDARD" row with benchmark values.
    - Combines everything and saves to FINAL_OUTPUT.xlsx in a sheet named 'Filtered HEGIS Codes'.

3. Process Instructional Effort Data (Parts 1 & 2): 

    - Reads INSTRUCTIONAL_EFFORT_PART_1.xlsx and INSTRUCTIONAL_EFFORT_PART_2.xlsx.
    - Extracts relevant columns from their 'Summary Table' sheets.
    - Merges both parts on 'HEGIS Code'.
    - Calculates total instructional effort scores.
    - Adds a "STANDARD" row (max of each column).
    - Merges into the 'Filtered HEGIS Codes' sheet.
    - Formats everything to 2 decimal places.

4. Process Success Data (Parts 1 & 2): 

    - Reads SUCCESS_PART_1.xlsx and SUCCESS_PART_2.xlsx.
    - Extracts summary data from each.
    - Merges and calculates total success scores.
    - Adds a "STANDARD" row.
    - Appends and formats results into 'Filtered HEGIS Codes'.

5. Process Engagement Data (FS_A and HIP_B): 

    - Reads FS_A_updated.xlsx and HIP_B.xlsx.
    - Extracts values from 'Flattened Data' sheets.
    - Combines and calculates engagement scores.
    - Adds a "STANDARD" row.
    - Updates and formats the data in 'Filtered HEGIS Codes'.

6. Final Calculation of Rubric Total Scores: 

    Sums the Instructional Effort, Success, and Engagement scores into:

        - Rubric Total Score
        - Rubric Total Score - HBG
        - Rubric Total Score - USMGC

    Standardizes each score by dividing by the "STANDARD" values.

    Outputs final percentages:

        - Rubric Standardized Score
        - Rubric Standardized Score - HBG
        - Rubric Standardized Score - USMGC

    Writes all the results back to the Excel file with full formatting.

Key Concepts: 

  - Filtering & Merging: Extracts only relevant HEGIS Codes and combines multiple datasets across different files.
  - Benchmarking: "STANDARD" rows serve as max-value benchmarks for comparison.
  - Scoring: Calculates both raw and standardized scores to compare programs or departments.
  - Excel Automation: Eliminates the need for manual copying, formatting, and calculation across spreadsheets.

What This Script Helps With

✅ Automates the aggregation and scoring of academic HEGIS Code data
✅ Combines multiple performance measures into one unified dataset
✅ Provides clean benchmarks for comparative evaluation
✅ Produces a fully formatted Excel report ready for presentation or review

Output:

A single, clean Excel file named FINAL_OUTPUT.xlsx

   Sheets include:

     - Filtered HEGIS Codes
     - Optionally, Updated HEGIS Codes with additional final metrics
