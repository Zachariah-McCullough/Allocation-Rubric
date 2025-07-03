# Allocation Rubric Processor

## Overview  
This project automates the aggregation, scoring, and benchmarking of academic data across multiple Excel files tied to HEGIS Codes (Higher Education General Information Survey). It calculates and standardizes metrics related to Instructional Effort, Student Success, and Student Engagement — producing a clean, presentation-ready Excel report.

> ⏱️ **Eliminates 1.5 to 2 weeks of manual data processing** — transforming what once required days of copy-pasting and formatting into a seamless operation completed in minutes.

---

## How It Works

### 1. Directory Setup  
- Prompts user to enter a base folder path  
- Searches for Excel files based on specific naming patterns

### 2. Filter HEGIS Codes  
- Parses files like `INSTRUCTIONAL_FTE_*.xlsx`  
- Extracts rows from the `'Pivot Table NEW CALC FTE'` sheet  
- Filters by a predefined list of HEGIS Codes  
- Adds a **"STANDARD" row** for benchmarking  
- Saves to `FINAL_OUTPUT.xlsx` → `'Filtered HEGIS Codes'` sheet

### 3. Instructional Effort  
- Reads `INSTRUCTIONAL_EFFORT_PART_1.xlsx` and `_PART_2.xlsx`  
- Extracts and merges columns from `'Summary Table'` sheets  
- Calculates total Instructional Effort score  
- Adds a "STANDARD" benchmark row  
- Merges results and formats to 2 decimal places

### 4. Student Success  
- Processes `SUCCESS_PART_1.xlsx` and `_PART_2.xlsx`  
- Merges and calculates Success scores  
- Adds "STANDARD" row and formats data

### 5. Student Engagement  
- Reads `FS_A_updated.xlsx` and `HIP_B.xlsx`  
- Extracts from `'Flattened Data'` sheets  
- Merges and calculates Engagement scores  
- Adds benchmark row and updates report

### 6. Final Scoring  
- Calculates:
  - **Rubric Total Score**
  - **Rubric Total Score – HBG**
  - **Rubric Total Score – USMGC**

- Standardizes scores relative to "STANDARD" row  
- Outputs final metrics:
  - `Rubric Standardized Score`  
  - `Rubric Standardized Score – HBG`  
  - `Rubric Standardized Score – USMGC`

- Final results written to `FINAL_OUTPUT.xlsx`

---

## Key Concepts  

- **🔍 Filtering & Merging**: Combines only the relevant HEGIS Codes across datasets  
- **📊 Scoring & Benchmarking**: Calculates and standardizes key academic metrics  
- **⚙️ Excel Automation**: Eliminates the need for manual work across dozens of files  
- **📈 Reporting-Ready Output**: Produces clean, fully formatted Excel sheets

---

## What This Script Helps With

✅ Automates aggregation and scoring of HEGIS Code–based academic data  
✅ Combines multiple performance measures into one unified Excel output  
✅ Provides benchmark values for comparative program evaluation  
✅ Saves up to **2 weeks of manual spreadsheet labor**

---

## Output  
**🗂 FINAL_OUTPUT.xlsx**  
- `Filtered HEGIS Codes`  
- `Updated HEGIS Codes` *(if applicable)*  
- Fully calculated and formatted scoring columns

---

## Technologies Used  
- Python  
- Pandas  
- OpenPyXL  
- pathlib / os

---

*Designed to empower data-driven decision-making in academic administration through automation and accuracy.*
