# 🎓 School Data Processing Repository

## 📜 Description

This repository contains a set of Python scripts designed for processing, analyzing, and consolidating school data from Excel files (.xlsx). The primary goal of these scripts is to help efficiently extract and analyze information from multiple files, ensuring that the data is cleaned and properly structured for further analysis. These scripts are used to process student, faculty, and specialization information from various universities, consolidating the data into a single file and generating relevant statistics for educational decisions.

## 🗂️ Folder Structure

The repository contains two main folders:

1. **academic_year_analysis**  
   - **extract_academic_data.py**: This script processes Excel files in a folder, extracts necessary data (STUDENT ID, MEDIUM OF ORIGIN, YEAR OF STUDIES, FACULTY, SPECIALIZATION), adds the academic year based on the file name, and generates a consolidated output file.
   - **faculty_analysis_report.py**: This script reads the files generated by the previous script, extracts statistics for each faculty, specialization, and academic year, and generates a final statistical report.

2. **school_data_processing**  
   - **aggregate_excel_sheets.py**: This script merges multiple Excel files, extracts and cleans the data (for students, faculties, counties, and countries), and saves it into a consolidated file.
   - **school_data_normalizer.py**: This script normalizes and groups similar schools by counties, removing duplicates and aggregating relevant data for each educational unit.

## ⚙️ Detailed Functionality

### 1. **Extract and Process Academic Data**  
   - **extract_academic_data.py**:
     - Extracts the academic year from the file name.
     - Selects only the relevant columns from the Excel files.
     - Cleans the data by removing incomplete rows.
     - Saves the processed data into a consolidated file per folder.
   
   - **faculty_analysis_report.py**:
     - Reads the files generated by the first script and extracts key statistics based on faculty, specialization, and academic year.
     - Groups the data and generates a final report with the statistics.

### 2. **Merge and Normalize School Data**  
   - **aggregate_excel_sheets.py**:
     - Merges multiple Excel files, extracts the necessary columns (NAME, SURNAME, FACULTY, COUNTY, COUNTRY), and cleans the data.
     - If the COUNTY is missing and COUNTRY is not "Romania", the value from COUNTRY is retained.
     - Saves the merged data into a single output file.
   
   - **school_data_normalizer.py**:
     - Normalizes and compares school names, grouping similar schools by county.
     - Aggregates data based on school similarity and calculates total students per county.
     - Saves the grouped data into an Excel file with counties merged in a single cell.

## 🕒 Time Saved

Given the large dataset of 30 Excel files, each containing 2000 entries and 130 columns, running the scripts normally takes **2-3 minutes**. Here’s an estimated **time-saving** analysis:

- **Data Processing Without Scripts**:  
  If we were to manually process and clean each file, ensuring that the required data is extracted, cleaned, and consolidated, the time required would be much higher. Assuming it takes **2 minutes** to process each file manually:
  - 30 files x 2 minutes = **60 minutes** (1 hour).

- **Time with Scripts**:  
  The scripts run in approximately **2-3 minutes** for the entire 30 files. This results in a time-saving of around:
  - **Approximately 58-59 minutes** for all files.

- **Time Saved (Linear)**:
  - **98% Time Saved** from manual processing.

By automating this task, we save **over 98%** of the time, allowing the team to focus on other more critical activities!

## 📂 How to Use

1. Clone or download this repository to your local machine.
2. Install the necessary dependencies (e.g., `pandas`, `openpyxl`, `fuzzywuzzy`).
3. Ensure your Excel files are correctly placed in the specified folders.
4. Run the respective Python scripts based on your needs:
   - For academic year extraction and faculty analysis: Run `extract_academic_data.py` and `faculty_analysis_report.py`.
   - For school data merging and normalization: Run `aggregate_excel_sheets.py` and `school_data_normalizer.py`.

## 🧑‍💻 Dependencies

Make sure you have the following Python libraries installed:

- `pandas`
- `openpyxl`
- `fuzzywuzzy`
- `unicodedata` (for text normalization)

To install the dependencies, you can use the following:

```bash
pip install pandas openpyxl fuzzywuzzy
