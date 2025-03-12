
# 📚 Academic Data Processing Repository

This repository contains several Python scripts for processing and analyzing academic data. Below are the details of each script and instructions for usage.

## ⚙️ Usage Instructions

### 1. Extract Academic Data (`extract_academic_data.py`)

This script processes all `.xlsx` files in a specified folder and consolidates them into a single Excel file.

- **To Run**: 
  - Ensure the folder path is set correctly in the script (`main_folder = r'C:\Users\UAB\Desktop\Situatii'`).
  - It will create an output file named after the folder (e.g., `Situatii.xlsx`).
  
---

### 2. Faculty Analysis Report (`faculty_analysis_report.py`)

This script reads data from multiple final files and generates statistical reports grouped by faculty, specialization, and academic year.

- **To Run**: 
  - Make sure the `main_folder` is correctly set in the script.
  - It generates statistics based on the columns: `FACULTATE`, `SPECIALIZARE`, `AN DE STUDII`, and `AN UNIVERSITAR`.
  - The final report is saved to `Statistica_Finala.xlsx`.

---

### 3. Aggregate Excel Sheets (`aggregate_excel_sheets.py`)

This script aggregates data from multiple Excel files into one, ensuring the necessary columns are available and that the data is clean.

- **To Run**: 
  - Provide a list of input files (`input_files`) and specify the output file name (`output_file`).
  - The final output will be saved as `final_result.xlsx`.

---

### 4. School Data Normalizer (`school_data_normalizer.py`)

This script normalizes and groups school data from Excel files, checking for similar schools and ensuring the proper grouping.

- **To Run**: 
  - Input the folder and filename when prompted.
  - The result will be saved as a new Excel file with grouped school data.
  - It uses fuzzy matching to group similar schools within the same county.

---

## 🛠️ Installation

1. **Clone the Repository**:
   - You can clone the repository to your local machine using Git:
   ```bash
   git clone https://github.com/yourusername/yourrepository.git
   ```

2. **Install Dependencies**:
   - Make sure you have Python 3.8+ installed.
   - Install the required dependencies using `pip`:
   ```bash
   pip install -r requirements.txt
   ```
   - Required dependencies include:
     - `pandas`
     - `openpyxl`
     - `fuzzywuzzy`
     - `python-Levenshtein`

---

## 📋 Usage Instructions

1. **Run the Scripts**:
   - Navigate to the script folder on your local machine and execute the desired script. For example:
   ```bash
   python extract_academic_data.py
   ```

2. **Customize**:
   - Update the folder paths, filenames, or any other variables in the scripts to match your use case.
   - You can modify the list of input files or customize the output filenames as needed.

---

## 📂 Folder Structure

- `extract_academic_data.py`: Script to process academic data and consolidate it into a single Excel file.
- `faculty_analysis_report.py`: Script to generate faculty analysis reports based on academic data.
- `aggregate_excel_sheets.py`: Script to aggregate data from multiple Excel files.
- `school_data_normalizer.py`: Script to normalize and group school data using fuzzy matching.

