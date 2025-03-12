
# 📚 UAB Automation Scripts 🚀

## ✨ Description

This repository contains a series of Python scripts designed to automate various administrative and data processing tasks related to academic data at **University 1 Decembrie 1918** in Alba Iulia. These scripts save time and effort by automating processes that would otherwise need to be done manually.

## 🛠️ Included Scripts

1. **`extract_an_universitar.py`**  
   Extracts the academic year from Excel file names and adds this information to a consolidated output file.

2. **`process_data.py`**  
   Processes Excel files, extracts relevant data (e.g., study year, faculty, specialization), and generates a final file consolidating information from multiple sources.

3. **`generate_statistics.py`**  
   Generates statistics related to students' provenance environment, grouped by faculties, specializations, study years, and academic years. Saves the statistics to an Excel file.

4. **`merge_excels.py`**  
   Merges multiple Excel files containing student and school data, adds a "No." column, and cleans the data to make it more uniform.

5. **`group_similar_schools.py`**  
   Groups similar schools based on fuzzy name matching and saves the results in an Excel file, including total counts per county.

---

## 🚀 Installation

To run the scripts above, you need to have the following Python libraries installed:

```bash
pip install pandas openpyxl fuzzywuzzy
```

---

## 📝 Usage Instructions

### 1. 📂 Add Input Files  
Ensure your Excel files are placed in the appropriate folders under the `Situatii` directory. Each script has specific instructions for reading input files from a particular folder.

### 2. ▶️ Running a Script  
After adding the input files, you can run the desired script. For example, to process the files and generate a final file, run:

```bash
python process_data.py
```

### 3. 📊 Generating Statistics  
Once the files are processed, you can generate statistics using:

```bash
python generate_statistics.py
```

The statistics will be saved in the `Situatii` directory.

---

## ⏱️ Time Saved

### 💡 Manual Processing Time Estimate

Given that there are 30 Excel files, each with 2000 entries and 130 columns of data, manual processing of these files would be extremely time-consuming. Each file would require at least 30 minutes to extract relevant data, analyze it, and save the final file.

- **Estimated manual time**: 30 files x 30 minutes = 900 minutes (~15 hours)
- **Time saved with automation**: The automation process can be done much faster (approximately 1-2 minutes per file, depending on system performance).

By automating the process, you can save **90%** of the time that would have been spent on manual processing.

### ⏳ Time Saved Example

- **Manual Time**: 900 minutes (15 hours)
- **Automated Time**: 3 minutes
- **Time Saved**: ~99.67% 🙌

---

## 🤝 Contributing

If you'd like to contribute to this project, feel free to create a pull request. Any improvements or bug fixes are welcome!

