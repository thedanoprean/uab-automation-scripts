import os
import pandas as pd


def extract_an_universitar(file_name):
    """Extracts the academic year from the file name."""
    try:
        an_map = {
            "19": "2019-2020",
            "20": "2020-2021",
            "21": "2021-2022",
            "22": "2022-2023",
            "23": "2023-2024",
            "24": "2024-2025"
        }

        # Extract the last two digits before the .xlsx extension
        an_code = file_name[-7:-5]  # Ensures correct identification of the academic year
        return an_map.get(an_code, "Unknown")
    except Exception as e:
        print(f"Error extracting academic year from {file_name}: {e}")
        return "Unknown"


def read_and_process_file(file_path):
    """Reads the Excel file and extracts only the required columns, adding the academic year."""
    try:
        df = pd.read_excel(file_path, sheet_name="Sheet1", header=5)
        required_columns = ["NR MATRICOL", "MEDIU PROVENIENTA", "AN DE STUDII", "FACULTATE", "SPECIALIZARE"]
        df = df[required_columns]
        df = df.dropna(subset=required_columns, how="any")

        # Add the "AN UNIVERSITAR" (academic year) column
        an_universitar = extract_an_universitar(os.path.basename(file_path))
        df["AN UNIVERSITAR"] = an_universitar

        return df
    except Exception as e:
        print(f"Error reading file {file_path}: {e}")
        return pd.DataFrame()


def process_folder(folder_path):
    """Processes all files in a folder and creates the final consolidated file."""
    files = [f for f in os.listdir(folder_path) if f.endswith('.xlsx')]
    all_data = pd.DataFrame()

    for file in files:
        file_path = os.path.join(folder_path, file)
        print(f"Processing file: {file_path}")
        df = read_and_process_file(file_path)
        if not df.empty:
            all_data = pd.concat([all_data, df], ignore_index=True)

    output_file = os.path.join(folder_path, f"{os.path.basename(folder_path)}.xlsx")
    if not all_data.empty:
        all_data.to_excel(output_file, index=False)
        print(f"Final file for {os.path.basename(folder_path)} has been saved: {output_file}")
    else:
        print(f"No valid data found for processing in folder {os.path.basename(folder_path)}.")


def main():
    main_folder = r'C:\\Users\\UAB\\Desktop\\Situatii'
    folders = [f for f in os.listdir(main_folder) if os.path.isdir(os.path.join(main_folder, f))]

    for folder in folders:
        folder_path = os.path.join(main_folder, folder)
        print(f"Processing folder: {folder_path}")
        process_folder(folder_path)


if __name__ == "__main__":
    main()
