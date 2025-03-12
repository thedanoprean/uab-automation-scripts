import os
import pandas as pd


def read_and_process_file(file_path):
    """Reads the Excel file and extracts only the required columns."""
    try:
        df = pd.read_excel(file_path)

        # Required columns for analysis
        required_columns = ["NR MATRICOL", "MEDIU PROVENIENTA", "AN DE STUDII", "FACULTATE", "SPECIALIZARE",
                            "AN UNIVERSITAR"]

        if not all(col in df.columns for col in required_columns):
            print(f"File {file_path} does not contain all required columns.")
            return pd.DataFrame()

        # Extract only the necessary columns
        df = df[required_columns]

        # Remove rows with missing values
        df = df.dropna(subset=required_columns, how="any")

        # Ignore students without a valid 'MEDIU PROVENIENTA'
        df = df[df['MEDIU PROVENIENTA'].notna() & (df['MEDIU PROVENIENTA'] != '-')]

        return df
    except Exception as e:
        print(f"Error reading file {file_path}: {e}")
        return pd.DataFrame()


def read_final_files():
    """Loads the final files from the corresponding folders."""
    main_folder = r'C:\\Users\\UAB\\Desktop\\Situatii'
    final_files_folders = ['FII', 'FILSE', 'FTO', 'FSE', 'FDSS']
    all_data = pd.DataFrame()

    for folder in final_files_folders:
        file_path = os.path.join(main_folder, folder, f"{folder}.xlsx")
        print(f"Checking file {file_path}")

        if os.path.exists(file_path):
            df = read_and_process_file(file_path)
            all_data = pd.concat([all_data, df], ignore_index=True)
        else:
            print(f"File {file_path} does not exist in folder {folder}.")

    return all_data


def generate_statistics(df):
    """Generates statistics grouped by faculty, specialization, study year, and academic year."""
    statistics = df.pivot_table(
        index=['FACULTATE', 'SPECIALIZARE', 'AN DE STUDII', 'AN UNIVERSITAR'],
        columns='MEDIU PROVENIENTA',
        aggfunc='size',
        fill_value=0
    )

    return statistics


def save_statistics(statistics):
    """Saves the statistics to an Excel file."""
    output_file = r'C:\\Users\\UAB\\Desktop\\Situatii\\Statistica_Finala.xlsx'
    statistics.to_excel(output_file)
    print(f"Final statistics have been saved to {output_file}")


def main():
    all_data = read_final_files()

    if not all_data.empty:
        statistics = generate_statistics(all_data)
        save_statistics(statistics)
    else:
        print("No valid data found for processing.")


if __name__ == "__main__":
    main()
