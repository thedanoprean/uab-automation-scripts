import pandas as pd


def merge_excels(input_files, output_file):
    dataframes = []

    for file in input_files:
        try:
            # Read data from Sheet1
            df = pd.read_excel(file, sheet_name="Sheet1", dtype=str)

            # Check if the file contains data
            if df.empty:
                print(f"Warning: The file {file} is empty and will be ignored.")
                continue

            # Check if the file contains valid columns
            if df.columns.empty:
                print(f"Warning: The file {file} does not have valid columns and will be ignored.")
                continue

            # Display columns available for debugging
            print(f"File: {file} -> Available columns: {df.columns.tolist()}")

            # Clean column names (remove unnecessary spaces)
            df.columns = df.columns.astype(str).str.strip()

            # Check if all required columns exist
            selected_columns = ["NAME", "SURNAME", "FACULTY", "COUNTY", "COUNTRY"]
            missing_columns = [col for col in selected_columns if col not in df.columns]

            if missing_columns:
                print(f"Warning: The file {file} is missing the columns: {missing_columns} and will be ignored.")
                continue  # Skip files that are incomplete

            # Select only the necessary columns
            df = df[selected_columns]

            # If COUNTY is empty and COUNTRY is not "Romania", retain the value from COUNTRY
            df.loc[df["COUNTY"].isna() | (df["COUNTY"].str.strip() == ""), "COUNTRY"] = df["COUNTRY"]

            dataframes.append(df)

        except Exception as e:
            print(f"Error processing file {file}: {e}")
            continue  # Avoid stopping the program in case of an error with one file

    if dataframes:
        final_df = pd.concat(dataframes, ignore_index=True)

        # Add the "No." column as the first column
        final_df.insert(0, "No.", range(1, len(final_df) + 1))

        final_df.to_excel(output_file, index=False, engine='openpyxl')
        print(f"The final file has been saved as {output_file}")
    else:
        print("Error: No valid files were found.")


# List of input files
input_files = [
    "export_data_students1740574949840.xlsx",
    "export_data_students1740575392150.xlsx",
    "export_data_students1740575440654.xlsx",
    "export_data_students1740575484669.xlsx",
    "export_data_students1740575526963.xlsx"
]

# Output file name
output_file = "final_result.xlsx"

# Call the function
merge_excels(input_files, output_file)
