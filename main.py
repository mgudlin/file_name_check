import os
import pandas as pd

def check_files_exist(excel_file_path, folder_path):
    # Read the Excel table into a pandas DataFrame
    df = pd.read_excel(excel_file_path)

    # Extract the file names from the DataFrame
    file_names = df["File Name"].tolist()

    # Check if each file exists in the folder
    missing_files = []
    for file_name in file_names:
        file_path = os.path.join(folder_path, file_name)
        if not os.path.exists(file_path):
            missing_files.append(file_name)

    # Add a new column "Exists" with value "N" for missing files
    exists_column = ["Y" if file_name not in missing_files else "N" for file_name in file_names]
    df["Exists"] = exists_column

    return df

if __name__ == "__main__":
    excel_file_path = "check_table-repaired.xlsx"
    folder_path = "fotografije_sve"

    updated_df = check_files_exist(excel_file_path, folder_path)

    # Save the updated DataFrame back to the Excel file
    updated_df.to_excel("updated_excel_file.xlsx", index=False)