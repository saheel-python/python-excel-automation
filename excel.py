import os
import pandas as pd

# Folder containing your Excel files
folder_path = "/Users/saroos/Desktop/excel_files"

# Output merged file
output_file = os.path.join(folder_path, "merged.xlsx")

print("MERGE SCRIPT RUNNING")
print(f"Looking inside: {folder_path}")

# List all Excel files
excel_files = [f for f in os.listdir(folder_path) if f.endswith(".xlsx") and not f.startswith("~$")]

all_data = []

# Columns to keep (change as needed, or use None to keep all)
columns_to_keep = None  # Example: ["Name", "Email", "Age"]

for file_name in excel_files:
    file_path = os.path.join(folder_path, file_name)
    print(f"Reading: {file_name}")

    # Read Excel file, specifying engine
    try:
        df = pd.read_excel(file_path, engine='openpyxl')
    except Exception as e:
        print(f"Error reading {file_name}: {e}")
        continue

    if df.empty:
        print(f"Skipping empty file: {file_name}")
        continue

    if columns_to_keep:
        df = df[columns_to_keep]

    df['Source File'] = file_name  # Add filename column
    all_data.append(df)

# Merge all data
if all_data:
    merged_df = pd.concat(all_data, ignore_index=True)
    merged_df.to_excel(output_file, index=False, engine='openpyxl')
    print(f"All files merged successfully into {output_file}")
else:
    print("No data found to merge.")