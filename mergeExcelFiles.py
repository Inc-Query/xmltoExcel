import pandas as pd
import os
import sys

# Check if the input directory is provided as a command-line argument
if len(sys.argv) != 2:
    print("Usage: python3 merge_excel.py <input_directory>")
    sys.exit(1)

# Get the input directory from the command-line argument
input_dir = sys.argv[1]

# Validate the input directory
if not os.path.isdir(input_dir):
    print(f"Error: The directory '{input_dir}' does not exist.")
    sys.exit(1)

# Output file will be saved in the same directory
output_file = os.path.join(input_dir, "merged_output.xlsx")

# Get all Excel files in the directory
excel_files = [f for f in os.listdir(input_dir) if f.endswith('.xlsx') or f.endswith('.xls')]

# Check if there are any Excel files in the directory
if not excel_files:
    print(f"No Excel files found in '{input_dir}'.")
    sys.exit(1)

# Create a Pandas ExcelWriter object
with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    for file in excel_files:
        # Read the first sheet of the Excel file
        df = pd.read_excel(os.path.join(input_dir, file), sheet_name=0)
        # Write the sheet to the merged file, named after the original file
        sheet_name = os.path.splitext(file)[0]  # Remove file extension
        df.to_excel(writer, sheet_name=sheet_name, index=False)

print(f"Merged {len(excel_files)} files into {output_file}")