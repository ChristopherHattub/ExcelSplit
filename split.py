import pandas as pd
import os
import math
import glob

def split_excel_file(file_path, number_of_splits):
    try:
        # Read the Excel file
        df = pd.read_excel(file_path, engine='openpyxl')
        print(f"File read successfully. Total rows: {len(df)}")

        # Ensure the number of splits is not greater than the number of rows
        if number_of_splits > len(df):
            print("The number of splits is greater than the number of rows in the file.")
            return

        # Calculate the number of rows per split file
        rows_per_file = math.ceil(len(df) / number_of_splits)
        print(f"Rows per split file: {rows_per_file}")

        # Splitting the file
        for i in range(number_of_splits):
            start_row = i * rows_per_file
            end_row = min(start_row + rows_per_file, len(df))
            split_df = df.iloc[start_row:end_row]

            # Construct new file name
            base_name = os.path.splitext(file_path)[0]
            new_file_name = f"{base_name}SplitPart{i + 1}.xlsx"

            # Save the new file
            split_df.to_excel(new_file_name, index=False, engine='openpyxl')
            print(f"Created {new_file_name} ({i + 1}/{number_of_splits})")

        print("File splitting completed.")
    except Exception as e:
        print(f"An error occurred: {e}")

def list_excel_files():
    # Set the directory to the script's directory
    os.chdir(os.path.dirname(os.path.abspath(__file__)))

    # List all .xlsx files in the current directory
    excel_files = glob.glob('*.xlsx')
    if excel_files:
        print("Available Excel files:")
        for file in excel_files:
            print(file)
    else:
        print("No available Excel files found in the directory.")
def main():
    # Display available Excel files
    list_excel_files()

    while True:
        # User input for file name
        file_name = input("Enter the name of the Excel file (including .xlsx): ")

        # Check if file exists in the current directory
        if os.path.exists(file_name):
            break
        else:
            print(f"File {file_name} not found. Please try again.")

    # User input for number of splits
    number_of_splits = int(input("Enter the number of files to split into: "))

    # Process the file splitting
    split_excel_file(file_name, number_of_splits)

if __name__ == "__main__":
    main()

