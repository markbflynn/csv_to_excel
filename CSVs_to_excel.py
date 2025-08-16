import pandas as pd
import os
import glob
from pathlib import Path


def combine_csv_files_to_excel(folder_path: str, output_file_name: str = "combined_data.xlsx") -> None:
    """
    Combines all CSV files in a folder into a single Excel file.
    Each CSV becomes a separate worksheet in the Excel file.

    Args:
    folder_path (str): Path to the folder containing CSV files
    output_file_name (str): Name of the output Excel file
    """

    # Convert folder path to Path object for easier handling
    folder_path = Path(folder_path)

    # Check if folder exists
    if not folder_path.exists():
        raise FileNotFoundError(f"Folder '{folder_path}' does not exist.")

    # Find all CSV files in the folder
    csv_files = list(folder_path.glob("*.csv"))

    if not csv_files:
        raise FileNotFoundError(f"No CSV files found in '{folder_path}'.")

    # Create Excel writer object
    output_path = folder_path / output_file_name

    try:
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            for csv_file in csv_files:
                try:
                    # Read CSV file
                    dataframe = pd.read_csv(csv_file)

                    # Create worksheet name from filename (without extension)
                    worksheet_name = csv_file.stem

                    # Truncate worksheet name if it exceeds 31 characters
                    worksheet_name = worksheet_name[:31]

                    # Remove invalid characters for Excel sheet names
                    invalid_chars = ['[', ']', '*', '?', ':', '/', '\\']
                    for char in invalid_chars:
                        worksheet_name = worksheet_name.replace(char, '_')

                    # Write to Excel sheet
                    dataframe.to_excel(writer, sheet_name=worksheet_name, index=False)

                except Exception as e:
                    print(f"Error processing '{csv_file.name}': {e}")
                    continue

        print(f"Success! Combined Excel file saved as: {output_path}")

    except Exception as e:
        raise Exception(f"Error creating Excel file: {e}")



def main():
    # Prompt user for input and validate folder path
    folder_path = validate_folder_path(input("Enter the folder path containing CSV files: ").strip())

    # Prompt user for output filename, default to combined_data.xlsx
    output_filename = validate_output_filename(input("Enter the output Excel filename (default: combined_data.xlsx): ").strip())

    # Combine CSV files to Excel
    combine_csv_files_to_excel(folder_path, output_filename)


def validate_folder_path(folder_path: str) -> str:
    """
    Validates the folder path and raises an exception if it does not exist.

    Args:
    folder_path (str): The folder path to validate.

    Returns:
    str: The validated folder path.

    Raises:
    FileNotFoundError: If the folder does not exist.
    """
    folder_path = Path(folder_path)
    if not folder_path.exists():
        raise FileNotFoundError(f"Folder '{folder_path}' does not exist.")
    return str(folder_path)


def validate_output_filename(filename: str) -> str:
    """
    Validates the output filename and appends .xlsx if it does not have an extension.

    Args:
    filename (str): The output filename to validate.

    Returns:
    str: The validated output filename.
    """
    if not filename.endswith('.xlsx'):
        filename += '.xlsx'
    return filename

if __name__ == "__main__":
    main()