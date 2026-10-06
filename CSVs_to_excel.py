import pandas as pd
import os
import glob
from pathlib import Path


EXCEL_SHEET_NAME_LIMIT = 31
INVALID_SHEET_NAME_CHARACTERS = ['[', ']', '*', '?', ':', '/', '\\']


def _get_worksheet_name(csv_file: Path, used_names: set[str]) -> str:
    """Return a valid, unique Excel worksheet name for ``csv_file``."""
    worksheet_name = csv_file.stem[:EXCEL_SHEET_NAME_LIMIT]

    for char in INVALID_SHEET_NAME_CHARACTERS:
        worksheet_name = worksheet_name.replace(char, '_')

    candidate = worksheet_name
    suffix_number = 2
    while candidate.casefold() in used_names:
        suffix = f"_{suffix_number}"
        candidate = (
            f"{worksheet_name[:EXCEL_SHEET_NAME_LIMIT - len(suffix)]}{suffix}"
        )
        suffix_number += 1

    used_names.add(candidate.casefold())
    return candidate


def combine_csv_files_to_excel(folder_path: str, output_file_name: str = "combined_data.xlsx") -> None:
    """
    Combines all CSV files in a folder into a single Excel file.
    Each CSV becomes a separate worksheet in the Excel file.

    Args:
    folder_path (str): Path to the folder containing CSV files
    output_file_name (str): Name of the output Excel file
    """

    folder_path = Path(folder_path)

    if not folder_path.exists():
        raise FileNotFoundError(f"Folder '{folder_path}' does not exist.")

    # Sort so that collision suffixes are assigned deterministically.
    csv_files = sorted(folder_path.glob("*.csv"), key=lambda path: (path.name.casefold(), path.name))

    if not csv_files:
        raise FileNotFoundError(f"No CSV files found in '{folder_path}'.")

    output_path = folder_path / output_file_name

    try:
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            used_worksheet_names = set()
            for csv_file in csv_files:
                try:
                    dataframe = pd.read_csv(csv_file)
                    worksheet_name = _get_worksheet_name(
                        csv_file, used_worksheet_names
                    )
                    dataframe.to_excel(writer, sheet_name=worksheet_name, index=False)

                except Exception as e:
                    print(f"Error processing '{csv_file.name}': {e}")
                    continue

        print(f"Success! Combined Excel file saved as: {output_path}")

    except Exception as e:
        raise Exception(f"Error creating Excel file: {e}")


def main():
    folder_path = validate_folder_path(input("Enter the folder path containing CSV files: ").strip())
    output_filename = validate_output_filename(input("Enter the output Excel filename (default: combined_data.xlsx): ").strip())
    combine_csv_files_to_excel(folder_path, output_filename)


def validate_folder_path(folder_path: str) -> str:
    """Validate the folder path and return it as a string."""
    folder_path = Path(folder_path)
    if not folder_path.exists():
        raise FileNotFoundError(f"Folder '{folder_path}' does not exist.")
    return str(folder_path)


def validate_output_filename(filename: str) -> str:
    """Append .xlsx when the output filename has no extension."""
    if not filename.endswith('.xlsx'):
        filename += '.xlsx'
    return filename


if __name__ == "__main__":
    main()
