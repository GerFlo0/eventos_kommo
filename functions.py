import pandas as pd
import json


def import_xlsx(file_path: str, sheet: int = 0) -> pd.DataFrame:
    """
    Imports data from an Excel file and returns a pandas DataFrame.

    Parameters:
    file_path (str): The path to the Excel file.
    sheet (int): The index of the sheet to import.
    if the sheet index is not provided, the first sheet will be imported.

    Returns:
    pd.DataFrame: A DataFrame containing the imported data.
    """
    try:
        df = pd.read_excel(file_path, sheet_name=sheet)
        return df
    except Exception as e:
        print(f"Error importing Excel file: {e}")
        return None

def import_json(file_path: str) -> dict:
    """
    Imports data from a JSON file and returns a dictionary.

    Parameters:
    file_path (str): The path to the JSON file.

    Returns:
    dict: A dictionary containing the imported data.
    """
    try:
        with open(file_path) as f:
            data = json.load(f)
        return data
    except Exception as e:
        print(f"Error importing JSON file: {e}")
        return None