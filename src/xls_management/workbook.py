import pandas as pd
import os
from pathlib import Path
from typing import Generator

class Workbook:
    def __init__(self, file_path:Path|str):
        """
        :param file_path: Path to the Excel file (.xls, .xlsx, .xlsm, etc.)
        """

        if isinstance(file_path, Path):
            self.file_path = file_path
        else:
            self.file_path = Path(file_path)

    def sheet_names(self) -> list[int|str]:
        """
        Returns a list of sheet names from an Excel file.
        :return: List of sheet names or None if error occurs
        """
        try:
            # Validate file existence
            if not os.path.isfile(self.file_path):
                raise FileNotFoundError(f"File not found: {self.file_path}")
            sheet_names: list[int|str]
            # Load Excel file
            with pd.ExcelFile(self.file_path, 'openpyxl') as excel_file:
                sheet_names =  excel_file.sheet_names
            # Return list of sheet names
            return sheet_names

        except FileNotFoundError as fnf_err:
            print(f"Error: {fnf_err}")
        except ValueError as val_err:
            print(f"Invalid file format: {val_err}")
        except Exception as e:
            print(f"Unexpected error: {e}")

        return None

    def sheets(self, index:int|str) -> pd.DataFrame:
        try:
            # Load the workbook (read-only mode for efficiency)
            names = self.sheet_names()
            df:pd.DataFrame|None = None
            name:str = ""
            if isinstance(index,str):
                name = index
            elif index < len(names):
                name = names[index]
            if name in names:
                df = pd.read_excel(
                    self.file_path,
                    sheet_name=name,
                    dtype=str,
                )
            else:
                raise IndexError(f"{index} is wrong value for {', '.join(names)}")
        except FileNotFoundError:
            print(f"Error: File '{self.file_path}' not found.")
        except Exception as e:
            print(f"Error reading file: {e}")
        return df
    
    def all_sheets(self) -> Generator[tuple[str,pd.DataFrame],None,None]:
        """iterator """
        for name in self.sheet_names():
            df:pd.DataFrame = pd.read_excel(self.file_path, sheet_name=name, dtype=str)
            yield name, df
    
    def append_worksheet(self, data_frame:pd.DataFrame, name:str):
        try:
            if os.path.exists(self.file_path):
                kvargs = {
                    'mode':'a',
                    'engine':'openpyxl',
                    'if_sheet_exists':'replace',
                }
            else:
                kvargs = {
                    'mode':'w',
                    'engine':'xlsxwriter',
                }
            # Append to existing workbook
            with pd.ExcelWriter(
                self.file_path,
                **kvargs,
            ) as writer:
                df = data_frame.replace('nan','')
                df.to_excel(
                    writer,
                    sheet_name=name,
                    index=False,
                freeze_panes=(1,3),
                autofilter=True,
                )
                writer.sheets[name].autofit()
            #else:
            #    if not os.path.exists(self.file_path.parent):
            #        os.makedirs(self.file_path.parent, exist_ok=True)
            #    # Create new workbook
            #    data_frame.to_excel(
            #        self.file_path,
            #        sheet_name=name,
            #        index=False,
            #        freeze_panes=(1,3),
            #        autofilter=True,
            #    )

            print(f"DataFrame saved to '{name}' in {self.file_path}")

        except PermissionError:
            print("Error: The file is open in another program. Please close it and try again.")
        except Exception as e:
            print(f"An error occurred: {e}")  