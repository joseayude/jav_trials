import csv

from openpyxl.styles import Alignment
import pandas as pd
import os
from pathlib import Path
from typing import Generator

from xls_management.utils.tools import col_name_from, get_slices

#CODEC = 'cp1252'
CODEC = 'iso-8859-1'

class Workbook:
    def __init__(self, file_path:Path|str, engine:str='openpyxl'):
        """
        :param file_path: Path to the Excel file (.xls, .xlsx, .xlsm, etc.)
        """

        if isinstance(file_path, Path):
            self.file_path = file_path
        else:
            self.file_path = Path(file_path)
        self.engine = engine

    def writer(self):
        return pd.ExcelWriter(self.file_path, engine=self.engine)
    
    def reader(self):
        return pd.ExcelFile(self.file_path, engine=self.engine)
    
    def append_worksheet(self, writer, data_frame:pd.DataFrame, name:str):
        try:
            df = data_frame.replace('nan','')
            df.to_excel(
                writer,
                sheet_name=name,
                index=False,
                freeze_panes=(1,3),
                engine=self.engine,
                autofilter=True,
                na_rep='',
            )
            worksheet = writer.sheets[name]
            for index in range(len(data_frame.columns)):
                col_name =col_name_from(index)
                columns = worksheet.column_dimensions[col_name]
                columns.alignment = Alignment(wrap_text=True)
                columns.width =70
            print(f"DataFrame saved to '{name}' in {self.file_path}")
        except Exception as e:
            print(f"An error occurred: {e}")

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
    
    def sheet(self, index:int|str) -> pd.DataFrame:
        try:
            # Load the workbook (read-only mode for efficiency)
            names = self.sheet_names()
            df:pd.DataFrame|None = None
            name:str = ""
            if isinstance(index,str):
                name = index
            elif index < len(names):
                name = names[index]
            with self.reader() as xls:
                df = pd.read_excel(xls, sheet_name=name,dtype=str, engine=self.engine)
                #fix: each value _x000D_ is replaced by \r
                df = df.replace(to_replace='_x000D_', value='\r', regex=True)
                df.fillna(value="",inplace=True)
        except FileNotFoundError:
            print(f"Error: File '{self.file_path}' not found.")
        except Exception as e:
            print(f"Error reading file: {e}")
        return df
    
    def all_sheets(self) -> Generator[tuple[str,pd.DataFrame],None,None]:
        """iterator """
        for name in self.sheet_names():
            df:pd.DataFrame = pd.read_excel(self.file_path, sheet_name=name, dtype=str, engine=self.engine)
            #fix: each value _x000D_ is replaced by \r
            df = df.replace(to_replace='_x000D_', value='\r', regex=True)
            df.fillna(value="",inplace=True)
            yield name, df

    def to_csv(
            self,
            skiprows:int=0,
            csv_path:str='{workdir}/{preffix}_{sheet_name}.csv',
            sheet_name:str|int=0,
            slice_size:int=0,
        ):
        """
        Convert an Excel worksheet to CSV.

        :param csv_path: Path to save the output CSV file
        :param sheet_name: Sheet name or index (default=0 for first sheet)
        """
        try:
            
            # Validate file existence
            if not self.file_path.is_file():
                raise FileNotFoundError(f"Excel file not found: {self.file_path}")
            # Read the Excel file
            kvargs={'sheet_name':sheet_name, 'dtype':str, 'engine':self.engine}
            if skiprows > 0:
                kvargs['skiprows'] =skiprows
            with self.reader() as reader:
                df = pd.read_excel(reader, **kvargs)
                #fix: each value _x000D_ is replaced by \r
                df = df.replace(to_replace='_x000D_', value='\r', regex=True)
            preffix = self.file_path.name.replace(self.file_path.suffix,'')
            csv_path = Path(
                csv_path.format(workdir=self.file_path.parent, preffix=preffix, sheet_name=sheet_name)
            )
            # Ensure output directory exists
            os.makedirs(csv_path.parent, exist_ok=True)

            # Save as CSV without index
            if slice_size == 0:
                df.to_csv(
                    csv_path,
                    index=False,
                    encoding=CODEC,
                    sep=';',
                    quoting=csv.QUOTE_ALL,
                )
            else:
                preffix = csv_path.name.replace(csv_path.suffix,'')
                assert slice_size > 0
                n:int = len(df)
                top = n - n % slice_size - slice_size
                for i, start, top in get_slices(0,len(df),slice_size):
                    slice_name = f'{preffix}_{i:03}.csv'
                    df_i = df.iloc[start:top]
                    df_i.to_csv(
                        csv_path.with_name(slice_name),
                        index=False,
                        encoding=CODEC,
                        sep=';',
                        quoting=csv.QUOTE_ALL,
                    )


            print(f"✅ Successfully saved '{sheet_name}' from '{self.file_path}' to '{csv_path}'")

        except FileNotFoundError as fnf_err:
            print(f"❌ Error: {fnf_err}")
        except ValueError as val_err:
            print(f"❌ Sheet error: {val_err}")
        except Exception as e:
            print(f"❌ Unexpected error: {e}")
    