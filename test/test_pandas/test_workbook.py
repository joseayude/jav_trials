from pathlib import Path
from test import working_path
import pandas as pd
from xls_management.workbook import Workbook


def test_open():
    file_path: Path = working_path / "test/data/example.xlsx"
    df:pd.DataFrame = pd.read_excel(file_path)
    for name in ('Age', 'City', 'Name'):
        assert name in df.keys()

def test_sheet_names():
    file_path: Path = working_path / "test/data/example.xlsx"
    excel_file =pd.ExcelFile(file_path)
    for name in ('People', 'Cars'):
        assert name in excel_file.sheet_names

def test_sheet():
    file_path: Path = working_path / "test/data/example.xlsx"
    df:pd.DataFrame =pd.read_excel(file_path, sheet_name='Cars')
    for name in ('License plate', 'Brand', 'Modell'):
        assert name in df.keys()
