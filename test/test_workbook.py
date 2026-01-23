from pathlib import Path
from xls_management.workbook import Workbook
from test import working_path

def test_workbook_sheets():
    file_path: Path = working_path / "test/data/example.xlsx"
    w: Workbook = Workbook(file_path)
    result = w.sheet_names()
    assert 'People' in result
    assert 'Cars' in result
    assert len(result)

def test_protected_workbook_sheets():
    file_path: Path = working_path / "test/data/example02.xlsx"
    w: Workbook = Workbook(file_path)
    result = w.xlsm_sheet_names()
    assert result is None
