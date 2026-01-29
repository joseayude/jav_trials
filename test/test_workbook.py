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

def test_workbook_sheet():
    file_path: Path = working_path / "test/data/example.xlsx"
    w: Workbook = Workbook(file_path)
    

def test_protected_workbook_sheets():
    file_path: Path = working_path / "test/data/example02.xlsx"
    w: Workbook = Workbook(file_path)
    result = w.xlsm_sheet_names()
    assert result is None

def test_all_sheets():
    expected = {
        'People':('Name', 'Age', 'City'),
        'Cars':('License plate', 'Brand', 'Modell'),
    }

    for file_name in ("example.xlsx", "example01.xlsx"):
        file_path: Path = working_path / f"test/data/{file_name}"
        workbook:Workbook = Workbook(file_path)
        for name, df in workbook.all_sheets():
            assert name in ('People', 'Cars')
            for col in df.keys():
                assert col in expected[name]
