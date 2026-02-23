from pathlib import Path
from shutil import copy
import pandas as pd

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
    result = w.sheet_names()
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


def test_add_worksheet(tmp_path):
    """Copy example01.xlsx to a temp location, append a sheet using pandas, and verify it appears."""
    src: Path = working_path / "test/data/example01.xlsx"
    dest: Path = tmp_path / "example01.xlsx"
    copy(src, dest)

    # create a simple DataFrame and append as a new sheet
    df = pd.DataFrame({"col1": [1, 2], "col2": ["a", "b"]})
    with pd.ExcelWriter(dest, engine="openpyxl", mode="a") as writer:
        df.to_excel(writer, sheet_name="NewSheet", index=False)

    # verify via Workbook wrapper
    w: Workbook = Workbook(dest)
    assert "NewSheet" in w.sheet_names()

    # Cleanup is handled by tmp_path fixture, so no need to delete the file manually
