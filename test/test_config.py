import os
import test
from test import working_path
from xls_management import HOMEPATH
from pathlib import Path

def test_config_create(monkeypatch):
    clean("test/data/config.yml")
    monkeypatch.setattr("xls_management.ROOTPATH",working_path / "test/data")
    from xls_management.config import ATEConfig
    ate = ATEConfig()
    assert 'BsM_workbook' in ate.config.keys()
    assert ate.config['BsM_workbook'] == (
        HOMEPATH /
        'vw/data/ATE-Status_Berichtsversion.xlsm'
    ).as_uri()

def test_config_open(monkeypatch):
    clean("test/data/config.yml")
    assert (working_path / "test/data/config.yml").exists() is False
    monkeypatch.setattr("xls_management.ROOTPATH",working_path / "test/data")
    from xls_management.config import ATEConfig
    ate = ATEConfig()
    assert 'BsM_workbook' in ate.config.keys()
    assert ate.config['BsM_workbook'] == (
        HOMEPATH /
        'vw/data/ATE-Status_Berichtsversion.xlsm'
    ).as_uri()
    del(ate)
    assert (working_path / "test/data/config.yml").exists() is True
    ate = ATEConfig()
    assert 'BsM_workbook' in ate.config.keys()
    assert ate.config['BsM_workbook'] == (
        HOMEPATH /
        'vw/data/ATE-Status_Berichtsversion.xlsm'
    ).as_uri()


def clean(working_file):
    file_path = working_path / working_file
    if file_path.exists():
        os.remove(file_path)