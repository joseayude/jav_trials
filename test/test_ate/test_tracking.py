
from xls_management.shell.ate import ATEStatus

def test_ATEStatus_config():
    ate_status = ATEStatus()
    assert ate_status.config is not None
    file_path_BsM:str|None = ate_status.config.get('workbook_path_BsM')
    assert file_path_BsM is not None

#def test_ATEStatus_perform_status():
#    ate_status = ATEStatus()
#    assert ate_status.config is not None
#    ate_status.perform_status()

def test_ATEStatus_read_blacklist():
    ate_status:ATEStatus = ATEStatus()
    ate_status.read_blacklist_LAHB()
    assert ate_status.blacklist_LAHB is not None
    assert isinstance(ate_status.blacklist_LAHB, tuple) is True
