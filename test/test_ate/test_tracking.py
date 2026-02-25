
from contextlib import redirect_stdout
from unittest.mock import patch
from pathlib import Path
from test import working_path

import sys

def fake_msgbox_no(msg:str)->bool:
    print(f'{msg}\nNo')
    sys.stdout.flush()
    return False

def fake_print(*vargs,**kvargs):
    print(*vargs,**kvargs)
    sys.stdout.flush()

def test_ATEStatus_perform_status():
    # ensure fresh imports so patched functions are picked up by modules
    old_stdout = sys.stdout
    to_remove = [name for name in sys.modules if name.startswith("xls_management")]
    for name in to_remove:
        del sys.modules[name]

    # prepare a list of file paths to be returned by the file picker
    file_path = working_path / "../ATEStatus_perfom_status.txt"
    # side_effect list long enough for repeated calls
    file_list = [
        working_path / '../in/MEB21_Statistik_Testing.xlsx',
        working_path / '../in/Alle Verifikationskriterien.xlsx',
        working_path / '../in/Alle Absicherungsaufträge.xlsx',
        working_path / '../in/Alle Testfälle.xlsx',
        working_path / '../in/MasterFeatureplan.xlsx',
        working_path / '../in/trial_Master.xlsx',
    ]
    with file_path.open("w") as f:
        sys.std_out = f
        with(
            patch('xls_management.tui.file_picker.path_from_file_picker', side_effect=file_list),
            patch('xls_management.ate.project.project_combo_box', return_value=('MEB21', False)),
            patch('xls_management.tui.msgbox.msgbox', new=fake_print),
            patch('xls_management.tui.yes_no_form.yes_no_msgbox', new=fake_msgbox_no),
        ):
            fake_print(f'....{__name__}')
            # import after patches so module-level imports pick up the patched functions
            from xls_management.shell.ate import ATEStatus

            ate_status = ATEStatus()
            #sheets = ate_status.output_workbook.sheet_names()
            #fake_print(', '.join(sheets))
            fake_print('..starting perform_status')
            ate_status.perform_status()
            fake_print('..perform_status ended')

            # project and flag were set by the mocked combo box
            assert ate_status.project == 'MEB21'
            assert ate_status.use_predecessor_ids is False
            
            #sheets_now = ate_status.output_workbook.sheet_names()
            #fake_print(','.join(sheets_now))
            #assert len(sheets_now) == len(sheets) + 1
        sys.stdout = old_stdout

def test_ATEStatus_config():
    from xls_management.shell.ate import ATEStatus

    ate_status = ATEStatus()
    assert ate_status.config is not None
    file_path_BsM:str|None = ate_status.config.get('workbook_path_BsM')
    assert file_path_BsM is not None

#def test_ATEStatus_perform_status():
#    ate_status = ATEStatus()
#    assert ate_status.config is not None
#    ate_status.perform_status()

def test_ATEStatus_read_blacklist():
    from xls_management.shell.ate import ATEStatus

    ate_status:ATEStatus = ATEStatus()
    ate_status.read_blacklist_LAHB()
    assert ate_status.blacklist_LAHB is not None
    assert isinstance(ate_status.blacklist_LAHB, tuple) is True
