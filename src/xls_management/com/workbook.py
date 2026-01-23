import win32com.client as win32
from pathlib import Path

class Workbook:  
    def __init__(self, file_path:Path|str):
        """
        :param file_path: Path to the Excel file (.xls, .xlsx, .xlsm, etc.)
        """

        if isinstance(file_path, Path):
            self.file_path = file_path
        else:
            self.file_path = Path(file_path)

    def sheet_names(self):
        """
        Opens an RMS-protected Excel file using Excel COM automation
        and returns the sheet names.
        Requires Excel installed and user to have RMS access.
        """
        excel = win32.Dispatch("Excel.Application")
        excel.Visible = False  # Keep Excel hidden

        try:
            # Open the workbook (Excel will handle RMS authentication)
            wb = excel.Workbooks.Open(self.file_path, ReadOnly=True)

            # Get sheet names
            sheet_names = [sheet.Name for sheet in wb.Sheets]

            # Close workbook without saving
            wb.Close(SaveChanges=False)
            return sheet_names

        except Exception as e:
            raise RuntimeError(f"Failed to open RMS-protected workbook: {e}")
        finally:
            excel.Quit()
