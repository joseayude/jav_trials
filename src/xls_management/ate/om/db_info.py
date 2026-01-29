from pathlib import Path
from xls_management.tui.file_picker import path_from_file_picker
from xls_management.workbook import Workbook
import pandas as pd


class DBInfo:
    def __init__(
        self,                        #
        #workbook:Workbook,           # ByRef wbImport As Workbook
        #sheet_name:str,              # ByRef wksImport As Worksheet
           # sheet_name and workbook is enougth to be able to load the data
        #columns:pd.DataFrame,        # ByRef rngAttribute() As Range
        attributes: tuple[str]= (),  # ByRef strAttribute() As String
    ):
        self.workbook:Workbook|None = None
        self.sheet_name:str = ""
        self.attributes = attributes
        self.columns:pd.DataFrame|None = None

    def str_attributes(self, separator:str=", "):
        return separator.join(self.attributes)

#   Public Function EinlesenDatei(ByVal strTitel As String, ByRef strAttribute() As String, ByRef rngAttribute() As Range, ByRef wbImport As Workbook, ByRef wksImport As Worksheet, ByRef strFehler As String, ByRef strDateinamen As String) As Boolean
    def einlesen_datei(self, titel:str) -> tuple[bool,str]:
        """
        user chooses a workbook using a file picker widget
        True,"" is returned if each expected attribute is in one of the workbook sheets;
                self.sheet_name is set with the name of the sheet containg those attributes
        False, error_trace is returned elsewhere; 
                being error trace a trace of missing attributes
           
        """ 
        error_trace = ""
        import_file_path = path_from_file_picker(location=".", title= f"{titel} auswählen")
        if import_file_path is not None:
            workbook: Workbook = Workbook(import_file_path)
            import_file_name = Path(workbook.file_path).name
            for self.sheet_name, self.columns in workbook.all_sheets():
                missing_attributes = [attribute for attribute in self.attributes if attribute not in self.columns]
                if len(missing_attributes) == 0:
                    # all expected attributes have been found; self.sheet_name and self.columns are set
                    return True, ""
                else:
                    error_trace += (
                        f"The following attributes are missing from {self.sheet_name} in {import_file_name}: "
                        f"{', '.join(missing_attributes)}\n"
                    )
        return False, error_trace
