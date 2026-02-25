import re
import pandas as pd
from xls_management.ate.data_de import TDSafeGuardsAttribute


class Absicherungsauftrag:
#Option Explicit
#

    def __init__(
        self,
        #testinstanz:str,
        #testumgebungstyp:str,
        #abs_status:str,
        #abs_ID:str,
        columns:pd.DataFrame,
        row:int,
    ):
#       Public testinstanz As String
#       Public Testumgebungstyp As String
#       Public abs_status As String
#       Public abs_ID As String
        #self.testinstanz = testinstanz
        #self.testumgebungstyp = testumgebungstyp
        #self.abs_status = abs_status
        #self.abs_ID = abs_ID
###### From Sub EinlesenAbsicherungsaufträge() --initialization from a data_frame row#               'Testinstanz einlesen
#               absicherungsAuftr.testinstanz = rngTDAAAttribute(4).Offset(lngZeile, 0).Value
        self.testinstanz = columns[TDSafeGuardsAttribute.TestInstance][row]
#               'Testumgebung einlesen
#               absicherungsAuftr.Testumgebungstyp = Replace(rngTDAAAttribute(5).Offset(lngZeile, 0).Value, "Testumgebungstyp: ", "")
        self.testumgebungstyp = str(columns[TDSafeGuardsAttribute.TestEnvironmentType][row]).replace('Testumgebungstyp: ', '')
#               'Status des Absicherungsauftrages einlesen
#               absicherungsAuftr.abs_status = rngTDAAAttribute(3).Offset(lngZeile, 0).Value
        self.abs_status = columns[TDSafeGuardsAttribute.Status][row]
#               'ID des Absicherungsauftrages einlesen, Entfernung der zusätzlichen Zeichen "?" und "r"
#               absicherungsAuftr.abs_ID = Replace(Replace(rngTDAAAttribute(1).Offset(lngZeile, 0).Value, "?", ""), "r", "")
        self.abs_id = re.sub(r'[\?r]', '', str(columns[TDSafeGuardsAttribute.ID][row]))
