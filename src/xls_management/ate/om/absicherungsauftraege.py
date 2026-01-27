class Absicherungsauftraege:
#Option Explicit
#

    def __init__(
        self,
        testinstanz:str,
        testumgebungstyp:str,
        abs_status:str,
        abs_ID:str,
    ):
#       Public testinstanz As String
#       Public Testumgebungstyp As String
#       Public abs_status As String
#       Public abs_ID As String
        self.testinstanz = testinstanz
        self.testumgebungstyp = testumgebungstyp
        self.abs_status = abs_status
        self.abs_ID = abs_ID
