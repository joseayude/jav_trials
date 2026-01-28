class DBInfo:
    def __init__(
            self,
            workbook:str,
            worksheet:str,
            attributes: tuple[str]= (),
            ranges:tuple[str]=(),
        ):
        self.workbook = workbook
        self.worksheet = worksheet
        self.attributes = attributes
        self.ranges = ranges

    def str_attributes(self, separator:str=", "):
        return separator.join(self.attributes)

#Public Function EinlesenDatei(ByVal strTitel As String, ByRef strAttribute() As String, ByRef rngAttribute() As Range, ByRef wbImport As Workbook, ByRef wksImport As Worksheet, ByRef strFehler As String, ByRef strDateinamen As String) As Boolean
#Dim blnAttributeImport As Boolean   'Flag für Attribute
#Dim intWksImport As Integer         'Zähler für Worksheets
#Dim strImportPfad As String         'String für Dateipfad
#Dim strImportDatei As String        'String für Dateinamen
#Dim strFehlerGesamt As String       'Sting für Gesamtfehler
#Dim intAttributeZaehler As Integer  'Zähler für Attribute
#Dim strFehlerAttribute As String    'String für fehlende Attribute
#
#On Error Resume Next
#
#blnAttributeImport = False
#strFehlerGesamt = ""
#strDateinamen = ""
#
#strImportPfad = Application.GetOpenFilename(FileFilter:="Excel-Dateien (*.xls; *.xlsx; *.xlm; *.xlsm), *.xls; *.xlsx; *.xlm; *.xlsm", FilterIndex:=1, Title:=strTitel & " auswählen")
#If Trim(strImportPfad) <> "Falsch" Then
#    'Datei öffnen
#    Workbooks.Open strImportPfad
#    'Dateinamen extrahieren
#    strImportDatei = Right(strImportPfad, Len(strImportPfad) - InStrRev(strImportPfad, "\"))
#    strDateinamen = strImportDatei
#    'Workbook zuweisen
#    Set wbImport = Workbooks(strImportDatei)
#    'Worksheets nach Attributen durchsuchen
#    intWksImport = 0
#    Do
#        'Rangeobjekte zurücksetzen
#        ReDim rngAttribute(LBound(strAttribute, 1) To UBound(strAttribute, 1))
#        'Zähler für Arbeitsblatt erhöhen
#        intWksImport = intWksImport + 1
#        'Worksheet zuweisen
#        Set wksImport = wbImport.Sheets(intWksImport)
#        'Attribute suchen
#        For intAttributeZaehler = LBound(strAttribute, 1) To UBound(strAttribute, 1)
#            Set rngAttribute(intAttributeZaehler) = wksImport.Cells.Find(strAttribute(intAttributeZaehler), lookat:=xlWhole)
#        Next intAttributeZaehler
#        'Suche auswerten
#        strFehlerAttribute = ""
#        If RangeObjekteVorhandenFehlerausgabe(rngAttribute, strAttribute, strFehlerAttribute) Then
#            blnAttributeImport = True
#        Else
#            If strFehlerGesamt = "" Then
#                strFehlerGesamt = wksImport.Name & ": " & strFehlerAttribute
#            Else
#                strFehlerGesamt = strFehlerGesamt & "; " & wksImport.Name & ": " & strFehlerAttribute
#            End If
#        End If
#    Loop While intWksImport < wbImport.Worksheets.Count And blnAttributeImport = False
#End If
#
#'Rückgabewert übernehmen
#EinlesenDatei = blnAttributeImport
#'Fehlerwert übernehmen
#strFehler = strFehlerGesamt
#End Function