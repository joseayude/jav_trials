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
    
#   Public Function EinlesenDatei_Projektspezifisch(ByVal strTitel As String, ByRef strAttribute() As String, ByRef rngAttribute() As Range, ByRef wbImport As Workbook, ByRef wksImport As Worksheet, ByRef strFehler As String, ByRef strDateinamen As String, _
#                                                       ByVal strProjekt As String, ByRef strAttributeProjekt() As String, ByRef rngAttributeProjekt() As Range) As Boolean
#       Dim blnAttributeImport As Boolean   'Flag für Attribute
#       Dim blnAttributeProjektImport As Boolean    'Flag für projektspezifische Attribute
#       Dim intWksImport As Integer         'Zähler für Worksheets
#       Dim strImportPfad As String         'String für Dateipfad
#       Dim strImportDatei As String        'String für Dateinamen
#       Dim strFehlerGesamt As String       'Sting für Gesamtfehler
#       Dim intAttributeZaehler As Integer  'Zähler für Attribute
#       Dim strFehlerAttribute As String    'String für fehlende Attribute
#       Dim intAttributeProjektZaehler As Integer   'Zähler für Projekt-Attribute
#       Dim strFehlerProjektAttribute As String 'String für fehlende Projekt-Attribute
#       Dim blnErsterFund As Boolean        'Flag für erstes Auffinden des vollständigen Attributesatzes
#       Dim rngAttributeErsterFund() As Range   'Range-Array für erstes Auffinden des vollständigen allgemeinen Attributesatzes
#       Dim rngAttributeProjektErsterFund() As Range    'Range-Array für erstes Auffinden des vollständigen projektspezifischen Attributesatzes
#       Dim wksImportErsterFund As Worksheet    'Worksheet für erstes Auffinden des vollständigen Attributesatzes
#       Dim lngZeilenZahlerErsterFund As Long   'Long für letzte Zeile des ersten Fundes
#       Dim lngWeitererFundZaehler As Long  'Long für Startzeile des zusätzlichen Datensatzen
#       
#       On Error Resume Next
#       
#       blnAttributeImport = False
#       blnAttributeProjektImport = False
#       strFehlerGesamt = ""
#       strDateinamen = ""
#       blnErsterFund = False
#       
#       strImportPfad = Application.GetOpenFilename(FileFilter:="Excel-Dateien (*.xls; *.xlsx; *.xlm; *.xlsm), *.xls; *.xlsx; *.xlm; *.xlsm", FilterIndex:=1, Title:=strTitel & " auswählen")
#       If Trim(strImportPfad) <> "Falsch" Then
#           'Datei öffnen
#           Workbooks.Open Filename:=strImportPfad, ReadOnly:=True
#           'Dateinamen extrahieren
#           strImportDatei = Right(strImportPfad, Len(strImportPfad) - InStrRev(strImportPfad, "\"))
#           strDateinamen = strImportDatei
#           'Workbook zuweisen
#           Set wbImport = Workbooks(strImportDatei)
#           'Worksheets nach Attributen durchsuchen
#           intWksImport = 0
#           Do
#               'Rangeobjekte zurücksetzen
#               ReDim rngAttribute(LBound(strAttribute, 1) To UBound(strAttribute, 1))
#               ReDim rngAttributeProjekt(LBound(strAttributeProjekt, 1) To UBound(strAttributeProjekt, 1))
#               
#               'Zähler für Arbeitsblatt erhöhen
#               intWksImport = intWksImport + 1
#               'Worksheet zuweisen
#               Set wksImport = wbImport.Sheets(intWksImport)
#               
#               'Allgemeine Attribute suchen
#               For intAttributeZaehler = LBound(strAttribute, 1) To UBound(strAttribute, 1)
#                   Set rngAttribute(intAttributeZaehler) = wksImport.Cells.Find(strAttribute(intAttributeZaehler), lookat:=xlWhole)
#               Next intAttributeZaehler
#               
#               'Projekt-Attribute suchen
#               For intAttributeProjektZaehler = LBound(strAttributeProjekt, 1) To UBound(strAttributeProjekt, 1)
#                   Set rngAttributeProjekt(intAttributeProjektZaehler) = wksImport.Cells.Find(strAttributeProjekt(intAttributeProjektZaehler), lookat:=xlWhole)
#               Next intAttributeProjektZaehler
#               
#               'Suche der allgemeinen Attribute auswerten
#               strFehlerAttribute = ""
#               If RangeObjekteVorhandenFehlerausgabe(rngAttribute, strAttribute, strFehlerAttribute) Then
#                   blnAttributeImport = True
#               Else
#                   If strFehlerGesamt = "" Then
#                       strFehlerGesamt = wksImport.Name & ": " & strFehlerAttribute
#                   Else
#                       strFehlerGesamt = strFehlerGesamt & "; " & wksImport.Name & ": " & strFehlerAttribute
#                   End If
#               End If
#               
#               'Suche der Projekt-Attribute auswerten
#               strFehlerProjektAttribute = ""
#               If RangeObjekteVorhandenFehlerausgabe(rngAttributeProjekt, strAttributeProjekt, strFehlerProjektAttribute) Then
#                   blnAttributeProjektImport = True
#               Else
#                   If strFehlerGesamt = "" Then
#                       strFehlerGesamt = wksImport.Name & ": " & strProjekt & " - " & strFehlerProjektAttribute
#                   Else
#                       strFehlerGesamt = strFehlerGesamt & "; " & wksImport.Name & ": " & strProjekt & " - " & strFehlerProjektAttribute
#                   End If
#               End If
#               
#               'Auswertung beider Suchen (allgemein und projektspezifisch) und Unterscheidung, ob es das erste Auffinden des vollständigen Attributsatzes ist oder weitere
#               'Beim ersten Auffinden werden die beiden Range-Arays gespeichert
#               'Ab dem zweitem Auffinden des vollständigen Attributsatzes erfolgt eine Abfrage, eine Kopiervorganges der dazugehörigen Daten zum ersten Datensatz erfolgen soll und ggf. Umsetzung
#               If blnAttributeImport = True And blnAttributeProjektImport = True Then
#                   'Zurücksetzen der Flags für den einzelnen Durchlauf
#                   blnAttributeImport = False
#                   blnAttributeProjektImport = False
#                   
#                   If blnErsterFund = False Then
#                       blnErsterFund = True
#                       'Felder dimensionieren
#                       ReDim rngAttributeErsterFund(LBound(strAttribute, 1) To UBound(strAttribute, 1))
#                       ReDim rngAttributeProjektErsterFund(LBound(strAttributeProjekt, 1) To UBound(strAttributeProjekt, 1))
#                       'Übernahme allgemeine Attribute
#                       For intAttributeZaehler = LBound(strAttribute, 1) To UBound(strAttribute, 1)
#                           Set rngAttributeErsterFund(intAttributeZaehler) = rngAttribute(intAttributeZaehler)
#                       Next intAttributeZaehler
#                       'Übernahme projektspezifische Attribute
#                       For intAttributeProjektZaehler = LBound(strAttributeProjekt, 1) To UBound(strAttributeProjekt, 1)
#                           Set rngAttributeProjektErsterFund(intAttributeProjektZaehler) = rngAttributeProjekt(intAttributeProjektZaehler)
#                       Next intAttributeProjektZaehler
#                       'Übernahme Arbeitsblatt
#                       Set wksImportErsterFund = wksImport
#                   Else
#                       If MsgBox("Weiterer Datensatz gefunden!" & vbCrLf & "Erster Fund: " & wksImportErsterFund.Name & vbCrLf & "Aktueller Fund: " & wksImport.Name & vbCrLf & vbCrLf & "Sollen die Datensätze zusammengeführt werden?", vbYesNo) = vbYes Then
#                           'Alle Zeilen mit Einträgen durchlaufen und kopieren
#                           For lngWeitererFundZaehler = 1 To wksImport.UsedRange.Rows.Count - rngAttribute(1).Row
#                               If rngAttribute(1).Offset(lngWeitererFundZaehler, 0).Value <> "" Then
#                                   'Zielzeile ermitteln
#                                   lngZeilenZahlerErsterFund = rngAttributeErsterFund(1).End(xlDown).Row + 1
#                                   
#                                   'Allgemeine Werte kopieren
#                                   For intAttributeZaehler = LBound(strAttribute, 1) To UBound(strAttribute, 1)
#                                       wksImportErsterFund.Cells(lngZeilenZahlerErsterFund, rngAttributeErsterFund(intAttributeZaehler).Column).Value = rngAttribute(intAttributeZaehler).Offset(lngWeitererFundZaehler).Value
#                                   Next intAttributeZaehler
#                                   
#                                   'Projektspezifische Werte kopieren
#                                   For intAttributeProjektZaehler = LBound(strAttributeProjekt, 1) To UBound(strAttributeProjekt, 1)
#                                       wksImportErsterFund.Cells(lngZeilenZahlerErsterFund, rngAttributeProjektErsterFund(intAttributeProjektZaehler).Column).Value = rngAttributeProjekt(intAttributeProjektZaehler).Offset(lngWeitererFundZaehler).Value
#                                   Next intAttributeProjektZaehler
#                               End If
#                           Next lngWeitererFundZaehler
#                       End If
#                   End If
#               End If
#               
#           'Loop While intWksImport < wbImport.Worksheets.Count And blnAttributeImport = False And blnAttributeProjektImport = False
#           Loop While intWksImport < wbImport.Worksheets.Count
#       End If
#       
#       
#       'Rückgabe allgemeine Attribute
#       For intAttributeZaehler = LBound(strAttribute, 1) To UBound(strAttribute, 1)
#           Set rngAttribute(intAttributeZaehler) = rngAttributeErsterFund(intAttributeZaehler)
#       Next intAttributeZaehler
#       'Rückgabe projektspezifische Attribute
#       For intAttributeProjektZaehler = LBound(strAttributeProjekt, 1) To UBound(strAttributeProjekt, 1)
#           Set rngAttributeProjekt(intAttributeProjektZaehler) = rngAttributeProjektErsterFund(intAttributeProjektZaehler)
#       Next intAttributeProjektZaehler
#       'Rückgabe Arbeitsblatt
#       Set wksImport = wksImportErsterFund
#       
#       'Rückgabewert übernehmen
#       EinlesenDatei_Projektspezifisch = blnErsterFund
#       'Fehlerwert übernehmen
#       strFehler = strFehlerGesamt
#   End Function
