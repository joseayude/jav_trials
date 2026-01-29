from xls_management.tui.file_picker import path_from_file_picker

#Option Explicit
#
#'Dokumentenabfrage einzeln \\vw.vwg\vwdfs\K-E\EF\1508\Groups\EFBS2_Konsulter\Testmanagement_EFDB\Projekt MQB48W\Testdesign\Statistik_TD_TS\
#
#'Klasse Verifikationskriterien mit Absicherungsaufträgen
#Public verifikationKritList As New Collection
#'Klasse AVW-Rohdaten
#Public BsMDatenList As New Collection
#'Klasse Testfälle
#Public testfallList As New Collection
#'Klasse FRU_Timing
#Public FRUTimingList As New Collection
#'Klasse AVWVorgaenger
#Public AVWVorgaengerList As New Collection
#'Flag für die Berücksichtigung von Vorgänger-IDs bei den AVW-Rohdaten
#Public blnAVWVorgaengerIDsVerwenden As Boolean
#
#Public Sub ATE_Status()
#'BsM_Status
#Dim wbBsM As Workbook                       'Workbook für BsM_Status
#Dim wksBsM As Worksheet                     'Worksheet für BsM_Status
#Dim strBsMAttribute() As String             'String-Array mit Attributen des Arbeitsblatts BsM_Status
#Dim rngBsMAttribute() As Range              'Range-Array mit Attributen des Arbeitsblatts BsM_Status
#'TD_Status
#Dim wbTD As Workbook                        'Workbook für TD_Status
#Dim wksTD As Worksheet                      'Worksheet für TD_Status
#Dim strTDAttribute() As String              'String-Array mit Attributen des Arbeitsblatts TD_Status
#Dim rngTDAttribute() As Range               'Range-Array mit Attributen des Arbeitsblatts TD_Status
#'AVW_Rohdaten - Projekt
#Dim wbAVW As Workbook                       'Workbook für AVW_Rohdaten
#Dim wksAVW As Worksheet                     'Worksheet für AVW_Rohdaten
#Dim strAVWAttribute() As String             'String-Array mit Attributen des Arbeitsblatts AVW_Rohdaten
#Dim rngAVWAttribute() As Range              'Range-Array mit Attributen des Arbeitsblatts AVW_Rohdaten
#Dim strAVWAttributeMEB21() As String        'String-Array mit Attributen des Arbeitsblatts AVW_Rohdaten für MEB21
#Dim rngAVWAttributeMEB21() As Range         'Range-Array mit Attributen des Arbeitsblatts AVW_Rohdaten für MEB21
#'AVW_Rohdaten - Master
#Dim wbAVWMaster As Workbook                 'Workbook für AVWMaster_Rohdaten
#Dim wksAVWMaster As Worksheet               'Worksheet für AVWMaster_Rohdaten
#Dim strAVWMasterAttribute() As String       'String-Array mit Attributen des Arbeitsblatts AVWMaster_Rohdaten
#Dim rngAVWMasterAttribute() As Range        'Range-Array mit Attributen des Arbeitsblatts AVWMaster_Rohdaten
#'TDVKs (Tesdesigns - Verifikationskriterium)
#Dim wbTDVK As Workbook                      'Workbook für TDs - Verifikationskriterien
#Dim wksTDVK As Worksheet                    'Worksheet für TDs - Verifikationskriterium
#Dim strTDVKAttribute() As String            'String-Array mit Attributen des Arbeitsblatts TDs - Verifikationskriterium
#Dim rngTDVKAttribute() As Range             'Range-Array mit Attributen des Arbeitsblatts TDs - Verifikationskriterium
#'TDAAs (Tesdesigns - Absicherungsaufträge)
#Dim wbTDAA As Workbook                      'Workbook für TDs - Absicherungsaufträge
#Dim wksTDAA As Worksheet                    'Worksheet für TDs - Absicherungsaufträge
#Dim strTDAAAttribute() As String            'String-Array mit Attributen des Arbeitsblatts TDs - Absicherungsaufträge
#Dim rngTDAAAttribute() As Range             'Range-Array mit Attributen des Arbeitsblatts TDs - Absicherungsaufträge
#'TF (Testfälle)
#Dim wbTF As Workbook                        'Workbook für Testfälle
#Dim wksTF As Worksheet                      'Worksheet für Testfälle
#Dim strTFAttribute() As String              'String-Array mit Attributen des Arbeitsblatts Testfälle
#Dim rngTFAttribute() As Range               'Range-Array mit Attributen des Arbeitsblatts Testfälle
#'FRUTiming
#Dim wbFRUTiming As Workbook                 'Workbook für FRU-Timing
#Dim wksFRUTiming As Worksheet               'Worksheet für FRU-Timing
#Dim strFRUTimingAttribute() As String       'String-Array mit Attributen des Arbeitsblattes für FRU-Timing
#Dim rngFRUTimingAttribute() As Range        'Range-Array mit Attributen des Arbeitsblattes für FRU-Timing
#'Allgemein
#Dim strFehlerGesamt As String               'String für Gesamtfehlerausgabe
#Dim strWeitereTUsAusgabe As String          'String für Ausgabe weiterer Testumgebungstypen
#Dim strLAHBlacklist() As String             'String-Array für einzulesende LAH-Blacklist
#Dim strVersionMakro As String               'String-Array für Makro-Version
#Dim strDateinamen(1 To 6) As String         'String-Array für die Namen der eingelesenen Dateien
#Dim strProjekte(0 To 5) As String           'String-Array für die auswählbaren Fahrzeugprojekte
#Dim strProjekt As String                    'String des ausgewählten Fahrzeugprojekts
#'Verlauf
#Dim strFehlerATEVerlauf As String           'String für Rückgabewert der Befüllung von ATE_Status_Verlauf
#Dim strFehlerTDVerlauf As String            'String für Rückgabewert der Befüllung von TD_Status_Verlauf
#Dim strFehlerVerlauf As String              'String für gemeinsamen Rückgabewert der Befüllung von ATE/TD_Status_Verlauf
#
#'Makro-Version
#strVersionMakro = "ATE-Status V015F6" & vbCrLf & "Programmiert von Alexander Kuhlicke, Tagueri AG 2024"
#
#'BsM-Status wird separat im Workbook des Makros erzeugt
#Set wbBsM = ThisWorkbook
#
#'Befüllung der Projektliste
#strProjekte(0) = "leer"
#strProjekte(1) = "MQB48W"
#strProjekte(2) = "MQB37W PA"
#strProjekte(3) = "MEB UNECE"
#strProjekte(4) = "MEB21"
#strProjekte(5) = "Andere"
#BoxAuswahlProjekt.ComboBox1.list() = strProjekte
#BoxAuswahlProjekt.ComboBox1.ListIndex = 0
#
#'Abfrage Projekt und Nutzung Master-Bereich
#BoxAuswahlProjekt.Caption = "ATE-Status " & strVersionMakro
#BoxAuswahlProjekt.Show
#
#If boolAuswahlGetroffen Then
#    'Eingabe Projekt
#    strProjekt = BoxAuswahlProjekt.ComboBox1.list(BoxAuswahlProjekt.ComboBox1.ListIndex)
#    'Eingabe Verwendung Master-IDs
#    If BoxAuswahlProjekt.OptionButton1.Value = True Then
#        blnAVWVorgaengerIDsVerwenden = True
#    Else
#        blnAVWVorgaengerIDsVerwenden = False
#    End If
#    
#    'Einlesen der Attribute der Rohdaten
#    If ATE_Status_Initializer(wbAVW, wbAVWMaster, wbTDVK, wbTDAA, wbTF, wbFRUTiming, _
#                              wksAVW, wksAVWMaster, wksTDVK, wksTDAA, wksTF, wksFRUTiming, _
#                              strAVWAttribute, strAVWMasterAttribute, strTDVKAttribute, strTDAAAttribute, strTFAttribute, strFRUTimingAttribute, _
#                              rngAVWAttribute, rngAVWMasterAttribute, rngTDVKAttribute, rngTDAAAttribute, rngTFAttribute, rngFRUTimingAttribute, _
#                              strFehlerGesamt, strDateinamen, blnAVWVorgaengerIDsVerwenden, _
#                              strProjekt, strAVWAttributeMEB21, rngAVWAttributeMEB21) Then
#        'LAH-Blacklist einlesen
#        Call EinlesenLAHBlacklist(wbBsM, strLAHBlacklist)
#        'Testdesigns - Verifikationskriterien einlesen
#        Call EinlesenTDVKs(wksTDVK, strTDVKAttribute, rngTDVKAttribute)
#        'Testdesigns - Absicherungsaufträge einlesen
#        Call EinlesenTDAAs(wksTDAA, strTDAAAttribute, rngTDAAAttribute)
#        'Testfälle
#        Call EinlesenTFs(wksTF, strTFAttribute, rngTFAttribute)
#        'FRU-Timing
#        Call EinlesenFRUTiming(wksFRUTiming, strFRUTimingAttribute, rngFRUTimingAttribute)
#        'Anforderungen
#        If blnAVWVorgaengerIDsVerwenden = False Then
#            'Anforderungsstatistik Projekt
#            Call EinlesenAVWRohdaten(wksAVW, strAVWAttribute, rngAVWAttribute, strLAHBlacklist, strProjekt, strAVWAttributeMEB21, rngAVWAttributeMEB21)
#        Else
#            'Anforderungsstatistik Masterbereich
#            Call EinlesenAVWVorgaengerRohdaten(wksAVWMaster, strAVWMasterAttribute, rngAVWMasterAttribute)
#            'Anforderungsstatistik Projekt
#            Call EinlesenAVWNachfolgerRohdaten(wksAVW, strAVWAttribute, rngAVWAttribute, strLAHBlacklist, strProjekt, strAVWAttributeMEB21, rngAVWAttributeMEB21)
#        End If
#        'Ausgabe ATE-Status
#        Call AusgabeATEStatus(wbBsM, wksBsM, strBsMAttribute, rngBsMAttribute, strWeitereTUsAusgabe, strDateinamen, strProjekt)
#        'Ausgabe TD-Status
#        Call AusgabeTDStatus(wbBsM, wksTD, strTDAttribute, rngTDAttribute, strDateinamen, strProjekt)
#        'Geöffnete Dateien schliessen
#        Call SchliessenWb(wbBsM, wbAVW, wbAVWMaster, wbTDVK, wbTDAA, wbTF, wbFRUTiming)
#        
#        'Verläufe ATE_Status_Verlauf und TD_Status_Verlauf befüllen und Rückgabewerte zusammenführen
#        Call AusgabeVerlauf(wksBsM, strFehlerATEVerlauf, 1)
#        Call AusgabeVerlauf(wksTD, strFehlerTDVerlauf, 2)
#        If strFehlerATEVerlauf <> "" And strFehlerTDVerlauf <> "" Then
#            strFehlerVerlauf = strFehlerATEVerlauf & vbCrLf & strFehlerTDVerlauf
#        ElseIf strFehlerATEVerlauf <> "" Then
#            strFehlerVerlauf = strFehlerATEVerlauf
#        ElseIf strFehlerTDVerlauf <> "" Then
#            strFehlerVerlauf = strFehlerTDVerlauf
#        Else
#            strFehlerVerlauf = ""
#        End If
#        
#        'Abschlussmeldung
#        If strWeitereTUsAusgabe = "" Then
#            If strFehlerVerlauf = "" Then
#                MsgBox "ATE-Status erstellt!" & vbCrLf & vbCrLf & "-----" & vbCrLf & strVersionMakro
#            Else
#                MsgBox "ATE-Status erstellt!" & vbCrLf & vbCrLf & strFehlerVerlauf & vbCrLf & vbCrLf & "-----" & vbCrLf & strVersionMakro
#            End If
#        Else
#            If strFehlerVerlauf = "" Then
#                MsgBox "ATE-Status erstellt!" & vbCrLf & vbCrLf & "Folgende weitere Testumgebungstypen wurden erkannt, aber nicht für den Vergleich berücksichtigt:" & vbCrLf & vbCrLf & strWeitereTUsAusgabe & vbCrLf & vbCrLf & "-----" & vbCrLf & strVersionMakro
#            Else
#                MsgBox "ATE-Status erstellt!" & vbCrLf & vbCrLf & "Folgende weitere Testumgebungstypen wurden erkannt, aber nicht für den Vergleich berücksichtigt:" & vbCrLf & vbCrLf & strWeitereTUsAusgabe & vbCrLf & vbCrLf & strFehlerVerlauf & vbCrLf & vbCrLf & "-----" & vbCrLf & strVersionMakro
#            End If
#        End If
#        
#    Else
#        MsgBox strFehlerGesamt, Buttons:=vbExclamation, Title:="Fehler beim Import für ATE-Tracking"
#    End If
#    
#    'Worksheets und Klassenmodule zurücksetzen
#    Call ATE_Status_Deinitializer(wksBsM, wksAVW, wksTDVK, wksTDAA, wksTF, wksFRUTiming)
#End If
#
#Unload BoxAuswahlProjekt
#End Sub
#
from xls_management.ate.data import AVW_ATTRIBUTE_DE


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

class ATEStatus():
    def __init__(
            self,
            predecessorIdUsed:bool,
            project:str,
        ):
        self.predecessorIdUsed = predecessorIdUsed
        self.project = project
#   Private Function ATE_Status_Initializer(ByRef wbAVW As Workbook, ByRef wbAVWMaster As Workbook, ByRef wbTDVK As Workbook, ByRef wbTDAA As Workbook, ByRef wbTF As Workbook, ByRef wbFRUTiming As Workbook, _
#           ByRef wksAVW As Worksheet, ByRef wksAVWMaster As Worksheet, ByRef wksTDVK As Worksheet, ByRef wksTDAA As Worksheet, ByRef wksTF As Worksheet, ByRef wksFRUTiming As Worksheet, _
#           ByRef strAVWAttribute() As String, ByRef strAVWMasterAttribute() As String, ByRef strTDVKAttribute() As String, ByRef strTDAAAttribute() As String, ByRef strTFAttribute() As String, ByRef strFRUTimingAttribute() As String, _
#           ByRef rngAVWAttribute() As Range, ByRef rngAVWMasterAttribute() As Range, ByRef rngTDVKAttribute() As Range, ByRef rngTDAAAttribute() As Range, ByRef rngTFAttribute() As Range, ByRef rngFRUTimingAttribute() As Range, _
#           ByRef strFehlerGesamt As String, ByRef strDateinamen() As String, ByVal blnAVWVorgaengerIDsVerwenden As Boolean, _
#           ByVal strProjekt As String, ByRef strAVWAttributeMEB21() As String, ByRef rngAVWAttributeMEB21() As Range) As Boolean
#           
#       Dim strFehlerBsM As String                  'String für Fehlerausgabe bei nicht vorhandenen Rangeobjekten
#       Dim strFehlerAVW As String                  'String für Fehlerausgabe bei nicht vorhandenen Rangeobjekten
#       Dim strFehlerAVWMaster As String            'String für Fehlerausgabe bei nicht vorhandenen Rangeobjekten
#       Dim strFehlerTDAA As String                 'String für Fehlerausgabe bei nicht vorhandenen Rangeobjekten
#       Dim strFehlerTDVK As String                 'String für Fehlerausgabe bei nicht vorhandenen Rangeobjekten
#       Dim strFehlerTF As String                   'String für Fehlerausgabe bei nicht vorhandenen Rangeobjekten
#       Dim strFehlerFRUTiming As String            'String für Fehlerausgabe bei nicht vorhandenen Rangeobjekten
#       Dim strAttributeAVW As String               'String für alle Attribute
#       Dim strAttributeAVWMaster As String         'String für alle Attribute
#       Dim strAttributeTDAA As String              'String für alle Attribute
#       Dim strAttributeTDVK As String              'String für alle Attribute
#       Dim strAttributeTF As String                'String für alle Attribute
#       Dim strAttributeFRUTiming As String         'String für alle Attribute
#       Dim blnImportAttribute(1 To 6) As Boolean   'Flag-Array für korrektes Einlesen der Arbeitsblätter #1: AVW-Rohdaten, #2: TDVKs, #3: TDAAs, #4: Testfälle, #5: FRU-Timing, #6: AVWMaster-Rohdaten
#       Dim i As Integer                            'Laufvariable
#       
#       strFehlerGesamt = ""
        self.strFehlerGesamt = ""
#       
#       'Arbeitsblatt AVW_Rohdaten
#       'AVW: #1: ID, #2: Dokument-ID, #3: Basis für Testdesign, #4: Typ, #5: Kategorie, #6: Status, #7: Feature, #8: Reifegrad, #9: Umsetzer, #10: ASIL
#       '#11: BSM-SaFuSi Bewertung, #12: BSM-ZZ Bewertung, #13: BSM-ED Bewertung, #14: BSM-FFF Bewertung, #15: BSM-O Bewertung, #16: BSM-Se Bewertung, #17: MV AVWMV
#       '#18: Cluster Testing, #19: Dokument, #20: Kommentar Redaktionskreis, #21: temp1_Text, #22: Abgezweigt aus (vorher: Vorgänger ID, nur bei Projekt+Master)
#       'Weiche für Erfassung der Nachfolger-IDs
#       If blnAVWVorgaengerIDsVerwenden = False Then
#           ReDim strAVWAttribute(1 To 22)
#       Else
#           ReDim strAVWAttribute(1 To 23)
#       End If
#       ReDim rngAVWAttribute(LBound(strAVWAttribute, 1) To UBound(strAVWAttribute, 1))
#       strAVWAttribute(1) = "ID"
#       strAVWAttribute(2) = "Dokument-ID"
#       strAVWAttribute(3) = "Basis für Testdesign"
#       strAVWAttribute(4) = "Typ"
#       strAVWAttribute(5) = "Kategorie"
#       strAVWAttribute(6) = "Status"
#       strAVWAttribute(7) = "Feature"
#       strAVWAttribute(8) = "Reifegrad"
#       strAVWAttribute(9) = "Umsetzer"
#       strAVWAttribute(10) = "ASIL"
#       strAVWAttribute(11) = "BSM-SaFuSi Bewertung"
#       strAVWAttribute(12) = "BSM-ZZ Bewertung"
#       strAVWAttribute(13) = "BSM-ED Bewertung"
#       strAVWAttribute(14) = "BSM-FFF Bewertung"
#       strAVWAttribute(15) = "BSM-O Bewertung"
#       strAVWAttribute(16) = "BSM-Se Bewertung"
#       strAVWAttribute(17) = "MV"
#       strAVWAttribute(18) = "Cluster Testing"
#       strAVWAttribute(19) = "Dokument"
#       strAVWAttribute(20) = "Kommentar Redaktionskreis"
#       strAVWAttribute(21) = "temp1_Text"
#       strAVWAttribute(22) = "Anforderungsverantwortliche"
#       If blnAVWVorgaengerIDsVerwenden Then
#           strAVWAttribute(23) = "Abgezweigt aus"  'strAVWAttribute(22) = "Abgezweigt aus"
#       End If
        if self.predecessorIdUsed:
            self.info_AVW = DBInfo(attributes=AVW_ATTRIBUTE_DE)
        else:
            self.info_AVW = DBInfo(attributes=AVW_ATTRIBUTE_DE[:-1])
#
#       'Dateiauswahl und Zuordnung
#       'Projektspezifisch (MEB21 oder MQB48W) oder allgemein
#       If strProjekt = "MEB21" Or strProjekt = "MQB48W" Then
        if self.project in ("MEB21", "MQB48W"):
#           ReDim strAVWAttributeMEB21(1 To 1)
#           ReDim rngAVWAttributeMEB21(1 To 1)
#           strAVWAttributeMEB21(1) = "Temp11_Auswahlfeld"
            self.str_AVW_attributeBEB21 = "Temp11_Auswahlfeld"
#           If EinlesenDatei_Projektspezifisch("Anforderungen Projekt " & strProjekt, strAVWAttribute, rngAVWAttribute, wbAVW, wksAVW, strFehlerAVW, strDateinamen(1), strProjekt, strAVWAttributeMEB21, rngAVWAttributeMEB21) Then
#               blnImportAttribute(1) = True
            self.import_attribute[0] = self.AVW.EinlesenDatei_Projektspezifisch()
#           Else
#               'Sammlung aller gesuchten allgemeinen Attribute erzeugen
#               For i = LBound(strAVWAttribute, 1) To UBound(strAVWAttribute, 1)
#                   If strAttributeAVW = "" Then
#                       strAttributeAVW = strAVWAttribute(i)
#                   Else
#                       strAttributeAVW = strAttributeAVW & ", " & strAVWAttribute(i)
#                   End If
#               Next
            if not self.import_attribute[0]:
                #strAttributeAVW = self.info_AVW.str_attributes()
#               'Sammlung aller gesuchten projektspezifischen Attribute erzeugen
#               For i = LBound(strAVWAttributeMEB21, 1) To UBound(strAVWAttributeMEB21, 1)
#                   If strAttributeAVW = "" Then
#                       strAttributeAVW = strAVWAttributeMEB21(i)
#                   Else
#                       strAttributeAVW = strAttributeAVW & ", " & strAVWAttributeMEB21(i)
#                   End If
#               Next i
                #strAttributeAVN += f", {self.str_AVW_attributeBEB21}" 
#               'Zusammenführung der gesuchten Attribute
#               If strFehlerGesamt = "" Then
#                   strFehlerGesamt = "Anforderungen können nicht eingelesen werden!" & vbCrLf & "(Benötigt: " & strProjekt & " - " & strAttributeAVW & ")"
#               Else
#                   strFehlerGesamt = strFehlerGesamt & vbCrLf & vbCrLf & "Anforderungen können nicht eingelesen werden!" & vbCrLf & "(Benötigt: " & strProjekt & " - " & strAttributeAVW & ")"
#               End If
#               blnImportAttribute(1) = False
                self.collect_errors(
                    "Anforderungen können nicht eingelesen werden!", 
                    f"{self.info_AVW.str_attributes()}, {self.str_AVW_attributeBEB21}"
                )
#           End If
#       ElseIf EinlesenDatei("Anforderungen Projektbereich", strAVWAttribute, rngAVWAttribute, wbAVW, wksAVW, strFehlerAVW, strDateinamen(1)) Then
#           ImportAttribute(1) = True
        else:
            self.import_attribute[0] = self.einlessen_datei()
        if not self.import_attribute[0]:
#           'Sammlung aller gesuchten Attribute erzeugen
#           strAttributeAVW = ""
#           For i = LBound(strAVWAttribute, 1) To UBound(strAVWAttribute, 1)
#               If strAttributeAVW = "" Then
#                   strAttributeAVW = strAVWAttribute(i)
#               Else
#                   strAttributeAVW = strAttributeAVW & ", " & strAVWAttribute(i)
#               End If
#           Next i
            strAttributeAVW = self.info_AVW.str_attributes()
#           If strFehlerGesamt = "" Then
#               strFehlerGesamt = "Anforderungen können nicht eingelesen werden!" & vbCrLf & "(Benötigt: " & strAttributeAVW & ")"
#           Else
#               strFehlerGesamt = strFehlerGesamt & vbCrLf & vbCrLf & "Anforderungen können nicht eingelesen werden!" & vbCrLf & "(Benötigt: " & strAttributeAVW & ")"
#           End If
            self.collecte_errors("Anforderungen können nicht eingelesen werden!", self.info_AVW)
#       End If
#       
#       If blnImportAttribute(1) Then
#           'Arbeitsblatt TDs - Verifikationskriterium
#           'TDVKs: #1: ID, #2: Basierend auf der Anforderung, #3: Status, #4: Temp1_Text, #5: Aktion
#           ReDim strTDVKAttribute(1 To 5)
#           ReDim rngTDVKAttribute(LBound(strTDVKAttribute, 1) To UBound(strTDVKAttribute, 1))
#           strTDVKAttribute(1) = "ID"
#           strTDVKAttribute(2) = "Basierend auf der Anforderung"
#           strTDVKAttribute(3) = "Status"
#           strTDVKAttribute(4) = "Temp1_Text"
#           strTDVKAttribute(5) = "Aktion"
        if self.import_attribute[0]:
            self.info_TDVK = DBInfo(
                attributes = (
                    "ID",
                    "Basierend auf der Anforderung",
                    "Status",
                    "Temp1_Text",
                    "Aktion",
                )
            )
#           'Dateiauswahl und Zuordnung
#           If EinlesenDatei("Verifikationskriterien", strTDVKAttribute, rngTDVKAttribute, wbTDVK, wksTDVK, strFehlerTDVK, strDateinamen(2)) Then
#               blnImportAttribute(2) = True
            self.import_attribute[1] = self.info_TDVK.einlesen_datei()
#           Else
#               'Sammlung aller gesuchten Attribute erzeugen
#               strAttributeTDVK = ""
#               For i = LBound(strTDVKAttribute, 1) To UBound(strTDVKAttribute, 1)
#                   If strAttributeTDVK = "" Then
#                       strAttributeTDVK = strTDVKAttribute(i)
#                   Else
#                       strAttributeTDVK = strAttributeTDVK & ", " & strTDVKAttribute(i)
#                   End If
#               Next i
            if not self.import_attribute[1]:
                # str_attribute_TDVK = self.info_TDVK.str_attributes()
#               If strFehlerGesamt = "" Then
#                   strFehlerGesamt = "Verifikationskriterien können nicht eingelesen werden!" & vbCrLf & "(Benötigt: " & strAttributeTDVK & ")"
#               Else
#                   strFehlerGesamt = strFehlerGesamt & vbCrLf & vbCrLf & "Verifikationskriterien können nicht eingelesen werden!" & vbCrLf & "(Benötigt: " & strAttributeTDVK & ")"
#               End If
                self.collect_errors("Verifikationskriterien können nicht eingelesen werden!", self.info_TDVK.str_attributes())
#               blnImportAttribute(2) = False
#           End If
#       End If
#       
#       If blnImportAttribute(2) Then
#           'Arbeitsblatt TDs - Absicherungsaufträge
#           'TDAAs: #1: ID, #2: Enthalten in, #3: Status, #4: Testinstanz, #5: Testumgebungstyp
#           ReDim strTDAAAttribute(1 To 5)
#           ReDim rngTDAAAttribute(LBound(strTDAAAttribute, 1) To UBound(strTDAAAttribute, 1))
        if self.import_attribute[1]:
#           strTDAAAttribute(1) = "ID"
#           strTDAAAttribute(2) = "Enthalten in"
#           strTDAAAttribute(3) = "Status"
#           strTDAAAttribute(4) = "Testinstanz"
#           strTDAAAttribute(5) = "Testumgebungstyp"
            self.info_TDAA = DBInfo(
                attributes = (
                    "ID",
                    "Enthalten in",
                    "Status",
                    "Testinstanz",
                    "Testumgebungstyp" ,
                )
            )
#           'Dateiauswahl und Zuordnung
#           If EinlesenDatei("Absicherungsaufträge", strTDAAAttribute, rngTDAAAttribute, wbTDAA, wksTDAA, strFehlerTDAA, strDateinamen(3)) Then
#               blnImportAttribute(3) = True
            self.import_attribute[3] = self.info_TDAA.einlessen_datei()
#           Else
#               'Sammlung aller gesuchten Attribute erzeugen
#               strAttributeTDAA = ""
#               For i = LBound(strTDAAAttribute, 1) To UBound(strTDAAAttribute, 1)
#                   If strAttributeTDAA = "" Then
#                       strAttributeTDAA = strTDAAAttribute(i)
#                   Else
#                       strAttributeTDAA = strAttributeTDAA & ", " & strTDAAAttribute(i)
#                   End If
#               Next i
            if not self.import_attribute[2]:
                #str_attribute_TDAA = self.info_TDAA.str_attributes()
#               If strFehlerGesamt = "" Then
#                   strFehlerGesamt = "Absicherungsaufträge können nicht eingelesen werden!" & vbCrLf & "(Benötigt: " & strAttributeTDAA & ")"
#               Else
#                   strFehlerGesamt = strFehlerGesamt & vbCrLf & vbCrLf & "Absicherungsaufträge können nicht eingelesen werden!" & vbCrLf & "(Benötigt: " & strAttributeTDAA & ")"
#               End If
                self.collect_errors("Absicherungsaufträge können nicht eingelesen werden!", self.info_TDAA.str_attributes())
#               blnImportAttribute(3) = False
#           End If
#       End If
#         
#       If blnImportAttribute(3) Then
#           'Arbeitsblatt Testfälle
#           'TFs: #1: ID, #2: Status, #3: Testfallname, #4: Sonstige-Varianten, #5: Basierend auf Testdesign, #6: verifiziert, #7: Testinstanz
#           ReDim strTFAttribute(1 To 7)
#           ReDim rngTFAttribute(LBound(strTFAttribute, 1) To UBound(strTFAttribute, 1))
#           strTFAttribute(1) = "ID"
#           strTFAttribute(2) = "Status"
#           strTFAttribute(3) = "Testfallname"
#           strTFAttribute(4) = "Sonstige-Varianten"
#           strTFAttribute(5) = "Basierend auf Testdesign"
#           strTFAttribute(6) = "verifiziert"
#           strTFAttribute(7) = "Testinstanz"
        if self.import_attribute[2]:
            self.info_TF = DBInfo(
                attributes = (
                    "ID",
                    "Status",
                    "Testfallname",
                    "Sonstige-Varianten",
                    "Basierend auf Testdesign",
                    "verifiziert",
                    "Testinstanz",
                )
            )
#           'Dateiauswahl und Zuordnung
#           If EinlesenDatei("Testfälle", strTFAttribute, rngTFAttribute, wbTF, wksTF, strFehlerTF, strDateinamen(4)) Then
#               blnImportAttribute(4) = True
            self.import_attribute[3] = self.info_TF.einlessen_datei()
#           Else
            if not self.import_attribute[3]:
#               'Sammlung aller gesuchten Attribute erzeugen
#               strAttributeTF = ""
#               For i = LBound(strTFAttribute, 1) To UBound(strTFAttribute, 1)
#                   If strAttributeTF = "" Then
#                       strAttributeTF = strTFAttribute(i)
#                   Else
#                       strAttributeTF = strAttributeTF & ", " & strTFAttribute(i)
#                   End If
#               Next i
                #str_attribute_TF = self.info_TF.str_attributes()
#               If strFehlerGesamt = "" Then
#                   strFehlerGesamt = "Testfälle können nicht eingelesen werden!" & vbCrLf & "(Benötigt: " & strAttributeTF & ")"
#               Else
#                   strFehlerGesamt = strFehlerGesamt & vbCrLf & vbCrLf & "Testfälle können nicht eingelesen werden!" & vbCrLf & "(Benötigt: " & strAttributeTF & ")"
#               End If
                self.collect_errors("Testfälle können nicht eingelesen werden!", self.info_TF.str_attributes())
#               blnImportAttribute(4) = False
#           End If
#       End If
#                   
#       If blnImportAttribute(4) Then
        if self.import_attribute[3]:
#           'Arbeitsblatt FRU-Timing
#           'FRUTiming: #1: FeatureName, #2: RG, #3: Umsetzer, #4: Zuordnung zu I-Stufe
#           ReDim strFRUTimingAttribute(1 To 4)
#           ReDim rngFRUTimingAttribute(LBound(strFRUTimingAttribute, 1) To UBound(strFRUTimingAttribute, 1))
#           strFRUTimingAttribute(1) = "FeatureName"
#           strFRUTimingAttribute(2) = "Reifegrad"  'vorher "RG"
#           strFRUTimingAttribute(3) = "Umsetzer"
#           strFRUTimingAttribute(4) = "FE_Meilenstein" 'vorher "Zuordnung zu I-Stufe"
            self.info_fru_timming = DBInfo(
                attributes = (
                    "FeatureName",
                    "Reifegrad",  #vorher "RG"
                    "Umsetzer",
                    "FE_Meilenstein", #vorher "Zuordnung zu I-Stufe",
                )
            )
#           'Dateiauswahl und Zuordnung
#           If EinlesenDatei("FRU-Timing", strFRUTimingAttribute, rngFRUTimingAttribute, wbFRUTiming, wksFRUTiming, strFehlerFRUTiming, strDateinamen(5)) Then
#               blnImportAttribute(5) = True
            self.import_attribute[4] = self.info_fru_timming.einlessen_datei()
#           Else
            if not self.import_attribute[4]:
#               'Sammlung aller gesuchten Attribute erzeugen
#               strAttributeFRUTiming = ""
#               For i = LBound(strFRUTimingAttribute, 1) To UBound(strFRUTimingAttribute, 1)
#                   If strAttributeFRUTiming = "" Then
#                       strAttributeFRUTiming = strFRUTimingAttribute(i)
#                   Else
#                       strAttributeFRUTiming = strAttributeFRUTiming & ", " & strFRUTimingAttribute(i)
#                   End If
#               Next i
                #str_attribute_fru_timming = self.info_FRUTimming.str_attributes()
#               If strFehlerGesamt = "" Then
#                   strFehlerGesamt = "FRU-Timing kann nicht eingelesen werden!" & vbCrLf & "(Benötigt: " & strAttributeFRUTiming & ")"
#               Else
#                   strFehlerGesamt = strFehlerGesamt & vbCrLf & vbCrLf & "FRU-Timing kann nicht eingelesen werden!" & vbCrLf & "(Benötigt: " & strAttributeFRUTiming & ")"
#               End If
                self.collect_errors("FRU-Timing kann nicht eingelesen werden!", self.info_AVW.str_attributes())
#               blnImportAttribute(5) = False
#           End If
#       End If
#       
#       If blnAVWVorgaengerIDsVerwenden = True Then
        if self.predecessorIdUsed:
#           If blnImportAttribute(5) = True Then
            if self.import_attribute[4]:
#               'Arbeitsblatt AVWMaster_Rohdaten
#               'AVW: #1: ID, #2: temp1_Text, #3: Kommentar Redaktionskreis
#               'Weiche für Erfassung der Nachfolger-IDs
#               ReDim strAVWMasterAttribute(1 To 3)
#               ReDim rngAVWMandyAttribute(LBound(strAVWMasterAttribute, 1) To UBound(strAVWMasterAttribute, 1))
#               strAVWMasterAttribute(1) = "ID"
#               strAVWMasterAttribute(2) = "temp1_Text"
#               strAVWMasterAttribute(3) = "Kommentar Redaktionskreis"
                self.info_AVW_master = DBInfo(
                    attributes = (
                        "ID",
                        "temp1_Text",
                        "Kommentar Redaktionskreis",
                    )
                )
#               'Dateiauswahl und Zuordnung
#               If EinlesenDatei("Anforderungen Masterbereich", strAVWMasterAttribute, rngAVWMasterAttribute, wbAVWMaster, wksAVWMaster, strFehlerAVWMaster, strDateinamen(6)) Then
#                   blnImportAttribute(6) = True
                self.import_attribute[5] = self.info_AVW_master.einlessen_datei()
#               Else
                if not self.import_attribute[5]:
#                   'Sammlung aller gesuchten Attribute erzeugen
#                   strAttributeAVWMaster = ""
#                   For i = LBound(strAVWMasterAttribute, 1) To UBound(strAVWMasterAttribute, 1)
#                       If strAttributeAVWMaster = "" Then
#                           strAttributeAVWMaster = strAVWMasterAttribute(i)
#                       Else
#                           strAttributeAVWMaster = strAttributeAVWMaster & ", " & strAVWMasterAttribute(i)
#                       End If
#                   Next i
                    # strAttributeAVWMaster = self.info_AVW_master.str_attributes()
#                   If strFehlerGesamt = "" Then
#                       strFehlerGesamt = "Anforderungen aus dem Masterbereich können nicht eingelesen werden!" & vbCrLf & "(Benötigt: " & strAttributeAVWMaster & ")"
#                   Else
#                       strFehlerGesamt = strFehlerGesamt & vbCrLf & vbCrLf & "Anforderungen aus dem Masterbereich können nicht eingelesen werden!" & vbCrLf & "(Benötigt: " & strAttributeAVWMaster & ")"
#                   End If
                    self.collect_errors(
                        "Anforderungen aus dem Masterbereich können nicht eingelesen werden!",
                        self.info_AVW_master.str_attributes(),
                    )
#                   blnImportAttribute(6) = False
#               End If
#           End If
#       End If
#       
#       'Rückgabewert
#       If blnAVWVorgaengerIDsVerwenden = False Then
#           If blnImportAttribute(1) And blnImportAttribute(2) And blnImportAttribute(3) And blnImportAttribute(4) And blnImportAttribute(5) Then
#               ATE_Status_Initializer = True
#           Else
#               ATE_Status_Initializer = False
#           End If
#       Else
#           If blnImportAttribute(1) And blnImportAttribute(2) And blnImportAttribute(3) And blnImportAttribute(4) And blnImportAttribute(5) And blnImportAttribute(6) Then
#               ATE_Status_Initializer = True
#           Else
#               ATE_Status_Initializer = False
#           End If
#       End If
        
        return all(self.import_attribute[:-1]) and (not self.predecessorIdUsed or self.import_attribute[5])
#   End Function

    def _status_initializer(self) -> bool:
        return True
    
    def collect_errors(self, error: str, attributes:str):
        if self.errors == "":
            self.errors = f"{error}\n(Benötigt: {self.project} - {attributes})"
        else:        
            self.errors += f"\n\n{error}\n(Benötigt: {self.project} - {attributes})"

#
#Private Sub ATE_Status_Deinitializer(ByRef wksBsM As Worksheet, ByRef wksAVW As Worksheet, ByRef wksTDVK As Worksheet, ByRef wksTDAA As Worksheet, ByRef wksTF As Worksheet, ByRef wksFRUTiming As Worksheet)
#'Worksheets identifizieren
#Set wksBsM = Nothing
#Set wksAVW = Nothing
#Set wksTDVK = Nothing
#Set wksTDAA = Nothing
#Set wksTF = Nothing
#Set wksFRUTiming = Nothing
#Set verifikationKritList = Nothing
#Set BsMDatenList = Nothing
#Set testfallList = Nothing
#Set FRUTimingList = Nothing
#Set AVWVorgaengerList = Nothing
#End Sub
#
# Funtion EinlesenDatei moved to om/db_info.py
#
#Public Function EinlesenDatei_Projektspezifisch(ByVal strTitel As String, ByRef strAttribute() As String, ByRef rngAttribute() As Range, ByRef wbImport As Workbook, ByRef wksImport As Worksheet, ByRef strFehler As String, ByRef strDateinamen As String, _
#                                                ByVal strProjekt As String, ByRef strAttributeProjekt() As String, ByRef rngAttributeProjekt() As Range) As Boolean
#Dim blnAttributeImport As Boolean   'Flag für Attribute
#Dim blnAttributeProjektImport As Boolean    'Flag für projektspezifische Attribute
#Dim intWksImport As Integer         'Zähler für Worksheets
#Dim strImportPfad As String         'String für Dateipfad
#Dim strImportDatei As String        'String für Dateinamen
#Dim strFehlerGesamt As String       'Sting für Gesamtfehler
#Dim intAttributeZaehler As Integer  'Zähler für Attribute
#Dim strFehlerAttribute As String    'String für fehlende Attribute
#Dim intAttributeProjektZaehler As Integer   'Zähler für Projekt-Attribute
#Dim strFehlerProjektAttribute As String 'String für fehlende Projekt-Attribute
#Dim blnErsterFund As Boolean        'Flag für erstes Auffinden des vollständigen Attributesatzes
#Dim rngAttributeErsterFund() As Range   'Range-Array für erstes Auffinden des vollständigen allgemeinen Attributesatzes
#Dim rngAttributeProjektErsterFund() As Range    'Range-Array für erstes Auffinden des vollständigen projektspezifischen Attributesatzes
#Dim wksImportErsterFund As Worksheet    'Worksheet für erstes Auffinden des vollständigen Attributesatzes
#Dim lngZeilenZahlerErsterFund As Long   'Long für letzte Zeile des ersten Fundes
#Dim lngWeitererFundZaehler As Long  'Long für Startzeile des zusätzlichen Datensatzen
#
#On Error Resume Next
#
#blnAttributeImport = False
#blnAttributeProjektImport = False
#strFehlerGesamt = ""
#strDateinamen = ""
#blnErsterFund = False
#
#strImportPfad = Application.GetOpenFilename(FileFilter:="Excel-Dateien (*.xls; *.xlsx; *.xlm; *.xlsm), *.xls; *.xlsx; *.xlm; *.xlsm", FilterIndex:=1, Title:=strTitel & " auswählen")
#If Trim(strImportPfad) <> "Falsch" Then
#    'Datei öffnen
#    Workbooks.Open Filename:=strImportPfad, ReadOnly:=True
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
#        ReDim rngAttributeProjekt(LBound(strAttributeProjekt, 1) To UBound(strAttributeProjekt, 1))
#        
#        'Zähler für Arbeitsblatt erhöhen
#        intWksImport = intWksImport + 1
#        'Worksheet zuweisen
#        Set wksImport = wbImport.Sheets(intWksImport)
#        
#        'Allgemeine Attribute suchen
#        For intAttributeZaehler = LBound(strAttribute, 1) To UBound(strAttribute, 1)
#            Set rngAttribute(intAttributeZaehler) = wksImport.Cells.Find(strAttribute(intAttributeZaehler), lookat:=xlWhole)
#        Next intAttributeZaehler
#        
#        'Projekt-Attribute suchen
#        For intAttributeProjektZaehler = LBound(strAttributeProjekt, 1) To UBound(strAttributeProjekt, 1)
#            Set rngAttributeProjekt(intAttributeProjektZaehler) = wksImport.Cells.Find(strAttributeProjekt(intAttributeProjektZaehler), lookat:=xlWhole)
#        Next intAttributeProjektZaehler
#        
#        'Suche der allgemeinen Attribute auswerten
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
#        
#        'Suche der Projekt-Attribute auswerten
#        strFehlerProjektAttribute = ""
#        If RangeObjekteVorhandenFehlerausgabe(rngAttributeProjekt, strAttributeProjekt, strFehlerProjektAttribute) Then
#            blnAttributeProjektImport = True
#        Else
#            If strFehlerGesamt = "" Then
#                strFehlerGesamt = wksImport.Name & ": " & strProjekt & " - " & strFehlerProjektAttribute
#            Else
#                strFehlerGesamt = strFehlerGesamt & "; " & wksImport.Name & ": " & strProjekt & " - " & strFehlerProjektAttribute
#            End If
#        End If
#        
#        'Auswertung beider Suchen (allgemein und projektspezifisch) und Unterscheidung, ob es das erste Auffinden des vollständigen Attributsatzes ist oder weitere
#        'Beim ersten Auffinden werden die beiden Range-Arays gespeichert
#        'Ab dem zweitem Auffinden des vollständigen Attributsatzes erfolgt eine Abfrage, eine Kopiervorganges der dazugehörigen Daten zum ersten Datensatz erfolgen soll und ggf. Umsetzung
#        If blnAttributeImport = True And blnAttributeProjektImport = True Then
#            'Zurücksetzen der Flags für den einzelnen Durchlauf
#            blnAttributeImport = False
#            blnAttributeProjektImport = False
#            
#            If blnErsterFund = False Then
#                blnErsterFund = True
#                'Felder dimensionieren
#                ReDim rngAttributeErsterFund(LBound(strAttribute, 1) To UBound(strAttribute, 1))
#                ReDim rngAttributeProjektErsterFund(LBound(strAttributeProjekt, 1) To UBound(strAttributeProjekt, 1))
#                'Übernahme allgemeine Attribute
#                For intAttributeZaehler = LBound(strAttribute, 1) To UBound(strAttribute, 1)
#                    Set rngAttributeErsterFund(intAttributeZaehler) = rngAttribute(intAttributeZaehler)
#                Next intAttributeZaehler
#                'Übernahme projektspezifische Attribute
#                For intAttributeProjektZaehler = LBound(strAttributeProjekt, 1) To UBound(strAttributeProjekt, 1)
#                    Set rngAttributeProjektErsterFund(intAttributeProjektZaehler) = rngAttributeProjekt(intAttributeProjektZaehler)
#                Next intAttributeProjektZaehler
#                'Übernahme Arbeitsblatt
#                Set wksImportErsterFund = wksImport
#            Else
#                If MsgBox("Weiterer Datensatz gefunden!" & vbCrLf & "Erster Fund: " & wksImportErsterFund.Name & vbCrLf & "Aktueller Fund: " & wksImport.Name & vbCrLf & vbCrLf & "Sollen die Datensätze zusammengeführt werden?", vbYesNo) = vbYes Then
#                    'Alle Zeilen mit Einträgen durchlaufen und kopieren
#                    For lngWeitererFundZaehler = 1 To wksImport.UsedRange.Rows.Count - rngAttribute(1).Row
#                        If rngAttribute(1).Offset(lngWeitererFundZaehler, 0).Value <> "" Then
#                            'Zielzeile ermitteln
#                            lngZeilenZahlerErsterFund = rngAttributeErsterFund(1).End(xlDown).Row + 1
#                            
#                            'Allgemeine Werte kopieren
#                            For intAttributeZaehler = LBound(strAttribute, 1) To UBound(strAttribute, 1)
#                                wksImportErsterFund.Cells(lngZeilenZahlerErsterFund, rngAttributeErsterFund(intAttributeZaehler).Column).Value = rngAttribute(intAttributeZaehler).Offset(lngWeitererFundZaehler).Value
#                            Next intAttributeZaehler
#                            
#                            'Projektspezifische Werte kopieren
#                            For intAttributeProjektZaehler = LBound(strAttributeProjekt, 1) To UBound(strAttributeProjekt, 1)
#                                wksImportErsterFund.Cells(lngZeilenZahlerErsterFund, rngAttributeProjektErsterFund(intAttributeProjektZaehler).Column).Value = rngAttributeProjekt(intAttributeProjektZaehler).Offset(lngWeitererFundZaehler).Value
#                            Next intAttributeProjektZaehler
#                        End If
#                    Next lngWeitererFundZaehler
#                End If
#            End If
#        End If
#        
#    'Loop While intWksImport < wbImport.Worksheets.Count And blnAttributeImport = False And blnAttributeProjektImport = False
#    Loop While intWksImport < wbImport.Worksheets.Count
#End If
#
#
#'Rückgabe allgemeine Attribute
#For intAttributeZaehler = LBound(strAttribute, 1) To UBound(strAttribute, 1)
#    Set rngAttribute(intAttributeZaehler) = rngAttributeErsterFund(intAttributeZaehler)
#Next intAttributeZaehler
#'Rückgabe projektspezifische Attribute
#For intAttributeProjektZaehler = LBound(strAttributeProjekt, 1) To UBound(strAttributeProjekt, 1)
#    Set rngAttributeProjekt(intAttributeProjektZaehler) = rngAttributeProjektErsterFund(intAttributeProjektZaehler)
#Next intAttributeProjektZaehler
#'Rückgabe Arbeitsblatt
#Set wksImport = wksImportErsterFund
#
#'Rückgabewert übernehmen
#EinlesenDatei_Projektspezifisch = blnErsterFund
#'Fehlerwert übernehmen
#strFehler = strFehlerGesamt
#End Function
#
#Private Sub SchliessenWb(ByVal wbBsM As Workbook, ByVal wbAVW As Workbook, ByVal wbAVWMaster As Workbook, ByVal wbTDVK As Workbook, ByVal wbTDAA As Workbook, ByVal wbTF As Workbook, ByRef wbFRUTiming As Workbook)
#'wbBsM.Close SaveChanges:=False
#wbAVW.Close SaveChanges:=False
#If Not wbAVWMaster Is Nothing Then
#    wbAVWMaster.Close SaveChanges:=False
#End If
#wbTDVK.Close SaveChanges:=False
#wbTDAA.Close SaveChanges:=False
#wbTF.Close SaveChanges:=False
#wbFRUTiming.Close SaveChanges:=False
#End Sub
#
#Private Function RangeObjekteVorhandenFehlerausgabe(ByRef rngObjekte() As Range, ByRef strObjekteNamen() As String, ByRef strNichtVorhandeneObjekte As String) As Boolean
#Dim i As Integer
#
#RangeObjekteVorhandenFehlerausgabe = True
#strNichtVorhandeneObjekte = ""
#For i = LBound(rngObjekte) To UBound(rngObjekte)
#    If rngObjekte(i) Is Nothing Then
#        RangeObjekteVorhandenFehlerausgabe = False
#        If strNichtVorhandeneObjekte = "" Then
#            strNichtVorhandeneObjekte = strObjekteNamen(i)
#        Else
#            strNichtVorhandeneObjekte = strNichtVorhandeneObjekte & ", " & strObjekteNamen(i)
#        End If
#    End If
#Next i
#End Function
#
#Private Sub EinlesenLAHBlacklist(ByVal wbBsM As Workbook, ByRef strLAHBlacklist() As String)
#Dim strWKSBlacklist As String           'String für Namen des Blacklist-Worksheets
#Dim wksBlacklist As Worksheet           'Worksheet für Blacklist
#Dim strAttributBlacklist As String      'String für Attribut der Blacklist
#Dim rngBlacklist As Range               'Range für Blacklist
#Dim lngBlacklist As Long                'Zählvariable für Blacklist
#Dim lngBlacklistErfasst As Long         'Zählvariable für bereits erfasste Blacklist-Einträge
#Dim blnBlacklistItemErfasst As Boolean  'Flag für bereits erfasste Blacklist-Items
#Dim lngZeile As Long                    'Zeilenzähler
#
#On Error Resume Next
#
#strWKSBlacklist = "Blacklist"
#strAttributBlacklist = "LAH, die ignoriert werden sollen"
#lngBlacklist = 0
#ReDim strLAHBlacklist(0)
#
#Set wksBlacklist = wbBsM.Sheets(strWKSBlacklist)
#If Not wksBlacklist Is Nothing Then
#    Set rngBlacklist = wksBlacklist.Cells.Find(strAttributBlacklist, lookat:=xlWhole)
#    If Not rngBlacklist Is Nothing Then
#        For lngZeile = 1 To wksBlacklist.UsedRange.Rows.Count - rngBlacklist.Row
#            blnBlacklistItemErfasst = False
#            If rngBlacklist.Offset(lngZeile, 0) <> "" Then
#                If lngBlacklist > 0 Then
#                    For lngBlacklistErfasst = LBound(strLAHBlacklist, 1) To UBound(strLAHBlacklist, 1)
#                        If strLAHBlacklist(lngBlacklistErfasst) = rngBlacklist.Offset(lngZeile, 0) Then
#                            blnBlacklistItemErfasst = True
#                            Exit For
#                        End If
#                    Next lngBlacklistErfasst
#                    If blnBlacklistItemErfasst = False Then
#                        lngBlacklist = lngBlacklist + 1
#                        ReDim Preserve strLAHBlacklist(1 To lngBlacklist)
#                        strLAHBlacklist(lngBlacklist) = rngBlacklist.Offset(lngZeile, 0)
#                    End If
#                Else
#                    lngBlacklist = 1
#                    ReDim strLAHBlacklist(1 To lngBlacklist)
#                    strLAHBlacklist(lngBlacklist) = rngBlacklist.Offset(lngZeile, 0)
#                End If
#            End If
#        Next lngZeile
#    Else
#        MsgBox "Attribut der Blacklist """ & strAttributBlacklist & """ ist nicht vorhanden!"
#    End If
#Else
#    MsgBox "Arbeitsblatt """ & strWKSBlacklist & """ ist nicht vorhanden!"
#End If
#End Sub
#
#Private Sub EinlesenTDVKs(ByVal wksTDVK As Worksheet, ByRef strTDVKAttribute() As String, ByRef rngTDVKAttribute() As Range)
#Dim anfIDs As String                            'String für eingelesene Anforderungs-IDs
#Dim verifikationKrit As Verifikationskriterium  'Klasse Verifikationskriterium
#Dim idList As Collection                        'Liste mit Anforderungs-IDs, die zu einem Verifikationskriterium eingelesen werden
#Dim lngZeile As Long                            'Long-Zähler für aktuell einzulesende Zeile
#Dim strVerifikationsID As String                'String für ID des Verifikationskriteriums
#
#'Verifikationskriterien einlesen
#'TDVKs: #1: ID, #2: Basierend auf der Anforderung, #3: Status, #4: Temp1_Text, #5: Aktion
#For lngZeile = 1 To wksTDVK.UsedRange.Rows.Count - rngTDVKAttribute(1).Row
#'    If rngTDVKAttribute(3).Offset(lngZeile, 0).Value = "Fachlich abgestimmt" Then
#        'Neues Verifikationskriterium anlegen
#        Set verifikationKrit = New Verifikationskriterium
#        'ID des Verifikationsauftrags einlesen, Entfernung der zusätzlichen Zeichen "?" und "r"
#        strVerifikationsID = Replace(Replace(rngTDVKAttribute(1).Offset(lngZeile, 0).Value, "?", ""), "r", "")
#        'ID des Verifikationsauftrags erfassen
#        verifikationKrit.VK_ID = strVerifikationsID
#        'Anforderungs-IDs einlesen
#        anfIDs = rngTDVKAttribute(2).Offset(lngZeile, 0).Value
#        'Anforderungs-IDs nach Kommas trennen
#        Set idList = EinlesenGetrennteWerteKomma(anfIDs)
#        'Alle mit dem aktuellen Verifikationskriterium verknüpften Anforderungs-IDs erfassen
#        Set verifikationKrit.anf_ids = idList
#        'Status des Verifikationskriteriums einlesen
#        verifikationKrit.VK_status = rngTDVKAttribute(3).Offset(lngZeile, 0).Value
#        'Absicherungsaufträge für dieses Verifikationskriterium anlegen
#        Set verifikationKrit.Absicherungsauftraege = New Collection
#        'Sammlung für Testfälle vorbereiten
#        Set verifikationKrit.VK_Testfaelle = New Collection
#        'Sammlung für I-Stufen vorbereiten
#        Set verifikationKrit.anf_IStufen = New Collection
#        'Sammlung für Umsetzer vorbereiten
#        Set verifikationKrit.anf_Umsetzer = New Collection
#        'Sammlung für BsM-Relevanz vorbereiten
#        Set verifikationKrit.anf_BsMRelevanz = New Collection
#        'Sammlung für ASIL vorbereiten
#        Set verifikationKrit.anf_ASIL = New Collection
#        'Sammlung für Feature vorbereiten
#        Set verifikationKrit.anf_Feature = New Collection
#        'Sammlung für Reifegrad vorbereiten
#        Set verifikationKrit.anf_Reifegrad = New Collection
#        'Sammlung für Modulverantwortliche vorbereiten
#        Set verifikationKrit.anf_MV = New Collection
#        'Sammlung für LAH-ID vorbereiten
#        Set verifikationKrit.anf_LAHID = New Collection
#        'Sammlung für LAH-Namen vorbereiten
#        Set verifikationKrit.anf_LAHNamen = New Collection
#        'Sammlung für Cluster Testing vorbereiten
#        Set verifikationKrit.anf_ClusterTesting = New Collection
#        'Sammlung für Anforderungsverantwortliche vorbereiten
#        Set verifikationKrit.anf_Anforderungsverantwortliche = New Collection
#        'Sammlung für Temp11_Auswahlfeld vorbereiten
#        Set verifikationKrit.anf_Temp11_Auswahlfeld = New Collection
#        'temp1_text einlesen
#        verifikationKrit.VK_temp1Text = rngTDVKAttribute(4).Offset(lngZeile, 0).Value
#        'Aktion einlesen
#        verifikationKrit.VK_Aktion = rngTDVKAttribute(5).Offset(lngZeile, 0).Value
#
#        'Erfasstes Verifikationskriterium in globaler Verifikationskriterien-Liste hinzufügen
#        verifikationKritList.Add Item:=verifikationKrit, Key:=verifikationKrit.VK_ID
#'    End If
#    'Fortschritt anzeigen
#    If lngZeile Mod 100 = 0 Then
#        Debug.Print "Verifikationskriterien einlesen: " & lngZeile & "/" & wksTDVK.UsedRange.Rows.Count - rngTDVKAttribute(1).Row
#    End If
#Next lngZeile
#End Sub
#
#Private Sub EinlesenTDAAs(ByVal wksTDAA As Worksheet, ByRef strTDAAAttribute() As String, ByRef rngTDAAAttribute() As Range)
#Dim anfIDs As String                            'String für eingelesene Anforderungs-IDs
#Dim absicherungsAuftr As Absicherungsauftraege  'Klasse Absicherungsauftraege
#Dim lngZeile As Long                            'Long-Zähler für aktuell einzulesende Zeile
#Dim strVerifikationsID As String                'String für ID des Verifikationskriteriums
#Dim Verifikationskriterium As Verifikationskriterium    'Verifikationskriterium für Zuordnung
#'Absicherungsaufträge einlesen
#'TDAAs: #1: ID, #2: Enthalten in, #3: Status, #4: Testinstanz, #5: Testumgebungstyp
#For lngZeile = 1 To wksTDAA.UsedRange.Rows.Count - rngTDAAAttribute(1).Row
#    If rngTDAAAttribute(3).Offset(lngZeile, 0).Value = "Fachlich abgestimmt" Or rngTDAAAttribute(3).Offset(lngZeile, 0).Value = "In Review" Or rngTDAAAttribute(3).Offset(lngZeile, 0).Value = "In Bearbeitung" Then
#        'Neuen Absicherungsauftrag anlegen
#        Set absicherungsAuftr = New Absicherungsauftraege
#        'Testinstanz einlesen
#        absicherungsAuftr.testinstanz = rngTDAAAttribute(4).Offset(lngZeile, 0).Value
#        'Testumgebung einlesen
#        absicherungsAuftr.Testumgebungstyp = Replace(rngTDAAAttribute(5).Offset(lngZeile, 0).Value, "Testumgebungstyp: ", "")
#        'Status des Absicherungsauftrages einlesen
#        absicherungsAuftr.abs_status = rngTDAAAttribute(3).Offset(lngZeile, 0).Value
#        'ID des übergeordneten Verifikationsauftrags einlesen, Entfernung der zusätzlichen Zeichen "?" und "r"
#        strVerifikationsID = Replace(Replace(rngTDAAAttribute(2).Offset(lngZeile, 0).Value, "?", ""), "r", "")
#        'ID des Absicherungsauftrages einlesen, Entfernung der zusätzlichen Zeichen "?" und "r"
#        absicherungsAuftr.abs_ID = Replace(Replace(rngTDAAAttribute(1).Offset(lngZeile, 0).Value, "?", ""), "r", "")
#        
#        'Zuordnung zu Verifikationskriterium in globaler Verifikationskriterien-Liste
#        Set Verifikationskriterium = New Verifikationskriterium
#        Set Verifikationskriterium = FindeVK(verifikationKritList, strVerifikationsID)
#        If Not Verifikationskriterium Is Nothing Then
#            Verifikationskriterium.Absicherungsauftraege.Add Item:=absicherungsAuftr, Key:=absicherungsAuftr.abs_ID
#        End If
#    End If
#    'Fortschritt anzeigen
#    If lngZeile Mod 100 = 0 Then
#        Debug.Print "Absicherungsaufträge einlesen: " & lngZeile & "/" & wksTDAA.UsedRange.Rows.Count - rngTDAAAttribute(1).Row
#    End If
#Next lngZeile
#End Sub
#
#Private Sub EinlesenTFs(ByVal wksTF As Worksheet, ByRef strTFAttribute() As String, ByRef rngTFAttribute() As Range)
#Dim lngZeile As Long                            'Long-Zähler für aktuell einzulesende Zeile
#Dim testfall As Testfaelle                      'Klasse Testfaelle
#Dim anfIDs As String                            'String für eingelesene Anforderungs-IDs
#Dim idList As Collection                        'Liste mit Anforderungs-IDs
#Dim varErfassteVKItem As Variant                'Variant für Item in der globalen Verifikationskriterien-Liste
#Dim varAnfID As Variant                         'Anforderungs-ID aus varErfassteVKItem
#Dim Verifikationskriterium As Verifikationskriterium    'Verifikationskriterium für Zuordnung
#
#'Testfälle einlesen
#'TFs: #1: ID, #2: Status, #3: Testfallname, #4: Sonstige-Varianten, #5: Basierend auf Testdesign, #6: verifiziert, #7: Testinstanz
#For lngZeile = 1 To wksTF.UsedRange.Rows.Count - rngTFAttribute(1).Row
#'    If rngTFAttribute(2).Offset(lngZeile, 0).Value = "Operativ" Then
#        'Neuen Testfall anlegen
#        Set testfall = New Testfaelle
#        'Testfall-ID einlesen, Entfernung der zusätzlichen Zeichen "?" und "r"
#        testfall.TF_ID = Replace(Replace(rngTFAttribute(1).Offset(lngZeile, 0).Value, "?", ""), "r", "")
#        'Status Testfall einlesen
#        testfall.TF_Status = rngTFAttribute(2).Offset(lngZeile, 0).Value
#        'Testfall-Name einlesen
#        testfall.TF_Name = rngTFAttribute(3).Offset(lngZeile, 0).Value
#        'Testinstanz einlesen
#        testfall.TF_Testinstanz = rngTFAttribute(7).Offset(lngZeile, 0).Value
#        'Testumgebungstyp einlesen
#        testfall.TF_Testumgebungstyp = Replace(rngTFAttribute(4).Offset(lngZeile, 0).Value, "Testumgebungstyp: ", "")
#        
#        'Alle direkt mit dem aktuellen Testfall verknüpften Anforderungs-IDs erfassen
#        'direkte Testfälle nicht berücksichtigen!
#        'Anforderungs-IDs einlesen
#        'anfIDs = rngTFAttribute(6).Offset(lngZeile, 0).Value
#        'Anforderungs-IDs nach Kommas trennen
#        'Set idList = EinlesenGetrennteWerteKomma(anfIDs)
#        'Anforderungs-IDs übernehmen
#        'Set testfall.TF_anfIDs = idList
#        'Neue Sammlung für Anforderungs-IDs anlegen - Notwendig, wenn Liste der direkten Testfälle nicht übernommen wird
#        Set testfall.TF_anfIDs = New Collection
#        
#        'Alle über das Testdesign mit dem aktuellen Testfall verknüpften Anforderungs-IDs erfassen
#        'ID des übergeordneten Verifikationsauftrags einlesen, Entfernung der zusätzlichen Zeichen "?" und "r"
#        testfall.TF_VK_ID = Replace(Replace(rngTFAttribute(5).Offset(lngZeile, 0).Value, "?", ""), "r", "")
#        'Zuordnung zu Verifikationskriterium in globaler Verifikationskriterien-Liste
#        Set Verifikationskriterium = New Verifikationskriterium
#        Set Verifikationskriterium = FindeVK(verifikationKritList, testfall.TF_VK_ID)
#        If Not Verifikationskriterium Is Nothing Then
#            'Anforderungs-ID aufnehmen
#            For Each varAnfID In Verifikationskriterium.anf_ids
#                testfall.addElementID (varAnfID)
#            Next varAnfID
#            'Testfall aufnehmen
#            Verifikationskriterium.VK_Testfaelle.Add Item:=testfall, Key:=testfall.TF_ID
#        End If
#        
#        'Erfassten Testfall in globaler Testfall-Liste hinzufügen
#        testfallList.Add Item:=testfall, Key:=testfall.TF_ID
#'    End If
#    'Fortschritt anzeigen
#    If lngZeile Mod 100 = 0 Then
#        Debug.Print "Testfälle einlesen: " & lngZeile & "/" & wksTF.UsedRange.Rows.Count - rngTFAttribute(1).Row
#    End If
#Next lngZeile
#End Sub
#
#Private Sub EinlesenFRUTiming(ByVal wksFRUTiming As Worksheet, ByRef strFRUTimingAttribute() As String, ByRef rngFRUTimingAttribute() As Range)
#Dim lngZeile As Long                    'Long-Zähler für aktuell einzulesende Zeile
#Dim FRUTiming As FRUTiming              'Klasse FRUTiming
#Dim strFRUKey As String                 'Key für Item in FRUTiming
#
#'Doppelte Einträge im FRU-Import ignorieren
#On Error Resume Next
#
#'FRU-Timing eonlesen
#'FRUTiming: #1: FeatureName, #2: RG, #3: Umsetzer, #4: Zuordnung zu I-Stufe
#For lngZeile = 1 To wksFRUTiming.UsedRange.Rows.Count - rngFRUTimingAttribute(1).Row
#    If rngFRUTimingAttribute(4).Offset(lngZeile, 0).Value <> "" Then
#        'Neues FRUTiming anlegen
#        Set FRUTiming = New FRUTiming
#        'Feature einlesen
#        FRUTiming.Feature = rngFRUTimingAttribute(1).Offset(lngZeile, 0).Value
#        'Reifegrad einlesen
#        FRUTiming.Reifegrad = rngFRUTimingAttribute(2).Offset(lngZeile, 0).Value
#        'Umsetzer einlesen
#        FRUTiming.Umsetzer = rngFRUTimingAttribute(3).Offset(lngZeile, 0).Value
#        'I-Stufe einlesen
#        FRUTiming.IStufe = rngFRUTimingAttribute(4).Offset(lngZeile, 0).Value
#        'FRU-Key erzeugen
#        strFRUKey = FRUTiming.Feature & FRUTiming.Reifegrad & FRUTiming.Umsetzer
#        'Erfasstes FRU-Timing in globaler FRUTiming-Liste hinzufügen
#        FRUTimingList.Add Item:=FRUTiming, Key:=strFRUKey
#    End If
#    'Fortschritt anzeigen
#    If lngZeile Mod 100 = 0 Then
#        Debug.Print "FRU-Timing einlesen: " & lngZeile & "/" & wksFRUTiming.UsedRange.Rows.Count - rngFRUTimingAttribute(1).Row
#    End If
#Next lngZeile
#End Sub
#
#Private Sub EinlesenAVWRohdaten(ByVal wksAVW As Worksheet, ByRef strAVWAttribute() As String, ByRef rngAVWAttribute() As Range, ByRef strLAHBlacklist() As String, _
#                                ByVal strProjekt As String, ByRef strAVWAttributeMEB21() As String, ByRef rngAVWAttributeMEB21() As Range)
#Dim lngZeile As Long                    'Long-Variable für aktuell einzulesende Zeile
#Dim BSMDatensatz As BSMDaten            'Klasse BSMDaten
#Dim strBsMRelevanz As String            'String-Variable für Zusammenfassung der BsM-Relevanz
#Const strBsMVorhanden As String = "ja"  'Konstante für Angabe aus AVW für vorhandenes BsM-Attribut
#Dim varErfassteVKItem As Variant        'Variant für Item in der globalen Verifikationskriterien-Liste
#Dim varVKAnfID As Variant               'Anforderungs-ID aus varErfassteVKItem
#Dim blnVKZugeordnet As Boolean          'Flag für zugeordnetes Verifikationskriterium
#Dim varErfassteTFItem As Variant        'Variant für Item in der globalen Testfälle-Liste
#Dim varTFAnfID As Variant               'Anforderungs-ID aus varErfassteTFItem
#Dim blnTFZugeordnet As Boolean          'Flag für zugeorndeten Testfall
#Dim varUmsetzer As Variant              'Variant-Array für Zerlegung der Umsetzer
#Dim intUmsetzer As Integer              'Zählvariable für Umsetzer
#Dim strIStufe As String                 'String für gefundene I-Stufe
#Dim strIStufeMin As String              'String für früheste I-Stufe
#Dim lngLAHBlacklist As Long             'Long für Zähler der Blacklist
#
#'Fehlerbehandlung ausschalten für evtl. fehlende Keys in der Collection FRUTimingList
#On Error Resume Next
#
#'Anforderungen einlesen
#'AVW: #1: ID, #2: Dokument-ID, #3: Basis für Testdesign, #4: Typ, #5: Kategorie, #6: Status, #7: Feature, #8: Reifegrad, #9: Umsetzer, #10: ASIL
#'#11: BSM-SaFuSi Bewertung, #12: BSM-ZZ Bewertung, #13: BSM-ED Bewertung, #14: BSM-FFF Bewertung, #15: BSM-O Bewertung, #16: BSM-Se Bewertung, #17: MV
#'#18: Cluster Testing, #19: Dokument, #20: Kommentar Redaktionskreis, #21: temp1_Text
#For lngZeile = 1 To wksAVW.UsedRange.Rows.Count - rngAVWAttribute(1).Row
#    'Nur Datensätze bei vorhandener Anforderungs-ID übernehmen
#    If rngAVWAttribute(1).Offset(lngZeile, 0).Value <> "" Then
#        'Nur LAH aufnehmen, die nicht auf der Blacklist stehen
#        If AuswertungLAHBlacklist(strLAHBlacklist, rngAVWAttribute(19).Offset(lngZeile, 0).Value) Then
#            'Nur Fachlich abgestimmte (AVW) oder gültige (DOORS) Anforderungen aufnehmen
#    '        If rngAVWAttribute(6).Offset(lngZeile, 0).Value = "Fachlich abgestimmt" Or rngAVWAttribute(6).Offset(lngZeile, 0).Value = "gültig" Then
#                'Attribut Cluster Testing => nicht relevant soll ignoriert werden
#    '            If rngAVWAttribute(18).Offset(lngZeile, 0).Value <> "nicht relevant" Then
#                    'Neuen AVW-Datensatz anlegen
#                    Set BSMDatensatz = New BSMDaten
#                    'Neues Verifikationskriterium anlegen
#                    Set BSMDatensatz.Verifikationskriterium = New Collection
#                    'Neuen Testfall anlegen
#                    Set BSMDatensatz.Testfaelle = New Collection
#                    'Feature einlesen
#                    BSMDatensatz.AVWFeature = CStr(rngAVWAttribute(7).Offset(lngZeile, 0).Value)
#                    'Reifegrad einlesen
#                    BSMDatensatz.AVWReifegrad = CStr(rngAVWAttribute(8).Offset(lngZeile, 0).Value)
#                    'Umsetzer einlesen
#                    BSMDatensatz.AVWUmsetzer = CStr(rngAVWAttribute(9).Offset(lngZeile, 0).Value)
#                    'Dokument-ID einlesen, Entfernung der zusätzlichen Zeichen "?" und "r"
#                    BSMDatensatz.AVWDokumentID = Replace(Replace(CStr(rngAVWAttribute(2).Offset(lngZeile, 0).Value), "?", ""), "r", "")
#                    'Dokument-Name einlesen
#                    BSMDatensatz.AVWDokumentName = CStr(rngAVWAttribute(19).Offset(lngZeile, 0).Value)
#                    'Modulverantwortlichen einlesen
#                    BSMDatensatz.AVWMV = CStr(rngAVWAttribute(17).Offset(lngZeile, 0).Value)
#                    'Anforderungs-ID einlesen, Entfernung der zusätzlichen Zeichen "?" und "r"
#                    BSMDatensatz.AVWID = Replace(Replace(CStr(rngAVWAttribute(1).Offset(lngZeile, 0).Value), "?", ""), "r", "")
#                    'Status einlesen
#                    BSMDatensatz.AVWStatus = CStr(rngAVWAttribute(6).Offset(lngZeile, 0).Value)
#                    'Typ einlesen
#                    BSMDatensatz.AVWTyp = CStr(rngAVWAttribute(4).Offset(lngZeile, 0).Value)
#                    'Kategorie einlesen
#                    BSMDatensatz.AVWKategorie = CStr(rngAVWAttribute(5).Offset(lngZeile, 0).Value)
#                    'BsM-Status einlesen
#                    BSMDatensatz.AVWBsMSaFuSi = CStr(rngAVWAttribute(11).Offset(lngZeile, 0).Value)
#                    BSMDatensatz.AVWBsMZZ = CStr(rngAVWAttribute(12).Offset(lngZeile, 0).Value)
#                    BSMDatensatz.AVWBsMED = CStr(rngAVWAttribute(13).Offset(lngZeile, 0).Value)
#                    BSMDatensatz.AVWBsMFFF = CStr(rngAVWAttribute(14).Offset(lngZeile, 0).Value)
#                    BSMDatensatz.AVWBsMO = CStr(rngAVWAttribute(15).Offset(lngZeile, 0).Value)
#                    BSMDatensatz.AVWBsMSe = CStr(rngAVWAttribute(16).Offset(lngZeile, 0).Value)
#                    'Zusammenführung BsM-Relevanz
#                    strBsMRelevanz = ""
#                    If CStr(BSMDatensatz.AVWBsMSaFuSi) = strBsMVorhanden Then
#                        strBsMRelevanz = "BsM-SaFuSi"
#                    End If
#                    If CStr(BSMDatensatz.AVWBsMZZ) = strBsMVorhanden Then
#                        If strBsMRelevanz = "" Then strBsMRelevanz = "BsM-ZZ" Else strBsMRelevanz = strBsMRelevanz & ",BsM-ZZ"
#                    End If
#                    If CStr(BSMDatensatz.AVWBsMED) = strBsMVorhanden Then
#                        If strBsMRelevanz = "" Then strBsMRelevanz = "BsM-ED" Else strBsMRelevanz = strBsMRelevanz & ",BsM-ED"
#                    End If
#                    If CStr(BSMDatensatz.AVWBsMFFF) = strBsMVorhanden Then
#                        If strBsMRelevanz = "" Then strBsMRelevanz = "BsM-FFF" Else strBsMRelevanz = strBsMRelevanz & ",BsM-FFF"
#                    End If
#                    If CStr(BSMDatensatz.AVWBsMO) = strBsMVorhanden Then
#                        If strBsMRelevanz = "" Then strBsMRelevanz = "BsM-O" Else strBsMRelevanz = strBsMRelevanz & ",BsM-O"
#                    End If
#                    If CStr(BSMDatensatz.AVWBsMSe) = strBsMVorhanden Then
#                        If strBsMRelevanz = "" Then strBsMRelevanz = "BsM-Se" Else strBsMRelevanz = strBsMRelevanz & ",BsM-Se"
#                    End If
#                    BSMDatensatz.BSMRelevanz = strBsMRelevanz
#                    'ASIL einlesen
#                    BSMDatensatz.AVWASIL = CStr(rngAVWAttribute(10).Offset(lngZeile, 0).Value)
#                    'Kommentar Redaktionskreis und temp1_Text einlesen
#                    If InStr(CStr(UCase(rngAVWAttribute(20).Offset(lngZeile, 0).Value)), "#ABGELEHNT_NICHT_TESTBAR") > 0 Or InStr(CStr(UCase(rngAVWAttribute(21).Offset(lngZeile, 0).Value)), "#ABGELEHNT_NICHT_TESTBAR") > 0 Then
#                        BSMDatensatz.AVWAbgelehntNichtTestbar = "x"
#                    End If
#                    
#                    'Geplante I-Stufe einlesen
#                    strIStufe = ""
#                    strIStufeMin = ""
#                    If BSMDatensatz.AVWUmsetzer <> "" Then
#                        varUmsetzer = Split(BSMDatensatz.AVWUmsetzer, ",", , vbBinaryCompare)
#                        For intUmsetzer = 0 To UBound(varUmsetzer, 1)
#                            strIStufe = FRUTimingList.Item(BSMDatensatz.AVWFeature & BSMDatensatz.AVWReifegrad & Trim(varUmsetzer(intUmsetzer))).IStufe
#                            If strIStufeMin = "" Then
#                                strIStufeMin = strIStufe
#                            ElseIf InStr(strIStufe, "IS") > 0 Then
#                                If strIStufe < strIStufeMin Then
#                                    strIStufeMin = strIStufe
#                                End If
#                            End If
#                        Next intUmsetzer
#                    End If
#                    BSMDatensatz.IStufe = strIStufeMin
#                    
#                    'Cluster Testing einlesen
#                    BSMDatensatz.ClusterTesting = CStr(rngAVWAttribute(18).Offset(lngZeile, 0).Value)
#                    
#                    'Anforderungsverantwortliche einlesen
#                    BSMDatensatz.AVWAnforderungsverantwortliche = CStr(rngAVWAttribute(22).Offset(lngZeile, 0).Value)
#                    
#                    'Projekt MEB21 - Temp11_Auswahlfeld einlesen
#                    If strProjekt = "MEB21" Or strProjekt = "MQB48W" Then
#                        BSMDatensatz.AVWTemp11_Auswahlfeld = CStr(rngAVWAttributeMEB21(1).Offset(lngZeile, 0).Value)
#                    End If
#                    
#                    'Verifikationskriterium zuordnen
#                    'Zuordnung zu Verifikationskriterium in globaler Verifikationskriterien-Liste
#                    blnVKZugeordnet = False
#                    For Each varErfassteVKItem In verifikationKritList
#                        'Abgleich über Element-ID
#                        For Each varVKAnfID In varErfassteVKItem.anf_ids
#                            If varVKAnfID = BSMDatensatz.AVWID Then
#                                'Zugehörigkeit des Verifikationskriteriums zu aktuellen Anforderungen kennzeichnen
#                                varErfassteVKItem.AnforderungVorhanden = True
#                                'VK-ID aufnehmen
#                                BSMDatensatz.Verifikationskriterium.Add Item:=varErfassteVKItem
#                                'Geplante I-Stufe für Verifikationskritierum erfassen
#                                If BSMDatensatz.IStufe <> "" Then
#                                    varErfassteVKItem.anf_IStufen.Add Item:=BSMDatensatz.IStufe
#                                End If
#                                'Umsetzer für Verifikationskritierum erfassen
#                                If BSMDatensatz.AVWUmsetzer <> "" Then
#                                    varErfassteVKItem.anf_Umsetzer.Add Item:=BSMDatensatz.AVWUmsetzer
#                                End If
#                                'BsM-Relevanz für Verifikationskritierum erfassen
#                                If BSMDatensatz.BSMRelevanz <> "" Then
#                                    varErfassteVKItem.anf_BsMRelevanz.Add Item:=BSMDatensatz.BSMRelevanz
#                                End If
#                                'ASIL für Verifikationskritierum erfassen
#                                If BSMDatensatz.AVWASIL <> "" Then
#                                    varErfassteVKItem.anf_ASIL.Add Item:=BSMDatensatz.AVWASIL
#                                End If
#                                'Feature für Verifikationskritierum erfassen
#                                If BSMDatensatz.AVWFeature <> "" Then
#                                    varErfassteVKItem.anf_Feature.Add Item:=BSMDatensatz.AVWFeature
#                                End If
#                                'Reifegrad für Verifikationskritierum erfassen
#                                If BSMDatensatz.AVWReifegrad <> "" Then
#                                    varErfassteVKItem.anf_Reifegrad.Add Item:=BSMDatensatz.AVWReifegrad
#                                End If
#                                'Modulverantwortliche für Verifikationskritierum erfassen
#                                If BSMDatensatz.AVWMV <> "" Then
#                                    varErfassteVKItem.anf_MV.Add Item:=BSMDatensatz.AVWMV
#                                End If
#                                'LAH-ID für Verifikationskritierum erfassen
#                                If BSMDatensatz.AVWDokumentID <> "" Then
#                                    varErfassteVKItem.anf_LAHID.Add Item:=BSMDatensatz.AVWDokumentID
#                                End If
#                                'LAH-Namen für Verifikationskritierum erfassen
#                                If BSMDatensatz.AVWDokumentName <> "" Then
#                                    varErfassteVKItem.addLAHName (BSMDatensatz.AVWDokumentName)
#                                End If
#                                'Cluster Testing für Verifikationskriterium erfassen
#                                If BSMDatensatz.ClusterTesting <> "" Then
#                                    varErfassteVKItem.anf_ClusterTesting.Add Item:=BSMDatensatz.ClusterTesting
#                                End If
#                                'Anforderungsverantwortliche für Verifikationskriterium erfassen
#                                If BSMDatensatz.AVWAnforderungsverantwortliche <> "" Then
#                                    varErfassteVKItem.anf_Anforderungsverantwortliche.Add Item:=BSMDatensatz.AVWAnforderungsverantwortliche
#                                End If
#                                'Temp11_Auswahlfeld für Verifikationskriterium erfassen
#                                If BSMDatensatz.AVWTemp11_Auswahlfeld <> "" Then
#                                    varErfassteVKItem.anf_Temp11_Auswahlfeld.Add Item:=BSMDatensatz.AVWTemp11_Auswahlfeld
#                                End If
#                                
#                                'Innere Schleife beenden, da es zu jeder Anforderung nur ein Verifikationskriterium gibt
#                                blnVKZugeordnet = True
#                                Exit For
#                            End If
#                        Next varVKAnfID
#                        'Äußere Schleife beenden, da es zu jeder Anforderung nur ein Verifikationskriterium gibt
#                        If blnVKZugeordnet Then Exit For
#                    Next varErfassteVKItem
#                    
#                    'Testfall zuordnen
#                    blnTFZugeordnet = False
#                    For Each varErfassteTFItem In testfallList
#                        'Abgleich über Element-ID
#                        For Each varTFAnfID In varErfassteTFItem.TF_anfIDs
#                            If varTFAnfID = BSMDatensatz.AVWID Then
#                                BSMDatensatz.Testfaelle.Add Item:=varErfassteTFItem
#                                blnTFZugeordnet = True
#                                Exit For
#                            End If
#                        Next varTFAnfID
#                    Next varErfassteTFItem
#                    
#                    'Erfasste AVW-Rohdaten in globaler AVW-Rohdaten-Liste hinzufügen
#                    BsMDatenList.Add Item:=BSMDatensatz, Key:=BSMDatensatz.AVWID
#    '            End If
#    '        End If
#        End If
#    End If
#
#    'Fortschritt anzeigen
#    If lngZeile Mod 100 = 0 Then
#        Debug.Print "Anforderungen einlesen: " & lngZeile & "/" & wksAVW.UsedRange.Rows.Count - rngAVWAttribute(1).Row
#    End If
#Next lngZeile
#End Sub
#
#Private Sub EinlesenAVWNachfolgerRohdaten(ByVal wksAVW As Worksheet, ByRef strAVWAttribute() As String, ByRef rngAVWAttribute() As Range, ByRef strLAHBlacklist() As String, _
#                                          ByVal strProjekt As String, ByRef strAVWAttributMEB21() As String, ByRef rngAVWAttributMEB21() As Range)
#Dim lngZeile As Long                    'Long-Variable für aktuell einzulesende Zeile
#Dim BSMDatensatz As BSMDaten            'Klasse BSMDaten
#Dim strBsMRelevanz As String            'String-Variable für Zusammenfassung der BsM-Relevanz
#Const strBsMVorhanden As String = "ja"  'Konstante für Angabe aus AVW für vorhandenes BsM-Attribut
#Dim varErfassteVKItem As Variant        'Variant für Item in der globalen Verifikationskriterien-Liste
#Dim varVKAnfID As Variant               'Anforderungs-ID aus varErfassteVKItem
#Dim blnVKZugeordnet As Boolean          'Flag für zugeordnetes Verifikationskriterium
#Dim varErfassteTFItem As Variant        'Variant für Item in der globalen Testfälle-Liste
#Dim varTFAnfID As Variant               'Anforderungs-ID aus varErfassteTFItem
#Dim blnTFZugeordnet As Boolean          'Flag für zugeorndeten Testfall
#Dim varUmsetzer As Variant              'Variant-Array für Zerlegung der Umsetzer
#Dim intUmsetzer As Integer              'Zählvariable für Umsetzer
#Dim strIStufe As String                 'String für gefundene I-Stufe
#Dim strIStufeMin As String              'String für früheste I-Stufe
#Dim lngLAHBlacklist As Long             'Long für Zähler der Blacklist
#Dim AVWVorgaenger As AVWVorgaenger      'Klasse AVWVorgaenger
#
#'Fehlerbehandlung ausschalten für evtl. fehlende Keys in der Collection FRUTimingList
#On Error Resume Next
#
#'Anforderungen einlesen
#'AVW: #1: ID, #2: Dokument-ID, #3: Basis für Testdesign, #4: Typ, #5: Kategorie, #6: Status, #7: Feature, #8: Reifegrad, #9: Umsetzer, #10: ASIL
#'#11: BSM-SaFuSi Bewertung, #12: BSM-ZZ Bewertung, #13: BSM-ED Bewertung, #14: BSM-FFF Bewertung, #15: BSM-O Bewertung, #16: BSM-Se Bewertung, #17: MV
#'#18: Cluster Testing, #19: Dokument, #20: Kommentar Redaktionskreis, #21: temp1_Text, #22: ID der Vorgänger-Anforderung
#For lngZeile = 1 To wksAVW.UsedRange.Rows.Count - rngAVWAttribute(1).Row
#    'Nur Datensätze bei vorhandener Anforderungs-ID übernehmen
#    If rngAVWAttribute(1).Offset(lngZeile, 0).Value <> "" Then
#        'Nur LAH aufnehmen, die nicht auf der Blacklist stehen
#        If AuswertungLAHBlacklist(strLAHBlacklist, rngAVWAttribute(19).Offset(lngZeile, 0).Value) Then
#            'Nur Fachlich abgestimmte (AVW) oder gültige (DOORS) Anforderungen aufnehmen
#    '        If rngAVWAttribute(6).Offset(lngZeile, 0).Value = "Fachlich abgestimmt" Or rngAVWAttribute(6).Offset(lngZeile, 0).Value = "gültig" Then
#                'Attribut Cluster Testing => nicht relevant soll ignoriert werden
#    '            If rngAVWAttribute(18).Offset(lngZeile, 0).Value <> "nicht relevant" Then
#                    'Neuen AVW-Datensatz anlegen
#                    Set BSMDatensatz = New BSMDaten
#                    'Neues Verifikationskriterium anlegen
#                    Set BSMDatensatz.Verifikationskriterium = New Collection
#                    'Neuen Testfall anlegen
#                    Set BSMDatensatz.Testfaelle = New Collection
#                    'Feature einlesen
#                    BSMDatensatz.AVWFeature = CStr(rngAVWAttribute(7).Offset(lngZeile, 0).Value)
#                    'Reifegrad einlesen
#                    BSMDatensatz.AVWReifegrad = CStr(rngAVWAttribute(8).Offset(lngZeile, 0).Value)
#                    'Umsetzer einlesen
#                    BSMDatensatz.AVWUmsetzer = CStr(rngAVWAttribute(9).Offset(lngZeile, 0).Value)
#                    'Dokument-ID einlesen, Entfernung der zusätzlichen Zeichen "?" und "r"
#                    BSMDatensatz.AVWDokumentID = Replace(Replace(CStr(rngAVWAttribute(2).Offset(lngZeile, 0).Value), "?", ""), "r", "")
#                    'Dokument-Name einlesen
#                    BSMDatensatz.AVWDokumentName = CStr(rngAVWAttribute(19).Offset(lngZeile, 0).Value)
#                    'Modulverantwortlichen einlesen
#                    BSMDatensatz.AVWMV = CStr(rngAVWAttribute(17).Offset(lngZeile, 0).Value)
#                    'Anforderungs-ID einlesen, Entfernung der zusätzlichen Zeichen "?" und "r"
#                    BSMDatensatz.AVWID = Replace(Replace(CStr(rngAVWAttribute(1).Offset(lngZeile, 0).Value), "?", ""), "r", "")
#                    'Status einlesen
#                    BSMDatensatz.AVWStatus = CStr(rngAVWAttribute(6).Offset(lngZeile, 0).Value)
#                    'Typ einlesen
#                    BSMDatensatz.AVWTyp = CStr(rngAVWAttribute(4).Offset(lngZeile, 0).Value)
#                    'Kategorie einlesen
#                    BSMDatensatz.AVWKategorie = CStr(rngAVWAttribute(5).Offset(lngZeile, 0).Value)
#                    'BsM-Status einlesen
#                    BSMDatensatz.AVWBsMSaFuSi = CStr(rngAVWAttribute(11).Offset(lngZeile, 0).Value)
#                    BSMDatensatz.AVWBsMZZ = CStr(rngAVWAttribute(12).Offset(lngZeile, 0).Value)
#                    BSMDatensatz.AVWBsMED = CStr(rngAVWAttribute(13).Offset(lngZeile, 0).Value)
#                    BSMDatensatz.AVWBsMFFF = CStr(rngAVWAttribute(14).Offset(lngZeile, 0).Value)
#                    BSMDatensatz.AVWBsMO = CStr(rngAVWAttribute(15).Offset(lngZeile, 0).Value)
#                    BSMDatensatz.AVWBsMSe = CStr(rngAVWAttribute(16).Offset(lngZeile, 0).Value)
#                    'Zusammenführung BsM-Relevanz
#                    strBsMRelevanz = ""
#                    If CStr(BSMDatensatz.AVWBsMSaFuSi) = strBsMVorhanden Then
#                        strBsMRelevanz = "BsM-SaFuSi"
#                    End If
#                    If CStr(BSMDatensatz.AVWBsMZZ) = strBsMVorhanden Then
#                        If strBsMRelevanz = "" Then strBsMRelevanz = "BsM-ZZ" Else strBsMRelevanz = strBsMRelevanz & ",BsM-ZZ"
#                    End If
#                    If CStr(BSMDatensatz.AVWBsMED) = strBsMVorhanden Then
#                        If strBsMRelevanz = "" Then strBsMRelevanz = "BsM-ED" Else strBsMRelevanz = strBsMRelevanz & ",BsM-ED"
#                    End If
#                    If CStr(BSMDatensatz.AVWBsMFFF) = strBsMVorhanden Then
#                        If strBsMRelevanz = "" Then strBsMRelevanz = "BsM-FFF" Else strBsMRelevanz = strBsMRelevanz & ",BsM-FFF"
#                    End If
#                    If CStr(BSMDatensatz.AVWBsMO) = strBsMVorhanden Then
#                        If strBsMRelevanz = "" Then strBsMRelevanz = "BsM-O" Else strBsMRelevanz = strBsMRelevanz & ",BsM-O"
#                    End If
#                    If CStr(BSMDatensatz.AVWBsMSe) = strBsMVorhanden Then
#                        If strBsMRelevanz = "" Then strBsMRelevanz = "BsM-Se" Else strBsMRelevanz = strBsMRelevanz & ",BsM-Se"
#                    End If
#                    BSMDatensatz.BSMRelevanz = strBsMRelevanz
#                    
#                    'ASIL einlesen
#                    BSMDatensatz.AVWASIL = CStr(rngAVWAttribute(10).Offset(lngZeile, 0).Value)
#                    
#                    'Vorgänger-ID einlesen
#                    If blnAVWVorgaengerIDsVerwenden Then
#                        BSMDatensatz.AVWVorgaengerID = CStr(rngAVWAttribute(23).Offset(lngZeile, 0).Value)  'BSMDatensatz.AVWVorgaengerID = CStr(rngAVWAttribute(22).Offset(lngZeile, 0).Value)
#                    End If
#                    
#                    'Kommentar Redaktionskreis und temp1_Text aus AVW-Rohdaten einlesen
#                    If InStr(CStr(UCase(rngAVWAttribute(20).Offset(lngZeile, 0).Value)), "#ABGELEHNT_NICHT_TESTBAR") > 0 Or InStr(CStr(UCase(rngAVWAttribute(21).Offset(lngZeile, 0).Value)), "#ABGELEHNT_NICHT_TESTBAR") > 0 Then
#                        BSMDatensatz.AVWAbgelehntNichtTestbar = "x"
#                    End If
#                    'Kommentar Redaktionskreis und temp1_Text aus AVW-Vorgänger einlesen
#                    Set AVWVorgaenger = New AVWVorgaenger
#                    Set AVWVorgaenger = FindeAVWVorgaenger(AVWVorgaengerList, BSMDatensatz.AVWVorgaengerID)
#                    If Not AVWVorgaenger Is Nothing Then
#                        If AVWVorgaenger.AbgelehntNichtTestbar = "x" Then
#                            If BSMDatensatz.AVWAbgelehntNichtTestbar = "" Then
#                                BSMDatensatz.AVWAbgelehntNichtTestbar = "x (Master)"
#                            Else
#                                BSMDatensatz.AVWAbgelehntNichtTestbar = BSMDatensatz.AVWAbgelehntNichtTestbar & vbCrLf & "x (Master)"
#                            End If
#                        End If
#                    End If
#                    
#                    'Geplante I-Stufe einlesen
#                    strIStufe = ""
#                    strIStufeMin = ""
#                    If BSMDatensatz.AVWUmsetzer <> "" Then
#                        varUmsetzer = Split(BSMDatensatz.AVWUmsetzer, ",", , vbBinaryCompare)
#                        For intUmsetzer = 0 To UBound(varUmsetzer, 1)
#                            strIStufe = FRUTimingList.Item(BSMDatensatz.AVWFeature & BSMDatensatz.AVWReifegrad & Trim(varUmsetzer(intUmsetzer))).IStufe
#                            If strIStufeMin = "" Then
#                                strIStufeMin = strIStufe
#                            ElseIf InStr(strIStufe, "IS") > 0 Then
#                                If strIStufe < strIStufeMin Then
#                                    strIStufeMin = strIStufe
#                                End If
#                            End If
#                        Next intUmsetzer
#                    End If
#                    BSMDatensatz.IStufe = strIStufeMin
#                    
#                    'Cluster Testing einlesen
#                    BSMDatensatz.ClusterTesting = CStr(rngAVWAttribute(18).Offset(lngZeile, 0).Value)
#
#                    'Anforderungsverantwortliche einlesen
#                    BSMDatensatz.AVWAnforderungsverantwortliche = CStr(rngAVWAttribute(22).Offset(lngZeile, 0).Value)
#                    
#                    'Projekt MEB21 - Temp11_Auswahlfeld einlesen
#                    If strProjekt = "MEB21" Or strProjekt = "MQB48W" Then
#                        BSMDatensatz.AVWTemp11_Auswahlfeld = CStr(rngAVWAttributeMEB21(1).Offset(lngZeile, 0).Value)
#                    End If
#                    
#                    'Verifikationskriterium zuordnen
#                    'Zuordnung zu Verifikationskriterium in globaler Verifikationskriterien-Liste
#                    blnVKZugeordnet = False
#                    For Each varErfassteVKItem In verifikationKritList
#                        'Abgleich über Element-ID des Vorgängers
#                        For Each varVKAnfID In varErfassteVKItem.anf_ids
#                            If varVKAnfID = BSMDatensatz.AVWVorgaengerID Then
#                                'Zugehörigkeit des Verifikationskriteriums zu aktuellen Anforderungen kennzeichnen
#                                varErfassteVKItem.AnforderungVorhanden = True
#                                'VK-ID aufnehmen
#                                BSMDatensatz.Verifikationskriterium.Add Item:=varErfassteVKItem
#                                'Geplante I-Stufe für Verifikationskritierum erfassen
#                                If BSMDatensatz.IStufe <> "" Then
#                                    varErfassteVKItem.anf_IStufen.Add Item:=BSMDatensatz.IStufe
#                                End If
#                                'Umsetzer für Verifikationskritierum erfassen
#                                If BSMDatensatz.AVWUmsetzer <> "" Then
#                                    varErfassteVKItem.anf_Umsetzer.Add Item:=BSMDatensatz.AVWUmsetzer
#                                End If
#                                'BsM-Relevanz für Verifikationskritierum erfassen
#                                If BSMDatensatz.BSMRelevanz <> "" Then
#                                    varErfassteVKItem.anf_BsMRelevanz.Add Item:=BSMDatensatz.BSMRelevanz
#                                End If
#                                'ASIL für Verifikationskritierum erfassen
#                                If BSMDatensatz.AVWASIL <> "" Then
#                                    varErfassteVKItem.anf_ASIL.Add Item:=BSMDatensatz.AVWASIL
#                                End If
#                                'Feature für Verifikationskritierum erfassen
#                                If BSMDatensatz.AVWFeature <> "" Then
#                                    varErfassteVKItem.anf_Feature.Add Item:=BSMDatensatz.AVWFeature
#                                End If
#                                'Reifegrad für Verifikationskritierum erfassen
#                                If BSMDatensatz.AVWReifegrad <> "" Then
#                                    varErfassteVKItem.anf_Reifegrad.Add Item:=BSMDatensatz.AVWReifegrad
#                                End If
#                                'Modulverantwortliche für Verifikationskritierum erfassen
#                                If BSMDatensatz.AVWMV <> "" Then
#                                    varErfassteVKItem.anf_MV.Add Item:=BSMDatensatz.AVWMV
#                                End If
#                                'LAH-ID für Verifikationskritierum erfassen
#                                If BSMDatensatz.AVWDokumentID <> "" Then
#                                    varErfassteVKItem.anf_LAHID.Add Item:=BSMDatensatz.AVWDokumentID
#                                End If
#                                'LAH-Namen für Verifikationskritierum erfassen
#                                If BSMDatensatz.AVWDokumentName <> "" Then
#                                    varErfassteVKItem.addLAHName (BSMDatensatz.AVWDokumentName)
#                                End If
#                                'Cluster Testing für Verifikationskriterium erfassen
#                                If BSMDatensatz.ClusterTesting <> "" Then
#                                    varErfassteVKItem.anf_ClusterTesting.Add Item:=BSMDatensatz.ClusterTesting
#                                End If
#                                'Anforderungsverantwortliche für Verifikationskriterium erfassen
#                                If BSMDatensatz.AVWAnforderungsverantwortliche <> "" Then
#                                    varErfassteVKItem.anf_Anforderungsverantwortliche.Add Item:=BSMDatensatz.AVWAnforderungsverantwortliche
#                                End If
#                                'Temp11_Auswahlfeld für Verifikationskriterium erfassen
#                                If BSMDatensatz.AVWTemp11_Auswahlfeld <> "" Then
#                                    varErfassteVKItem.anf_Temp11_Auswahlfeld.Add Item:=BSMDatensatz.AVWTemp11_Auswahlfeld
#                                End If
#                                
#                                'Innere Schleife beenden, da es zu jeder Anforderung nur ein Verifikationskriterium gibt
#                                blnVKZugeordnet = True
#                                Exit For
#                            End If
#                        Next varVKAnfID
#                        'Äußere Schleife beenden, da es zu jeder Anforderung nur ein Verifikationskriterium gibt
#                        If blnVKZugeordnet Then Exit For
#                    Next varErfassteVKItem
#                    
#                    'Testfall zuordnen
#                    blnTFZugeordnet = False
#                    For Each varErfassteTFItem In testfallList
#                        'Abgleich über Element-ID
#                        For Each varTFAnfID In varErfassteTFItem.TF_anfIDs
#                            If varTFAnfID = BSMDatensatz.AVWVorgaengerID Then
#                                BSMDatensatz.Testfaelle.Add Item:=varErfassteTFItem
#                                blnTFZugeordnet = True
#                                Exit For
#                            End If
#                        Next varTFAnfID
#                    Next varErfassteTFItem
#                    
#                    'Erfasste AVW-Rohdaten in globaler AVW-Rohdaten-Liste hinzufügen
#                    BsMDatenList.Add Item:=BSMDatensatz, Key:=BSMDatensatz.AVWID
#    '            End If
#    '        End If
#        End If
#    End If
#
#    'Fortschritt anzeigen
#    If lngZeile Mod 100 = 0 Then
#        Debug.Print "Anforderungen einlesen: " & lngZeile & "/" & wksAVW.UsedRange.Rows.Count - rngAVWAttribute(1).Row
#    End If
#Next lngZeile
#End Sub
#
#Private Sub EinlesenAVWVorgaengerRohdaten(ByVal wksAVWMaster As Worksheet, ByRef strAVWMasterAttribute() As String, ByRef rngAVWMasterAttribute() As Range)
#Dim lngZeile As Long                    'Long-Variable für aktuell einzulesende Zeile
#Dim AVWVorgaenger As AVWVorgaenger      'Klasse AVWVorgaenger
#
#'Anforderungen einlesen
#'AVW: #1: ID, #2: temp1_Text, #3: Kommentar Redaktionskreis
#For lngZeile = 1 To wksAVWMaster.UsedRange.Rows.Count - rngAVWMasterAttribute(1).Row
#    'Nur Datensätze bei vorhandener Anforderungs-ID übernehmen
#    If rngAVWMasterAttribute(1).Offset(lngZeile, 0).Value <> "" Then
#        'Neuen AVW-Datensatz anlegen
#        Set AVWVorgaenger = New AVWVorgaenger
#        'Master-ID einlesen
#        AVWVorgaenger.ID = CStr(rngAVWMasterAttribute(1).Offset(lngZeile, 0).Value)
#        'Kommentar Redaktionskreis und temp1_Text einlesen
#        If InStr(CStr(UCase(rngAVWMasterAttribute(2).Offset(lngZeile, 0).Value)), "#ABGELEHNT_NICHT_TESTBAR") > 0 Or InStr(CStr(UCase(rngAVWMasterAttribute(3).Offset(lngZeile, 0).Value)), "#ABGELEHNT_NICHT_TESTBAR") > 0 Then
#            AVWVorgaenger.AbgelehntNichtTestbar = "x"
#        End If
#        
#        'Erfasste AVW-Vorgänger in globaler AVW-Vorgaenger-Liste hinzufügen
#        AVWVorgaengerList.Add Item:=AVWVorgaenger, Key:=AVWVorgaenger.ID
#    End If
#
#    'Fortschritt anzeigen
#    If lngZeile Mod 100 = 0 Then
#        Debug.Print "Anforderungen aus Masterbereich einlesen: " & lngZeile & "/" & wksAVWMaster.UsedRange.Rows.Count - rngAVWMasterAttribute(1).Row
#    End If
#Next lngZeile
#End Sub
#
#Function EinlesenGetrennteWerteKomma(ByVal lahIDs As String) As Collection
#Dim idCollection As Collection
#Dim subStrings() As String
#Set idCollection = New Collection
#Dim newString As String
#Dim x As Integer
#
#newString = Replace(lahIDs, "?", "")
#subStrings = Split(newString, ",")
#For x = LBound(subStrings) To UBound(subStrings)
#    idCollection.Add Trim(subStrings(x))
#Next
#Set EinlesenGetrennteWerteKomma = idCollection
#End Function
#
#Private Function AusgabeSammlungLF(ByRef list As Collection) As String
#Dim strTemp As String
#Dim i As Integer
#
#strTemp = ""
#If list.Count > 0 Then
#    For i = 1 To list.Count
#        If (strTemp = "") Then
#            strTemp = list(i)
#        Else
#            strTemp = strTemp & vbCrLf & list(i)
#        End If
#    Next
#End If
#AusgabeSammlungLF = strTemp
#End Function
#
#Private Function AusgabeSammlungLFEinfach(ByRef list As Collection) As String
#Dim strTemp As String
#Dim i As Integer
#
#strTemp = ""
#If list.Count > 0 Then
#    For i = 1 To list.Count
#        If (strTemp = "") Then
#            strTemp = list(i)
#        Else
#            If InStr(strTemp, list(i)) = 0 Then
#                strTemp = strTemp & vbCrLf & list(i)
#            End If
#        End If
#    Next
#End If
#AusgabeSammlungLFEinfach = strTemp
#End Function
#
#Private Function AusgabeSammlungKomma(ByRef list As Collection) As String
#Dim strTemp As String
#Dim i As Integer
#
#strTemp = ""
#If list.Count > 0 Then
#    For i = 1 To list.Count
#        If (strTemp = "") Then
#            strTemp = list(i)
#        Else
#            strTemp = strTemp & ", " & list(i)
#        End If
#    Next
#End If
#AusgabeSammlungKomma = strTemp
#End Function
#
#Private Function AuswertungLAHBlacklist(ByRef strLAHBlacklist() As String, ByVal strLAH As String)
#Dim lngBlacklistZaehler As Long     'Long-Zähler für Blacklist
#
#AuswertungLAHBlacklist = True
#If UBound(strLAHBlacklist, 1) > 0 Then
#    For lngBlacklistZaehler = LBound(strLAHBlacklist, 1) To UBound(strLAHBlacklist, 1)
#        If strLAHBlacklist(lngBlacklistZaehler) = strLAH Then
#            AuswertungLAHBlacklist = False
#            Exit For
#        End If
#    Next lngBlacklistZaehler
#End If
#End Function
#
#Private Function FindeVK(ByRef Liste As Collection, ByVal strKey As String) As Verifikationskriterium
#Dim ListenObjekt As Verifikationskriterium
#    
#On Error GoTo err
#Set ListenObjekt = Liste.Item(strKey)
#Set FindeVK = ListenObjekt
#Exit Function
#
#err:
#    Set ListenObjekt = Nothing
#    Set FindeVK = ListenObjekt
#End Function
#
#Private Function FindeAVWVorgaenger(ByRef Liste As Collection, ByVal strKey As String) As AVWVorgaenger
#Dim ListenObjekt As AVWVorgaenger
#    
#On Error GoTo err
#Set ListenObjekt = Liste.Item(strKey)
#Set FindeAVWVorgaenger = ListenObjekt
#Exit Function
#
#err:
#    Set ListenObjekt = Nothing
#    Set FindeAVWVorgaenger = ListenObjekt
#End Function
#
#Private Sub AusgabeATEStatus(ByVal wbBsM As Workbook, ByRef wksBsM As Worksheet, ByRef strBsMAttribute() As String, ByRef rngBsMAttribute() As Range, ByRef strWeitereTUsAusgabe As String, ByRef strDateinamen() As String, ByVal strProjekt As String)
#Dim lngDatensatz As Long                        'Long-Variable für aktuell zu schreibenden Datensatz
#Dim varErfassteBsMDatensatzItem As Variant      'Variant für Item im globalen BsM-Datensatz
#Dim varErfassteTDAAItem As Variant              'Variant für Item aus den jeweiligen Absicherungsaufträgen
#Dim strTDAA As String                           'String für Sammlung der Absicherungsaufträge
#Dim strTDTiTu As String                         'String für Sammlung der Ti:Tu-Kombinationen
#Dim varErfassteTFItem As Variant                'Variant für Item aus den jeweiligen Testfällen
#Dim strTestfaelle As String                     'String für Sammlung der Testfälle
#Dim strAbgleichTUs() As String                  'String-Array für die Namen der abzugleichenden Testumgebungen bei TDs und TFs
#Dim intAbgleichTUs() As Integer                 'Integer-Array für Erfassung der Testumgebungen bei TDs und TFs
#Dim i As Long                                   'Laufvariable
#Dim intAuswertungTUs() As Integer               'Integer-Array für die Ergebnisse des Tu-Abgleichs
#Dim strAuswertungTUs() As String                'String-Array für die Ergebnisse des Tu-Abgleichs
#Dim strAusgabeAuswertungTUs As String           'String für Ausgabe des Tu-Abgleichs
#Dim intAusgabeAuswertungTUs As Integer          'Integer für Ausgabe des Tu-Abgleichs
#Dim strAuswertungTUsFehlendeAAs As String       'String für Ausgabe der fehlenden TUs bei TD-AAs
#Dim strAuswertungTUsFehlendeTFs As String       'String für Ausgabe der fehlenden TUs bei TFs
#Dim strAusgabeAuswertungTUsDetails As String    'String für Ausgabe des Tu-Abgleichs mit Details
#Dim intWeitereTUs As Integer                    'Zählvariable für weitere TUs
#Dim strWeitereTUs() As String                   'String-Array für weitere TUs
#Dim strBekannteTUs() As String                  'String-Array für alle bekannten TUs
#Dim blnTFTUZugeordnet As Boolean                'Flag für zugeordnete TU
#Dim blnAATUZugeordnet As Boolean                'Flag für zugeordnete TU
#Dim intWeitereTUsZaehler As Integer             'Laufvariable für weitere TUs
#Dim blnTIUnerlaubt As Boolean                   'Flag für unerlaubte Testinstanzen (Erlaubt: "eigene Organisationseinheit", "Dauerlauf Gesamtfahrzeug", "Gesamtverbundintegration" sowie alle eingetragenen Umsetzer)
#Dim intUmsetzer As Integer                      'Zähler für Umsetzer aus AVW
#Dim varUmsetzer As Variant                      'Variant-Array für Umsetzer aus AVW
#Dim blnUmsetzer() As Boolean                    'Flag-Array für Abgleich der Umsetzer AVW<->AA
#Dim intZielspalte As Integer                    'Integer für Ausgabespalte
#Dim intRelevantekTUs As Integer                 'Integer für Anzahl der relevanten TUs
#Dim dblTDVKAnzahlUseCases As Double             'Anzahl der Vorkommen der Use-Case-Begriffe
#Dim strTDVKAktion As String                     'String zur Bearbeitung der TDVK-Aktion
#
#'Tabelle erzeugen
#'Neues Worksheet erzeugen
#Set wksBsM = wbBsM.Sheets.Add(after:=wbBsM.Worksheets(wbBsM.Worksheets.Count))
#wksBsM.Name = "ATE_Status_" & "Today" & "_" & Replace(Time, ":", "")
#'Arbeitsblatt BsM_Status
#'BsM-Status: #1: ID, #2: Dokument-ID, #3: BsM-Relevanz, #4: BSM-SaFuSi Bewertung, #5: BSM-ZZ Bewertung, #6: BSM-ED Bewertung, #7: BSM-FFF Bewertung, #8: BSM-O Bewertung,
#'#9: BSM-Se Bewertung, #10: ASIL, #11: Feature, #12: Reifegrad, #13: Umsetzer, #14: Status, #15: TD-VK, #16: TD-AA, #17: TD-TI:TU, #18: Testfälle, #19: Vergleich TUs,
#'#20: MV, #21: Kategorie, #22: Dokument, #23: #abgelehnt_nicht_testbar, #24: Zugeordnete I-Stufe, #25: Status TD-VK, #26: Fehlende TUs bei TD-AAs, #27: Fehlende TUs bei TFs,
#'#28: Erläuterungen zum Vergleich, #29: Cluster Testing, #30: Projekt, #31: TD-VK temp1_Text, #32: TD-VK Effort Estimation, #33: Anforderungsverantwortliche,#34: KW
#'Projektspezifisch - MEB21
#'#35: Temp11_Auswahlfeld
#
#lngDatensatz = 1
#If blnAVWVorgaengerIDsVerwenden = False Then
#    intZielspalte = 0
#    ReDim strBsMAttribute(1 To 34)
#Else
#    intZielspalte = 1
#    ReDim strBsMAttribute(0 To 34)
#End If
#ReDim rngBsMAttribute(LBound(strBsMAttribute, 1) To UBound(strBsMAttribute, 1))
#'Name und Position der Tabellenattribute
#If blnAVWVorgaengerIDsVerwenden Then
#    strBsMAttribute(0) = "Abgezweigt aus"
#    Set rngBsMAttribute(0) = wksBsM.Cells(lngDatensatz, intZielspalte)
#End If
#strBsMAttribute(34) = "KW Datenauswertung"
#Set rngBsMAttribute(34) = wksBsM.Cells(lngDatensatz, intZielspalte + 1)
#strBsMAttribute(1) = "ID"
#Set rngBsMAttribute(1) = wksBsM.Cells(lngDatensatz, intZielspalte + 2)
#strBsMAttribute(2) = "Dokument-ID"
#Set rngBsMAttribute(2) = wksBsM.Cells(lngDatensatz, intZielspalte + 3)
#strBsMAttribute(22) = "Dokument"
#Set rngBsMAttribute(22) = wksBsM.Cells(lngDatensatz, intZielspalte + 4)
#strBsMAttribute(21) = "Kategorie"
#Set rngBsMAttribute(21) = wksBsM.Cells(lngDatensatz, intZielspalte + 5)
#strBsMAttribute(11) = "Feature"
#Set rngBsMAttribute(11) = wksBsM.Cells(lngDatensatz, intZielspalte + 6)
#strBsMAttribute(12) = "Reifegrad"
#Set rngBsMAttribute(12) = wksBsM.Cells(lngDatensatz, intZielspalte + 7)
#strBsMAttribute(13) = "Umsetzer"
#Set rngBsMAttribute(13) = wksBsM.Cells(lngDatensatz, intZielspalte + 8)
#strBsMAttribute(3) = "BsM-Relevanz"
#Set rngBsMAttribute(3) = wksBsM.Cells(lngDatensatz, intZielspalte + 9)
#strBsMAttribute(4) = "BSM-SaFuSi Bewertung"
#Set rngBsMAttribute(4) = wksBsM.Cells(lngDatensatz, intZielspalte + 10)
#strBsMAttribute(5) = "BSM-ZZ Bewertung"
#Set rngBsMAttribute(5) = wksBsM.Cells(lngDatensatz, intZielspalte + 11)
#strBsMAttribute(6) = "BSM-ED Bewertung"
#Set rngBsMAttribute(6) = wksBsM.Cells(lngDatensatz, intZielspalte + 12)
#strBsMAttribute(7) = "BSM-FFF Bewertung"
#Set rngBsMAttribute(7) = wksBsM.Cells(lngDatensatz, intZielspalte + 13)
#strBsMAttribute(8) = "BSM-O Bewertung"
#Set rngBsMAttribute(8) = wksBsM.Cells(lngDatensatz, intZielspalte + 14)
#strBsMAttribute(9) = "BSM-Se Bewertung"
#Set rngBsMAttribute(9) = wksBsM.Cells(lngDatensatz, intZielspalte + 15)
#strBsMAttribute(10) = "ASIL"
#Set rngBsMAttribute(10) = wksBsM.Cells(lngDatensatz, intZielspalte + 16)
#strBsMAttribute(14) = "Status"
#Set rngBsMAttribute(14) = wksBsM.Cells(lngDatensatz, intZielspalte + 17)
#strBsMAttribute(29) = "Cluster Testing"
#Set rngBsMAttribute(29) = wksBsM.Cells(lngDatensatz, intZielspalte + 18)
#strBsMAttribute(23) = "#abgelehnt_nicht_testbar"
#Set rngBsMAttribute(23) = wksBsM.Cells(lngDatensatz, intZielspalte + 19)
#strBsMAttribute(20) = "MV"
#Set rngBsMAttribute(20) = wksBsM.Cells(lngDatensatz, intZielspalte + 20)
#strBsMAttribute(33) = "Anforderungsverantwortlicher"
#Set rngBsMAttribute(33) = wksBsM.Cells(lngDatensatz, intZielspalte + 21)
#strBsMAttribute(15) = "TD-VK"
#Set rngBsMAttribute(15) = wksBsM.Cells(lngDatensatz, intZielspalte + 22)
#strBsMAttribute(25) = "Status TD-VK"
#Set rngBsMAttribute(25) = wksBsM.Cells(lngDatensatz, intZielspalte + 23)
#strBsMAttribute(31) = "TD-VK temp1_Text"
#Set rngBsMAttribute(31) = wksBsM.Cells(lngDatensatz, intZielspalte + 24)
#strBsMAttribute(32) = "TD-VK Effort Estimation"
#Set rngBsMAttribute(32) = wksBsM.Cells(lngDatensatz, intZielspalte + 25)
#strBsMAttribute(16) = "TD-AA"
#Set rngBsMAttribute(16) = wksBsM.Cells(lngDatensatz, intZielspalte + 26)
#strBsMAttribute(17) = "TD-TI:TU"
#Set rngBsMAttribute(17) = wksBsM.Cells(lngDatensatz, intZielspalte + 27)
#strBsMAttribute(18) = "Testfälle"
#Set rngBsMAttribute(18) = wksBsM.Cells(lngDatensatz, intZielspalte + 28)
#strBsMAttribute(19) = "Vergleich TUs (TD-TF) - operativ"
#Set rngBsMAttribute(19) = wksBsM.Cells(lngDatensatz, intZielspalte + 29)
#strBsMAttribute(28) = "Erläuterungen zum Vergleich"
#Set rngBsMAttribute(28) = wksBsM.Cells(lngDatensatz, intZielspalte + 30)
#strBsMAttribute(26) = "Fehlende TUs bei TD-AAs"
#Set rngBsMAttribute(26) = wksBsM.Cells(lngDatensatz, intZielspalte + 31)
#strBsMAttribute(27) = "Fehlende TUs bei TFs"
#Set rngBsMAttribute(27) = wksBsM.Cells(lngDatensatz, intZielspalte + 32)
#strBsMAttribute(24) = "Zugeordnete I-Stufe"
#Set rngBsMAttribute(24) = wksBsM.Cells(lngDatensatz, intZielspalte + 33)
#strBsMAttribute(30) = "Projekt"
#Set rngBsMAttribute(30) = wksBsM.Cells(lngDatensatz, intZielspalte + 34)
#
#'Ergänzung projektspezifischer Attribute
#If strProjekt = "MEB21" Or strProjekt = "MQB48W" Then
#    ReDim Preserve strBsMAttribute(LBound(strBsMAttribute, 1) To UBound(strBsMAttribute, 1) + 1)
#    ReDim Preserve rngBsMAttribute(LBound(strBsMAttribute, 1) To UBound(strBsMAttribute, 1))
#    strBsMAttribute(35) = "Temp11_Auswahlfeld"
#    Set rngBsMAttribute(35) = wksBsM.Cells(lngDatensatz, intZielspalte + 35)    ' => nachträglich an richtige Stelle verschieben?
#End If
#
#'Tabellenkopf anlegen
#For i = LBound(strBsMAttribute, 1) To UBound(strBsMAttribute, 1)
#    With rngBsMAttribute(i)
#        .Value = strBsMAttribute(i)
#        .Font.Bold = True
#        .Interior.Color = RGB(217, 217, 217)
#    End With
#Next i
#
#'Bekannte Testumgebungen
#ReDim strBekannteTUs(1 To 17)
#intRelevantekTUs = 9
#strBekannteTUs(1) = "BRS-HiL_Laborplatz_automatisiert"
#strBekannteTUs(2) = "BRS-HiL_Basis-Funktion"
#strBekannteTUs(3) = "BRS-HiL_Kunden-Funktion"
#strBekannteTUs(4) = "BRS-HiL_Bremssystem"
#strBekannteTUs(5) = "BRS-Fahrversuch_Kunden-Funktion"
#strBekannteTUs(6) = "BRS-Fahrversuch_Basis-Funktion"
#strBekannteTUs(7) = "Vernetzter-Fahrwerks-HiL_Kundenfunktion"
#strBekannteTUs(8) = "BRS-HiL_Basisdienst_Halten"
#strBekannteTUs(9) = "BRS-HiL_Basisdienst_Verzoegern"
#'ab hier nicht mehr relevant
#strBekannteTUs(10) = "BRS-SiL_Kunden-Funktion"
#strBekannteTUs(11) = "Code-Review"
#strBekannteTUs(12) = "Design-Review"
#strBekannteTUs(13) = "Dokumenten-Review"
#strBekannteTUs(14) = "Prozess-Review"
#strBekannteTUs(15) = "Entscheidung_liegt_bei_Testinstanz"
#strBekannteTUs(16) = "BRS-Fahrversuch_Applikation"
#strBekannteTUs(17) = "BRS-Fahrversuch_Erprobung"
#'Statuswerte intBekannteTUs:
#'   TF \ VK                         kein VK     VK vorhanden
#'   kein TF                         0           1
#'   TF operativ                     10          11
#'   TF nicht operativ               20          21
#'   TF operativ und nicht operativ  30          31
#
#'Zähler für weitere Testumgebungen
#intWeitereTUs = 0
#strWeitereTUsAusgabe = ""
#
#'Tabelle füllen
#'Relevante Testumgebungen für Abgleich zwischen TDs und TFs
#ReDim strAbgleichTUs(1 To intRelevantekTUs)
#For i = 1 To intRelevantekTUs
#    strAbgleichTUs(i) = strBekannteTUs(i)
#Next i
#
#lngDatensatz = 0
#'BsM-Daten ausgeben
#For Each varErfassteBsMDatensatzItem In BsMDatenList
#    'Zähler für Datensatz/Zeile
#    lngDatensatz = lngDatensatz + 1
#    'Kalenderwoche der Datenauswertung
#    If WorksheetFunction.WeekNum(Date, 2) < 10 Then
#        rngBsMAttribute(34).Offset(lngDatensatz, 0).Value = CStr(Year(Date) & "/" & "0" & WorksheetFunction.WeekNum(Date, 2))
#    Else
#        rngBsMAttribute(34).Offset(lngDatensatz, 0).Value = CStr(Year(Date) & "/" & WorksheetFunction.WeekNum(Date, 2))
#    End If
#    'Vorgänger ID
#    If blnAVWVorgaengerIDsVerwenden Then
#        rngBsMAttribute(0).Offset(lngDatensatz, 0).Value = varErfassteBsMDatensatzItem.AVWVorgaengerID
#    End If
#    'Ausgabe ID
#    rngBsMAttribute(1).Offset(lngDatensatz, 0).Value = varErfassteBsMDatensatzItem.AVWID
#    'Ausgabe Dokument-ID
#    rngBsMAttribute(2).Offset(lngDatensatz, 0).Value = varErfassteBsMDatensatzItem.AVWDokumentID
#    'Ausgabe BsM-Relevanz
#    rngBsMAttribute(3).Offset(lngDatensatz, 0).Value = varErfassteBsMDatensatzItem.BSMRelevanz
#    'Ausgabe BSM-SaFuSi
#    rngBsMAttribute(4).Offset(lngDatensatz, 0).Value = varErfassteBsMDatensatzItem.AVWBsMSaFuSi
#    'Ausgabe BSM-ZZ
#    rngBsMAttribute(5).Offset(lngDatensatz, 0).Value = varErfassteBsMDatensatzItem.AVWBsMZZ
#    'Ausgabe BSM-ED
#    rngBsMAttribute(6).Offset(lngDatensatz, 0).Value = varErfassteBsMDatensatzItem.AVWBsMED
#    'Ausgabe BSM-FFF
#    rngBsMAttribute(7).Offset(lngDatensatz, 0).Value = varErfassteBsMDatensatzItem.AVWBsMFFF
#    'Ausgabe BSM-O
#    rngBsMAttribute(8).Offset(lngDatensatz, 0).Value = varErfassteBsMDatensatzItem.AVWBsMO
#    'Ausgabe BSM-Se
#    rngBsMAttribute(9).Offset(lngDatensatz, 0).Value = varErfassteBsMDatensatzItem.AVWBsMSe
#    'Ausgabe ASIL
#    rngBsMAttribute(10).Offset(lngDatensatz, 0).Value = varErfassteBsMDatensatzItem.AVWASIL
#    'Ausgabe Feature
#    rngBsMAttribute(11).Offset(lngDatensatz, 0).Value = varErfassteBsMDatensatzItem.AVWFeature
#    'Ausgabe Reifegrad
#    rngBsMAttribute(12).Offset(lngDatensatz, 0).Value = varErfassteBsMDatensatzItem.AVWReifegrad
#    'Ausgabe Umsetzer
#    rngBsMAttribute(13).Offset(lngDatensatz, 0).Value = varErfassteBsMDatensatzItem.AVWUmsetzer
#    'Ausgabe Status
#    rngBsMAttribute(14).Offset(lngDatensatz, 0).Value = varErfassteBsMDatensatzItem.AVWStatus
#    'Ausgabe MV
#    rngBsMAttribute(20).Offset(lngDatensatz, 0).Value = varErfassteBsMDatensatzItem.AVWMV
#    'Ausgabe Kategorie
#    rngBsMAttribute(21).Offset(lngDatensatz, 0).Value = varErfassteBsMDatensatzItem.AVWKategorie
#    'Ausgabe Dokumentenname
#    rngBsMAttribute(22).Offset(lngDatensatz, 0).Value = varErfassteBsMDatensatzItem.AVWDokumentName
#    'Ausgabe #abgelehnt_nicht_testbar
#    rngBsMAttribute(23).Offset(lngDatensatz, 0).Value = varErfassteBsMDatensatzItem.AVWAbgelehntNichtTestbar
#    'Ausgabe Zugeordnete I-Stufe
#    rngBsMAttribute(24).Offset(lngDatensatz, 0).Value = varErfassteBsMDatensatzItem.IStufe
#    'Ausgabe Cluster Testing
#    rngBsMAttribute(29).Offset(lngDatensatz, 0).Value = varErfassteBsMDatensatzItem.ClusterTesting
#    'Ausgabe Projekt
#    rngBsMAttribute(30).Offset(lngDatensatz, 0).Value = strProjekt
#    'Ausgabe Anforderungsverantwortliche
#    rngBsMAttribute(33).Offset(lngDatensatz, 0).Value = varErfassteBsMDatensatzItem.AVWAnforderungsverantwortliche
#
#    'Rücksetzen der Variablen für TU-Abgleich
#    ReDim intAbgleichTUs(LBound(strAbgleichTUs, 1) To UBound(strAbgleichTUs, 1))
#    ReDim intAuswertungTUs(1 To 31)
#    ReDim strAuswertungTUs(1 To 31)
#    strAuswertungTUsFehlendeAAs = ""
#    strAuswertungTUsFehlendeTFs = ""
#    intAusgabeAuswertungTUs = 0
#    strAusgabeAuswertungTUs = ""
#    strAusgabeAuswertungTUsDetails = ""
#
#    'Auswertung TF
#    strTestfaelle = ""
#    If varErfassteBsMDatensatzItem.Testfaelle.Count > 0 Then
#        For Each varErfassteTFItem In varErfassteBsMDatensatzItem.Testfaelle
#            'Testfälle zusammenführen
#            If strTestfaelle = "" Then
#                strTestfaelle = varErfassteTFItem.TF_ID & " - " & varErfassteTFItem.TF_Status & " - " & varErfassteTFItem.TF_Testinstanz & " - " & varErfassteTFItem.TF_Testumgebungstyp
#            Else
#                strTestfaelle = strTestfaelle & vbCrLf & varErfassteTFItem.TF_ID & " - " & varErfassteTFItem.TF_Status & " - " & varErfassteTFItem.TF_Testinstanz & " - " & varErfassteTFItem.TF_Testumgebungstyp
#            End If
#
#            'Erfassung der vorhandenen relevanten Testumgebungen
#            blnTFTUZugeordnet = False
#            For i = LBound(strAbgleichTUs, 1) To UBound(strAbgleichTUs, 1)
#                If varErfassteTFItem.TF_Testumgebungstyp = strAbgleichTUs(i) Then
#                    blnTFTUZugeordnet = True
#                    'Unterscheidung nach Status des Testfalls
#                    If varErfassteTFItem.TF_Status = "Operativ" Then
#                        'Nicht operative Testfälle bereits erfasst?
#                        If intAbgleichTUs(i) = 0 Then
#                            intAbgleichTUs(i) = 10
#                        ElseIf intAbgleichTUs(i) = 20 Then
#                            intAbgleichTUs(i) = 30
#                        End If
#                    Else
#                        'Operative Testfälle bereits erfasst?
#                        If intAbgleichTUs(i) = 0 Then
#                            intAbgleichTUs(i) = 20
#                        ElseIf intAbgleichTUs(i) = 10 Then
#                            intAbgleichTUs(i) = 30
#                        End If
#                    End If
#                End If
#            Next i
#'            'Restliche bekannte TUs abgleichen
#'            If blnTFTUZugeordnet = False Then
#'                For i = intRelevantekTUs + 1 To UBound(strBekannteTUs, 1)
#'                    If varErfassteTFItem.TF_Testumgebungstyp = strBekannteTUs(i) Then
#'                        blnTFTUZugeordnet = True
#'                        Exit For
#'                    End If
#'                Next i
#'            End If
#            'Weitere TUs erfassen
#            If blnTFTUZugeordnet = False Then
#                If intWeitereTUs > 0 Then
#                    For intWeitereTUsZaehler = 1 To intWeitereTUs
#                        If strWeitereTUs(intWeitereTUsZaehler) = varErfassteTFItem.TF_Testumgebungstyp Then
#                            blnTFTUZugeordnet = True
#                            Exit For
#                        End If
#                    Next intWeitereTUsZaehler
#                    If blnTFTUZugeordnet = False Then
#                        intWeitereTUs = intWeitereTUs + 1
#                        ReDim Preserve strWeitereTUs(1 To intWeitereTUs)
#                        strWeitereTUs(intWeitereTUs) = varErfassteTFItem.TF_Testumgebungstyp
#                    End If
#                Else
#                    intWeitereTUs = 1
#                    ReDim strWeitereTUs(1 To intWeitereTUs)
#                    strWeitereTUs(intWeitereTUs) = varErfassteTFItem.TF_Testumgebungstyp
#                End If
#            End If
#        Next varErfassteTFItem
#    End If
#    
#    'Auswertung TD
#    strTDAA = ""
#    strTDTiTu = ""
#    blnTIUnerlaubt = False
#    
#    'Trennung der Umsetzer für Abgleich der Testinstanzen
#    varUmsetzer = Split(varErfassteBsMDatensatzItem.AVWUmsetzer, ",", , vbBinaryCompare)
#    If varErfassteBsMDatensatzItem.AVWUmsetzer <> "" Then
#    ReDim blnUmsetzer(LBound(varUmsetzer, 1) To UBound(varUmsetzer, 1))
#    For intUmsetzer = LBound(varUmsetzer, 1) To UBound(varUmsetzer, 1)
#        blnUmsetzer(intUmsetzer) = False
#    Next intUmsetzer
#    End If
#    
#    If varErfassteBsMDatensatzItem.Verifikationskriterium.Count > 0 Then
#        'Ausgabe TD-VK inkl. Sicherheitsprüfung für mehrere Verifikationskriterien
#        If varErfassteBsMDatensatzItem.Verifikationskriterium.Count > 1 Then
#            rngBsMAttribute(15).Offset(lngDatensatz, 0).Value = "Achtung, mehrere Verifikationskriterien vorhanden! Ausgabe nur des ersten Items." & vbCrLf & _
#            varErfassteBsMDatensatzItem.Verifikationskriterium.Item(1).VK_ID
#        Else
#            rngBsMAttribute(15).Offset(lngDatensatz, 0).Value = varErfassteBsMDatensatzItem.Verifikationskriterium.Item(1).VK_ID
#        End If
#        'Ausgabe TD-VK Status
#        rngBsMAttribute(25).Offset(lngDatensatz, 0).Value = varErfassteBsMDatensatzItem.Verifikationskriterium.Item(1).VK_status
#        'Ausgabe TD-VK temp1_Text
#        rngBsMAttribute(31).Offset(lngDatensatz, 0).Value = varErfassteBsMDatensatzItem.Verifikationskriterium.Item(1).VK_temp1Text
#        'Ausgabe TD-VK Effort Estimation - Aufwandsschätzung auf Basis der Vorkommen von "Use-Case", "Step", "Aktion"
#        dblTDVKAnzahlUseCases = 1
#        strTDVKAktion = varErfassteBsMDatensatzItem.Verifikationskriterium.Item(1).VK_Aktion
#        If strTDVKAktion <> "" Then
#            strTDVKAktion = Replace(UCase(strTDVKAktion), "USE CASE", "USE-CASE")
#            strTDVKAktion = Replace(UCase(strTDVKAktion), "USECASE", "USE-CASE")
#            dblTDVKAnzahlUseCases = (Len(strTDVKAktion) - Len(Replace(UCase(strTDVKAktion), "USE-CASE", ""))) / Len("Use-Case")
#            'Anzahl 1 bei Befüllung ohne Vorkommen der Schlagwörter
#            If dblTDVKAnzahlUseCases = 0 Then dblTDVKAnzahlUseCases = 1
#        End If
#        rngBsMAttribute(32).Offset(lngDatensatz, 0).Value = dblTDVKAnzahlUseCases
#       
#        'Auswertung TD-AA
#        If varErfassteBsMDatensatzItem.Verifikationskriterium.Item(1).Absicherungsauftraege.Count > 0 Then
#            For Each varErfassteTDAAItem In varErfassteBsMDatensatzItem.Verifikationskriterium.Item(1).Absicherungsauftraege
#                'Absicherungsaufträge zusammenführen
#                If strTDAA = "" Then
#                    strTDAA = varErfassteTDAAItem.abs_ID
#                Else
#                    strTDAA = strTDAA & vbCrLf & varErfassteTDAAItem.abs_ID
#                End If
#                'Ti-Tu-Kombinationen zusammenführen
#                If strTDTiTu = "" Then
#                    strTDTiTu = varErfassteTDAAItem.testinstanz & ": " & varErfassteTDAAItem.Testumgebungstyp
#                Else
#                    strTDTiTu = strTDTiTu & vbCrLf & varErfassteTDAAItem.testinstanz & ": " & varErfassteTDAAItem.Testumgebungstyp
#                End If
#                
#                'Auswertung ob relevante Testinstanzen abgedeckt sind
#                If varErfassteTDAAItem.testinstanz <> "eigene Organisationseinheit" And varErfassteTDAAItem.testinstanz <> "Dauerlauf Gesamtfahrzeug" And varErfassteTDAAItem.testinstanz <> "Gesamtverbundintegration" And varErfassteTDAAItem.testinstanz <> "HMS" And varErfassteTDAAItem.testinstanz <> "VZM" Then
#                    For intUmsetzer = LBound(varUmsetzer, 1) To UBound(varUmsetzer, 1)
#                        If Trim(varErfassteTDAAItem.testinstanz) = Trim(varUmsetzer(intUmsetzer)) Then
#                            blnUmsetzer(intUmsetzer) = True
#                            Exit For
#                        End If
#                    Next intUmsetzer
#                End If
#                
#'                'Auswertung Testinstanz Erlaubt/Unerlaubt
#'                If varErfassteTDAAItem.testinstanz <> "eigene Organisationseinheit" And varErfassteTDAAItem.testinstanz <> "Dauerlauf Gesamtfahrzeug" And varErfassteTDAAItem.testinstanz <> "Gesamtverbundintegration" And varErfassteTDAAItem.testinstanz <> "HMS" And varErfassteTDAAItem.testinstanz <> "VZM" Then
#'                    If AbgleichUmsetzerTI(varErfassteTDAAItem.testinstanz, varErfassteBsMDatensatzItem.AVWUmsetzer) = False Then
#'                        blnTIUnerlaubt = True
#'                    End If
#'                End If
#                
#                'Abgleich der vorhandenen relevanten Testumgebungen
#                blnAATUZugeordnet = False
#                For i = LBound(strAbgleichTUs, 1) To UBound(strAbgleichTUs, 1)
#                    If varErfassteTDAAItem.Testumgebungstyp = strAbgleichTUs(i) Then
#                        blnAATUZugeordnet = True
#                        If intAbgleichTUs(i) = 0 Then
#                            'VK-TU vorhanden, kein Testfall vorhanden
#                            intAbgleichTUs(i) = 1
#                        ElseIf intAbgleichTUs(i) = 10 Then
#                            'VK-TU vorhanden, Testfälle operativ
#                            intAbgleichTUs(i) = 11
#                        ElseIf intAbgleichTUs(i) = 20 Then
#                            'VK-TU vorhanden, Testfälle nicht operativ
#                            intAbgleichTUs(i) = 21
#                        ElseIf intAbgleichTUs(i) = 30 Then
#                            'VK-TU vorhanden, Testfälle operativ und nicht operativ
#                            intAbgleichTUs(i) = 31
#                        End If
#                    End If
#                Next i
#'                'Restliche bekannte TUs abgleichen
#'                If blnAATUZugeordnet = False Then
#'                    For i = intRelevantekTUs + 1 To UBound(strBekannteTUs, 1)
#'                        If varErfassteTDAAItem.Testumgebungstyp = strBekannteTUs(i) Then
#'                            blnAATUZugeordnet = True
#'                            Exit For
#'                        End If
#'                    Next i
#'                End If
#                'Weitere TUs erfassen
#                If blnAATUZugeordnet = False Then
#                    If intWeitereTUs > 0 Then
#                        For intWeitereTUsZaehler = 1 To intWeitereTUs
#                            If strWeitereTUs(intWeitereTUsZaehler) = varErfassteTDAAItem.Testumgebungstyp Then
#                                blnAATUZugeordnet = True
#                                Exit For
#                            End If
#                        Next intWeitereTUsZaehler
#                        If blnAATUZugeordnet = False Then
#                            intWeitereTUs = intWeitereTUs + 1
#                            ReDim Preserve strWeitereTUs(1 To intWeitereTUs)
#                            strWeitereTUs(intWeitereTUs) = varErfassteTDAAItem.Testumgebungstyp
#                        End If
#                    Else
#                        intWeitereTUs = 1
#                        ReDim strWeitereTUs(1 To intWeitereTUs)
#                        strWeitereTUs(intWeitereTUs) = varErfassteTDAAItem.Testumgebungstyp
#                    End If
#                End If
#            Next varErfassteTDAAItem
#        
#            'Auswertung des Abgleichs der relevanten Testumgebungen
#            Call AuswertungTUAbgleich(intAbgleichTUs, strAbgleichTUs, intAuswertungTUs, strAuswertungTUs)
#            
#            'Erzeugung der Ausgabe für TU-Vergleich
#            Call AusgabeTUAbgleich(intAuswertungTUs, strAuswertungTUs, strAuswertungTUsFehlendeAAs, strAuswertungTUsFehlendeTFs, intAusgabeAuswertungTUs, strAusgabeAuswertungTUs, strAusgabeAuswertungTUsDetails)
#        Else
#            'Keine Absicherungsaufträge vorhanden
#            strAusgabeAuswertungTUs = "Kein Absicherungsauftrag vorhanden"
#            intAusgabeAuswertungTUs = 3
#        End If
#    Else
#        'Kein Verifikationskriterium vorhanden
#        strAusgabeAuswertungTUs = "Kein Verifikationskriterium vorhanden"
#        intAusgabeAuswertungTUs = 3
#    End If
#    
#    'Ausgabe Testfälle
#    rngBsMAttribute(18).Offset(lngDatensatz, 0).Value = strTestfaelle
#    'Ausgabe TD-AA
#    rngBsMAttribute(16).Offset(lngDatensatz, 0).Value = strTDAA
#    'Ausgabe TD-TI:TU
#    rngBsMAttribute(17).Offset(lngDatensatz, 0).Value = strTDTiTu
#    'Alle Umsetzer in den Testinstanzen abgedeckt?
#    blnTIUnerlaubt = False
#    For intUmsetzer = LBound(varUmsetzer, 1) To UBound(varUmsetzer, 1)
#        If blnUmsetzer(intUmsetzer) = False Then
#            blnTIUnerlaubt = True
#        End If
#    Next intUmsetzer
#    If blnTIUnerlaubt Then
#        rngBsMAttribute(17).Offset(lngDatensatz, 0).Interior.Color = RGB(255, 255, 102)
#    End If
#    'Ausgabe Vergleich TUs
#    With rngBsMAttribute(19).Offset(lngDatensatz, 0)
#        .Value = strAusgabeAuswertungTUs
#        If intAusgabeAuswertungTUs = 1 Then
#            'Grün
#            .Interior.Color = RGB(51, 204, 51)
#        ElseIf intAusgabeAuswertungTUs = 2 Then
#            'Gelb
#            .Interior.Color = RGB(255, 255, 102)
#        ElseIf intAusgabeAuswertungTUs = 3 Then
#            'Rot
#            .Interior.Color = RGB(255, 51, 0)
#        End If
#    End With
#    With rngBsMAttribute(28).Offset(lngDatensatz, 0)
#        .Value = strAusgabeAuswertungTUsDetails
#        If intAusgabeAuswertungTUs = 1 Then
#            'Grün
#            .Interior.Color = RGB(51, 204, 51)
#        ElseIf intAusgabeAuswertungTUs = 2 Then
#            'Gelb
#            .Interior.Color = RGB(255, 255, 102)
#        ElseIf intAusgabeAuswertungTUs = 3 Then
#            'Rot
#            .Interior.Color = RGB(255, 51, 0)
#        End If
#    End With
#    'Ausgabe Erläuterungen zum Vergleich
#    'Ausgabe fehlende TUs bei TD-AAs
#    rngBsMAttribute(26).Offset(lngDatensatz, 0).Value = strAuswertungTUsFehlendeAAs
#    'Ausgabe fehlende TUs bei TFs
#    rngBsMAttribute(27).Offset(lngDatensatz, 0).Value = strAuswertungTUsFehlendeTFs
#    
#    'Projektspezifische Ausgabe - MEB21
#    If strProjekt = "MEB21" Or strProjekt = "MQB48W" Then
#        rngBsMAttribute(35).Offset(lngDatensatz, 0).Value = varErfassteBsMDatensatzItem.AVWTemp11_Auswahlfeld
#    End If
#    
#Next varErfassteBsMDatensatzItem
#
#'Weitere TUs zusammenfassen
#If intWeitereTUs > 0 Then
#    For i = LBound(strWeitereTUs, 1) To UBound(strWeitereTUs, 1)
#        If strWeitereTUsAusgabe = "" Then
#            strWeitereTUsAusgabe = strWeitereTUs(i)
#        Else
#            strWeitereTUsAusgabe = strWeitereTUsAusgabe & vbCrLf & strWeitereTUs(i)
#        End If
#    Next i
#End If
#
#'Spaltenbreite anpassen
#With wksBsM.Cells
#    .Columns.AutoFit
#    .Rows.AutoFit
#End With
#
#'Projektspezifische Sortierung - MEB21
#If strProjekt = "MEB21" Then
#    wksBsM.Columns(35).Cut
#    wksBsM.Columns(18).Insert shift:=xlToRight
#End If
#
#'Filterung aktivieren
#wksBsM.Rows(1).AutoFilter
#
#'Dateinamen ausgeben und verstecken
#lngDatensatz = 1
#wksBsM.Rows(lngDatensatz).EntireRow.Insert shift:=xlDown
#wksBsM.Cells(lngDatensatz, 1) = "Anforderungen: " & strDateinamen(1) & vbCrLf & "Verifikationskriterien: " & strDateinamen(2) & vbCrLf & _
#                                "Absicherungsaufträge: " & strDateinamen(3) & vbCrLf & "Testfälle: " & strDateinamen(3) & vbCrLf & _
#                                "FRU-Timing: " & strDateinamen(5)
#wksBsM.Rows(lngDatensatz).EntireRow.Hidden = True
#End Sub
#
#Private Sub AusgabeTDStatus(ByVal wbBsM As Workbook, ByRef wksTD As Worksheet, ByRef strTDAttribute() As String, ByRef rngTDAttribute() As Range, ByRef strDateinamen() As String, ByVal strProjekt As String)
#Dim lngDatensatz As Long                        'Long-Variable für aktuell zu schreibenden Datensatz
#Dim Verifikationskriterium As Verifikationskriterium    'Verifikationskriterium
#Dim varErfassteTDAAItem As Variant              'Variant für Item aus den jeweiligen Absicherungsaufträgen
#Dim strTDAA As String                           'String für Sammlung der Absicherungsaufträge
#Dim strTDTiTu As String                         'String für Sammlung der Ti:Tu-Kombinationen
#Dim varErfassteTFItem As Variant                'Variant für Item aus den jeweiligen Testfällen
#Dim strTestfaelle As String                     'String für Sammlung der Testfälle
#Dim strAbgleichTUs() As String                  'String-Array für die Namen der abzugleichenden Testumgebungen bei TDs und TFs
#Dim intAbgleichTUs() As Integer                 'Integer-Array für Erfassung der Testumgebungen bei TDs und TFs
#Dim i As Long                                   'Laufvariable
#Dim intAuswertungTUs() As Integer               'Integer-Array für die Ergebnisse des Tu-Abgleichs
#Dim strAuswertungTUs() As String                'String-Array für die Ergebnisse des Tu-Abgleichs
#Dim strAusgabeAuswertungTUs As String           'String für Ausgabe des Tu-Abgleichs
#Dim intAusgabeAuswertungTUs As Integer          'Integer für Ausgabe des Tu-Abgleichs
#Dim strAuswertungTUsFehlendeAAs As String       'String für Ausgabe der fehlenden TUs bei TD-AAs
#Dim strAuswertungTUsFehlendeTFs As String       'String für Ausgabe der fehlenden TUs bei TFs
#Dim strAusgabeAuswertungTUsDetails As String    'String für Ausgabe des Tu-Abgleichs mit Details
#Dim intWeitereTUs As Integer                    'Zählvariable für weitere TUs
#Dim strWeitereTUs() As String                   'String-Array für weitere TUs
#Dim strBekannteTUs() As String                  'String-Array für alle bekannten TUs
#Dim blnTFTUZugeordnet As Boolean                'Flag für zugeordnete TU
#Dim blnAATUZugeordnet As Boolean                'Flag für zugeordnete TU
#Dim intWeitereTUsZaehler As Integer             'Laufvariable für weitere TUs
#Dim intRelevantekTUs As Integer                 'Integer für Anzahl der relevanten TUs
#Dim dblTDVKAnzahlUseCases As Double             'Anzahl der Vorkommen der Use-Case-Begriffe
#Dim strTDVKAktion As String                     'String zur Bearbeitung der TDVK-Aktion
#
#'Tabelle erzeugen
#'Neues Worksheet erzeugen
#Set wksTD = wbBsM.Sheets.Add(after:=wbBsM.Worksheets(wbBsM.Worksheets.Count))
#wksTD.Name = "TD_Status_" & "Today" & "_" & Replace(Time, ":", "")
#'Arbeitsblatt TD_Status
#'#1: TD-VK, #2: Status TD-VK, #3: TD-AA, #4: TD-TI:TU, #5: Testfälle, #6: Vergleich TUs (TD-TF) - operativ, #7: Erläuterungen zum Vergleich,
#'#8: Fehlende TUs bei TD-AAs, #9: Fehlende TUs bei TFs, #10: Anforderungs-IDs, #11: Zugeordnete I-Stufe, #12: Umsetzer, #13: BsM-Relevanz,
#'#14: ASIL, #15: Feature, #16: Reifegrad, #17: MV, #18: LAH-ID, #19: Dokumente (LAH), #20: Cluster Testing, #21: Projekt, #22: TD-VK temp1_Text,
#'#23: TD-VK Effort Estimation, #24: Anforderungsverantwortliche, #25: KW
#'Projektspezifisch - MEB21
#'#26: Temp11_Auswahlfeld
#ReDim strTDAttribute(1 To 25)
#ReDim rngTDAttribute(LBound(strTDAttribute, 1) To UBound(strTDAttribute, 1))
#'Name und Position der Tabellenattribute
#strTDAttribute(25) = "KW Datenauswertung"
#Set rngTDAttribute(25) = wksTD.Cells(1, 1)
#strTDAttribute(1) = "TD-VK"
#Set rngTDAttribute(1) = wksTD.Cells(1, 2)
#strTDAttribute(2) = "Status TD-VK"
#Set rngTDAttribute(2) = wksTD.Cells(1, 3)
#strTDAttribute(22) = "TD-VK temp1_Text"
#Set rngTDAttribute(22) = wksTD.Cells(1, 4)
#strTDAttribute(23) = "TD-VK Effort Estimation"
#Set rngTDAttribute(23) = wksTD.Cells(1, 5)
#strTDAttribute(3) = "TD-AA"
#Set rngTDAttribute(3) = wksTD.Cells(1, 6)
#strTDAttribute(4) = "TD-TI:TU"
#Set rngTDAttribute(4) = wksTD.Cells(1, 7)
#strTDAttribute(5) = "Testfälle"
#Set rngTDAttribute(5) = wksTD.Cells(1, 8)
#strTDAttribute(6) = "Vergleich TUs (TD-TF) - operativ"
#Set rngTDAttribute(6) = wksTD.Cells(1, 9)
#strTDAttribute(7) = "Erläuterungen zum Vergleich"
#Set rngTDAttribute(7) = wksTD.Cells(1, 10)
#strTDAttribute(8) = "Fehlende TUs bei TD-AAs"
#Set rngTDAttribute(8) = wksTD.Cells(1, 11)
#strTDAttribute(9) = "Fehlende TUs bei TFs"
#Set rngTDAttribute(9) = wksTD.Cells(1, 12)
#strTDAttribute(10) = "Anforderungs-IDs"
#Set rngTDAttribute(10) = wksTD.Cells(1, 13)
#strTDAttribute(20) = "Cluster Testing"
#Set rngTDAttribute(20) = wksTD.Cells(1, 14)
#strTDAttribute(14) = "ASIL (LAH)"
#Set rngTDAttribute(14) = wksTD.Cells(1, 15)
#strTDAttribute(13) = "BsM-Relevanz (LAH)"
#Set rngTDAttribute(13) = wksTD.Cells(1, 16)
#strTDAttribute(15) = "Feature (LAH)"
#Set rngTDAttribute(15) = wksTD.Cells(1, 17)
#strTDAttribute(16) = "Reifegrad (LAH)"
#Set rngTDAttribute(16) = wksTD.Cells(1, 18)
#strTDAttribute(12) = "Umsetzer (LAH)"
#Set rngTDAttribute(12) = wksTD.Cells(1, 19)
#strTDAttribute(17) = "MV (LAH)"
#Set rngTDAttribute(17) = wksTD.Cells(1, 20)
#strTDAttribute(24) = "Anforderungsverantwortliche (LAH)"
#Set rngTDAttribute(24) = wksTD.Cells(1, 21)
#strTDAttribute(18) = "LAH-ID"
#Set rngTDAttribute(18) = wksTD.Cells(1, 22)
#strTDAttribute(19) = "Dokumente (LAH)"
#Set rngTDAttribute(19) = wksTD.Cells(1, 23)
#strTDAttribute(11) = "Zugeordnete I-Stufe"
#Set rngTDAttribute(11) = wksTD.Cells(1, 24)
#strTDAttribute(21) = "Projekt"
#Set rngTDAttribute(21) = wksTD.Cells(1, 25)
#
#'Ergänzung projektspezifische Attribute
#If strProjekt = "MEB21" Or strProjekt = "MQB48W" Then
#    ReDim Preserve strTDAttribute(LBound(strTDAttribute, 1) To UBound(strTDAttribute, 1) + 1)
#    ReDim Preserve rngTDAttribute(LBound(strTDAttribute, 1) To UBound(strTDAttribute, 1))
#    strTDAttribute(26) = "Temp11_Auswahlfeld (LAH)"
#    Set rngTDAttribute(26) = wksTD.Cells(1, 26)
#End If
#
#'Tabellenkopf anlegen
#For i = LBound(strTDAttribute, 1) To UBound(strTDAttribute, 1)
#    With rngTDAttribute(i)
#        .Value = strTDAttribute(i)
#        .Font.Bold = True
#        .Interior.Color = RGB(217, 217, 217)
#    End With
#Next i
#
#'Bekannte Testumgebungen
#ReDim strBekannteTUs(1 To 17)
#intRelevantekTUs = 9
#strBekannteTUs(1) = "BRS-HiL_Laborplatz_automatisiert"
#strBekannteTUs(2) = "BRS-HiL_Basis-Funktion"
#strBekannteTUs(3) = "BRS-HiL_Kunden-Funktion"
#strBekannteTUs(4) = "BRS-HiL_Bremssystem"
#strBekannteTUs(5) = "BRS-Fahrversuch_Kunden-Funktion"
#strBekannteTUs(6) = "BRS-Fahrversuch_Basis-Funktion"
#strBekannteTUs(7) = "Vernetzter-Fahrwerks-HiL_Kundenfunktion"
#strBekannteTUs(8) = "BRS-HiL_Basisdienst_Halten"
#strBekannteTUs(9) = "BRS-HiL_Basisdienst_Verzoegern"
#'ab hier nicht mehr relevant
#strBekannteTUs(10) = "BRS-SiL_Kunden-Funktion"
#strBekannteTUs(11) = "Code-Review"
#strBekannteTUs(12) = "Design-Review"
#strBekannteTUs(13) = "Dokumenten-Review"
#strBekannteTUs(14) = "Prozess-Review"
#strBekannteTUs(15) = "Entscheidung_liegt_bei_Testinstanz"
#strBekannteTUs(16) = "BRS-Fahrversuch_Applikation"
#strBekannteTUs(17) = "BRS-Fahrversuch_Erprobung"
#'Statuswerte intBekannteTUs:
#'   TF \ VK                         kein VK     VK vorhanden
#'   kein TF                         0           1
#'   TF operativ                     10          11
#'   TF nicht operativ               20          21
#'   TF operativ und nicht operativ  30          31
#
#'Zähler für weitere Testumgebungen
#intWeitereTUs = 0
#
#'Tabelle füllen
#'Relevante Testumgebungen für Abgleich zwischen TDs und TFs
#ReDim strAbgleichTUs(1 To intRelevantekTUs)
#For i = 1 To intRelevantekTUs
#    strAbgleichTUs(i) = strBekannteTUs(i)
#Next i
#
#'TD-Daten ausgeben
#lngDatensatz = 0
#For Each Verifikationskriterium In verifikationKritList
#    If Verifikationskriterium.AnforderungVorhanden = True Then
#        'Zähler für Datensatz/Zeile
#        lngDatensatz = lngDatensatz + 1
#        'Kalenderwoche der Datenauswertung
#        If WorksheetFunction.WeekNum(Date, 2) < 10 Then
#            rngTDAttribute(25).Offset(lngDatensatz, 0).Value = CStr(Year(Date) & "/" & "0" & WorksheetFunction.WeekNum(Date, 2))
#        Else
#            rngTDAttribute(25).Offset(lngDatensatz, 0).Value = CStr(Year(Date) & "/" & WorksheetFunction.WeekNum(Date, 2))
#        End If
#        'Ausgabe VK-ID
#        rngTDAttribute(1).Offset(lngDatensatz, 0).Value = Verifikationskriterium.VK_ID
#        'Ausgabe VK-Status
#        rngTDAttribute(2).Offset(lngDatensatz, 0).Value = Verifikationskriterium.VK_status
#        'Ausgabe VK temp1_Text
#        rngTDAttribute(22).Offset(lngDatensatz, 0).Value = Verifikationskriterium.VK_temp1Text
#        'Ausgabe Anforderungs-IDs
#        rngTDAttribute(10).Offset(lngDatensatz, 0).Value = AusgabeSammlungLF(Verifikationskriterium.anf_ids)
#        'Ausgabe Zugeordnete I-Stufe
#        rngTDAttribute(11).Offset(lngDatensatz, 0).Value = AusgabeSammlungLFEinfach(Verifikationskriterium.anf_IStufen)
#        If AuswertungUnterschiedlicheIStufen(Verifikationskriterium.anf_IStufen) = True Then
#            rngTDAttribute(11).Offset(lngDatensatz, 0).Interior.Color = RGB(255, 255, 102)
#        End If
#        'Ausgabe Umsetzer
#        rngTDAttribute(12).Offset(lngDatensatz, 0).Value = AusgabeSammlungLFEinfach(Verifikationskriterium.anf_Umsetzer)
#        'Ausgabe BsM-Relevanz
#        rngTDAttribute(13).Offset(lngDatensatz, 0).Value = AusgabeSammlungLFEinfach(Verifikationskriterium.anf_BsMRelevanz)
#        'Ausgabe ASIL
#        rngTDAttribute(14).Offset(lngDatensatz, 0).Value = AusgabeSammlungLFEinfach(Verifikationskriterium.anf_ASIL)
#        'Ausgabe Feature
#        rngTDAttribute(15).Offset(lngDatensatz, 0).Value = AusgabeSammlungLFEinfach(Verifikationskriterium.anf_Feature)
#        'Ausgabe Reifegrad
#        rngTDAttribute(16).Offset(lngDatensatz, 0).Value = AusgabeSammlungLFEinfach(Verifikationskriterium.anf_Reifegrad)
#        'Ausgabe Modulverantwortlicher
#        rngTDAttribute(17).Offset(lngDatensatz, 0).Value = AusgabeSammlungLFEinfach(Verifikationskriterium.anf_MV)
#        'Ausgabe LAH-IDs
#        rngTDAttribute(18).Offset(lngDatensatz, 0).Value = AusgabeSammlungLFEinfach(Verifikationskriterium.anf_LAHID)
#        'Ausgabe LAH-Namen
#        rngTDAttribute(19).Offset(lngDatensatz, 0).Value = AusgabeSammlungLFEinfach(Verifikationskriterium.anf_LAHNamen)
#        'Ausgabe Cluster Testing
#        rngTDAttribute(20).Offset(lngDatensatz, 0).Value = AusgabeSammlungLFEinfach(Verifikationskriterium.anf_ClusterTesting)
#        'Ausgabe Projekt
#        rngTDAttribute(21).Offset(lngDatensatz, 0).Value = strProjekt
#        'Ausgabe Anforderungsverantwortliche
#        rngTDAttribute(24).Offset(lngDatensatz, 0).Value = AusgabeSammlungLFEinfach(Verifikationskriterium.anf_Anforderungsverantwortliche)
#        
#        'Ausgabe Aufwandsschätzung auf Basis der Vorkommen von "Use-Case", "Step", "Aktion"
#        dblTDVKAnzahlUseCases = 1
#        strTDVKAktion = Verifikationskriterium.VK_Aktion
#        If strTDVKAktion <> "" Then
#            strTDVKAktion = Replace(UCase(strTDVKAktion), "USE CASE", "USE-CASE")
#            strTDVKAktion = Replace(UCase(strTDVKAktion), "USECASE", "USE-CASE")
#            dblTDVKAnzahlUseCases = (Len(strTDVKAktion) - Len(Replace(UCase(strTDVKAktion), "USE-CASE", ""))) / Len("Use-Case")
#            'Anzahl 1 bei Befüllung ohne Vorkommen der Schlagwörter
#            If dblTDVKAnzahlUseCases = 0 Then dblTDVKAnzahlUseCases = 1
#        End If
#        rngTDAttribute(23).Offset(lngDatensatz, 0).Value = dblTDVKAnzahlUseCases
#        
#        'Rücksetzen der Variablen für TU-Abgleich
#        ReDim intAbgleichTUs(LBound(strAbgleichTUs, 1) To UBound(strAbgleichTUs, 1))
#        ReDim intAuswertungTUs(1 To 31)
#        ReDim strAuswertungTUs(1 To 31)
#        strAuswertungTUsFehlendeAAs = ""
#        strAuswertungTUsFehlendeTFs = ""
#        intAusgabeAuswertungTUs = 0
#        strAusgabeAuswertungTUs = ""
#        strAusgabeAuswertungTUsDetails = ""
#    
#        'Auswertung TF
#        strTestfaelle = ""
#        If Verifikationskriterium.VK_Testfaelle.Count > 0 Then
#            For Each varErfassteTFItem In Verifikationskriterium.VK_Testfaelle
#                'Testfälle zusammenführen
#                If strTestfaelle = "" Then
#                    strTestfaelle = varErfassteTFItem.TF_ID & " - " & varErfassteTFItem.TF_Status & " - " & varErfassteTFItem.TF_Testinstanz & " - " & varErfassteTFItem.TF_Testumgebungstyp
#                Else
#                    strTestfaelle = strTestfaelle & vbCrLf & varErfassteTFItem.TF_ID & " - " & varErfassteTFItem.TF_Status & " - " & varErfassteTFItem.TF_Testinstanz & " - " & varErfassteTFItem.TF_Testumgebungstyp
#                End If
#    
#                'Erfassung der vorhandenen relevanten Testumgebungen
#                blnTFTUZugeordnet = False
#                For i = LBound(strAbgleichTUs, 1) To UBound(strAbgleichTUs, 1)
#                    If varErfassteTFItem.TF_Testumgebungstyp = strAbgleichTUs(i) Then
#                        blnTFTUZugeordnet = True
#                        'Unterscheidung nach Status des Testfalls
#                        If varErfassteTFItem.TF_Status = "Operativ" Then
#                            'Nicht operative Testfälle bereits erfasst?
#                            If intAbgleichTUs(i) = 0 Then
#                                intAbgleichTUs(i) = 10
#                            ElseIf intAbgleichTUs(i) = 20 Then
#                                intAbgleichTUs(i) = 30
#                            End If
#                        Else
#                            'Operative Testfälle bereits erfasst?
#                            If intAbgleichTUs(i) = 0 Then
#                                intAbgleichTUs(i) = 20
#                            ElseIf intAbgleichTUs(i) = 10 Then
#                                intAbgleichTUs(i) = 30
#                            End If
#                        End If
#                    End If
#                Next i
#    '            'Restliche bekannte TUs abgleichen
#    '            If blnTFTUZugeordnet = False Then
#    '                For i = intRelevantekTUs + 1 To UBound(strBekannteTUs, 1)
#    '                    If varErfassteTFItem.TF_Testumgebungstyp = strBekannteTUs(i) Then
#    '                        blnTFTUZugeordnet = True
#    '                        Exit For
#    '                    End If
#    '                Next i
#    '            End If
#                'Weitere TUs erfassen
#                If blnTFTUZugeordnet = False Then
#                    If intWeitereTUs > 0 Then
#                        For intWeitereTUsZaehler = 1 To intWeitereTUs
#                            If strWeitereTUs(intWeitereTUsZaehler) = varErfassteTFItem.TF_Testumgebungstyp Then
#                                blnTFTUZugeordnet = True
#                                Exit For
#                            End If
#                        Next intWeitereTUsZaehler
#                        If blnTFTUZugeordnet = False Then
#                            intWeitereTUs = intWeitereTUs + 1
#                            ReDim Preserve strWeitereTUs(1 To intWeitereTUs)
#                            strWeitereTUs(intWeitereTUs) = varErfassteTFItem.TF_Testumgebungstyp
#                        End If
#                    Else
#                        intWeitereTUs = 1
#                        ReDim strWeitereTUs(1 To intWeitereTUs)
#                        strWeitereTUs(intWeitereTUs) = varErfassteTFItem.TF_Testumgebungstyp
#                    End If
#                End If
#            Next varErfassteTFItem
#        End If
#        
#        'Auswertung TD
#        strTDAA = ""
#        strTDTiTu = ""
#            
#        'Auswertung TD-AA
#        If Verifikationskriterium.Absicherungsauftraege.Count > 0 Then
#            For Each varErfassteTDAAItem In Verifikationskriterium.Absicherungsauftraege
#                'Absicherungsaufträge zusammenführen
#                If strTDAA = "" Then
#                    strTDAA = varErfassteTDAAItem.abs_ID
#                Else
#                    strTDAA = strTDAA & vbCrLf & varErfassteTDAAItem.abs_ID
#                End If
#                'Ti-Tu-Kombinationen zusammenführen
#                If strTDTiTu = "" Then
#                    strTDTiTu = varErfassteTDAAItem.testinstanz & ": " & varErfassteTDAAItem.Testumgebungstyp
#                Else
#                    strTDTiTu = strTDTiTu & vbCrLf & varErfassteTDAAItem.testinstanz & ": " & varErfassteTDAAItem.Testumgebungstyp
#                End If
#                    
#                'Abgleich der vorhandenen relevanten Testumgebungen
#                blnAATUZugeordnet = False
#                For i = LBound(strAbgleichTUs, 1) To UBound(strAbgleichTUs, 1)
#                    If varErfassteTDAAItem.Testumgebungstyp = strAbgleichTUs(i) Then
#                        blnAATUZugeordnet = True
#                        If intAbgleichTUs(i) = 0 Then
#                            'VK-TU vorhanden, kein Testfall vorhanden
#                            intAbgleichTUs(i) = 1
#                        ElseIf intAbgleichTUs(i) = 10 Then
#                            'VK-TU vorhanden, Testfälle operativ
#                            intAbgleichTUs(i) = 11
#                        ElseIf intAbgleichTUs(i) = 20 Then
#                            'VK-TU vorhanden, Testfälle nicht operativ
#                            intAbgleichTUs(i) = 21
#                        ElseIf intAbgleichTUs(i) = 30 Then
#                            'VK-TU vorhanden, Testfälle operativ und nicht operativ
#                            intAbgleichTUs(i) = 31
#                        End If
#                    End If
#                Next i
#    '                'Restliche bekannte TUs abgleichen
#    '                If blnAATUZugeordnet = False Then
#    '                    For i = intRelevantekTUs + 1 To UBound(strBekannteTUs, 1)
#    '                        If varErfassteTDAAItem.Testumgebungstyp = strBekannteTUs(i) Then
#    '                            blnAATUZugeordnet = True
#    '                            Exit For
#    '                        End If
#    '                    Next i
#    '                End If
#                'Weitere TUs erfassen
#                If blnAATUZugeordnet = False Then
#                    If intWeitereTUs > 0 Then
#                        For intWeitereTUsZaehler = 1 To intWeitereTUs
#                            If strWeitereTUs(intWeitereTUsZaehler) = varErfassteTDAAItem.Testumgebungstyp Then
#                                blnAATUZugeordnet = True
#                                Exit For
#                            End If
#                        Next intWeitereTUsZaehler
#                        If blnAATUZugeordnet = False Then
#                            intWeitereTUs = intWeitereTUs + 1
#                            ReDim Preserve strWeitereTUs(1 To intWeitereTUs)
#                            strWeitereTUs(intWeitereTUs) = varErfassteTDAAItem.Testumgebungstyp
#                        End If
#                    Else
#                        intWeitereTUs = 1
#                        ReDim strWeitereTUs(1 To intWeitereTUs)
#                        strWeitereTUs(intWeitereTUs) = varErfassteTDAAItem.Testumgebungstyp
#                    End If
#                End If
#            Next varErfassteTDAAItem
#            
#            'Auswertung des Abgleichs der relevanten Testumgebungen
#            Call AuswertungTUAbgleich(intAbgleichTUs, strAbgleichTUs, intAuswertungTUs, strAuswertungTUs)
#            
#            'Erzeugung der Ausgabe für TU-Vergleich
#            Call AusgabeTUAbgleich(intAuswertungTUs, strAuswertungTUs, strAuswertungTUsFehlendeAAs, strAuswertungTUsFehlendeTFs, intAusgabeAuswertungTUs, strAusgabeAuswertungTUs, strAusgabeAuswertungTUsDetails)
#        Else
#            'Keine Absicherungsaufträge vorhanden
#            strAusgabeAuswertungTUs = "Kein Absicherungsauftrag vorhanden"
#            intAusgabeAuswertungTUs = 3
#        End If
#             
#        'Ausgabe TD-AA
#        rngTDAttribute(3).Offset(lngDatensatz, 0).Value = strTDAA
#        'Ausgabe Testfälle
#        rngTDAttribute(5).Offset(lngDatensatz, 0).Value = strTestfaelle
#        'Ausgabe TD-TI:TU
#        rngTDAttribute(4).Offset(lngDatensatz, 0).Value = strTDTiTu
#        'Ausgabe Vergleich TUs
#        With rngTDAttribute(6).Offset(lngDatensatz, 0)
#            .Value = strAusgabeAuswertungTUs
#            If intAusgabeAuswertungTUs = 1 Then
#                'Grün
#                .Interior.Color = RGB(51, 204, 51)
#            ElseIf intAusgabeAuswertungTUs = 2 Then
#                'Gelb
#                .Interior.Color = RGB(255, 255, 102)
#            ElseIf intAusgabeAuswertungTUs = 3 Then
#                'Rot
#                .Interior.Color = RGB(255, 51, 0)
#            End If
#        End With
#        With rngTDAttribute(7).Offset(lngDatensatz, 0)
#            .Value = strAusgabeAuswertungTUsDetails
#            If intAusgabeAuswertungTUs = 1 Then
#                'Grün
#                .Interior.Color = RGB(51, 204, 51)
#            ElseIf intAusgabeAuswertungTUs = 2 Then
#                'Gelb
#                .Interior.Color = RGB(255, 255, 102)
#            ElseIf intAusgabeAuswertungTUs = 3 Then
#                'Rot
#                .Interior.Color = RGB(255, 51, 0)
#            End If
#        End With
#        'Ausgabe Erläuterungen zum Vergleich
#        'Ausgabe fehlende TUs bei TD-AAs
#        rngTDAttribute(8).Offset(lngDatensatz, 0).Value = strAuswertungTUsFehlendeAAs
#        'Ausgabe fehlende TUs bei TFs
#        rngTDAttribute(9).Offset(lngDatensatz, 0).Value = strAuswertungTUsFehlendeTFs
#    
#        'Projektspezifische Ausgabe - MEB21
#        If strProjekt = "MEB21" Or strProjekt = "MQB48W" Then
#            rngTDAttribute(26).Offset(lngDatensatz, 0).Value = AusgabeSammlungLFEinfach(Verifikationskriterium.anf_Temp11_Auswahlfeld)
#        End If
#    
#    End If
#    
#Next Verifikationskriterium
#
#'Spaltenbreite anpassen
#With wksTD.Cells
#    .Columns.EntireColumn.AutoFit
#    .Rows.AutoFit
#End With
#
#'Projektspezifische Sortierung - MEB21
#If strProjekt = "MEB21" Then
#    wksTD.Columns(26).Cut
#    wksTD.Columns(14).Insert shift:=xlToRight
#End If
#
#'Filterung aktivieren
#wksTD.Rows(1).AutoFilter
#
#'Dateinamen ausgeben und verstecken
#lngDatensatz = 1
#wksTD.Rows(lngDatensatz).EntireRow.Insert shift:=xlDown
#wksTD.Cells(lngDatensatz, 1) = "Anforderungen: " & strDateinamen(1) & vbCrLf & "Verifikationskriterien: " & strDateinamen(2) & vbCrLf & _
#                                "Absicherungsaufträge: " & strDateinamen(3) & vbCrLf & "Testfälle: " & strDateinamen(3) & vbCrLf & _
#                                "FRU-Timing: " & strDateinamen(5)
#wksTD.Rows(lngDatensatz).EntireRow.Hidden = True
#End Sub
#
#Private Sub AuswertungTUAbgleich(ByRef intAbgleichTUs() As Integer, ByRef strAbgleichTUs() As String, ByRef intAuswertungTUs() As Integer, ByRef strAuswertungTUs() As String)
#Dim i As Integer        'Laufvariable
#
#For i = LBound(strAbgleichTUs, 1) To UBound(strAbgleichTUs, 1)
#    If intAbgleichTUs(i) = 1 Then
#        '#TF nicht vorhanden, VK fachlich abgestimmt
#        intAuswertungTUs(1) = intAuswertungTUs(1) + 1
#        If strAuswertungTUs(1) = "" Then
#            strAuswertungTUs(1) = strAbgleichTUs(i)
#        Else
#            strAuswertungTUs(1) = strAuswertungTUs(1) & ", " & strAbgleichTUs(i)
#        End If
#        
#    ElseIf intAbgleichTUs(i) = 10 Then
#        '#TF operativ, VK nicht vorhanden
#        intAuswertungTUs(10) = intAuswertungTUs(10) + 1
#        If strAuswertungTUs(10) = "" Then
#            strAuswertungTUs(10) = strAbgleichTUs(i)
#        Else
#            strAuswertungTUs(10) = strAuswertungTUs(10) & ", " & strAbgleichTUs(i)
#        End If
#    ElseIf intAbgleichTUs(i) = 11 Then
#        '#TF operativ, VK fachlich abgestimmt
#        intAuswertungTUs(11) = intAuswertungTUs(11) + 1
#        If strAuswertungTUs(11) = "" Then
#            strAuswertungTUs(11) = strAbgleichTUs(i)
#        Else
#            strAuswertungTUs(11) = strAuswertungTUs(11) & ", " & strAbgleichTUs(i)
#        End If
#
#    ElseIf intAbgleichTUs(i) = 20 Then
#        '#TF nicht operativ, VK nicht vorhanden
#        intAuswertungTUs(20) = intAuswertungTUs(20) + 1
#        If strAuswertungTUs(20) = "" Then
#            strAuswertungTUs(20) = strAbgleichTUs(i)
#        Else
#            strAuswertungTUs(20) = strAuswertungTUs(20) & ", " & strAbgleichTUs(i)
#        End If
#    ElseIf intAbgleichTUs(i) = 21 Then
#        '#TF nicht operativ, VK fachlich abgestimmt
#        intAuswertungTUs(21) = intAuswertungTUs(21) + 1
#        If strAuswertungTUs(21) = "" Then
#            strAuswertungTUs(21) = strAbgleichTUs(i)
#        Else
#            strAuswertungTUs(21) = strAuswertungTUs(21) & ", " & strAbgleichTUs(i)
#        End If
#        
#    ElseIf intAbgleichTUs(i) = 30 Then
#        '#TF teilweise operativ und nicht operativ, VK nicht vorhanden
#        intAuswertungTUs(30) = intAuswertungTUs(30) + 1
#        If strAuswertungTUs(30) = "" Then
#            strAuswertungTUs(30) = strAbgleichTUs(i)
#        Else
#            strAuswertungTUs(30) = strAuswertungTUs(30) & ", " & strAbgleichTUs(i)
#        End If
#    ElseIf intAbgleichTUs(i) = 31 Then
#        '#TF teilweise operativ und nicht operativ, VK fachlich abgestimmt
#        intAuswertungTUs(31) = intAuswertungTUs(31) + 1
#        If strAuswertungTUs(31) = "" Then
#            strAuswertungTUs(31) = strAbgleichTUs(i)
#        Else
#            strAuswertungTUs(31) = strAuswertungTUs(31) & ", " & strAbgleichTUs(i)
#        End If
#    End If
#Next i
#End Sub
#
#Private Sub AusgabeTUAbgleich(ByRef intAuswertungTUs() As Integer, ByRef strAuswertungTUs() As String, _
#                                         ByRef strAuswertungTUsFehlendeAAs As String, ByRef strAuswertungTUsFehlendeTFs As String, _
#                                         ByRef intAusgabeAuswertungTUs As Integer, ByRef strAusgabeAuswertungTUs As String, ByRef strAusgabeAuswertungTUsDetails As String)
#
#'1) alle TF operativ und VK:TU = TF:TU
#'2) TF vorhanden, aber Status != operativ oder VK:TU != TF:TU
#'3) keine TF vorhanden
#
#'VKs sind obligatorisch
#If intAuswertungTUs(1) = 0 Then
#    'Keine relevanten VK-TUs ohne TF vorhanden
#    If intAuswertungTUs(11) > 0 Then
#        'Alle relevanten VK-TUs mit TF (operativ) abgedeckt
#        strAusgabeAuswertungTUs = "Alle relevanten Testumgebungstypen abgedeckt"
#        'strAusgabeAuswertungTUsDetails = "Alle relevanten Testumgebungstypen mit operativen Testfällen abgedeckt"
#        strAusgabeAuswertungTUsDetails = "Testfälle vollständig"
#        intAusgabeAuswertungTUs = 1
#    End If
#    If intAuswertungTUs(21) > 0 Then
#        'Alle relevanten VK-TUs mit TF (nicht operativ) abgedeckt
#        strAusgabeAuswertungTUs = "Alle relevanten Testumgebungstypen abgedeckt"
#        'strAusgabeAuswertungTUsDetails = "Alle relevanten Testumgebungstypen mit nicht operativen Testfällen abgedeckt"
#        strAusgabeAuswertungTUsDetails = "Testfälle unvollständig"
#        intAusgabeAuswertungTUs = 2
#    End If
#    If intAuswertungTUs(31) > 0 Then
#        'Alle relevanten VK-TUs mit TF (teilweise operativ) abgedeckt
#        strAusgabeAuswertungTUs = "Alle relevanten Testumgebungstypen abgedeckt"
#        'strAusgabeAuswertungTUsDetails = "Alle relevanten Testumgebungstypen mit operativen und nicht operativen Testfällen abgedeckt"
#        strAusgabeAuswertungTUsDetails = "Testfälle unvollständig"
#        intAusgabeAuswertungTUs = 2
#    End If
#    If intAuswertungTUs(11) = 0 And intAuswertungTUs(21) = 0 And intAuswertungTUs(31) = 0 Then
#        'Keine relevanten VK-TUs vorhanden
#        strAusgabeAuswertungTUs = "Keine relevanten Testumgebungstypen vorhanden"
#        intAusgabeAuswertungTUs = 1
#    End If
#    'Feldfarbe grün
#ElseIf intAuswertungTUs(1) > 0 Then
#    'Relevante VK-TUs ohne TF vorhanden
#    If intAuswertungTUs(11) > 0 Then
#        'Einige relevanten VK-TUs mit TF (operativ) abgedeckt
#        strAusgabeAuswertungTUs = "Relevante Testumgebungstypen teilweise abgedeckt"
#        'strAusgabeAuswertungTUsDetails = "Relevante Testumgebungstypen teilweise mit operativen Testfällen abgedeckt"
#        strAusgabeAuswertungTUsDetails = "Testfälle unvollständig"
#        'Feldfarbe gelb
#        intAusgabeAuswertungTUs = 2
#    End If
#    If intAuswertungTUs(21) > 0 Then
#        'Einige relevanten VK-TUs mit TF (nicht operativ) abgedeckt
#        strAusgabeAuswertungTUs = "Relevante Testumgebungstypen teilweise abgedeckt"
#        'strAusgabeAuswertungTUsDetails = "Relevante Testumgebungstypen teilweise mit nicht operativen Testfällen abgedeckt"
#        strAusgabeAuswertungTUsDetails = "Testfälle unvollständig"
#        'Feldfarbe gelb
#        intAusgabeAuswertungTUs = 2
#    End If
#    If intAuswertungTUs(31) > 0 Then
#        'Einige relevanten VK-TUs mit TF (teilweise operativ) abgedeckt
#        strAusgabeAuswertungTUs = "Relevante Testumgebungstypen teilweise abgedeckt"
#        'strAusgabeAuswertungTUsDetails = "Relevante Testumgebungstypen teilweise mit operativen und nicht operativen Testfällen abgedeckt"
#        strAusgabeAuswertungTUsDetails = "Testfälle unvollständig"
#        'Feldfarbe gelb
#        intAusgabeAuswertungTUs = 2
#    End If
#    If intAuswertungTUs(11) = 0 And intAuswertungTUs(21) = 0 And intAuswertungTUs(31) = 0 Then
#        'Keine relevanten VK-TUs abgedeckt
#        strAusgabeAuswertungTUs = "Relevante Testumgebungstypen nicht abgedeckt"
#        strAusgabeAuswertungTUsDetails = "Keine Testfälle vorhanden"
#        'Feldfarbe rot
#        intAusgabeAuswertungTUs = 3
#    End If
#    'Fehlende TUs bei TFs erfassen
#    strAuswertungTUsFehlendeTFs = strAuswertungTUs(1)
#End If
#
#'Weitere relevante TUs in TFs vorhanden?
#If intAuswertungTUs(10) > 0 Then
#    'TF operativ
#    'strAusgabeAuswertungTUsDetails = strAusgabeAuswertungTUsDetails & vbCrLf & "Weitere operative Testfälle für abweichende Testumgebungstypen vorhanden."
#    strAuswertungTUsFehlendeAAs = strAuswertungTUs(10)
#End If
#If intAuswertungTUs(20) > 0 Then
#    'TF nicht operativ
#    'strAusgabeAuswertungTUsDetails = strAusgabeAuswertungTUsDetails & vbCrLf & "Weitere nicht operative Testfälle für abweichende Testumgebungstypen vorhanden."
#    If strAuswertungTUsFehlendeAAs = "" Then
#        strAuswertungTUsFehlendeAAs = strAuswertungTUs(20)
#    Else
#        strAuswertungTUsFehlendeAAs = strAuswertungTUsFehlendeAAs & ", " & strAuswertungTUs(20)
#    End If
#End If
#If intAuswertungTUs(30) > 0 Then
#    'TF operativ und nicht operativ
#    'strAusgabeAuswertungTUsDetails = strAusgabeAuswertungTUsDetails & vbCrLf & "Weitere operative und nicht operative Testfälle für abweichende Testumgebungstypen vorhanden."
#    If strAuswertungTUsFehlendeAAs = "" Then
#        strAuswertungTUsFehlendeAAs = strAuswertungTUs(30)
#    Else
#        strAuswertungTUsFehlendeAAs = strAuswertungTUsFehlendeAAs & ", " & strAuswertungTUs(30)
#    End If
#    strAuswertungTUsFehlendeAAs = strAuswertungTUs(30)
#End If
#End Sub
#
#Private Function AuswertungUnterschiedlicheIStufen(ByRef IStufen As Collection) As Boolean
#Dim i As Integer
#
#AuswertungUnterschiedlicheIStufen = False
#
#If IStufen.Count > 1 Then
#    For i = 2 To IStufen.Count
#        If IStufen.Item(i) <> IStufen.Item(i - 1) Then
#            AuswertungUnterschiedlicheIStufen = True
#        End If
#    Next i
#End If
#End Function
#
#Private Sub AusgabeVerlauf(ByRef wksStatus As Worksheet, ByRef strFehlerVerlauf As String, ByVal intAuswahl As Integer)
#Dim wksVerlauf As Worksheet
#Dim rngAttributeVerlauf() As Range
#Dim intAttributeZaehler As Integer
#Dim intAttributeStatus As Integer
#Dim blnVerlaufAttribute As Boolean
#Dim lngVerlaufLetzteZeile As Long
#Dim lngStatusLetzteZeile As Long
#Dim intKW As Integer
#Dim rngVerlaufKWVorhanden As Range
#
#'Fehlermeldungen ausschalten
#On Error Resume Next
#
#blnVerlaufAttribute = False
#strFehlerVerlauf = ""
#
#'Arbeitsblatt ATE/TD_Status_Verlauf vorhanden?
#Select Case intAuswahl:
#    Case 1:
#        Set wksVerlauf = ThisWorkbook.Sheets("ATE_Status_Verlauf")
#    Case 2:
#        Set wksVerlauf = ThisWorkbook.Sheets("TD_Status_Verlauf")
#End Select
#
#'Arbeitsblatt ATE/TD_Status_Verlauf vorhanden
#If Not wksVerlauf Is Nothing Then
#    blnVerlaufAttribute = True
#    
#    'Attribute aus ATE/TD_Status in ATE/TD_Status_Verlauf suchen
#    intAttributeStatus = wksStatus.Cells(2, 1).End(xlToRight).Column
#
#    ReDim rngAttributeVerlauf(1 To intAttributeStatus)
#    
#    For intAttributeZaehler = 1 To intAttributeStatus
#        Set rngAttributeVerlauf(intAttributeZaehler) = wksVerlauf.Cells.Find(wksStatus.Cells(2, intAttributeZaehler).Value, lookat:=xlWhole)
#        
#        'Festhalten des Indexes für Attribut "KW Datenauswertung"
#        If wksStatus.Cells(2, intAttributeZaehler).Value = "KW Datenauswertung" Then
#            intKW = intAttributeZaehler
#        End If
#        
#        'Flag setzen, falls nicht alle Attribute aus ATE/TD_Status in ATE/TD_Status_Verlauf gefunden
#        If rngAttributeVerlauf(intAttributeZaehler) Is Nothing Then
#            blnVerlaufAttribute = False
#        End If
#    Next intAttributeZaehler
#    
#    'Prüfung, ob KW der Auswertung bereits im Verlauf vorhanden:
#    'KW-Eintrag vorhanden: Abbruch mit Fehlermeldung
#    'KW-Eintrag nicht vorhanden: Werte übernehmen
#    With wksVerlauf.Columns(rngAttributeVerlauf(intKW).Column)
#        Set rngVerlaufKWVorhanden = .Cells.Find(wksStatus.Cells(3, intKW), lookat:=xlWhole)
#    End With
#    
#    If Not rngVerlaufKWVorhanden Is Nothing Then
#    Select Case intAuswahl:
#    Case 1:
#        strFehlerVerlauf = "Eintrag für Kalenderwoche bereits im ATE_Status_Verlauf vorhanden."
#    Case 2:
#        strFehlerVerlauf = "Eintrag für Kalenderwoche bereits im TD_Status_Verlauf vorhanden."
#    End Select
#    
#    Else
#
#        'Werte von ATE/TD_Status in ATE/TD_Status_Verlauf übernehmen
#        If blnVerlaufAttribute = True Then
#            'Letzte Zeile im ATE/TD_Status_Verlauf ermitteln
#            If wksVerlauf.Cells(2, 1).Value <> "" Then
#                lngVerlaufLetzteZeile = wksVerlauf.Cells(1, 1).End(xlDown).Row
#            Else
#                lngVerlaufLetzteZeile = 1
#            End If
#            
#            'Letzte Zeile im ATE_Status ermitteln
#            If wksStatus.Cells(3, 1).Value <> "" Then
#                lngStatusLetzteZeile = wksStatus.Cells(2, 1).End(xlDown).Row
#            Else
#                lngStatusLetzteZeile = 2
#            End If
#    
#            'Falls Werte vorhanden, Werte spaltenweise übernehmen
#            If lngStatusLetzteZeile > 2 Then
#                For intAttributeZaehler = 1 To intAttributeStatus
#                    wksStatus.Range(wksStatus.Cells(3, intAttributeZaehler), wksStatus.Cells(lngStatusLetzteZeile, intAttributeZaehler)).Copy _
#                    Destination:=wksVerlauf.Range(wksVerlauf.Cells(lngVerlaufLetzteZeile + 1, rngAttributeVerlauf(intAttributeZaehler).Column), wksVerlauf.Cells(lngVerlaufLetzteZeile + lngStatusLetzteZeile - 2, rngAttributeVerlauf(intAttributeZaehler).Column))
#                Next intAttributeZaehler
#            End If
#        End If
#    
#    End If
#
#Else
#    Select Case intAuswahl:
#    Case 1:
#        strFehlerVerlauf = "Arbeitsblatt 'ATE_Status_Verlauf' nicht vorhanden"
#    Case 2:
#        strFehlerVerlauf = "Arbeitsblatt 'TD_Status_Verlauf' nicht vorhanden"
#    End Select
#End If
#
#End Sub
#