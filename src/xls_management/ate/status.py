from xls_management.ate.project import project_combo_box
from importlib.metadata import version, PackageNotFoundError

def status():
    """
    'BsM_Status
    Dim wbBsM As Workbook                       'Workbook für BsM_Status
    Dim wksBsM As Worksheet                     'Worksheet für BsM_Status
    Dim strBsMAttribute() As String             'String-Array mit Attributen des Arbeitsblatts BsM_Status
    Dim rngBsMAttribute() As Range              'Range-Array mit Attributen des Arbeitsblatts BsM_Status
    'TD_Status
    Dim wbTD As Workbook                        'Workbook für TD_Status
    Dim wksTD As Worksheet                      'Worksheet für TD_Status
    Dim strTDAttribute() As String              'String-Array mit Attributen des Arbeitsblatts TD_Status
    Dim rngTDAttribute() As Range               'Range-Array mit Attributen des Arbeitsblatts TD_Status
    'AVW_Rohdaten - Projekt
    Dim wbAVW As Workbook                       'Workbook für AVW_Rohdaten
    Dim wksAVW As Worksheet                     'Worksheet für AVW_Rohdaten
    Dim strAVWAttribute() As String             'String-Array mit Attributen des Arbeitsblatts AVW_Rohdaten
    Dim rngAVWAttribute() As Range              'Range-Array mit Attributen des Arbeitsblatts AVW_Rohdaten
    Dim strAVWAttributeMEB21() As String        'String-Array mit Attributen des Arbeitsblatts AVW_Rohdaten für MEB21
    Dim rngAVWAttributeMEB21() As Range         'Range-Array mit Attributen des Arbeitsblatts AVW_Rohdaten für MEB21
    'AVW_Rohdaten - Master
    Dim wbAVWMaster As Workbook                 'Workbook für AVWMaster_Rohdaten
    Dim wksAVWMaster As Worksheet               'Worksheet für AVWMaster_Rohdaten
    Dim strAVWMasterAttribute() As String       'String-Array mit Attributen des Arbeitsblatts AVWMaster_Rohdaten
    Dim rngAVWMasterAttribute() As Range        'Range-Array mit Attributen des Arbeitsblatts AVWMaster_Rohdaten
    'TDVKs (Tesdesigns - Verifikationskriterium)
    Dim wbTDVK As Workbook                      'Workbook für TDs - Verifikationskriterien
    Dim wksTDVK As Worksheet                    'Worksheet für TDs - Verifikationskriterium
    Dim strTDVKAttribute() As String            'String-Array mit Attributen des Arbeitsblatts TDs - Verifikationskriterium
    Dim rngTDVKAttribute() As Range             'Range-Array mit Attributen des Arbeitsblatts TDs - Verifikationskriterium
    'TDAAs (Tesdesigns - Absicherungsaufträge)
    Dim wbTDAA As Workbook                      'Workbook für TDs - Absicherungsaufträge
    Dim wksTDAA As Worksheet                    'Worksheet für TDs - Absicherungsaufträge
    Dim strTDAAAttribute() As String            'String-Array mit Attributen des Arbeitsblatts TDs - Absicherungsaufträge
    Dim rngTDAAAttribute() As Range             'Range-Array mit Attributen des Arbeitsblatts TDs - Absicherungsaufträge
    'TF (Testfälle)
    Dim wbTF As Workbook                        'Workbook für Testfälle
    Dim wksTF As Worksheet                      'Worksheet für Testfälle
    Dim strTFAttribute() As String              'String-Array mit Attributen des Arbeitsblatts Testfälle
    Dim rngTFAttribute() As Range               'Range-Array mit Attributen des Arbeitsblatts Testfälle
    'FRUTiming
    Dim wbFRUTiming As Workbook                 'Workbook für FRU-Timing
    Dim wksFRUTiming As Worksheet               'Worksheet für FRU-Timing
    Dim strFRUTimingAttribute() As String       'String-Array mit Attributen des Arbeitsblattes für FRU-Timing
    Dim rngFRUTimingAttribute() As Range        'Range-Array mit Attributen des Arbeitsblattes für FRU-Timing
    'Allgemein
    Dim strFehlerGesamt As String               'String für Gesamtfehlerausgabe
    Dim strWeitereTUsAusgabe As String          'String für Ausgabe weiterer Testumgebungstypen
    Dim strLAHBlacklist() As String             'String-Array für einzulesende LAH-Blacklist
    Dim strVersionMakro As String               'String-Array für Makro-Version
    Dim strDateinamen(1 To 6) As String         'String-Array für die Namen der eingelesenen Dateien
    Dim strProjekte(0 To 5) As String           'String-Array für die auswählbaren Fahrzeugprojekte
    Dim strProjekt As String                    'String des ausgewählten Fahrzeugprojekts
    'Verlauf
    Dim strFehlerATEVerlauf As String           'String für Rückgabewert der Befüllung von ATE_Status_Verlauf
    Dim strFehlerTDVerlauf As String            'String für Rückgabewert der Befüllung von TD_Status_Verlauf
    Dim strFehlerVerlauf As String              'String für gemeinsamen Rückgabewert der Befüllung von ATE/TD_Status_Verlauf

    """

    #'Makro-Version
    #strVersionMakro = "ATE-Status V015F6" & vbCrLf & "Programmiert von Alexander Kuhlicke, Tagueri AG 2024"
    #
    str_version_makro = f"ATE-Status {version('xls_management')}"
    #'BsM-Status wird separat im Workbook des Makros erzeugt
    #Set wbBsM = ThisWorkbook
    #TODO
    
    str_project, evalue_master_id = project_combo_box()
    #
    #If boolAuswahlGetroffen Then
    #    'Eingabe Projekt
    #    strProjekt = BoxAuswahlProjekt.ComboBox1.list(BoxAuswahlProjekt.ComboBox1.ListIndex)
    #    'Eingabe Verwendung Master-IDs
    #    blnAVWVorgaengerIDsVerwenden = BoxAuswahlProjekt.OptionButton1.Value
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


def box_auswahl_projekt_show() -> str:
    #'Befüllung der Projektliste
    #strProjekte(0) = "empty",
    #strProjekte(1) = "More Queries Broad48W",
    #strProjekte(2) = "More Queries Broad37W PA",
    #strProjekte(3) = "More Examples Broad UNECE",
    #strProjekte(4) = "More Examples Broad21",
    #strProjekte(5) = "Other"
    str_project:tuple[str] = (p for p in PROJECTS)
    #BoxAuswahlProjekt.ComboBox1.list() = strProjekte
    #BoxAuswahlProjekt.ComboBox1.ListIndex = 0
    #TODO

    #'Abfrage Projekt und Nutzung Master-Bereich
    #BoxAuswahlProjekt.Caption = "ATE-Status " & strVersionMakro
    #BoxAuswahlProjekt.Show
    return str_project[1]