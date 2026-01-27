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
            info_AVW:DBInfo,
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
        
        self.ATE_STATus_Initializer = all(self.import_attribute)
#   End Function

    def _status_initializer(self) -> bool:
        return True
    
    def collect_errors(self, error: str, attributes:str):
        if self.errors == "":
            self.errors = f"{error}\n(Benötigt: {self.project} - {attributes})"
        else:        
            self.errors += f"\n\n{error}\n(Benötigt: {self.project} - {attributes})"
    