
AVW_ATTRIBUTE: tuple[str] = (
    "ID",
#       strAVWAttribute(1) = "ID"
    "Dokument-ID",
#       strAVWAttribute(2) = "Dokument-ID"
    "Basis für Testdesign",
#       strAVWAttribute(3) = "Basis für Testdesign"
    "Typ",
#       strAVWAttribute(4) = "Typ"
    "Kategorie",
#       strAVWAttribute(5) = "Kategorie"
    "Status",
#       strAVWAttribute(6) = "Status"
    "Feature",
#       strAVWAttribute(7) = "Feature"
    "Reifegrad",
#       strAVWAttribute(8) = "Reifegrad"
    "Umsetzer",
#       strAVWAttribute(9) = "Umsetzer"
    "ASIL",
#       strAVWAttribute(10) = "ASIL"
    "BSM-SaFuSi Bewertung",
#       strAVWAttribute(11) = "BSM-SaFuSi Bewertung"
    "BSM-ZZ Bewertung",
#       strAVWAttribute(12) = "BSM-ZZ Bewertung"
    "BSM-ED Bewertung",
#       strAVWAttribute(13) = "BSM-ED Bewertung"
    "BSM-FFF Bewertung",
#       strAVWAttribute(14) = "BSM-FFF Bewertung"
    "BSM-O Bewertung",
#       strAVWAttribute(15) = "BSM-O Bewertung"
    "BSM-Se Bewertung",
#       strAVWAttribute(16) = "BSM-Se Bewertung"
    "MV",
#       strAVWAttribute(17) = "MV"
    "Cluster Testing",
#       strAVWAttribute(18) = "Cluster Testing"
    "Dokument",
#       strAVWAttribute(19) = "Dokument"
    "Kommentar Redaktionskreis",
#       strAVWAttribute(20) = "Kommentar Redaktionskreis"
    "temp1_Text",
#       strAVWAttribute(21) = "temp1_Text"
    "Anforderungsverantwortliche",
#       strAVWAttribute(22) = "Anforderungsverantwortliche"
    "Abgezweigt aus",
#       If blnAVWVorgaengerIDsVerwenden Then
#           strAVWAttribute(23) = "Abgezweigt aus"  'strAVWAttribute(22) = "Abgezweigt aus"
#       End If
)

OUTPUT_BSM_ATTRIBUTE: tuple[str] = (
    'Abgezweigt aus',
    'ID',
    'Dokument-ID',
    'BsM-Relevanz',
    'BSM-SaFuSi Bewertung',
    'BSM-ZZ Bewertung',
    'BSM-ED Bewertung',
    'BSM-FFF Bewertung',
    'BSM-O Bewertung',
    'BSM-Se Bewertung',
    'ASIL',
    'Feature',
    'Reifegrad',
    'Umsetzer',
    'Status',
    'TD-VK',
    'TD-AA',
    'TD-TI:TU',
    'Testfälle',
    'Vergleich TUs (TD-TF) - operativ',
    'MV',
    'Kategorie',
    'Dokument',
    '#abgelehnt_nicht_testbar',
    'Zugeordnete I-Stufe',
    'Status TD-VK',
    'Fehlende TUs bei TD-AAs',
    'Fehlende TUs bei TFs',
    'Erläuterungen zum Vergleich',
    'Cluster Testing',
    'Projekt',
    'TD-VK temp1_Text',
    'TD-VK Effort Estimation',
    'Anforderungsverantwortlicher',
    'KW Datenauswertung',
    'Temp11_Auswahlfeld',
)
        # From ouput_status() in tracking.py
#       'Bekannte Testumgebungen
        #known test environments (EN)
#       ReDim strBekannteTUs(1 To 17)
#       intRelevantekTUs = 9
RELEVANT_TEST_ENVIRONMENT_TOP = 9
KNOWN_TEST_ENVIRONMENTS: tuple[str] = (
#       strBekannteTUs(1) = "BRS-HiL_Laborplatz_automatisiert"
    'BRS-HiL_Laborplatz_automatisiert',
#       strBekannteTUs(3) = "BRS-HiL_Kunden-Funktion"
    'BRS-HiL_Kunden-Funktion',
#       strBekannteTUs(4) = "BRS-HiL_Bremssystem"
    'BRS-HiL_Bremssystem',
#       strBekannteTUs(5) = "BRS-Fahrversuch_Kunden-Funktion"
    'BRS-Fahrversuch_Kunden-Funktion',
#       strBekannteTUs(6) = "BRS-Fahrversuch_Basis-Funktion"
    'BRS-Fahrversuch_Basis-Funktion',
#       strBekannteTUs(7) = "Vernetzter-Fahrwerks-HiL_Kundenfunktion"
    'Vernetzter-Fahrwerks-HiL_Kundenfunktion', 
#       strBekannteTUs(8) = "BRS-HiL_Basisdienst_Halten"
    'BRS-HiL_Basisdienst_Halten',
#       strBekannteTUs(9) = "BRS-HiL_Basisdienst_Verzoegern"
    'BRS-HiL_Basisdienst_Verzoegern',
#       'ab hier nicht mehr relevant
#       strBekannteTUs(10) = "BRS-SiL_Kunden-Funktion"
    'BRS-SiL_Kunden-Funktion',
#       strBekannteTUs(11) = "Code-Review"
    'Code-Review',
#       strBekannteTUs(12) = "Design-Review"
    'Design-Review',
#       strBekannteTUs(13) = "Dokumenten-Review"
    'Dokumenten-Review',
#       strBekannteTUs(14) = "Prozess-Review"
    'Prozess-Review',
#       strBekannteTUs(15) = "Entscheidung_liegt_bei_Testinstanz"
    'Entscheidung_liegt_bei_Testinstanz',
#       strBekannteTUs(16) = "BRS-Fahrversuch_Applikation"
    'BRS-Fahrversuch_Applikation',
#       strBekannteTUs(17) = "BRS-Fahrversuch_Erprobung"
    'BRS-Fahrversuch_Erprobung',
)

TDAA_ATTRIBUTE = (
#   strTDAAAttribute(1) = "ID"
    "ID",
#   strTDAAAttribute(2) = "Enthalten in"
    "Enthalten in",
#   strTDAAAttribute(3) = "Status"
    "Status",
#   strTDAAAttribute(4) = "Testinstanz"
    "Testinstanz",
#   strTDAAAttribute(5) = "Testumgebungstyp"
    "Testumgebungstyp" ,
)

MASTER_ATTRIBUTE = (
#   strAVWMasterAttribute(1) = "ID"
    "ID",
#   strAVWMasterAttribute(2) = "temp1_Text"
    "temp1_Text",
#   strAVWMasterAttribute(3) = "Kommentar Redaktionskreis"
    "Kommentar Redaktionskreis",
)

TDVK_ATTRIBUTE = (
#   strTDVKAttribute(1) = "ID"
    "ID",
#   strTDVKAttribute(2) = "Basierend auf der Anforderung"
    "Basierend auf der Anforderung",
#   strTDVKAttribute(3) = "Status"
    "Status",
#   strTDVKAttribute(4) = "Temp1_Text"
    "Temp1_Text",
#   strTDVKAttribute(5) = "Aktion"
    "Aktion",
)

TF_ATTRIBUTE = (
#   strTFAttribute(1) = "ID"
    "ID",
#   strTFAttribute(2) = "Status"
    "Status",
#   strTFAttribute(3) = "Testfallname"
    "Testfallname",
#   strTFAttribute(4) = "Sonstige-Varianten"
    "Sonstige-Varianten",
#   strTFAttribute(5) = "Basierend auf Testdesign"
    "Basierend auf Testdesign",
#   strTFAttribute(6) = "verifiziert"
    "verifiziert",
#   strTFAttribute(7) = "Testinstanz"
    "Testinstanz",
)

FRU_TIMING_ATTRIBUTE = (
#   strFRUTimingAttribute(1) = "FeatureName"
    "FeatureName",
#   strFRUTimingAttribute(2) = "Reifegrad"  'vorher "RG"
    "Reifegrad",
#   strFRUTimingAttribute(3) = "Umsetzer"
    "Umsetzer",
#   strFRUTimingAttribute(4) = "FE_Meilenstein" 'vorher "Zuordnung zu I-Stufe"
    "FE_Meilenstein",
)
