
from enum import StrEnum


class RequirementAttribute(StrEnum):
    ID = "ID",
#       strAVWAttribute(1) = "ID"
    DocumentID = "Dokument-ID",
#       strAVWAttribute(2) = "Dokument-ID"
    TestDesignBase="Basis für Testdesign",
#       strAVWAttribute(3) = "Basis für Testdesign"
    Type = "Typ",
#       strAVWAttribute(4) = "Typ"
    Category = "Kategorie",
#       strAVWAttribute(5) = "Kategorie"
    Status = "Status",
#       strAVWAttribute(6) = "Status"
    Feature ="Feature",
#       strAVWAttribute(7) = "Feature"
    MaturityLevel ="Reifegrad",
#       strAVWAttribute(8) = "Reifegrad"
    Implementer = "Umsetzer",
#       strAVWAttribute(9) = "Umsetzer"
    ASIL = "ASIL",
#       strAVWAttribute(10) = "ASIL"
    BSMSaFuSiAssesment ="BSM-SaFuSi Bewertung",
#       strAVWAttribute(11) = "BSM-SaFuSi Bewertung"
    BSMZZAssesment = "BSM-ZZ Bewertung",
#       strAVWAttribute(12) = "BSM-ZZ Bewertung"
    BSMEDAssesment = "BSM-ED Bewertung",
#       strAVWAttribute(13) = "BSM-ED Bewertung"
    BSMFFFAssesment = "BSM-FFF Bewertung",
#       strAVWAttribute(14) = "BSM-FFF Bewertung"
    BSMOAssesment ="BSM-O Bewertung",
#       strAVWAttribute(15) = "BSM-O Bewertung"
    BSMSeAssesment = "BSM-Se Bewertung",
#       strAVWAttribute(16) = "BSM-Se Bewertung"
    MV = "MV",
#       strAVWAttribute(17) = "MV"
    TestingCluster = "Cluster Testing",
#       strAVWAttribute(18) = "Cluster Testing"
    Document = "Dokument",
#       strAVWAttribute(19) = "Dokument"
    EditorialTeamComent = "Kommentar Redaktionskreis",
#       strAVWAttribute(20) = "Kommentar Redaktionskreis"
####    Temp1_Text = "temp1_Text", ## Diff in file MEBwq_Statistik_Testing.xlsx
    Temp1_Text = 'Temp1_text',
#       strAVWAttribute(21) = "temp1_Text"
    RequirementOwners = "Anforderungsverantwortliche",
#       strAVWAttribute(22) = "Anforderungsverantwortliche"
    RedirectedFrom = "Abgezweigt aus",
#       If blnAVWVorgaengerIDsVerwenden Then
#           strAVWAttribute(23) = "Abgezweigt aus"  'strAVWAttribute(22) = "Abgezweigt aus"
#       End If

class AVWProjectAttribute(StrEnum):
    Temp11SelectionField ='Temp11_Auswahlfeld'

class TDProjectAttribute(StrEnum):
    Temp11SelectionField ='Temp11_Auswahlfeld (LAH)'

class OutputBSMAttribute(StrEnum):
    RedirectedFrom = 'Abgezweigt aus'   # Use predecessors ids attribute
    ID = 'ID'
    DocumentID = 'Dokument-ID'
    BSMRelevance = 'BsM-Relevanz'
    BSMSaFuSiAssesment ='BSM-SaFuSi Bewertung'
    BSMZZAssesment = 'BSM-ZZ Bewertung'
    BSMEDAssesment = 'BSM-ED Bewertung'
    BSMFFFAssesment = 'BSM-FFF Bewertung'
    BSMOAssesment = 'BSM-O Bewertung'
    BSMSeAssesment = 'BSM-Se Bewertung'
    ASIL = 'ASIL'
    Feature = 'Feature'
    MaturityLevel = 'Reifegrad'
    Implementer = 'Umsetzer'
    Status = 'Status'
    TDVC = 'TD-VK'
    TDSafeguards = 'TD-AA'
    TDTITE = 'TD-TI:TU'
    Temp11SelectionField ='Temp11_Auswahlfeld' # Project specific attribute
    TestCase = 'Testfälle'
    OperationalComparisonTEsTDTC = 'Vergleich TUs (TD-TF) - operativ'
    MV = 'MV'
    Category = 'Kategorie'
    Document = 'Dokument'
    RejectedNotTestable = '#abgelehnt_nicht_testbar'
    AssignedILevel = 'Zugeordnete I-Stufe'
    StatusTDVC = 'Status TD-VK'
    MissingTEInTDSafeguards = 'Fehlende TUs bei TD-AAs'
    MissingTEInTCs = 'Fehlende TUs bei TFs'
    ComparisonExplanations ='Erläuterungen zum Vergleich'
    TestingCluster = 'Cluster Testing'
    Project = 'Projekt'
    TDVCTemp1Text ='TD-VK temp1_Text'
    TDVCEffortEstimation ='TD-VK Effort Estimation'
    RequirementOwner = 'Anforderungsverantwortlicher'
    KWDataEvaluation = 'KW Datenauswertung'

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

class TDSafeGuardsAttribute(StrEnum):
#   strTDAAAttribute(1) = "ID"
    ID = "ID",
#   strTDAAAttribute(2) = "Enthalten in"
    IncludedIn = "Enthalten in",
#   strTDAAAttribute(3) = "Status"
    Status = "Status",
#   strTDAAAttribute(4) = "Testinstanz"
    TestInstance = "Testinstanz",
#   strTDAAAttribute(5) = "Testumgebungstyp"
    TestEnvironmentType = "Testumgebungstyp" 

class RequirementMasterAttribute(StrEnum):
#   strAVWMasterAttribute(1) = "ID"
    ID = "ID"
#   strAVWMasterAttribute(2) = "temp1_Text"
    Temp1Text = "temp1_Text"
#   strAVWMasterAttribute(3) = "Kommentar Redaktionskreis"
    EditorialTeamComent = "Kommentar Redaktionskreis"

class TDVCAttribute(StrEnum):
    ID = 'ID'
    RequirementBased = 'Basierend auf der Anforderung'
    Status = 'Status'
    Temp1Text = 'Temp1_text'
    Action = 'Aktion'

class TestCaseAttribute(StrEnum):
#   strTFAttribute(1) = "ID"
    ID = 'ID'
#   strTFAttribute(2) = "Status"
    Status = 'Status'
#   strTFAttribute(3) = "Testfallname"
    TestCaseName = 'Testfallname'
#   strTFAttribute(4) = "Sonstige-Varianten"
    OtherVariants = 'Sonstige-Varianten'
#   strTFAttribute(5) = "Basierend auf Testdesign"
    TestDesignBased = 'Basierend auf Testdesign'
#   strTFAttribute(6) = "verifiziert"
    Verified = 'verifiziert'
#   strTFAttribute(7) = "Testinstanz"
    TestInstance = 'Testinstanz'

class FRUTimingAttribute(StrEnum):
#   strFRUTimingAttribute(1) = "FeatureName"
    FeatureName = "FeatureName"
#   strFRUTimingAttribute(2) = "Reifegrad"  'vorher "RG"
    MaturityLevel = "Reifegrad"
#   strFRUTimingAttribute(3) = "Umsetzer"
    Implementer = "Umsetzer"
#   strFRUTimingAttribute(4) = "FE_Meilenstein" 'vorher "Zuordnung zu I-Stufe"
    FEMilestone = "FE_Meilenstein"

class TDAttribute(StrEnum):
#       strTDAttribute(1) = "TD-VK"
    TDVC ='TD-VK'
#       strTDAttribute(2) = "Status TD-VK"
    StatusTDVC ='Status TD-VK'
#       strTDAttribute(3) = "TD-AA"
    TDSafeGuards ='TD-AA'
#       strTDAttribute(4) = "TD-TI:TU"
    TDTITE ='TD-TI:TU'
#       strTDAttribute(5) = "Testfälle"
    TestCases ='Testfälle'
#       strTDAttribute(6) = "Vergleich TUs (TD-TF) - operativ"
    OperativeTEComparisonTDTC ='Vergleich TUs (TD-TF) - operativ'
#       strTDAttribute(7) = "Erläuterungen zum Vergleich"
    ComparisonExplanations ='Erläuterungen zum Vergleich'
#       strTDAttribute(8) = "Fehlende TUs bei TD-AAs"
    MissingTEsInTDSafeGuards ='Fehlende TUs bei TD-AAs'
#       strTDAttribute(9) = "Fehlende TUs bei TFs"
    MissingTEsInTCs ='Fehlende TUs bei TFs'
#       strTDAttribute(10) = "Anforderungs-IDs"
    RequirementIDs ='Anforderungs-IDs'
#       strTDAttribute(11) = "Zugeordnete I-Stufe"
    AsignedILevel ='Zugeordnete I-Stufe'
#       strTDAttribute(12) = "Umsetzer (LAH)"
    LAH_Implementer ='Umsetzer (LAH)'
#       strTDAttribute(13) = "BsM-Relevanz (LAH)"
    LAH_BsMRelevance ='BsM-Relevanz (LAH)'
    Temp11SelectionField ='Temp11_Auswahlfeld (LAH)' # Project Specific Attribute
#       strTDAttribute(14) = "ASIL (LAH)"
    LAH_ASIL ='ASIL (LAH)'
#       strTDAttribute(15) = "Feature (LAH)"
    LAH_Feature ='Feature (LAH)'
#       strTDAttribute(16) = "Reifegrad (LAH)"
    LAH_MaturityLevel ='Reifegrad (LAH)'
#       strTDAttribute(17) = "MV (LAH)"
    LAH_MV ='MV (LAH)'
#       strTDAttribute(18) = "LAH-ID"
    LAH_ID ='LAH-ID'
#       strTDAttribute(19) = "Dokumente (LAH)"
    LAH_Document ='Dokumente (LAH)'
#       strTDAttribute(20) = "Cluster Testing"
    TestingCluster ='Cluster Testing'
#       strTDAttribute(21) = "Projekt"
    Project ='Projekt'
#       strTDAttribute(22) = "TD-VK temp1_Text"
    TDVCTemp1Text ='TD-VK temp1_Text'
#       strTDAttribute(23) = "TD-VK Effort Estimation"
    TDVCEffortEstimation ='TD-VK Effort Estimation'
#       strTDAttribute(24) = "Anforderungsverantwortliche (LAH)"
    LAH_RequirementOwner ='Anforderungsverantwortliche (LAH)'
#       strTDAttribute(25) = "KW Datenauswertung"
    KWDataAnalysis ='KW Datenauswertung'
