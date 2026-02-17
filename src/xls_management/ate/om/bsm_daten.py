import re
import pandas as pd

from xls_management.ate.om.fru_timming import FRUTiming
from xls_management.ate.om.test_case import TestCase
from xls_management.ate.om.verificationskriterium import Verificationskriterium


class BSMDaten:
#Option Explicit
#

    def __init__(
        self,
        columns:pd.DataFrame,
        row:int,
        fru_timing_index: dict[str,FRUTiming],
        is_specific: bool = False,
    ):
        self.bsm_available:str = 'ja'
        self.fru_timing_index = fru_timing_index
#       'Neues Verifikationskriterium anlegen
#       Set BSMDatensatz.Verifikationskriterium = New Collection
        self.verifications_criteria: list[Verificationskriterium] = []
#       'Neuen Testfall anlegen
#       Set BSMDatensatz.Testfaelle = New Collection
        self.test_cases: list[TestCase] = []
#       'Feature einlesen
#       BSMDatensatz.AVWFeature = CStr(rngAVWAttribute(7).Offset(lngZeile, 0).Value)
        self.avw_feature = str(columns['Feature'][row])
#       'Reifegrad einlesen
#       BSMDatensatz.AVWReifegrad = CStr(rngAVWAttribute(8).Offset(lngZeile, 0).Value)
        self.avw_reifegrad = str(columns['Reifegrad'][row])
#       'Umsetzer einlesen
#       BSMDatensatz.AVWUmsetzer = CStr(rngAVWAttribute(9).Offset(lngZeile, 0).Value)
        self.avw_converter = str(columns['Umsetzer'][row])
#       'Dokument-ID einlesen, Entfernung der zusätzlichen Zeichen "?" und "r"
#       BSMDatensatz.AVWDokumentID = Replace(Replace(CStr(rngAVWAttribute(2).Offset(lngZeile, 0).Value), "?", ""), "r", "")
        self.avw_dokument_id = re.sub('[\?r]','',str(columns['DokumentID'][row]))
#       'Dokument-Name einlesen
#       BSMDatensatz.AVWDokumentName = CStr(rngAVWAttribute(19).Offset(lngZeile, 0).Value)
        self.avw_dokument_name = str(columns['Dokument'][row])
#       'Modulverantwortlichen einlesen
#       BSMDatensatz.AVWMV = CStr(rngAVWAttribute(17).Offset(lngZeile, 0).Value)
        self.avw_mv = str(columns['MV'][row])
#       'Anforderungs-ID einlesen, Entfernung der zusätzlichen Zeichen "?" und "r"
#       BSMDatensatz.AVWID = Replace(Replace(CStr(rngAVWAttribute(1).Offset(lngZeile, 0).Value), "?", ""), "r", "")
        self.avw_id = re.sub('[\?r]','',str(columns['ID'][row]))
#       'Status einlesen
#       BSMDatensatz.AVWStatus = CStr(rngAVWAttribute(6).Offset(lngZeile, 0).Value)
        self.avw_status = str(columns['Status'][row])
#       'Typ einlesen
#       BSMDatensatz.AVWTyp = CStr(rngAVWAttribute(4).Offset(lngZeile, 0).Value)
        self.avw_typ = str(columns['Typ'][row])
#       'Kategorie einlesen
#       BSMDatensatz.AVWKategorie = CStr(rngAVWAttribute(5).Offset(lngZeile, 0).Value)
        self.avw_kategorie = str(columns['Kategorie'][row])
#       'BsM-Status einlesen
#       BSMDatensatz.AVWBsMSaFuSi = CStr(rngAVWAttribute(11).Offset(lngZeile, 0).Value)
        self.avw_bsm_safusi = str(columns['BSM-SaFuSi Bewertung'][row])
#       BSMDatensatz.AVWBsMZZ = CStr(rngAVWAttribute(12).Offset(lngZeile, 0).Value)
        self.avw_bsm_zz = str(columns['BSM-ZZ Bewertung'][row])
#       BSMDatensatz.AVWBsMED = CStr(rngAVWAttribute(13).Offset(lngZeile, 0).Value)
        self.avw_bsm_ed = str(columns['BSM-ED Bewertung'][row])
#       BSMDatensatz.AVWBsMFFF = CStr(rngAVWAttribute(14).Offset(lngZeile, 0).Value)
        self.avw_bsm_fff = str(columns['BSM-FFF Bewertung'][row])
#       BSMDatensatz.AVWBsMO = CStr(rngAVWAttribute(15).Offset(lngZeile, 0).Value)
        self.avw_bsm_o = str(columns['BSM-O Bewertung'][row])   
#       BSMDatensatz.AVWBsMSe = CStr(rngAVWAttribute(16).Offset(lngZeile, 0).Value)
        self.avw_bsm_se = str(columns['BSM-Se Bewertung'][row])
        self.set_relevance()
        ### refactored as BSMDaten.set_relevance()
#       'ASIL einlesen
#       BSMDatensatz.AVWASIL = CStr(rngAVWAttribute(10).Offset(lngZeile, 0).Value)
        self.avw_asil = str(columns['ASIL'][row])
#       'Kommentar Redaktionskreis und temp1_Text einlesen
#       If InStr(CStr(UCase(rngAVWAttribute(20).Offset(lngZeile, 0).Value)), "#ABGELEHNT_NICHT_TESTBAR") > 0 Or InStr(CStr(UCase(rngAVWAttribute(21).Offset(lngZeile, 0).Value)), "#ABGELEHNT_NICHT_TESTBAR") > 0 Then
        if re.search('#abgelehnt_nicht_testbar', str(columns['Kommentar Redaktionskreis'][row]), re.IGNORECASE) or re.search('#abgelehnt_nicht_testbar', str(columns['Temp1_Text'][row]), re.IGNORECASE):
#           BSMDatensatz.AVWAbgelehntNichtTestbar = "x"
            self.avw_abgelehnt_nicht_testbar = 'x'
#       End If
        self.set_i_stufe()
#       'Cluster Testing einlesen
#       BSMDatensatz.ClusterTesting = CStr(rngAVWAttribute(18).Offset(lngZeile, 0).Value)
        self.cluster_testing = str(columns['Cluster Testing'][row])
#       
#       'Anforderungsverantwortliche einlesen
#       BSMDatensatz.AVWAnforderungsverantwortliche = CStr(rngAVWAttribute(22).Offset(lngZeile, 0).Value)
        self.avw_anforderungsverantwortliche = str(columns['Anforderungsverantwortliche'][row])
#       
#       'Projekt MEB21 - Temp11_Auswahlfeld einlesen
#       If strProjekt = "MEB21" Or strProjekt = "MQB48W" Then
        if is_specific:
#           BSMDatensatz.AVWTemp11_Auswahlfeld = CStr(rngAVWAttributeMEB21(1).Offset(lngZeile, 0).Value)
            self.avw_temp11_auswahlfeld = str(columns['Temp11_Auswahlfeld'][row])
#       End If
    
    def set_relevance(self):
#       'Zusammenführung BsM-Relevanz
#       strBsMRelevanz = ""
        relevance = []
#       If CStr(BSMDatensatz.AVWBsMSaFuSi) = strBsMVorhanden Then
#           If strBsMRelevanz = "" Then strBsMRelevanz = "BsM-SaFuSi" Else strBsMRelevanz = strBsMRelevanz & ",BsM-SaFuSi"
        if self.avw_bsm_safusi == self.bsm_available:
#           strBsMRelevanz = "BsM-SaFuSi"
            relevance.append('BsM-SaFuSi')
#       End If
#       If CStr(BSMDatensatz.AVWBsMZZ) = strBsMVorhanden Then
        if self.avw_bsm_zz == self.bsm_available:
#           If strBsMRelevanz = "" Then strBsMRelevanz = "BsM-ZZ" Else strBsMRelevanz = strBsMRelevanz & ",BsM-ZZ"
            relevance.append('BsM-ZZ')
#       End If
#       If CStr(BSMDatensatz.AVWBsMED) = strBsMVorhanden Then
        if self.avw_bsm_ed == self.bsm_available:
#           If strBsMRelevanz = "" Then strBsMRelevanz = "BsM-ED" Else strBsMRelevanz = strBsMRelevanz & ",BsM-ED"
            relevance.append('BsM-ED')
#       End If
#       If CStr(BSMDatensatz.AVWBsMFFF) = strBsMVorhanden Then
        if self.avw_bsm_fff == self.bsm_available:
#           If strBsMRelevanz = "" Then strBsMRelevanz = "BsM-FFF" Else strBsMRelevanz = strBsMRelevanz & ",BsM-FFF"
            relevance.append('BsM-FFF')
#       End If
#       If CStr(BSMDatensatz.AVWBsMO) = strBsMVorhanden Then
        if self.avw_bsm_o == self.bsm_available:
#           If strBsMRelevanz = "" Then strBsMRelevanz = "BsM-O" Else strBsMRelevanz = strBsMRelevanz & ",BsM-O"
            relevance.append('BsM-O')
#       End If
#       If CStr(BSMDatensatz.AVWBsMSe) = strBsMVorhanden Then
        if self.avw_bsm_se == self.bsm_available:
#           If strBsMRelevanz = "" Then strBsMRelevanz = "BsM-Se" Else strBsMRelevanz = strBsMRelevanz & ",BsM-Se"
            relevance.append('BsM-Se')
#       End If
#       BSMDatensatz.BSMRelevanz = strBsMRelevanz
        self.bsm_relevanz = ','.join(relevance)

    def set_i_stufe(self):
#       'Geplante I-Stufe einlesen
#       strIStufe = ""
        i_stufe = ''
#       strIStufeMin = ""
        min_i_stufe = ''
        if len(self.fru_timing_index) > 0:
#           If BSMDatensatz.AVWUmsetzer <> "" Then
            if self.avw_converter != '':
#               varUmsetzer = Split(BSMDatensatz.AVWUmsetzer, ",", , vbBinaryCompare)
                converter_list = [converter.strip() for converter in self.avw_converter.split(',')]     
#               For intUmsetzer = 0 To UBound(varUmsetzer, 1)
                for converter in converter_list:
#                   strIStufe = FRUTimingList.Item(BSMDatensatz.AVWFeature & BSMDatensatz.AVWReifegrad & Trim(varUmsetzer(intUmsetzer))).IStufe
                    fru_key = f'{self.avw_feature}{self.avw_reifegrad}{converter}'
                    try:
                        i_stufe = self.fru_timing_index[fru_key].i_stufe
                    except KeyError:
                        pass
#                   If strIStufeMin = "" Then
#                       strIStufeMin = strIStufe
#                   ElseIf InStr(strIStufe, "IS") > 0 Then
#                       If strIStufe < strIStufeMin Then
#                           strIStufeMin = strIStufe
#                       End If
#                   End If
                    if min_i_stufe == '' or ('IS' in i_stufe and i_stufe < min_i_stufe):
                        min_i_stufe = i_stufe
#               Next intUmsetzer
#           End If
#       BSMDatensatz.IStufe = strIStufeMin
        self.i_stufe = min_i_stufe
#

    def add_verification_criterion(self, verification_criterion:Verificationskriterium):
#       'VK-ID aufnehmen
#       BSMDatensatz.Verifikationskriterium.Add Item:=vabsm_dataset.add_verification_criterion(verification_criterion) 
        if verification_criterion not in self.verifications_criteria:
            self.verifications_criteria.append(verification_criterion)
#       'Geplante I-Stufe für Verifikationskritierum erfassen
#       If BSMDatensatz.IStufe <> "" Then
        if self.i_stufe != '' and self.i_stufe not in verification_criterion.anf_i_stufen:
#           varErfassteVKItem.anf_IStufen.Add Item:=BSMDatensatz.IStufe
            verification_criterion.anf_i_stufen.append(self.i_stufe)
#       End If
#       'Umsetzer für Verifikationskritierum erfassen
#       If BSMDatensatz.AVWUmsetzer <> "" Then
        if self.avw_converter != '' and self.avw_converter not in verification_criterion.anf_umsetzer:
#           varErfassteVKItem.anf_Umsetzer.Add Item:=BSMDatensatz.AVWUmsetzer
            verification_criterion.anf_umsetzer.append(self.avw_converter)
#       End If
#       'BsM-Relevanz für Verifikationskritierum erfassen
#       If BSMDatensatz.BSMRelevanz <> "" Then
        if self.bsm_relevanz != '' and self.bsm_relevanz not in verification_criterion.anf_bsm_relevanz:    
#           varErfassteVKItem.anf_BsMRelevanz.Add Item:=BSMDatensatz.BSMRelevanz
            verification_criterion.anf_bsm_relevanz.append(self.bsm_relevanz)
#       End If
#       'ASIL für Verifikationskritierum erfassen
#       If BSMDatensatz.AVWASIL <> "" Then
        if self.avw_asil != '' and self.avw_asil not in verification_criterion.anf_asil:
#           varErfassteVKItem.anf_ASIL.Add Item:=BSMDatensatz.AVWASIL
            verification_criterion.anf_asil.append(self.avw_asil)
#       End If
#       'Feature für Verifikationskritierum erfassen
#       If BSMDatensatz.AVWFeature <> "" Then
        if self.avw_feature != '' and self.avw_feature not in verification_criterion.anf_feature:
#           varErfassteVKItem.anf_Feature.Add Item:=BSMDatensatz.AVWFeature
            verification_criterion.anf_feature.append(self.avw_feature)
#       End If
#       'Reifegrad für Verifikationskritierum erfassen
#       If BSMDatensatz.AVWReifegrad <> "" Then
        if self.avw_reifegrad != '' and self.avw_reifegrad not in verification_criterion.anf_reifegrad:
#           varErfassteVKItem.anf_Reifegrad.Add Item:=BSMDatensatz.AVWReifegrad
            verification_criterion.anf_reifegrad.append(self.avw_reifegrad)
#       End If
#       'Modulverantwortliche für Verifikationskritierum erfassen
#       If BSMDatensatz.AVWMV <> "" Then
        if self.avw_mv != '' and self.avw_mv not in verification_criterion.anf_mv:
#           varErfassteVKItem.anf_MV.Add Item:=BSMDatensatz.AVWMV
            verification_criterion.anf_mv.append(self.avw_mv)
#       End If
#       'LAH-ID für Verifikationskritierum erfassen
#       If BSMDatensatz.AVWDokumentID <> "" Then
        if self.avw_dokument_id != '' and self.avw_dokument_id not in verification_criterion.anf_lah_id:
#           varErfassteVKItem.anf_LAHID.Add Item:=BSMDatensatz.AVWDokumentID
            verification_criterion.anf_lah_id.append(self.avw_dokument_id)
#       End If
#       'LAH-Namen für Verifikationskritierum erfassen
#       If BSMDatensatz.AVWDokumentName <> "" Then
        if self.avw_dokument_name != '' and self.avw_dokument_name not in verification_criterion.anf_lah_namen:
#           varErfassteVKItem.addLAHName (BSMDatensatz.AVWDokumentName)
            verification_criterion.anf_lah_namen.append(self.avw_dokument_name)
#       End If
#       'Cluster Testing für Verifikationskriterium erfassen
#       If BSMDatensatz.ClusterTesting <> "" Then
        if self.cluster_testing != '' and self.cluster_testing not in verification_criterion.anf_cluster_testing:
#           varErfassteVKItem.anf_ClusterTesting.Add Item:=BSMDatensatz.ClusterTesting
            verification_criterion.anf_cluster_testing.append(self.cluster_testing)
#       End If
#       'Anforderungsverantwortliche für Verifikationskriterium erfassen
#       If BSMDatensatz.AVWAnforderungsverantwortliche <> "" Then
        if (
            self.avw_anforderungsverantwortliche != '' and
            self.avw_anforderungsverantwortliche not in verification_criterion.anf_anforderungsverantwortliche
        ):   
#           varErfassteVKItem.anf_Anforderungsverantwortliche.Add Item:=BSMDatensatz.AVWAnforderungsverantwortliche
            verification_criterion.anf_anforderungsverantwortliche.append(self.avw_anforderungsverantwortliche)
#       End If
#       'Temp11_Auswahlfeld für Verifikationskriterium erfassen
#       If BSMDatensatz.AVWTemp11_Auswahlfeld <> "" Then
        if (
            self.is_specific and self.avw_temp11_auswahlfeld != '' and 
            self.avw_temp11_auswahlfeld not in verification_criterion.anf_temp11_auswahlfeld
        ):
#           varErfassteVKItem.anf_Temp11_Auswahlfeld.Add Item:=BSMDatensatz.AVWTemp11_Auswahlfeld
            verification_criterion.anf_temp11_auswahlfeld.append(self.avw_temp11_auswahlfeld)
#       End If
#       
#       'Innere Schleife beenden, da es zu jeder Anforderung nur ein Verifikationskriterium gibt

    def same_id(self, requirement_id:id) -> bool:
        return requirement_id == self.avw_id
