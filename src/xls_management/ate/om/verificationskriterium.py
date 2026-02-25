from __future__ import annotations
import pandas as pd
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from xls_management.ate.om.test_case import TestCase
from xls_management.utils.aux import list_from_comma_separated_str
from xls_management.ate.om.absicherungsauftraege import Absicherungsauftrag
from xls_management.ate.data_de import TDVCAttribute as VC # Verification Criterion


class Verificationskriterium:
#Option Explicit
#

    def __init__(
        self,
    #    id:str,                #Public TF_ID As String  'ID des Testfalls
    #    name:str,              #Public TF_Name As String    'Name des Testfalls
    #    status:str,            #Public TF_Status As String  'Status des Tesfalls
    #    testinstanz:str,       #Public TF_Testinstanz As String 'Testinstanz
    #    testumgebungstyp:str,  #Public TF_Testumgebungstyp As String    'Testumgebungstyp
    #    vk_id:str,             #Public TF_VK_ID As String   'Testdesign-ID auf dem der Testfall basiert
    #    anf_ids:list[str]=[],  #Public TF_anfIDs As Collection   'Sammlung der Anforderungen, die direkt oder indirekt mit dem Testfall verknüpft sind
         columns:pd.DataFrame,
         row:int,
    ):
    #    self.id = id
    #    self.name = name
    #    self.status = status
    #    self.testinstanz = testinstanz
    #    self.testumgebungstyp = testumgebungstyp
    #    self.vk_id = vk_id
    #    self.anf_ids = anf_ids

###### From Sub EinlesenVerifikationskriterien() --initialization from a data_frame row
#               'ID des Verifikationsauftrags einlesen, Entfernung der zusätzlichen Zeichen "?" und "r"
#               strVerifikationsID = Replace(Replace(rngTDVKAttribute(1).Offset(lngZeile, 0).Value, "?", ""), "r", "")
#               'ID des Verifikationsauftrags erfassen
#               verifikationKrit.VK_ID = strVerifikationsID
                self.vk_id = str(columns[VC.ID][row]).replace('?', '').replace('r', '')
#               'Anforderungs-IDs einlesen
#               anfIDs = rngTDVKAttribute(2).Offset(lngZeile, 0).Value
                anf_ids_str = str(columns[VC.RequirementBased][row])
#               'Anforderungs-IDs nach Kommas trennen
#               Set idList = EinlesenGetrennteWerteKomma(anfIDs)
#               'Alle mit dem aktuellen Verifikationskriterium verknüpften Anforderungs-IDs erfassen
#               Set verifikationKrit.anf_ids = idList
                self.anf_ids = list_from_comma_separated_str(anf_ids_str)
#               'Status des Verifikationskriteriums einlesen
#               verifikationKrit.VK_status = rngTDVKAttribute(3).Offset(lngZeile, 0).Value
                self.status = columns[VC.Status][row]
#               'Absicherungsaufträge für dieses Verifikationskriterium anlegen
#               Set verifikationKrit.Absicherungsauftraege = New Collection
                self.absicherungsauftraege:dict[str,Absicherungsauftrag] = {}
#               'Sammlung für Testfälle vorbereiten
#               Set verifikationKrit.VK_Testfaelle = New Collection
                self.test_cases: dict[str, "TestCase"] = {}
#               'Sammlung für I-Stufen vorbereiten
#               Set verifikationKrit.anf_IStufen = New Collection
                self.anf_i_stufen = []  
#               'Sammlung für Umsetzer vorbereiten
#               Set verifikationKrit.anf_Umsetzer = New Collection
                self.anf_umsetzer = []
#               'Sammlung für BsM-Relevanz vorbereiten
#               Set verifikationKrit.anf_BsMRelevanz = New Collection
                self.anf_bsm_relevanz = []
#               'Sammlung für ASIL vorbereiten
#               Set verifikationKrit.anf_ASIL = New Collection
                self.anf_asil = []
#               'Sammlung für Feature vorbereiten
#               Set verifikationKrit.anf_Feature = New Collection
                self.anf_feature = []
#               'Sammlung für Reifegrad vorbereiten
#               Set verifikationKrit.anf_Reifegrad = New Collection
                self.anf_reifegrad = []
#               'Sammlung für Modulverantwortliche vorbereiten
#               Set verifikationKrit.anf_MV = New Collection
                self.anf_mv = []
#               'Sammlung für LAH-ID vorbereiten
#               Set verifikationKrit.anf_LAHID = New Collection
                self.anf_lah_id = []
#               'Sammlung für LAH-Namen vorbereiten
#               Set verifikationKrit.anf_LAHNamen = New Collection
                self.anf_lah_namen = []
#               'Sammlung für Cluster Testing vorbereiten
#               Set verifikationKrit.anf_ClusterTesting = New Collection
                self.anf_cluster_testing = []
#               'Sammlung für Anforderungsverantwortliche vorbereiten
#               Set verifikationKrit.anf_Anforderungsverantwortliche = New Collection
                self.anf_anforderungsverantwortliche = []
#               'Sammlung für Temp11_Auswahlfeld vorbereiten
#               Set verifikationKrit.anf_Temp11_Auswahlfeld = New Collection
                self.anf_temp11_auswahlfeld = []
#               verifikationKrit.VK_temp1Text = rngTDVKAttribute(4).Offset(lngZeile, 0).Value
                self.temp1_text = columns[VC.Temp1Text][row]
#               'Aktion einlesen
#               verifikationKrit.VK_Aktion = rngTDVKAttribute(5).Offset(lngZeile, 0).Value
                self.aktion = columns[VC.Action][row]
#
#   Sub addElementID(ByVal elemID2 As String)
    def add_element_id(self, anf_id:str):
#       Dim elemID1 As Variant
#       Dim isContained As Boolean
#       
#       isContained = False
#       For Each elemID1 In Me.TF_anfIDs
#           If (elemID1 = elemID2) Then
#               isContained = True
#               Exit For
#           End If
#       Next elemID1
#       If (isContained = False) Then
#           TF_anfIDs.Add elemID2
#       End If
        if not anf_id in self.anf_ids:
            self.anf_ids.append(anf_id)
#   End Sub
