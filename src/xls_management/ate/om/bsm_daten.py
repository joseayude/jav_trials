class BSMDaten:
#Option Explicit
#

    def __init__(
        self,
        avw_ID:str,                          #AVWID
        avw_DokumentID:str,                  #AVWDokumentID
        avw_DokumentName:str,                #AVWDokumentName
        avw_Feature:str,                     #AVWFeature
        avw_Reifegrad:str,                   #AVWReifegrad
        avw_Umsetzer:str,                    #AVWUmsetzer
        avw_Status:str,                      #AVWStatus
        avw_ASIL:str,                        #AVWASIL
        avwBsMSaFuSi:str,                    #AVWBsMSaFuSi
        avwBsMZZ:str,                        #AVWBsMZZ
        avwBsMED:str,                        #AVWBsMED
        avwBsMFFF:str,                       #AVWBsMFFF
        avwBsMO:str,                         #AVWBsMO
        avwBsMSe:str,                        #AVWBsMSe
        avwMV:str,                           #AVWMV
        avwTyp:str,                          #AVWTyp
        avwKategorie:str,                    #AVWKategorie
        avwAbgelehntNichtTestbar:str,        #AVWAbgelehntNichtTestbar
        bsm_relevanz:str,                    #BSMRelevanz
        verifikationskriterium:str,          #Verifikationskriterium As 
        testfaelle:str,                      #Testfaelle As 
        i_stufe:str,                         #IStufe
        cluster_testing:str,                 #ClusterTesting
        avw_vorgaenger_id:str,               #AVWVorgaengerID
        avw_anforderungsverantwortliche:str, #AVWAnforderungsverantwortliche
        avw_temp11_auswahlfeld:str,          #AVWTemp11_Auswahlfeld
    ):
        #Public AVWID As String
        #Public AVWDokumentID As String
        #Public AVWDokumentName As String
        #Public AVWFeature As String
        #Public AVWReifegrad As String
        #Public AVWUmsetzer As String
        #Public AVWStatus As String
        #Public AVWASIL As String
        #Public AVWBsMSaFuSi As String
        #Public AVWBsMZZ As String
        #Public AVWBsMED As String
        #Public AVWBsMFFF As String
        #Public AVWBsMO As String
        #Public AVWBsMSe As String
        #Public AVWMV As String
        #Public AVWTyp As String
        #Public AVWKategorie As String
        #Public AVWAbgelehntNichtTestbar As String
        #Public bsm_Relevanz As String
        #Public Verifikationskriterium As Collection
        #Public Testfaelle As Collection
        #Public IStufe As String
        #Public ClusterTesting As String
        #Public AVWVorgaengerID As String
        #Public AVWAnforderungsverantwortliche As String
        #Public AVWTemp11_Auswahlfeld As String
        self.avw_ID = avw_ID
        self.avw_DokumentID = avw_DokumentID
        self.avw_DokumentName = avw_DokumentName
        self.avw_Feature = avw_Feature
        self.avw_Reifegrad = avw_Reifegrad
        self.avw_Umsetzer = avw_Umsetzer
        self.avw_Status = avw_Status
        self.avw_ASIL = avw_ASIL
        self.avwBsMSaFuSi = avwBsMSaFuSi
        self.avwBsMZZ = avwBsMZZ
        self.avwBsMED = avwBsMED
        self.avwBsMFFF = avwBsMFFF
        self.avwBsMO = avwBsMO
        self.avwBsMSe = avwBsMSe
        self.avwMV = avwMV
        self.avwTyp = avwTyp
        self.avwKategorie = avwKategorie
        self.avwAbgelehntNichtTestbar = avwAbgelehntNichtTestbar
        self.bsm_Relevanz = bsm_relevanz
        self.verifikationskriterium  = verifikationskriterium
        self.testfaelle = testfaelle 
        self.i_stufe = i_stufe
        self.cluster_testing = cluster_testing
        self.avw_vorgaenger_id = avw_vorgaenger_id
        self.avw_anforderungsverantwortliche = avw_anforderungsverantwortliche
        self.avw_Temp11_auswahlfeld = avw_temp11_auswahlfeld
