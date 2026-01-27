class Verificationskriterium:
#Option Explicit
#

    def __init(
        self,
        id:str,                #Public TF_ID As String  'ID des Testfalls
        name:str,              #Public TF_Name As String    'Name des Testfalls
        status:str,            #Public TF_Status As String  'Status des Tesfalls
        testinstanz:str,       #Public TF_Testinstanz As String 'Testinstanz
        testumgebungstyp:str,  #Public TF_Testumgebungstyp As String    'Testumgebungstyp
        vk_id:str,             #Public TF_VK_ID As String   'Testdesign-ID auf dem der Testfall basiert
        anf_ids:list[str]=[],  #Public TF_anfIDs As Collection   'Sammlung der Anforderungen, die direkt oder indirekt mit dem Testfall verknüpft sind
    ):
        self.id = id
        self.name = name
        self.status = status
        self.testinstanz = testinstanz
        self.testumgebungstyp = testumgebungstyp
        self.vk_id = vk_id
        self.anf_ids = anf_ids
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
        if not anf_id in self.tf_vanf_idsk_id:
            self.anf_ids.append(anf_id)
#   End Sub
