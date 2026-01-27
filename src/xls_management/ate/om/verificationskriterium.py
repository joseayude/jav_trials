#class Verificationskriterium
#Option Explicit
#
#Public TF_ID As String  'ID des Testfalls
#Public TF_Name As String    'Name des Testfalls
#Public TF_Status As String  'Status des Tesfalls
#Public TF_Testinstanz As String 'Testinstanz
#Public TF_Testumgebungstyp As String    'Testumgebungstyp
#Public TF_VK_ID As String   'Testdesign-ID auf dem der Testfall basiert
#Public TF_anfIDs As Collection   'Sammlung der Anforderungen, die direkt oder indirekt mit dem Testfall verknüpft sind
#
#Sub addElementID(ByVal elemID2 As String)
#Dim elemID1 As Variant
#Dim isContained As Boolean
#
#isContained = False
#For Each elemID1 In Me.TF_anfIDs
#    If (elemID1 = elemID2) Then
#        isContained = True
#        Exit For
#    End If
#Next elemID1
#If (isContained = False) Then
#    TF_anfIDs.Add elemID2
#End If
#End Sub
