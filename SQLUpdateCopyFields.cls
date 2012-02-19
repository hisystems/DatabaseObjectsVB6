VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SQLUpdateCopyFields"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' ___________________________________________________
'
'  © Hi-Integrity Systems 2007. All rights reserved.
'  www.hisystems.com.au - Toby Wicks
' ___________________________________________________
'

Option Explicit

Private pcolFields As Collection

Public Function Add(ByVal strSourceFieldName As String, ByVal strDestinationFieldName As String) As SQLUpdateCopyField

    Dim objCopyField As SQLUpdateCopyField
    Set objCopyField = New SQLUpdateCopyField
    
    objCopyField.Initialize strSourceFieldName, strDestinationFieldName
    
    pcolFields.Add objCopyField
    
    Set Add = objCopyField
    
End Function

Public Property Get Item(ByVal intIndex As Integer) As SQLUpdateCopyField
Attribute Item.VB_UserMemId = 0

    Set Item = pcolFields(intIndex)

End Property

Public Property Get Count() As Integer

    Count = pcolFields.Count

End Property

Private Sub Class_Initialize()

    Set pcolFields = New Collection

End Sub
