VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SQLIndexFields"
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

Public Function Add( _
    Optional ByVal strFieldName As String, _
    Optional ByVal eOrder As SQLOrderByEnum) As SQLIndexField
    
    Dim objField As SQLIndexField
    Set objField = New SQLIndexField
    
    With objField
        .Name = strFieldName
        .Order = eOrder
    End With

    pcolFields.Add objField
    Set Add = objField

End Function

Private Sub Class_Initialize()

    Set pcolFields = New Collection
    
End Sub

Friend Property Get SQL(ByVal eConnectionType As ConnectionTypeEnum) As String

    Const cstrSeperator As String = ", "

    Dim strSQL As String
    Dim objField As SQLIndexField
    
    For Each objField In pcolFields
        strSQL = strSQL & objField.SQL(eConnectionType) & cstrSeperator
    Next
        
    SQL = Left$(strSQL, Len(strSQL) - Len(cstrSeperator))    'remove the last comma and space

End Property
