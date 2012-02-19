VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SQLSelectTable"
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

Private pstrName As String

'Public Alias As String

Public Property Let Name(ByVal strValue As String)

    If Trim$(strValue) = vbNullString Then
        RaiseError dboErrorInvalidArgument, "Null"
    Else
        pstrName = strValue
    End If

End Property

Public Property Get Name() As String

    Name = pstrName

End Property

Friend Property Get SQL(ByVal eConnectionType As ConnectionTypeEnum) As String

    Dim strSQL As String

    If Trim$(Me.Name) = vbNullString Then
        RaiseError dboErrorGeneral, "Table has not been set."
    End If

    strSQL = SQLConvertIdentifierName(Me.Name, eConnectionType)
    
    'If Trim$(Me.Alias) <> vbNullString Then
    '    strSQL = strSQL & " AS " & SQLConvertIdentifierName(Me.Alias)
    'End If
    
    SQL = strSQL
    
End Property
