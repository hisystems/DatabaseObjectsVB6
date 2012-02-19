VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SQLInsert"
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

Implements ISQLStatement

Private pobjFields As SQLFieldValues

Public ConnectionType As ConnectionTypeEnum
Public TableName As String

Public Property Set Fields(ByVal objFields As SQLFieldValues)

    If objFields Is Nothing Then
        RaiseError dboErrorObjectIsNothing
    End If

    Set pobjFields = objFields

End Property

Public Property Get Fields() As SQLFieldValues

    Set Fields = pobjFields

End Property

Public Property Get SQL() As String
Attribute SQL.VB_UserMemId = 0

    Dim intIndex As Integer
    Dim strSQL As String
    Dim strFields As String
    Dim strFieldValues As String
    
    If Trim$(TableName) = vbNullString Then
        RaiseError dboErrorGeneral, "TableName property has not been set."
    End If
    
    If pobjFields.Count = 0 Then
        RaiseError dboErrorGeneral, "Field values have not been set."
    End If
    
    For intIndex = 1 To pobjFields.Count
        strFields = strFields & SQLConvertIdentifierName(pobjFields(intIndex).Name, Me.ConnectionType)
        If intIndex <> pobjFields.Count Then
            strFields = strFields & ","
        End If
    Next
    
    For intIndex = 1 To pobjFields.Count
        strFieldValues = strFieldValues & SQLConvertValue(pobjFields(intIndex).Value, Me.ConnectionType)
        If intIndex <> pobjFields.Count Then
            strFieldValues = strFieldValues & ","
        End If
    Next
    
    strSQL = _
        "INSERT INTO " & SQLConvertIdentifierName(Trim$(Me.TableName), Me.ConnectionType) & " " & _
        "(" & strFields & ") VALUES (" & strFieldValues & ")"
    
    SQL = strSQL
    
End Property

Private Sub Class_Initialize()
    
    Set pobjFields = New SQLFieldValues
    ConnectionType = modMisc.ConnectionType

End Sub

Private Property Get ISQLStatement_ConnectionType() As ConnectionTypeEnum
    ISQLStatement_ConnectionType = Me.ConnectionType
End Property

Private Property Let ISQLStatement_ConnectionType(ByVal RHS As ConnectionTypeEnum)
    Me.ConnectionType = RHS
End Property

Private Property Get ISQLStatement_SQL() As String
    ISQLStatement_SQL = Me.SQL
End Property
