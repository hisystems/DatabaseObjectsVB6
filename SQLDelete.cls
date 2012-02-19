VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SQLDelete"
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

Private pobjConditions As SQLConditions

Public ConnectionType As ConnectionTypeEnum
Public TableName As String

Public Property Get Where() As SQLConditions
    
    Set Where = pobjConditions

End Property

Public Property Set Where(ByVal objValue As SQLConditions)

    Set pobjConditions = objValue
    
End Property

Public Property Get SQL() As String
Attribute SQL.VB_UserMemId = 0

    Dim strSQL As String

    If Trim$(TableName) = vbNullString Then
        RaiseError dboErrorGeneral, "TableName property has not been set."
    End If

    strSQL = "DELETE FROM " & SQLConvertIdentifierName(TableName, Me.ConnectionType)

    If Not pobjConditions Is Nothing Then
        If pobjConditions.SQL(Me.ConnectionType) <> vbNullString Then
            strSQL = strSQL & " WHERE " & pobjConditions.SQL(Me.ConnectionType)
        End If
    End If
    
    SQL = strSQL

End Property

Private Sub Class_Initialize()

    Me.ConnectionType = modMisc.ConnectionType
    Set pobjConditions = New SQLConditions

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
