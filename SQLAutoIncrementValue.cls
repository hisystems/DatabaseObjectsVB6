VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SQLAutoIncrementValue"
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

Public ConnectionType As ConnectionTypeEnum

Private pstrReturnFieldName As String

Public Property Let ReturnFieldName(ByVal strValue As String)

    If Trim$(strValue) = vbNullString Then
        RaiseError dboErrorInvalidPropertyValue, "vbNullString"
    End If

    pstrReturnFieldName = strValue

End Property

Public Property Get ReturnFieldName() As String
    
    ReturnFieldName = pstrReturnFieldName

End Property

Public Property Get SQL() As String
Attribute SQL.VB_UserMemId = 0
    
    Select Case Me.ConnectionType
    Case dboConnectionTypeMicrosoftAccess, dboConnectionTypeSQLServer
        SQL = "SELECT @@IDENTITY AS " & SQLConvertIdentifierName(Me.ReturnFieldName, Me.ConnectionType)
    Case dboConnectionTypeMySQL
        'The @@IDENTITY function is supported by MySQL from version 3.23.25
        'but use the original function here just in case
        SQL = "SELECT LAST_INSERT_ID() AS " & SQLConvertIdentifierName(Me.ReturnFieldName, Me.ConnectionType)
    End Select

End Property

Private Sub Class_Initialize()
    
    ConnectionType = modMisc.ConnectionType
    Me.ReturnFieldName = "AutoIncrementValue"
    
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
