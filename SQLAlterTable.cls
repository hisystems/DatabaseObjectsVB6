VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SQLAlterTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ISQLStatement

Public ConnectionType As ConnectionTypeEnum
Public Name As String

Private pobjFields As SQLTableFields

Public Property Get Fields() As SQLTableFields

    Set Fields = pobjFields

End Property

Public Property Get SQL() As String

    Dim strSQL As String
    Dim strMode As String

    If Me.Name = vbNullString Then
        RaiseError dboErrorGeneral, "TableName has not been set."
    End If
    
    If Me.Fields.Count = 0 Then
        RaiseError dboErrorGeneral, "Fields have not been added."
    End If

    strSQL = _
        "ALTER TABLE " & SQLConvertIdentifierName(Me.Name, Me.ConnectionType) & " " & _
        pobjFields.SQL(Me.ConnectionType)
    
    SQL = strSQL
    
End Property

Private Sub Class_Initialize()
    
    ConnectionType = modMisc.ConnectionType
    Set pobjFields = New SQLTableFields
    
End Sub

Private Property Get ISQLStatement_ConnectionType() As ConnectionTypeEnum
    
    ISQLStatement_ConnectionType = ConnectionType

End Property

Private Property Let ISQLStatement_ConnectionType(ByVal RHS As ConnectionTypeEnum)
    
    ConnectionType = RHS

End Property

Private Property Get ISQLStatement_SQL() As String
    
    ISQLStatement_SQL = Me.SQL

End Property
