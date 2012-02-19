VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SQLCreateTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ISQLStatement

Public ConnectionType As ConnectionTypeEnum

Private pstrName As String
Private pobjFields As SQLTableFields

Public Property Let Name(ByVal strValue As String)

    pstrName = Trim$(strValue)

End Property

Public Property Get Name() As String

    Name = pstrName
    
End Property

Public Property Get Fields() As SQLTableFields

    Set Fields = pobjFields

End Property

Public Property Get SQL() As String
    
    Dim strSQL As String
    
    If Me.Name = vbNullString Then
        RaiseError dboErrorGeneral, "TableName has not been set."
    End If
    
    strSQL = _
        "CREATE TABLE " & _
        SQLConvertIdentifierName(Me.Name, Me.ConnectionType) & " (" & pobjFields.SQL(Me.ConnectionType, bIncludeMode:=False) & ")"
    
    'IIf(Me.ConnectionType = dboConnectionTypeMySQL And pbCheckDoesNotExist, "IF NOT EXISTS ", vbNullString) & _

    SQL = strSQL
    
End Property

Private Sub Class_Initialize()
    
    Me.ConnectionType = modMisc.ConnectionType
    
    Set pobjFields = New SQLTableFields
    pobjFields.AlterMode = dboTableFieldsModeAdd    'set that fields can only be added
    
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

