VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SQLIndexExists"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ISQLStatement

Public ConnectionType As ConnectionTypeEnum

Private pstrName As String

Public Property Let Name(ByVal strValue As String)

    pstrName = Trim$(strValue)

End Property

Public Property Get Name() As String

    Name = pstrName
    
End Property

Public Property Get SQL() As String
    
    'To use this class check whether the recordset is empty or not
    'This only works with Microsoft Access and SQL Server - if anyone knows how to get this to work with MySQL let me know!
    
    Dim strSQL As String
    Dim objSelect As SQLSelect
    
    Select Case Me.ConnectionType
        Case dboConnectionTypeMicrosoftAccess
            Set objSelect = New SQLSelect
            With objSelect
                .ConnectionType = Me.ConnectionType
                .Tables.Add "msysobjects"
                .Where.Add "Name", dboComparisonEqualTo, Me.Name
                .Where.Add "Type", dboComparisonEqualTo, 1
                strSQL = .SQL
            End With
        Case dboConnectionTypeMySQL
            RaiseError dboErrorNotSupported
        Case dboConnectionTypeSQLServer
        
'select * from sysindexes
  'where id=object_id('<table_name>') and name='<index_name>')
  
            Set objSelect = New SQLSelect
            With objSelect
                .ConnectionType = Me.ConnectionType
                .Tables.Add "sysindexes"
                .Where.Add
                .Where.Add "Name", dboComparisonEqualTo, Me.Name
                .Where.Add "XType", dboComparisonEqualTo, "U"       'U = User defined table
                strSQL = .SQL
            End With
    End Select
            
    SQL = strSQL

End Property

Private Sub Class_Initialize()
    Me.ConnectionType = modMisc.ConnectionType
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

