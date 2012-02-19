VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SQLTableExists"
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

Private pstrName As String

Public Property Let Name(ByVal strValue As String)

    pstrName = Trim$(strValue)

End Property

Public Property Get Name() As String

    Name = pstrName
    
End Property

Public Property Get SQL() As String
    
    'To use this class check whether the recordset is empty or not - this will work with all databases
    
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
            strSQL = "SHOW TABLES LIKE " & SQLConvertValue(Me.Name, Me.ConnectionType)
        Case dboConnectionTypeSQLServer
            Set objSelect = New SQLSelect
            With objSelect
                .ConnectionType = Me.ConnectionType
                .Tables.Add "sysobjects"
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
