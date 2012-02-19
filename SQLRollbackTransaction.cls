VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SQLRollbackTransaction"
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

Private Sub Class_Initialize()

    Me.ConnectionType = modMisc.ConnectionType

End Sub

Public Property Get SQL() As String

    Select Case Me.ConnectionType
        Case dboConnectionTypeSQLServer
            SQL = "ROLLBACK TRANSACTION"
        Case dboConnectionTypeMicrosoftAccess
            SQL = "ROLLBACK TRANSACTION"
        Case dboConnectionTypeMySQL
            SQL = "ROLLBACK"
        Case Else
            RaiseError dboErrorInvalidArgument
    End Select

End Property

Private Property Get ISQLStatement_ConnectionType() As ConnectionTypeEnum
    
    ISQLStatement_ConnectionType = Me.ConnectionType
    
End Property

Private Property Let ISQLStatement_ConnectionType(ByVal RHS As ConnectionTypeEnum)

    Me.ConnectionType = RHS

End Property

Private Property Get ISQLStatement_SQL() As String

    ISQLStatement_SQL = Me.SQL

End Property


