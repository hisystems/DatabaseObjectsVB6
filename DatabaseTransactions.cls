VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DatabaseTransactions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' ___________________________________________________
'
'  © Hi-Integrity Systems 2007. All rights reserved.
'  www.hisystems.com.au - Toby Wicks
' ___________________________________________________
'

Private pobjConnection As ConnectionController

Friend Sub Initialize(ByVal objConnection As ConnectionController)

    Set pobjConnection = objConnection

End Sub

Public Sub Begin()
    
    pobjConnection.BeginTransaction

End Sub

Public Sub Commit()

    pobjConnection.CommitTransaction

End Sub

Public Sub Rollback()

    pobjConnection.RollbackTransaction

End Sub

Public Sub Execute(ByVal objSQL As ISQLStatement)

    pobjConnection.Execute objSQL, adCmdUnknown

End Sub

Public Function ExeceuteStatements(ByRef objSQLStatements() As ISQLStatement) As adodb.Recordset
    
    Set ExeceuteStatements = pobjConnection.ExeceuteStatements(objSQLStatements, adCmdUnknown)
     
End Function

