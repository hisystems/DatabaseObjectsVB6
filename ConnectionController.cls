VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ConnectionController"
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

Option Explicit

Private pobjConnection As adodb.Connection
Private pstrConnectionString As String
Private peConnectionType As ConnectionTypeEnum
Private pintTransactionCount As Integer
Private pintConnectionCount As Integer

Friend Sub Initialize(ByVal strConnectionString As String, ByVal eConnectionType As ConnectionTypeEnum)

    pstrConnectionString = strConnectionString
    peConnectionType = eConnectionType
    modMisc.ConnectionType = eConnectionType

End Sub

Private Function CreateConnection() As adodb.Connection

    Dim objConnection As adodb.Connection
    
    Set objConnection = New adodb.Connection
    objConnection.ConnectionString = pstrConnectionString
    
    Set CreateConnection = objConnection

End Function

''' --------------------------------------------------------------------------------
''' <summary>
''' Indicates that either Execute or ExecuteNonQuery is going to be used
''' and that a connection needs to be opened if one is not already.
''' If in transaction mode (Transactions.Begin has been called) then the
''' current connection is left opened.
''' If not in transaction mode then a new connection is opened.
''' Always call Start before using Execute or ExecuteNonQuery whether in
''' transaction mode or not as the library will open the connection if necessary.
''' </summary>
''' <remarks>
''' This feature is particularly relevant when database records are locked
''' during transactions. If a second connection outside of the DatabaseObjects
''' library is used then a possible deadlock could occur. Using the Execute
''' and ExecuteNonQuery functions means that a new connection is opened if not
''' in transaction mode or the current transaction connection is used - thereby
''' avoiding potential deadlocks.
''' </remarks>
''' --------------------------------------------------------------------------------
Public Sub Start()

    ConnectionStart

End Sub

''' --------------------------------------------------------------------------------
''' <summary>
''' Indicates that either Execute or ExecuteNonQuery have been called and are not
''' going to be called again.
''' If in transaction mode (Transactions.Begin has been called) then the
''' connection is left open until Transactions.Commit or Rollback is called.
''' If not in transaction mode then the connection is closed.
''' Always call Finished when finished using the connection whether in
''' transaction mode or not as the library will close the connection if necessary.
''' </summary>
''' <remarks>
''' This feature is particularly relevant when database records are locked
''' during transactions. If a second connection outside of the DatabaseObjects
''' library is used then a possible deadlock could occur. Using the Execute
''' and ExecuteNonQuery functions means that a new connection is opened if not
''' in transaction mode or the current transaction connection is used - thereby
''' avoiding potential deadlocks.
''' </remarks>
''' --------------------------------------------------------------------------------
Public Sub Finished()

    ConnectionFinished

End Sub

Private Sub ConnectionStart()

    If pintConnectionCount = 0 Then
        Set pobjConnection = CreateConnection()
        pobjConnection.Open
    End If
    
    pintConnectionCount = pintConnectionCount + 1

End Sub

Private Sub ConnectionFinished(Optional ByVal intConnectionCountToSubtract As Integer = 1)
 
    pintConnectionCount = pintConnectionCount - intConnectionCountToSubtract
 
    If pintConnectionCount <= 0 Then
        pintConnectionCount = 0
        If pobjConnection Is Nothing Then
            RaiseError dboErrorGeneral, "Attempted to close a connection that is already closed"
        End If
        pobjConnection.Close
        Set pobjConnection = Nothing
    End If
 
End Sub

Friend Sub BeginTransaction()

    ConnectionStart

    If pintTransactionCount = 0 Then
        Execute New SQLBeginTransaction, adCmdText
    End If

    pintTransactionCount = pintTransactionCount + 1

End Sub

Friend Sub CommitTransaction()

    pintTransactionCount = pintTransactionCount - 1

    If pintTransactionCount = 0 Then
        Execute New SQLCommitTransaction, adCmdText
    End If

    ConnectionFinished

End Sub

Friend Sub RollbackTransaction()

    Execute New SQLRollbackTransaction, adCmdText

    'Remove the connection count by the number of transaction levels
    'that were rolled back.
    ConnectionFinished pintConnectionCount
    pintTransactionCount = 0

End Sub

Public Function Execute( _
    ByVal objSQL As ISQLStatement, _
    ByVal eCommandType As CommandTypeEnum) As adodb.Recordset
 
    Dim objStatements(1 To 1) As ISQLStatement
    Set objStatements(1) = objSQL
 
    Set Execute = ExecuteInternal(pobjConnection, objStatements, eCommandType)
 
End Function
 
Public Function ExeceuteStatements( _
    ByRef objSQLStatements() As ISQLStatement, _
    ByVal eCommandType As CommandTypeEnum) As adodb.Recordset
    
    Set ExeceuteStatements = ExecuteInternal(pobjConnection, objSQLStatements, eCommandType)
     
End Function
 
Private Function ExecuteInternal( _
    ByVal objConnection As adodb.Connection, _
    ByRef objSQLStatements() As ISQLStatement, _
    ByVal eCommandType As CommandTypeEnum) As adodb.Recordset
    
    If objConnection Is Nothing Then
        RaiseError dboErrorGeneral, "Connection is not open, call DBO.Connection.Start or DBO.Transactions.Begin"
    End If
    
    Dim strSQL As String
    Dim intIndex As Integer
    
    For intIndex = LBound(objSQLStatements) To UBound(objSQLStatements)
        objSQLStatements(intIndex).ConnectionType = peConnectionType
        strSQL = strSQL & objSQLStatements(intIndex).SQL & ";"
    Next
    
    #If DebugSQL Then
        Debug.Print strSQL
    #End If
    
    Set ExecuteInternal = objConnection.Execute(strSQL, , eCommandType)
 
End Function
