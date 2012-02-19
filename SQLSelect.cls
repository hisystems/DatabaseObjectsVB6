VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SQLSelect"
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

Public Enum SQLAggregateFunctionEnum
    dboAggregateAverage = 1
    dboAggregateCount
    dboAggregateSum
    dboAggregateMinimum
    dboAggregateMaximum
    dboAggregateStandardDeviation
    dboAggregateVariance
End Enum

Public ConnectionType As ConnectionTypeEnum

Private pobjFields As SQLSelectFields
Private pobjTables As SQLSelectTables
Private pobjConditions As SQLConditions
Private pobjOrderByFields As SQLSelectOrderByFields
Private pobjGroupByFields As SQLSelectGroupByFields
Private pbDistinct As Boolean
Private plngTop As Long
Private pbPerformLocking As Boolean

Public Property Let Top(ByVal lngValue As Long)

    If lngValue <= 0 Then
        RaiseError dboErrorInvalidArgument, lngValue
    End If
    
    plngTop = lngValue

End Property

Public Property Get Top() As Long

    Top = plngTop

End Property

''' <summary>
''' Indicates whether the rows that are selected are locked for reading and updating.
''' Equivalent to Serialiazable isolation level.
''' These rows cannot be read or updated until the lock is released.
''' Locks are released when the transaction has been committed or rolled back.
''' </summary>
Public Property Get PerformLocking() As Boolean

    PerformLocking = pbPerformLocking

End Property

Public Property Let PerformLocking(ByVal bValue As Boolean)

    pbPerformLocking = bValue

End Property
        
Public Property Get Distinct() As Boolean

    Distinct = pbDistinct

End Property

Public Property Let Distinct(ByVal bValue As Boolean)

    pbDistinct = bValue

End Property

Public Property Get Tables() As SQLSelectTables

    Set Tables = pobjTables

End Property

Public Property Set Tables(ByVal objTables As SQLSelectTables)

    If objTables Is Nothing Then
        RaiseError dboErrorObjectIsNothing
    Else
        Set pobjTables = objTables
    End If

End Property

Public Property Set Fields(ByVal objValue As SQLSelectFields)

    If objValue Is Nothing Then
        RaiseError dboErrorObjectIsNothing
    End If
    
    Set pobjFields = objValue

End Property

Public Property Get Fields() As SQLSelectFields

    Set Fields = pobjFields

End Property

Public Property Get Where() As SQLConditions

    Set Where = pobjConditions

End Property

Public Property Set Where(ByVal objValue As SQLConditions)
    
    Set pobjConditions = objValue
    
End Property

Public Property Get OrderBy() As SQLSelectOrderByFields

    Set OrderBy = pobjOrderByFields

End Property

Public Property Set OrderBy(ByVal objValue As SQLSelectOrderByFields)

    Set pobjOrderByFields = objValue

End Property

Public Property Get GroupBy() As SQLSelectGroupByFields

    Set GroupBy = pobjGroupByFields

End Property

Public Property Set GroupBy(ByVal objValue As SQLSelectGroupByFields)

    Set pobjGroupByFields = objValue

End Property

Public Property Get SQL() As String
Attribute SQL.VB_UserMemId = 0
    
    Dim intIndex As Integer
    Dim strSQL As String
    Dim strConditions As String
    Dim strOrderBy As String
    Dim strGroupBy As String

    If pobjTables.Count = 0 Then
        RaiseError dboErrorGeneral, "The table has not been set."
    End If
     
    strSQL = _
        "SELECT" & DistinctClause & TopClause & " " & pobjFields.SQL(Me.ConnectionType) & _
        " FROM " & pobjTables.SQL(Me.ConnectionType)

    If pbPerformLocking Then
        Select Case Me.ConnectionType
            Case dboConnectionTypeSQLServer
                strSQL = strSQL & " WITH (HOLDLOCK, ROWLOCK)"
            Case dboConnectionTypeMicrosoftAccess
                'Unsure of the MSAccess equivalent
                RaiseError dboErrorGeneral, "PerformLocking is unavailable for Microsoft Access"
        End Select
    End If
        
    If Not pobjConditions Is Nothing Then
        strConditions = pobjConditions.SQL(Me.ConnectionType)
        If strConditions <> vbNullString Then
            strSQL = strSQL & " WHERE " & strConditions
        End If
    End If

    If Not pobjGroupByFields Is Nothing Then
        strGroupBy = pobjGroupByFields.SQL(Me.ConnectionType)
        If strGroupBy <> vbNullString Then
            strSQL = strSQL & " GROUP BY " & strGroupBy
        End If
    End If

    If Not pobjOrderByFields Is Nothing Then
        strOrderBy = pobjOrderByFields.SQL(Me.ConnectionType)
        If strOrderBy <> vbNullString Then
            strSQL = strSQL & " ORDER BY " & strOrderBy
        End If
    End If

    If pbPerformLocking Then
        Select Case Me.ConnectionType
            Case dboConnectionTypeMySQL
                strSQL = strSQL & " FOR UPDATE"
        End Select
    End If
    
    SQL = strSQL

End Property

Private Function TopClause() As String

    If plngTop > 0 Then
        TopClause = " TOP " & plngTop
    End If

End Function

Private Function DistinctClause() As String

    If pbDistinct Then
        DistinctClause = "DISTINCT "
    Else
        DistinctClause = vbNullString
    End If

End Function

Private Sub Class_Initialize()
 
    Me.ConnectionType = modMisc.ConnectionType
    Set pobjFields = New SQLSelectFields
    Set pobjTables = New SQLSelectTables
    Set pobjConditions = New SQLConditions
    Set pobjOrderByFields = New SQLSelectOrderByFields
    Set pobjGroupByFields = New SQLSelectGroupByFields

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
