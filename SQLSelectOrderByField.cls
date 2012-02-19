VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SQLSelectOrderByField"
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

Private pstrName As String

Public Table As SQLSelectTable
Public Order As SQLOrderByEnum
Public AggregateFunction As SQLAggregateFunctionEnum

Public Property Let Name(ByVal strValue As String)

    pstrName = Trim$(strValue)

End Property

Public Property Get Name() As String

    Name = pstrName
    
End Property

Private Sub Class_Initialize()

    Me.Order = dboOrderAscending

End Sub

Friend Property Get SQL(ByVal eConnectionType As ConnectionTypeEnum) As String

    Dim strSQL As String

    If Me.Name = vbNullString Then
        RaiseError dboErrorGeneral, "Order By field has not been set."
    End If
    
    If Me.AggregateFunction > 0 Then
        strSQL = SQLConvertAggregate(Me.AggregateFunction) & "("
    End If

    strSQL = strSQL & SQLFieldNameAndTablePrefix(Me.Table, Me.Name, eConnectionType)

    If Me.AggregateFunction > 0 Then
        strSQL = strSQL & ")"
    End If

    Select Case Me.Order
        Case dboOrderAscending
        Case dboOrderDescending
            strSQL = strSQL & " DESC"
    End Select
    
    SQL = strSQL

End Property
