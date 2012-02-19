VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SQLSelectField"
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
Private pstrAlias As String

Public Table As SQLSelectTable
Public AggregateFunction As SQLAggregateFunctionEnum

Public Property Let Name(ByVal strValue As String)

    pstrName = Trim$(strValue)

End Property

Public Property Get Name() As String

    Name = pstrName
    
End Property

Public Property Let Alias(ByVal strValue As String)

    pstrAlias = Trim$(strValue)

End Property

Public Property Get Alias() As String

    Alias = pstrAlias
    
End Property

Friend Property Get SQL(ByVal eConnectionType As ConnectionTypeEnum) As String

    Dim strSQL As String

    If Me.AggregateFunction > 0 Then
        strSQL = SQLConvertAggregate(Me.AggregateFunction) & "("
    End If

    If Me.AggregateFunction = dboAggregateCount Then
        strSQL = strSQL & "*"
    Else
        strSQL = strSQL & SQLFieldNameAndTablePrefix(Me.Table, Me.Name, eConnectionType)
        If Me.Name = vbNullString Then
            RaiseError dboErrorGeneral, "Field Name has not been set."
        End If
    End If
    
    If Me.AggregateFunction > 0 Then
        strSQL = strSQL & ")"
    End If

    If Me.Alias <> vbNullString Then
        strSQL = strSQL & " AS " & SQLConvertIdentifierName(Me.Alias, eConnectionType)
    End If

    SQL = strSQL

End Property

