VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SQLConditionFieldCompare"
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

Public Table1 As SQLSelectTable
Public FieldName1 As String

Public Compare As SQLComparisonOperatorEnum

Public Table2 As SQLSelectTable
Public FieldName2 As String

Friend Property Get SQL(ByVal eConnectionType As ConnectionTypeEnum) As String

    If Trim$(FieldName1) = vbNullString Then
        RaiseError dboErrorInvalidPropertyValue, "FieldName1 not set."
    End If

    If Trim$(FieldName2) = vbNullString Then
        RaiseError dboErrorInvalidPropertyValue, "FieldName2 not set."
    End If

    SQL = _
        SQLFieldNameAndTablePrefix(Me.Table1, Me.FieldName1, eConnectionType) & " " & _
        SQLConvertCompare(Me.Compare) & " " & _
        SQLFieldNameAndTablePrefix(Me.Table2, Me.FieldName2, eConnectionType)

End Property
