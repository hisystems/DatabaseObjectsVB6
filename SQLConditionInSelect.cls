VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SQLConditionInSelect"
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

'This class represents an IN operation together with a SELECT statement
'i.e.  ... ProductID IN (SELECT ProductID FROM Product WHERE ...)

Public Table As SQLSelectTable
Public FieldName As String
Public SelectSet As SQLSelect

Friend Property Get SQL(ByVal eConnectionType As ConnectionTypeEnum) As String

    If Trim$(FieldName) = vbNullString Then
        RaiseError dboErrorGeneral, "FieldName not set."
    End If

    If SelectSet Is Nothing Then
        RaiseError dboErrorGeneral, "SelectSet not set."
    End If

    SelectSet.ConnectionType = eConnectionType

    SQL = _
        SQLFieldNameAndTablePrefix(Me.Table, Me.FieldName, eConnectionType) & _
        " IN (" & SelectSet.SQL & ")"

End Property
