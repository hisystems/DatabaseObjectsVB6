VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SQLSelectTableJoinCondition"
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

Public LeftTableFieldName As String
Public RightTableFieldName As String
Public Compare As SQLComparisonOperatorEnum

Private pobjParent As SQLSelectTableJoinConditions

Friend Property Set Parent(ByVal objValue As SQLSelectTableJoinConditions)

    Set pobjParent = objValue

End Property

Friend Property Get SQL(ByVal eConnectionType As ConnectionTypeEnum) As String

    Dim strSQL As String

    If Trim$(LeftTableFieldName) = vbNullString Then
        RaiseError dboErrorGeneral, "LeftTableFieldName has not been specified."
    End If

    If Trim$(RightTableFieldName) = vbNullString Then
        RaiseError dboErrorGeneral, "RightTableFieldName has not been specified."
    End If

    strSQL = _
            SQLFieldNameAndTablePrefix(pobjParent.Parent.LeftTable, LeftTableFieldName, eConnectionType) & " " & _
            SQLConvertCompare(Compare) & " " & _
            SQLFieldNameAndTablePrefix(pobjParent.Parent.RightTable, RightTableFieldName, eConnectionType)

    SQL = strSQL

End Property


