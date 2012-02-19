VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SQLCondition"
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

Public Enum SQLComparisonOperatorEnum
    dboComparisonEqualTo
    dboComparisonNotEqualTo
    dboComparisonLessThan
    dboComparisonLessThanOrEqualTo
    dboComparisonGreaterThan
    dboComparisonGreaterThanOrEqualTo
    dboComparisonLike
    dboComparisonNotLike
End Enum

Public Enum SQLLogicalOperatorEnum
    dboLogicalAnd
    dboLogicalOr
End Enum

Private pvValue As Variant

Public Table As SQLSelectTable
Public FieldName As String
Public compare As SQLComparisonOperatorEnum

Public Property Let Value(ByVal vValue As Variant)

    SQLConditionValue vValue, pvValue
    
End Property

Public Property Get Value() As Variant
 
    If VarType(pvValue) = vbObject Then
        Set Value = pvValue
    Else
        Value = pvValue
    End If
   
End Property

Friend Property Get SQL(ByVal eConnectionType As ConnectionTypeEnum) As String

    CompareValuePairAssertValid Me.compare, pvValue
    
    SQL = Condition(Me.Table, Me.FieldName, Me.compare, Me.Value, eConnectionType)

End Property

Private Function Condition( _
    ByVal objTable As SQLSelectTable, _
    ByVal strFieldName As String, _
    ByVal eCompare As SQLComparisonOperatorEnum, _
    ByVal vValue As Variant, _
    ByVal eConnectionType As ConnectionTypeEnum) As String

    SQLConvertBooleanValue vValue, eCompare

    Condition = _
        SQLFieldNameAndTablePrefix(objTable, strFieldName, eConnectionType) & " "

    'return 'IS NULL' rather than '= NULL'
    If SQLValueIsNull(vValue) Then
        If eCompare = dboComparisonEqualTo Then
            Condition = Condition & "IS " & SQLConvertValue(vValue, eConnectionType)
        ElseIf eCompare = dboComparisonNotEqualTo Then
            Condition = Condition & "IS NOT " & SQLConvertValue(vValue, eConnectionType)
        Else
            RaiseError dboErrorInvalidArgument, "vbNull or Nothing specified as an SQLCondition value without using the dboComparisonEqualTo or dboComparisonNotEqualTo operators"
        End If
    Else
        Condition = Condition & _
            SQLConvertCompare(eCompare) & " " & _
            SQLConvertValue(vValue, eConnectionType)
    End If

End Function

