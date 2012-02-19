VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SQLConditionSelect"
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

'This class allows a conditional value generated from an SELECT statement to be added as an SQL condition.
'i.e.  ... WHERE (SELECT MAX(StockOnHand) FROM Product WHERE Supplier.ProductID = Product.ProductID) > 1000

Private pvValue As Variant

Public SelectSet As SQLSelect
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

    If Me.SelectSet Is Nothing Then
        RaiseError dboErrorGeneral, "SelectSet not set."
    End If

    CompareValuePairAssertValid Me.compare, pvValue

    SelectSet.ConnectionType = eConnectionType
    SQL = Condition(SelectSet, Me.compare, Me.Value, eConnectionType)

End Property

Private Function Condition( _
    ByVal objSelect As SQLSelect, _
    ByVal eCompare As SQLComparisonOperatorEnum, _
    ByVal vValue As Variant, _
    ByVal eConnectionType As ConnectionTypeEnum) As String

    SQLConvertBooleanValue vValue, eCompare

    Condition = _
        "(" & objSelect.SQL & ") " & _
        SQLConvertCompare(eCompare) & " " & _
        SQLConvertValue(vValue, eConnectionType)

End Function

