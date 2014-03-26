VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SQLBitwiseExpression"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' ___________________________________________________
'
'  © Hi-Integrity Systems 2014. All rights reserved.
'  www.hisystems.com.au - Toby Wicks
' ___________________________________________________
'

Option Explicit

Implements ISQLExpression

Public Enum SQLBitwiseOperatorEnum
    dboBitwiseAnd
    dboBitwiseOr
    dboBitwiseXor
End Enum

Public Left As ISQLExpression
Public Right As ISQLExpression
Public Operator As SQLBitwiseOperatorEnum

Private Sub Class_Initialize()

    Set Left = New SQLValueExpression
    Set Right = New SQLValueExpression
    Operator = dboArithmeticAdd

End Sub

Private Property Get ISQLExpression_SQL(ByVal eConnection As ConnectionTypeEnum) As String

    ISQLExpression_SQL = "(" & Left.SQL(eConnection) & " " & OperatorToString(Operator) & " " & Right.SQL(eConnection) & ")"

End Property

Private Function OperatorToString(ByVal eOperator As SQLBitwiseOperatorEnum) As String

    Select Case eOperator
        Case dboBitwiseAnd
            OperatorToString = "&"
        Case dboBitwiseOr
            OperatorToString = "|"
        Case dboBitwiseXor
            OperatorToString = "^"
        Case Else
            RaiseError dboErrorInvalidArgument
    End Select

End Function

