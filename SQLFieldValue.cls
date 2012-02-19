VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SQLFieldValue"
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
Private pvValue As Variant

Public Property Let Name(ByVal strValue As String)

    pstrName = Trim$(strValue)

End Property

Public Property Get Name() As String

    Name = pstrName
    
End Property

Public Property Let Value(ByVal vValue As Variant)

    Select Case VarType(vValue)
    Case vbByte, vbCurrency, vbDecimal, vbDouble, vbInteger, vbLong
    Case vbSingle, vbString, vbDate, vbBoolean, vbNull
    'allow for byte arrays
    Case vbByte Or vbArray
    Case Else
        RaiseError dboErrorGeneral, "Invalid datatype"
    End Select

    pvValue = vValue

End Property

Public Property Get Value() As Variant
Attribute Value.VB_UserMemId = 0

    Value = pvValue
    
End Property
