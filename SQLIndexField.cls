VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SQLIndexField"
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
Private peOrder As SQLOrderByEnum

Public Property Let Name(ByVal strValue As String)

    pstrName = Trim$(strValue)

End Property

Public Property Get Name() As String

    Name = pstrName
    
End Property

Public Property Let Order(ByVal eValue As SQLOrderByEnum)

    peOrder = eValue

End Property

Public Property Get Order() As SQLOrderByEnum

    Order = peOrder

End Property

Friend Property Get SQL(ByVal eConnectionType As ConnectionTypeEnum)

    SQL = SQLConvertIdentifierName(Me.Name, eConnectionType) & IIf(peOrder = dboOrderDescending, " DESC", vbNullString)

End Property
