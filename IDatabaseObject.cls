VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IDatabaseObject"
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

Public Property Let IsSaved(ByVal bIsSaved As Boolean)

End Property

Public Property Get IsSaved() As Boolean

End Property

Public Property Let DistinctValue(ByVal vValue As Variant)

End Property

Public Property Get DistinctValue() As Variant

End Property

Public Sub Load(ByVal objFields As SQLFieldValues)

End Sub

Public Function Save() As SQLFieldValues

End Function
