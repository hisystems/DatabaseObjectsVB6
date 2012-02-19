VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SQLUpdateCopyField"
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

Private pstrSourceFieldName As String
Private pstrDestinationFieldName As String

Friend Sub Initialize( _
    ByVal strSourceFieldName As String, _
    ByVal strDestinationFieldName As String)
    
    If Trim$(strSourceFieldName) = vbNullString Then
        RaiseError dboErrorInvalidArgument, "SourceFieldName is null"
    End If
    
    If Trim$(strDestinationFieldName) = vbNullString Then
        RaiseError dboErrorInvalidArgument, "DestinationFieldName is null"
    End If
    
    pstrSourceFieldName = strSourceFieldName
    pstrDestinationFieldName = strDestinationFieldName
    
End Sub

Public Property Get SourceFieldName() As String

    SourceFieldName = pstrSourceFieldName

End Property

Public Property Get DestinationFieldName() As String

    DestinationFieldName = pstrDestinationFieldName

End Property

