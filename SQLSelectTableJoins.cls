VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SQLSelectTableJoins"
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

Public Enum SQLJoinEnum
    dboInnerJoin
    dboFullOuterJoin
    dboLeftOuterJoin
    dboRightOuterJoin
End Enum

Private pcolJoins As Collection

Public Function Add( _
    ByVal objLeftTable As SQLSelectTable, _
    ByVal eJoin As SQLJoinEnum, _
    ByVal objRightTable As SQLSelectTable) As SQLSelectTableJoin
   
    Dim objJoin As SQLSelectTableJoin
    Set objJoin = New SQLSelectTableJoin
    
    Set objJoin.Parent = Me
    Set objJoin.LeftTable = objLeftTable
    objJoin.TheType = eJoin
    Set objJoin.RightTable = objRightTable
    
    pcolJoins.Add objJoin
    Set Add = objJoin
   
End Function

Public Property Get Item(ByVal intIndex As Integer) As SQLSelectTableJoin
Attribute Item.VB_UserMemId = 0

    Set Item = pcolJoins(intIndex)

End Property

Public Function Exists(ByVal objTable As SQLSelectTable) As Boolean

    Dim intIndex As Integer
    
    For intIndex = 1 To pcolJoins.Count
        With Me.Item(intIndex)
            If .LeftTable Is objTable Or .RightTable Is objTable Then
                Exists = True
                Exit For
            End If
        End With
    Next

End Function

Public Sub Delete(ByVal objJoin As SQLSelectTableJoin)

    If Not CollectionRemoveItem(pcolJoins, objJoin) Then
        RaiseError dboErrorObjectDoesNotExist
    End If
    Set objJoin = Nothing

End Sub

Public Property Get Count() As Integer

    Count = pcolJoins.Count

End Property

Public Property Get Enumerator() As IUnknown
Attribute Enumerator.VB_UserMemId = -4

    Set Enumerator = pcolJoins.[_NewEnum]

End Property

Friend Property Get SQL(ByVal eConnectionType As ConnectionTypeEnum) As String

    If Me.Count > 0 Then
        'recurse through the joins from right to left
        SQL = Me.Item(Me.Count).SQL(eConnectionType)
    End If

'    Dim intIndex As Integer
'    Dim strSQL As String
'
'    For intIndex = 1 To Me.Count
'        strSQL = strSQL & Me.Item(intIndex).SQL
'        If intIndex <> Me.Count Then
'            strSQL = strSQL & " "
'        End If
'    Next
'
'    SQL = strSQL

End Property

Private Sub Class_Initialize()
    
    Set pcolJoins = New Collection
    
End Sub
