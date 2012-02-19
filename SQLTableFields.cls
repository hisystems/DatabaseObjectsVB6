VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SQLTableFields"
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

Private Const pcintAlterModeUninitialized As Integer = -1

Private pcolFields As Collection
Private peAlterMode As SQLTableFieldsAlterModeEnum

Friend Property Let AlterMode(ByVal eValue As SQLTableFieldsAlterModeEnum)
    
    peAlterMode = eValue
    
End Property

Friend Property Get AlterMode() As SQLTableFieldsAlterModeEnum

    AlterMode = peAlterMode

End Property

Public Function Add( _
    Optional ByVal strFieldName As String, _
    Optional ByVal eDataType As DataTypeEnum = dboDataTypeVariableCharacter, _
    Optional ByVal intSize As Integer = 0) As SQLTableField
    
    EnsureAlterModeValid dboTableFieldsModeAdd
    
    Dim objField As SQLTableField
    Set objField = New SQLTableField
    
    With objField
        .Name = strFieldName
        .DataType = eDataType
        If intSize > 0 Then
            .Size = intSize
        End If
    End With

    pcolFields.Add objField
    Set Add = objField

End Function

Public Property Get Item(ByVal strFieldName As String) As SQLTableField
Attribute Item.VB_UserMemId = 0

    EnsureAlterModeValid dboTableFieldsModeAlter
    
    Dim intIndex As Integer
    
    strFieldName = Trim$(strFieldName)
    intIndex = FieldNameIndex(strFieldName)
    
    If intIndex = 0 Then
        Set Item = New SQLTableField
        Item.Name = strFieldName
        pcolFields.Add Item
    Else
        Set Item = pcolFields(intIndex)
    End If
    
End Property

Public Sub Drop(ByVal strFieldName As String)

    EnsureAlterModeValid dboTableFieldsModeDrop
    
    Dim objField As SQLTableField
    Set objField = New SQLTableField
    
    objField.Name = strFieldName
    
    If FieldNameIndex(strFieldName) = 0 Then
        pcolFields.Add objField
    Else
        RaiseError dboErrorInvalidArgument, "Field '" & strFieldName & "' already exists"
    End If

End Sub

Public Property Get Count() As Integer

    Count = pcolFields.Count

End Property

Public Property Get Enumerator() As IUnknown
Attribute Enumerator.VB_UserMemId = -4

    Set Enumerator = pcolFields.[_NewEnum]

End Property

Private Function FieldNameIndex( _
    ByVal strFieldName As String) As Integer
    
    Dim intIndex As Integer
    
    strFieldName = Trim$(strFieldName)
    
    For intIndex = 1 To pcolFields.Count
        If StrComp(pcolFields(intIndex).Name, strFieldName, vbTextCompare) = 0 Then
            FieldNameIndex = intIndex
            Exit For
        End If
    Next
    
End Function

Private Sub Class_Initialize()

    Set pcolFields = New Collection
    peAlterMode = pcintAlterModeUninitialized
    
End Sub

Friend Property Get SQL(ByVal eConnectionType As ConnectionTypeEnum, Optional ByVal bIncludeMode As Boolean = True) As String

    Const cstrSeperator As String = ", "

    Dim strSQL As String
    Dim objField As SQLTableField
    Dim bOnlyFieldName As Boolean
    Dim strMode As String
    
    bOnlyFieldName = peAlterMode = dboTableFieldsModeDrop
    
    'Include mode when altering a table, otherwise when creating a table the mode is not required.
    If bIncludeMode Then
        Select Case peAlterMode
            Case dboTableFieldsModeAdd:
                strMode = "ADD"
            Case dboTableFieldsModeAlter:
                Select Case eConnectionType
                    Case dboConnectionTypeMySQL
                        strMode = "MODIFY COLUMN"
                    Case Else
                        strMode = "ALTER COLUMN"
                End Select
            Case dboTableFieldsModeDrop:
                strMode = "DROP COLUMN"
        End Select
    
        Select Case eConnectionType
            Case dboConnectionTypeMicrosoftAccess, dboConnectionTypeSQLServer
                strSQL = strMode & " "
        End Select
    End If
    
    For Each objField In pcolFields
        If bIncludeMode Then
            Select Case eConnectionType
                Case dboConnectionTypeMySQL
                    strSQL = strSQL & strMode & " "
            End Select
        End If
        strSQL = strSQL & objField.SQL(eConnectionType, bOnlyFieldName) & cstrSeperator
    Next
        
    SQL = Left$(strSQL, Len(strSQL) - Len(cstrSeperator))    'remove the last comma and space

End Property

Private Sub EnsureAlterModeValid(ByVal eAlterMode As SQLTableFieldsAlterModeEnum)

    'if the alter mode hasn't been set then any of the modes are valid
    If peAlterMode = pcintAlterModeUninitialized Then
        peAlterMode = eAlterMode
    Else
        If eAlterMode <> peAlterMode Then
            RaiseError dboErrorMethodOrPropertyLocked, "Cannot mix " & AlterModeDescription(peAlterMode) & " fields and " & AlterModeDescription(eAlterMode) & " fields into one SQL statement"
        End If
    End If

End Sub

Private Function AlterModeDescription(ByVal eAlterMode As SQLTableFieldsAlterModeEnum) As String
    
    Select Case eAlterMode
        Case dboTableFieldsModeAdd: AlterModeDescription = "adding"
        Case dboTableFieldsModeAlter: AlterModeDescription = "altering"
        Case dboTableFieldsModeDrop: AlterModeDescription = "dropping"
    End Select

End Function

