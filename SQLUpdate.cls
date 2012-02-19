VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SQLUpdate"
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

Implements ISQLStatement

Private pobjFields As SQLFieldValues
Private pobjCopyFields As SQLUpdateCopyFields
Private pobjConditions As SQLConditions

Public ConnectionType As ConnectionTypeEnum
Public TableName As String

Public Property Set Fields(ByVal objFields As SQLFieldValues)

    If objFields Is Nothing Then
        RaiseError dboErrorObjectIsNothing
    End If

    Set pobjFields = objFields

End Property

Public Property Get Fields() As SQLFieldValues

    Set Fields = pobjFields

End Property

Public Property Get CopyFields() As SQLUpdateCopyFields

    Set CopyFields = pobjCopyFields

End Property

Public Property Set CopyFields(ByVal objValue As SQLUpdateCopyFields)

    If objValue Is Nothing Then
        RaiseError dboErrorObjectIsNothing
    End If

    Set pobjCopyFields = objValue

End Property

Public Property Get Where() As SQLConditions
   
    Set Where = pobjConditions

End Property

Public Property Set Where(ByVal objValue As SQLConditions)

    Set pobjConditions = objValue
    
End Property

Public Property Get SQL() As String
Attribute SQL.VB_UserMemId = 0

    Dim intIndex As Integer
    Dim strSQL As String
    Dim strFieldValues As String
    Dim strConditions As String
    
    If Trim$(TableName) = vbNullString Then
        RaiseError dboErrorGeneral, "TableName property has not been set."
    End If
    
    If pobjFields.Count = 0 And pobjCopyFields.Count = 0 Then
        RaiseError dboErrorGeneral, "Field values have not been set."
    End If
    
    For intIndex = 1 To pobjFields.Count
        With pobjFields(intIndex)
            'Check the field name has been set. Can't really check whether the value has been set or not.
            If Trim$(.Name) = vbNullString Then
                RaiseError dboErrorGeneral, "Field name not set"
            End If
            strFieldValues = strFieldValues & _
                SQLConvertIdentifierName(.Name, Me.ConnectionType) & " = " & SQLConvertValue(.Value, Me.ConnectionType)
        End With
        If intIndex <> pobjFields.Count Then
            strFieldValues = strFieldValues & ", "
        End If
    Next
    
    If pobjFields.Count > 0 And pobjCopyFields.Count > 0 Then
        strFieldValues = strFieldValues & ", "
    End If
    
    For intIndex = 1 To pobjCopyFields.Count
        With pobjCopyFields(intIndex)
            strFieldValues = strFieldValues & _
                SQLConvertIdentifierName(.DestinationFieldName, Me.ConnectionType) & " = " & SQLConvertIdentifierName(.SourceFieldName, Me.ConnectionType)
        End With
        If intIndex <> pobjCopyFields.Count Then
            strFieldValues = strFieldValues & ", "
        End If
    Next
    
    strSQL = _
        "UPDATE " & SQLConvertIdentifierName(Trim$(Me.TableName), Me.ConnectionType) & " " & _
        "SET " & strFieldValues
    
    If Not pobjConditions Is Nothing Then
        strConditions = pobjConditions.SQL(Me.ConnectionType)
        If strConditions <> vbNullString Then
            strSQL = strSQL & " WHERE " & strConditions
        End If
    End If
        
    SQL = strSQL
    
End Property

Private Sub Class_Initialize()
    
    ConnectionType = modMisc.ConnectionType
    Set pobjFields = New SQLFieldValues
    Set pobjCopyFields = New SQLUpdateCopyFields
    Set pobjConditions = New SQLConditions
    
End Sub

Private Property Get ISQLStatement_ConnectionType() As ConnectionTypeEnum

    ISQLStatement_ConnectionType = Me.ConnectionType
    
End Property

Private Property Let ISQLStatement_ConnectionType(ByVal RHS As ConnectionTypeEnum)

    Me.ConnectionType = RHS
    
End Property

Private Property Get ISQLStatement_SQL() As String

    ISQLStatement_SQL = Me.SQL
    
End Property
