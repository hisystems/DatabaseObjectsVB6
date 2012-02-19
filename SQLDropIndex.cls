VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SQLDropIndex"
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

Public ConnectionType As ConnectionTypeEnum

Private pstrName As String
Private pstrTableName As String

Public Property Let Name(ByVal strValue As String)

    pstrName = Trim$(strValue)

End Property

Public Property Get Name() As String

    Name = pstrName
    
End Property

Public Property Get TableName() As String
    
    TableName = pstrTableName

End Property

Public Property Let TableName(ByVal strValue As String)

    pstrTableName = Trim$(strValue)

End Property

Public Property Get SQL() As String

    If Me.Name = vbNullString Then
        RaiseError dboErrorGeneral, "Index Name has not been set."
    End If

    If Me.TableName = vbNullString Then
        RaiseError dboErrorGeneral, "TableName has not been set."
    End If

    If Me.ConnectionType = dboConnectionTypeSQLServer Then
        SQL = "DROP INDEX " & SQLConvertIdentifierName(Me.TableName, Me.ConnectionType) & "." & SQLConvertIdentifierName(Me.Name, Me.ConnectionType)
    Else
        SQL = "DROP INDEX " & SQLConvertIdentifierName(Me.Name, Me.ConnectionType) & " ON " & SQLConvertIdentifierName(Me.TableName, Me.ConnectionType)
    End If

End Property

Private Sub Class_Initialize()
    
    Me.ConnectionType = modMisc.ConnectionType

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
