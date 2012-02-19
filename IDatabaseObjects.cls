VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IDatabaseObjects"
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

Public Function ItemInstance() As IDatabaseObject

End Function

Public Function Subset() As SQLConditions

End Function

Public Function TableName() As String

End Function

Public Function KeyFieldName() As String

End Function

Public Function DistinctFieldName() As String

End Function

Public Function DistinctFieldAutoIncrements() As Boolean

End Function

Public Function OrderBy() As SQLSelectOrderByFields

End Function

Public Function TableJoins(ByVal objPrimaryTable As SQLSelectTable, ByVal objTables As SQLSelectTables) As SQLSelectTableJoins

End Function
