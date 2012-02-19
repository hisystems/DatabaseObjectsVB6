VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SQLTableField"
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

Public Enum DataTypeEnum
    dboDataTypeTinyInteger              'Integer data from 0 through 255
    dboDataTypeSmallInteger             'Integer data from 2^15 (-32,768) through 2^15 - 1 (32,767)
    dboDataTypeInteger                  'Integer (whole number) data from -2^31 (-2,147,483,648) through 2^31 - 1 (2,147,483,647)
    dboDataTypeBigInteger               'Integer (whole number) data from -2^63 (-9223372036854775808) through 2^63-1 (9223372036854775807)
    dboDataTypeCharacter                '* SQLServer limitation: Fixed-length non-Unicode character data with a maximum length of 8,000 characters
    dboDataTypeUnicodeCharacter         '* SQLServer limitation: Fixed-length Unicode data with a maximum length of 4,000 characters
    dboDataTypeVariableCharacter        '* SQLServer limitation: Variable-length non-Unicode data with a maximum of 8,000 characters
    dboDataTypeUnicodeVariableCharacter '* SQLServer limitation: Variable-length Unicode data with a maximum length of 4,000 characters
    dboDataTypeDecimal                  '* Fixed precision and scale numeric data
    dboDataTypeReal                     'Floating precision number data from -3.40E + 38 through 3.40E + 38  (double)
    dboDataTypeFloat                    'Floating precision number data from -1.79E + 308 through 1.79E + 308 (single / float)
    dboDataTypeSmallMoney               'Monetary data values from -214,748.3648 through +214,748.3647, with accuracy to a ten-thousandth of a monetary unit
    dboDataTypeMoney                    'Monetary data values from -2^63 (-922,337,203,685,477.5808) through 2^63 - 1 (+922,337,203,685,477.5807), with accuracy to a ten-thousandth of a monetary unit
    dboDataTypeBoolean                  'Integer data with either a 1 or 0 value
    dboDataTypeSmallDateTime            'Date and time data from January 1, 1900, through June 6, 2079, with an accuracy of one minute
    dboDataTypeDateTime                 'Date and time data from January 1, 1753, through December 31, 9999, with an accuracy of three-hundredths of a second, or 3.33 milliseconds
    dboDataTypeTimeStamp
    dboDataTypeText                     'Variable-length non-Unicode data with a maximum length of 2^31 - 1 (2,147,483,647) characters
    dboDataTypeUnicodeText              'Variable-length Unicode data with a maximum length of 2^30 - 1 (1,073,741,823) characters
    dboDataTypeBinary                   '* SQLServer limitation: Fixed-length binary data with a maximum length of 8,000 bytes
    dboDataTypeVariableBinary           '* SQLServer limitation: Variable-length binary data with a maximum length of 8,000 bytes
    dboDataTypeImage                    'Variable-length binary data with a maximum length of 2^31 - 1 (2,147,483,647) bytes
End Enum
    
Public Enum KeyTypeEnum
    dboKeyTypeNone = 0
    dboKeyTypePrimary
    dboKeyTypeUnique
End Enum
    
Private pstrName As String
Private peType As DataTypeEnum
Private peKeyType As KeyTypeEnum
Private pintSize As Integer
Private pintScale As Integer
Private pintPrecision As Integer
Private pbAutoIncrements As Boolean
Private pbAcceptsNull As Boolean
Private pvDefault As Variant
'Private pbDefaultExistingRows As Boolean    'sets existing rows to pvDefault (which is the default value for new rows)

Public Property Let Name(ByVal strValue As String)

    pstrName = Trim$(strValue)

End Property

Public Property Get Name() As String

    Name = pstrName
    
End Property

Public Property Get DataType() As DataTypeEnum

    DataType = peType

End Property

Public Property Let DataType(ByVal eValue As DataTypeEnum)

    peType = eValue

End Property

Public Property Let KeyType(ByVal eValue As KeyTypeEnum)

    peKeyType = eValue

    If eValue = dboKeyTypePrimary Then
        'need to set to false as true is the default - see class_initialize
        pbAcceptsNull = False
    End If

End Property

Public Property Get KeyType() As KeyTypeEnum
    
    KeyType = peKeyType
    
End Property

Public Property Let Size(ByVal intValue As Integer)

    EnsureCharacterDataType

    If intValue <= 0 Then
        RaiseError dboErrorInvalidArgument
    End If

    pintSize = intValue
    
End Property

Public Property Get Size() As Integer
    
    EnsureCharacterDataType
    Size = pintSize
    
End Property

'Public Property Let DefaultExistingRows(ByVal bValue As Boolean)
'
'    pbDefaultExistingRows = bValue
'
'End Property
'
'Public Property Get DefaultExistingRows() As Boolean
'
'    DefaultExistingRows = pbDefaultExistingRows
'
'End Property

Public Property Let ScaleLength(ByVal intValue As Integer)

    EnsureDecimalDataType

    If intValue <= 0 Then
        RaiseError dboErrorInvalidArgument
    End If

    pintScale = intValue

End Property

Public Property Get ScaleLength() As Integer
    
    EnsureDecimalDataType
    ScaleLength = pintScale

End Property

Public Property Let Precision(ByVal intValue As Integer)
    
    EnsureDecimalDataType

    If intValue <= 0 Then
        RaiseError dboErrorInvalidArgument
    End If
    
    pintPrecision = intValue
    
End Property

Public Property Get Precision() As Integer
    
    EnsureDecimalDataType
    Precision = pintPrecision
    
End Property

Public Property Let AutoIncrements(ByVal bValue As Boolean)

    EnsureIntegerDataType
    pbAutoIncrements = bValue
    pbAcceptsNull = Not bValue

End Property

Public Property Get AutoIncrements() As Boolean

    EnsureIntegerDataType
    AutoIncrements = pbAutoIncrements

End Property

Public Property Let AcceptsNull(ByVal bValue As Boolean)

    pbAcceptsNull = bValue

End Property

Public Property Get AcceptsNull() As Boolean

    AcceptsNull = pbAcceptsNull

End Property

Public Property Let Default(ByVal vValue As Variant)

    pvDefault = vValue
    
    'default to updating the existing rows with the default value also
    'pbDefaultExistingRows = True

End Property

Public Property Get Default() As Variant

    Default = pvDefault

End Property

Private Function DataTypeString( _
    ByVal eConnection As ConnectionTypeEnum, _
    ByVal eDataType As DataTypeEnum) As String
    
    Dim strDataType As String
    
    Select Case eDataType
        Case dboDataTypeTinyInteger
            Select Case eConnection
                Case dboConnectionTypeMicrosoftAccess
                    strDataType = "BYTE"
                Case dboConnectionTypeMySQL
                    strDataType = "TINYINT UNSIGNED"
                Case dboConnectionTypeSQLServer
                    strDataType = "TINYINT"
            End Select
        Case dboDataTypeSmallInteger
            Select Case eConnection
                Case dboConnectionTypeMicrosoftAccess
                    strDataType = "SMALLINT"
                Case dboConnectionTypeMySQL
                    strDataType = "SMALLINT"
                Case dboConnectionTypeSQLServer
                    strDataType = "SMALLINT"
            End Select
        Case dboDataTypeInteger
            Select Case eConnection
                Case dboConnectionTypeMicrosoftAccess
                    strDataType = "INTEGER"
                Case dboConnectionTypeMySQL
                    strDataType = "INT"
                Case dboConnectionTypeSQLServer
                    strDataType = "INTEGER"
            End Select
        Case dboDataTypeBigInteger
            Select Case eConnection
                Case dboConnectionTypeMicrosoftAccess
                    strDataType = "NUMERIC(19,0)"
                Case dboConnectionTypeMySQL
                    strDataType = "BIGINT"
                Case dboConnectionTypeSQLServer
                    strDataType = "BIGINT"
            End Select
        Case dboDataTypeCharacter
            Select Case eConnection
                Case dboConnectionTypeMicrosoftAccess
                    strDataType = "TEXT(" & pintSize & ")"
                Case dboConnectionTypeMySQL
                    strDataType = "CHAR(" & pintSize & ")"
                Case dboConnectionTypeSQLServer
                    strDataType = "CHAR(" & pintSize & ")"
            End Select
        Case dboDataTypeUnicodeCharacter
            Select Case eConnection
                Case dboConnectionTypeMicrosoftAccess
                    strDataType = "TEXT(" & pintSize & ")"  'Unicode is only supported in Microsoft Access 2000+
                Case dboConnectionTypeMySQL
                    strDataType = "NCHAR(" & pintSize & ")"
                Case dboConnectionTypeSQLServer
                    strDataType = "NCHAR(" & pintSize & ")"
            End Select
        Case dboDataTypeVariableCharacter
            Select Case eConnection
                Case dboConnectionTypeMicrosoftAccess
                    strDataType = "TEXT(" & pintSize & ")"
                Case dboConnectionTypeMySQL
                    strDataType = "VARCHAR(" & pintSize & ")"
                Case dboConnectionTypeSQLServer
                    strDataType = "VARCHAR(" & pintSize & ")"
            End Select
        Case dboDataTypeUnicodeVariableCharacter
            Select Case eConnection
                Case dboConnectionTypeMicrosoftAccess
                    strDataType = "TEXT(" & pintSize & ")"                    'Unicode is only supported in Microsoft Access 2000+
                Case dboConnectionTypeMySQL
                    strDataType = "NVARCHAR(" & pintSize & ")"
                Case dboConnectionTypeSQLServer
                    strDataType = "NVARCHAR(" & pintSize & ")"
            End Select
        Case dboDataTypeDecimal
            Select Case eConnection
                Case dboConnectionTypeMicrosoftAccess
                    strDataType = "NUMERIC(" & pintPrecision & "," & pintScale & ")"
                Case dboConnectionTypeMySQL
                    strDataType = "DECIMAL(" & pintPrecision & "," & pintScale & ")"
                Case dboConnectionTypeSQLServer
                    strDataType = "NUMERIC(" & pintPrecision & "," & pintScale & ")"
            End Select
        Case dboDataTypeReal
            Select Case eConnection
                Case dboConnectionTypeMicrosoftAccess
                    strDataType = "REAL"
                Case dboConnectionTypeMySQL
                    strDataType = "FLOAT"
                Case dboConnectionTypeSQLServer
                    strDataType = "REAL"
            End Select
        Case dboDataTypeFloat
            Select Case eConnection
                Case dboConnectionTypeMicrosoftAccess
                    strDataType = "FLOAT"
                Case dboConnectionTypeMySQL
                    strDataType = "DOUBLE"
                Case dboConnectionTypeSQLServer
                    strDataType = "FLOAT"
            End Select
        Case dboDataTypeSmallMoney
            Select Case eConnection
                Case dboConnectionTypeMicrosoftAccess
                    strDataType = "NUMERIC(10,4)"
                Case dboConnectionTypeMySQL
                    strDataType = "DECIMAL(10,4)"
                Case dboConnectionTypeSQLServer
                    strDataType = "SMALLMONEY"
            End Select
        Case dboDataTypeMoney
            Select Case eConnection
                Case dboConnectionTypeMicrosoftAccess
                    strDataType = "NUMERIC(19,4)"
                Case dboConnectionTypeMySQL
                    strDataType = "DECIMAL(19,4)"
                Case dboConnectionTypeSQLServer
                    strDataType = "MONEY"
            End Select
        Case dboDataTypeBoolean
            Select Case eConnection
                Case dboConnectionTypeMicrosoftAccess
                    strDataType = "YESNO"
                Case dboConnectionTypeMySQL
                    strDataType = "BIT"
                Case dboConnectionTypeSQLServer
                    strDataType = "BIT"
            End Select
        Case dboDataTypeSmallDateTime
            Select Case eConnection
                Case dboConnectionTypeMicrosoftAccess
                    strDataType = "DATETIME"
                Case dboConnectionTypeMySQL
                    strDataType = "DATETIME"
                Case dboConnectionTypeSQLServer
                    strDataType = "SMALLDATETIME"
            End Select
        Case dboDataTypeDateTime
            Select Case eConnection
                Case dboConnectionTypeMicrosoftAccess
                    strDataType = "DATETIME"
                Case dboConnectionTypeMySQL
                    strDataType = "DATETIME"
                Case dboConnectionTypeSQLServer
                    strDataType = "DATETIME"
            End Select
        Case dboDataTypeTimeStamp
            Select Case eConnection
                Case dboConnectionTypeMicrosoftAccess
                    RaiseError dboErrorNotSupported, "TIMESTAMP"
                Case dboConnectionTypeMySQL
                    strDataType = "TIMESTAMP"
                Case dboConnectionTypeSQLServer
                    strDataType = "TIMESTAMP"
            End Select
        Case dboDataTypeText
            Select Case eConnection
                Case dboConnectionTypeMicrosoftAccess
                    strDataType = "MEMO"
                Case dboConnectionTypeMySQL
                    strDataType = "LONGTEXT"
                Case dboConnectionTypeSQLServer
                    strDataType = "TEXT"
            End Select
        Case dboDataTypeUnicodeText
            Select Case eConnection
                Case dboConnectionTypeMicrosoftAccess
                    strDataType = "MEMO"    'Unicode is only supported in Microsoft Access 2000+
                Case dboConnectionTypeMySQL
                    strDataType = "LONGTEXT CHARACTER SET UTF8"
                Case dboConnectionTypeSQLServer
                    strDataType = "NTEXT"
            End Select
        Case dboDataTypeBinary
            Select Case eConnection
                Case dboConnectionTypeMicrosoftAccess
                    strDataType = "OLEOBJECT"
                Case dboConnectionTypeMySQL
                    strDataType = "BLOB"
                Case dboConnectionTypeSQLServer
                    strDataType = "BINARY"
            End Select
        Case dboDataTypeVariableBinary
            Select Case eConnection
                Case dboConnectionTypeMicrosoftAccess
                    strDataType = "OLEOBJECT"
                Case dboConnectionTypeMySQL
                    strDataType = "BLOB"
                Case dboConnectionTypeSQLServer
                    strDataType = "VARBINARY"
            End Select
        Case dboDataTypeImage
            Select Case eConnection
                Case dboConnectionTypeMicrosoftAccess
                    strDataType = "IMAGE"
                Case dboConnectionTypeMySQL
                    strDataType = "LONGBLOB"
                Case dboConnectionTypeSQLServer
                    strDataType = "IMAGE"
            End Select
    End Select
    
    DataTypeString = strDataType
    
End Function

Friend Property Get SQL(ByVal eConnectionType As ConnectionTypeEnum, ByVal bOnlyFieldName As Boolean) As String

    Dim strName As String
    Dim strDataType As String
    Dim strColumnOptions As String
    Dim strSQL As String

    If Me.Name = vbNullString Then
        RaiseError dboErrorGeneral, "Field Name has not been set."
    End If
    
    strName = SQLConvertIdentifierName(Me.Name, eConnectionType)
    
    If bOnlyFieldName Then
        strSQL = strName
    Else
        strDataType = DataTypeString(eConnectionType, peType)
        strColumnOptions = ColumnOptions(eConnectionType)
        strSQL = strName & " " & strDataType & strColumnOptions
    End If
    
    SQL = strSQL
 
End Property

Private Function ColumnOptions( _
    ByVal eConnection As ConnectionTypeEnum) As String

    Dim colOptions As Collection
    Dim vOption As Variant
    Dim strOptions As String
    
    Set colOptions = New Collection
    
    If pbAcceptsNull Then
        colOptions.Add "NULL"
    Else
        colOptions.Add "NOT NULL"
    End If
    
    If DefaultValueIsSet Then
        colOptions.Add "DEFAULT " & SQLConvertValue(pvDefault, eConnection)
    End If
    
    If IsIntegerDataType And pbAutoIncrements Then
        Select Case eConnection
            Case dboConnectionTypeMicrosoftAccess
                colOptions.Add "IDENTITY"
            Case dboConnectionTypeMySQL
                colOptions.Add "AUTO_INCREMENT"
                'must be set to a key
                If peKeyType = dboKeyTypeNone Then
                    peKeyType = dboKeyTypeUnique
                End If
            Case dboConnectionTypeSQLServer
                colOptions.Add "IDENTITY"
        End Select
    End If

    Select Case peKeyType
        Case dboKeyTypePrimary
            colOptions.Add "PRIMARY KEY"
        Case dboKeyTypeUnique
            colOptions.Add "UNIQUE"
    End Select


    For Each vOption In colOptions
        strOptions = strOptions & " " & vOption
    Next
    
    ColumnOptions = strOptions

End Function

Friend Property Get DefaultValueIsSet() As Boolean
    
    DefaultValueIsSet = Not IsEmpty(pvDefault)
    
End Property

Private Sub EnsureDecimalDataType()

    If peType <> dboDataTypeDecimal Then
        RaiseError dboErrorMethodOrPropertyLocked, "First set TheType to dboDataTypeDecimal"
    End If

End Sub

Private Sub EnsureCharacterDataType()

    Select Case peType
        Case dboDataTypeCharacter, dboDataTypeUnicodeCharacter, _
             dboDataTypeVariableCharacter, dboDataTypeUnicodeVariableCharacter
        Case Else
            RaiseError dboErrorMethodOrPropertyLocked, "Data Type is not character based"
    End Select

End Sub

Private Sub EnsureIntegerDataType()

    If Not IsIntegerDataType Then
        RaiseError dboErrorMethodOrPropertyLocked
    End If

End Sub

Private Property Get IsIntegerDataType() As Boolean

    Select Case peType
        Case dboDataTypeBigInteger, dboDataTypeTinyInteger, dboDataTypeSmallInteger, dboDataTypeInteger
            IsIntegerDataType = True
        Case Else
            IsIntegerDataType = False
    End Select

End Property

Private Sub Class_Initialize()
    
    peType = dboDataTypeVariableCharacter
    pintSize = 1
    pintPrecision = 18
    pintScale = 0
    pbAcceptsNull = True

End Sub
