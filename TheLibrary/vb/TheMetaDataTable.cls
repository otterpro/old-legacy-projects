VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TheMetaDataTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=============================================================================
'
'TODO: Convert to TheConfigTable  and move the meta-data related stuff to
' TheDB

'filename "config.ini", "config" in table (aka metadata)
' overview: we first define the meta-data in the first few rows.
            ' We know it is a metadata if type# (aka category#) is -1.
            ' Then key becomes type#, and value becomes description of type.
'Table Fields:
'id# = autonumber
'type# = long
    'type#=-1, means category
    'type#=0 means it is no category, ie it represnets category.
    'type#= >0, it is user-defined category.
'key = varchar(64)
'value = varchar(255)
'
'id#    type#    key         value
'0      -1      1           Fruit   'cat1=Fruit type
'1      -1      2           Candy   'cat2=Candy type
'2      -1      5           Meat    'cat5=meat type
'3      0       test        Nothing 'no category, test=Nothing
'4      2       snack       Snickers    'snack=Snickers, category=Candy(2)
'5      5       lunch       BBQ Ribs    'lunch=BBQ Ribs, category=Meat(5)
'
'To use it as db meta-data table
'0      -1      "1"         "fieldName,Type for product Table"
'1      -1      "2"         "fieldName,Type for employee table"
'2      -1      "10"        "config for this app"
'3      1       "id"        "COUNTER PRIMARY KEY"
'4      1       "name"      "VARCHAR(64)"
'5      1       "note"      "TEXT"
'6      2       "id"        "COUNTER PRIMARY KEY"
'7      2       "name"      "VARCHAR(128)"
'8      2       "position"  "INTEGER"
'9      2       "dob"       "DATE"
'10     10      "path"      "c:/app/data.dat"
'11     10      "backup"    "6/1/2003"

' TO Represent Simple TREE structure
'0      -1      "1"         "Root"
'1      -1      "2"         "Node"
'2      -1      "3"         "Special"
'3      1       "Root"      ""
'
'
'
'
'
'
'
'

'=============================================================================
' THis is a member in TheDB. TheTable should be able to access it.
Option Explicit
Private Const ERROR_CODE = 1001
Private Const META_DATA_TABLE = "metaData"  'looks for metaData table.

' META_DATA TABLE's FIELD NAMES. Table must have 3 fields.
Private Const TABLE_NAME = "tableName"
Private Const FIELD_NAME = "fieldName"
Private Const FIELD_TYPE = "fieldType"

Private myDb As Object
    'db object such as TheAdoDb.
Private myTable As TheAdoTable2
    ' table obj such as TheAdoTable

    
Public Function openFile(db, Optional tableName As String)
    On Error GoTo OPEN_FILE_FAILED
    Dim sqlStatement As String
    Set myDb = db
    If IsEmpty(myDb) Or IsNull(myDb) Then
        eprint "TheMetaDataTable::openFile() myDB is not defined. Must define it before calling this."
        Err.Raise ERROR_CODE
    End If
    
    openFile = True
    Exit Function
OPEN_FILE_FAILED:
    openFile = False
    eprint "MetaDataTable::open() FAILED. MetaData table might be missing"
End Function

'Public Function closeFile() As Boolean
    'closeFile = myDb.closeTable(myTable)
'End Function

'aka getColumnName()
'getFieldAndType("employee")    'get metadata for the employee table from "metadata" table.
Public Function getFieldAndType(ByVal tableName As String) As TheDictionary
    Dim i
    Dim dict As New TheDictionary
    
    openTable tableName
    myTable.gotoFirst
    Dim fieldName As String, fieldType As String
    While Not myTable.isEof()
        fieldName = myTable.getValue(FIELD_NAME)
        fieldType = myTable.getValue(FIELD_TYPE)
        dict.setValue fieldName, fieldType
        myTable.gotoNext
    Wend
    Set getFieldAndType = dict
    closeTable
End Function

Private Sub openTable(ByVal tableName As String)
    Dim sqlStatement As String
    sqlStatement = "SELECT * FROM " & META_DATA_TABLE & " WHERE " & _
                                TABLE_NAME & " = '" & tableName & "'"
    Set myTable = myDb.openTable(sqlStatement)
End Sub

Private Sub closeTable()
    myTable.closeFile
End Sub
