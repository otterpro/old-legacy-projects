VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TheAdoDb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=============================================================================
'
'=============================================================================
' Requires TheFileSystem.class, the.bas

'createTable(tableName, %fields)
'dropTable(tableName)

'insert() not impl yet
'update() not impl yet


'Low-Level
'execute(sqlStmt)
    'Try to use createTable(), dropTable(), insert(), and other
    'high-level method instead
    'execute("CREATE TABLE myTbl (id INDEX, [name] VARCHAR(30));")
Option Explicit

'member var
Private myDb As ADODB.Connection
Private myCommand As ADODB.Command
Private Const CONFIG_KEY_DEFAULT = "key"
Private Const CONFIG_VALUE_DEFAULT = "value"
'Dim myDbIsOpen As Boolean   ' true if db is currently open.

Private myTables As New TheList
    ' all the tables that are contained in this Database
    ' many DB format can have more than 1 table (MDB (tables), Excel (Sheets))
    
Private myCurrentTable As TheAdoTable2
    'reference to the currently selected table.
    ' If there is only 1 table, this will always point to that even if
    ' the user tries to change it.

Private mySelectTableName As String 'SELECT ... FROM Table
Private mySelectFieldName As String 'SELECT name,date FROM ...
Private mySelectWhere As String     'SELECT ... WHERE a>0

Private myMetaDataTable As New TheMetaDataTable

' EXCEL -related
Const DefaultSheet = "Sheet1"

Public Enum MERGE_OPERATION
    MERGE_ADD
    MERGE_SUBTRACT
    'MERGE_INTERSECT
    'MERGE_INVERSE_INTERSECT
End Enum

Private Sub Class_Initialize()
    'myDbIsOpen = False
End Sub

Private Sub Class_Terminate()
    closeFile
End Sub

Public Function getConnection() As ADODB.Connection
    Set getConnection = myDb
End Function

Public Function openFile(ByVal filename As String) As Boolean
    closeFile 'make sure to close any db if open.
    On Error GoTo OPEN_DB_FAILED
    If Dir(filename) = "" Then
        eprint "WARNING: Database file " & filename & " was not found. " _
            & "This application will not function properly without it."
        Exit Function
    End If

    Set myDb = New ADODB.Connection
    
    ' determine the type of db, whether Excel, MDB(Access), or CSV(not impl)
    ' by looking at the filename's extension
    Dim fileSys As New TheFileSystem
    If fileSys.hasExtension(filename, ".xls") Then
        myDb.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
        "Data Source=" & filename & ";" & _
            "Extended Properties=Excel 8.0;"
        '"DATABASE=" & filename & _
        ''"Excel 5.0;" & "TABLE=" & DefaultSheet & ";HDR=YES;"
        
    ElseIf fileSys.hasExtension(filename, ".mdb") Then
        'ACCESS
        myDb.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" _
                & "Data Source=" & filename & ";Persist Security Info=False;" _
                & "Mode=Read|Write"
        'samsDb.CursorLocation = adUseClient
    Else
        eprint "TheAdoDb::openDb() - The file is an unknown type."
        Exit Function   ' unknown type
    End If
    
    myDb.Open
    myMetaDataTable.openFile Me
    ' check for error
    Dim state As Boolean
    state = IIf((myDb.state = adStateClosed), False, True)
    openFile = state 'T/F depending on the State
    ' MsgBox samsDb.State
OPEN_DB_FAILED:
    openFile = False

End Function

Public Sub closeFile()
On Error GoTo DB_NOT_OPEN
    'If myDbIsOpen Then
        'myMetaDataTable.closeFile
        myDb.Close
        Set myDb = Nothing
'        myDbIsOpen = False
    'End If
DB_NOT_OPEN:
    ' do nothing since it was never open.
End Sub

Public Function openTable(ByVal sqlStatement As String _
                         ) As TheAdoTable2
    'If IsMissing(sqlStatement) Then
    '    sqlStatement = "SELECT * FROM " & tableName
    'End If
    Dim table As New TheAdoTable2
    Dim result As Boolean
    result = table.openFile(sqlStatement, Me)
    Debug.Assert result
    myTables.add table
    Set openTable = table
End Function

Public Function closeTable(table As TheAdoTable2)
    myTables.remove table
    closeTable = table.closeFile
End Function

'table= either the table name or the TheTable object.
' Howver, deleting obj table is not implemented
Public Function dropTable(table) As Boolean
    Dim sqlText As String
    If isString(table) Then
        sqlText = "DROP TABLE " & table
        dropTable = execute(sqlText)
    Else
        dropTable = False
        eprint "dropTable(table), dropping TheAdoTable object is not implemented yet."
    End If
End Function

Public Function createTable(ByVal tableName As String, fieldDict As TheDictionary) As Boolean
    Dim sqlText As String
    sqlText = TheSQLTool.createTable(tableName, fieldDict)
    createTable = execute(sqlText)
End Function

Public Function getMetaData(ByVal tableName As String) As TheDictionary
    Set getMetaData = myMetaDataTable.getFieldAndType(tableName)
End Function

Public Function execute(ByVal sqlStatement As String) As Boolean
    'create myCommand everytime since in order to save the action, it
    'needs to go out of scope or be set to Nothing.
    On Error GoTo EXECUTE_FAILED
    Set myCommand = New ADODB.Command
    Set myCommand.ActiveConnection = myDb
    myCommand.CommandText = sqlStatement
    myCommand.execute
    Set myCommand = Nothing     'must flush the cache since
                                'it doesn't save until cmd is gone.
    execute = True
    Exit Function
    
EXECUTE_FAILED:
    execute = False
    warn "TheAdoDB::execute(" & sqlStatement & ") Failed. ErrorMsg=" & _
                            Err.Description
End Function

'=============================================================================
' HELPER METHODS --
'=============================================================================
' Put them in Table
'Public Function findAndGet(ByVal fieldName, ByVal match, ByVal fieldToReturn)
'    If findRecord(fieldName, "=", match) Then
'        findAndGet = getValue(fieldToReturn)
'    Else
'        findAndGet = Null
'    End If
'End Function
'
'
'Public Sub findAndSet(ByVal fieldName, ByVal match, ByVal fieldToSet, ByVal value)
'    If findRecord(fieldName, "=", match) Then
'        setValue fieldToSet, value
'    End If
'End Sub


'=============================================================================
' CONFIG Table based on ADO TABLE
'TODO: create TheAdoConfigTable and rename them get(), set()
'=============================================================================


'Public Function getConfig(ByVal key)
'    getConfig = findAndGet(CONFIG_KEY_DEFAULT, key, CONFIG_VALUE_DEFAULT)
'End Function
'
'Public Sub setConfig(ByVal key, ByVal value)
'    findAndSet CONFIG_KEY_DEFAULT, key, CONFIG_VALUE_DEFAULT, value
'End Sub
'
''move this to TheAdoDB / TheDB
'Public Function createDuplicate(ByVal tableName As String, _
'                            srcTableName As Variant, _
'                            Optional db As TheAdoDb, _
'                            Optional copyData = False) As Boolean
'End Function

