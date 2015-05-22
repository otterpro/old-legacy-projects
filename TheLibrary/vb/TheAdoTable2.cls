VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TheAdoTable2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=============================================================================
'
'=============================================================================

' There are 2 ways of doing things. 1. SQL. 2. ADO.
' ADO allows you to have a cursor that you can move around, sometimes
'   even backward. Also allows to count the # of record, etc.
'
' Executes the SQL Statement and stores the result in myTable.
' ADO allows the data to be updated in the result (myTable).
'
'addRow()
    'adds new empty record.
'addRow([keys],[values])
    'currently only supports "key","value" STRING pair or
    '(keys, values), keys and values are collections or arrays
    'Remember, it is NOT add([key1,value1],[key2,value2],...)
    '    Rather , it Is add([key1,key2,..], [value1,value2,...])
    'addRow("name","Dave")
    'addRow(myFields, myValues)
    'addRow     'just adds new record.
'copyCurrentRow(srcTable,[ignoreEmptyText=T])
    ' copies the content of current row from srcTable.
    ' It does not add new row. Instead, it only changes the curent row
    ' of this table. If adding new row, use add() instead.
    ' must match field name. myTable."name" = foreignTable."name".
    ' if the field doesn't exist, it ignores it and prints warning msg.
'copy(srcTable)
    'deep copy
    'This is different than copyRow(). copyRow() copies 1 row.
    'This copies all rows. Same as .clear() + append(srcTable)
    'to create an entirely new table and copy, use .createDuplicate() instead.
'clone()
    'ADO's clone() gets a reference copy but it is not synchrnoized after original
    'table does .Requery().
    'sort of like shallow copy. Get a reference copy of itself.
    'Don't need to openFile() if cloning.
    'set myNewTable = currentTable.clone() <-not tested
    'or
    'myNewTable.clone(currentTable) <- not tested.
    '
'append(srcTable)
    'appends all the srcTable's data to this data.
    'this = this + srcTable
'createDuplicate(newTableName, srcTable [,db] [,copyData=F])
    'Can't implement it as a generic method because the meta-data (schema) is not easily accessible.
    'Although it can be obtained using ado.openSchema(), it is too complicated
    'and non-portable. Instead, using a makeshift solution. It is expecting
    'the meta-data in the table called "metadata". See TheMetaData Class.
    'Create a new table and copy the content of this table (if copyData=T)
    ' srcTable = name of source table or the reference to currently opened TheAdoTable
    ' copyData = False (Default) copies only the structure of the table
    '           True = copies structure and data.
    ' db = optional TheAdoDb. If not defined, it uses srcTable's adoDB.
    ' createDuplicate("newEmployee","employee")   ' looks for "employee" table and copies it.
    ' createDuplicate("newSales", mySalesTable) ' mySalesTable is an opened AdoTable
    ' createDuplicate("AnotherCar","car",copyData=True) 'dup and copy in 1.
'merge()
    ' Either adds or subtracts "quantity" field from the two table. Used for updating the quantity.
    ' Should return the # of rows(records) that were affected.
    ' Will change the cursor position, usually to 1st but not always!!!
    ' Returns the balance (after performing add/subtract) to the table.
'commit()
    'must be called to save the changes/add record. Problem
    'is that Errors are raised when recset.Update() is called too
    'often.Don 't know why it happens but it is ok when only 1 update.
    'exception: add (arrays,arrays) because arrays of
    'field and values are assumed to be batched so only 1
    'updates it.
'bool findRecord(fieldName,operatorText As String, find)
'   fieldName can't contain "#" char. # is a delimiter in the RecSet.Find()
'   This moves the cursor of the matching unique key.
'   returns TRUE if found, false if not found.
'   Different than SQL "SELECT ... WHERE x=y"
'   For finding multiple-results, use "SELECT ... WHERE..." in openFile().
'
'getSize()
'getColumnSize()
'getSum("field")= SELECT COUNT(field)

' META DATA
'   isAutoIncrement(fieldName) Returns true if field is an autoincrement.
'       may not work on non-Access DB source.
'   isPrimaryKey(fieldName)
'   getPrimaryKey()
'
' TODO/OPTIMIZATION :
'   in merge(), for optimization,, change cursor type. Apparently after adding array of fields,
'       it loses the cursor. It must then do .Requery() in order to be proper.
'       In OpenDynamic cursor, the opening of the table is very fast. Perhaps
'       I can use that instead of OpenKeyset. Or, use SQL to add (ie INSERT INTO)
'       This might fix the problem.
'addRow(srcTable)
   'adds a new row. Then copy the current row from SrcTable.
'create(keys, types)
    ' creates new table.
'
Option Explicit
Private Const CONFIG_KEY_DEFAULT = "key"
Private Const CONFIG_VALUE_DEFAULT = "value"


Private WithEvents myTable As ADODB.Recordset
Attribute myTable.VB_VarHelpID = -1
Private myDb As TheAdoDb
Private mySqlStatement As String    'full sql statement used to open
                                    ' this table ("SELECT * ...")
Private myTableName As String   'name of table ("Employee", etc)
Private myAutoCommit As Boolean 'aka autoSave
Private myCursorType As ADODB.CursorTypeEnum
Private myCursorLocation As ADODB.CursorLocationEnum
Private myLockType As ADODB.LockTypeEnum
Event cursorMoved()
    ' called whenever goto___(), addRow(), etc. are performed and
    ' the cursor gets moved.
Public Function getString() As String
    getString = "TheAdoTable:" & myTableName & " RecSize=" & CStr(getSize())
End Function

Public Property Get sqlStatement() As String
    sqlStatement = mySqlStatement
End Property

Public Property Get DBConnection() As TheAdoDb
    Set DBConnection = myDb
End Property
'private sub openTable(

'openFile("cat")    'opens cat table using default "select * from cat"
'openFile("select * from dog")
' not really opening a file. It is opening and executing a query.
' Should rename to openIt()? VB can't use open().
Public Function openFile(ByVal openStatement As String, _
                        db As TheAdoDb, _
                        Optional cloneTable) As Boolean
    On Error GoTo TABLE_OPEN_FAILED
    closeFile
    Set myDb = db
    If InStr(1, openStatement, "SELECT ", vbTextCompare) <= 0 Then
        'only tablename was defined. Generate SELECT * statement
        myTableName = openStatement
        openStatement = TheSQLTool.selectFrom(myTableName)
    Else
        myTableName = TheSQLTool.getTableName(openStatement)
    End If
    mySqlStatement = openStatement
    'Debug.Assert (Not IsNull(db))
    If IsMissing(cloneTable) Then 'IsMissing(cloneTable) Or IsNull(cloneTable) Then
        Set myTable = New ADODB.Recordset
        myTable.cursorLocation = myCursorLocation
        myTable.Open openStatement, db.getConnection(), myCursorType, myLockType
    Else
        Set myTable = cloneTable.getRecordSet.clone(myLockType)
        'Sometimes, this should not actually open the table.
        'Usually done when cloning.
    End If
    openFile = True
    Exit Function
TABLE_OPEN_FAILED:
    openFile = False
    eprint "TheAdoTable::openFile() failed for " & openStatement
End Function

Public Function clone(Optional srcTable As TheAdoTable2) As TheAdoTable2
    'Dim cloneTable As TheAdoTable2
    On Error GoTo CLONE_FAILED
    
    If IsMissing(srcTable) Then 'Returning new table
        Dim cloneTable As New TheAdoTable2
        cloneTable.clone Me
        Set clone = cloneTable
    Else
        Me.openFile srcTable.sqlStatement, srcTable.DBConnection, _
            cloneTable:=srcTable
        
    End If
    Exit Function
CLONE_FAILED:
    eprint "TheAdoTable()::Clone() failed"
    'Set clone = Nothing
End Function


Private Sub Class_Initialize()
    setMode 'set default mode.
End Sub

Private Sub Class_Terminate()
    closeFile
End Sub

'Declare Events

    

Public Function closeFile()
    On Error GoTo TABLE_IS_NOT_OPEN
    myTable.Close  'don't close while data is being written.
    Set myTable = Nothing   ' remove from memory
    closeFile = True
    Exit Function
TABLE_IS_NOT_OPEN:
    ' do nothing since nothing was opened.
    closeFile = False
End Function

'Private Const DEFAULT_CURSOR_TYPE = adOpenStatic
'Private Const DEFAULT_LOCK_TYPE = adLockOptimistic
'Private Const DEFAULT_CURSOR_LOCATION = adUseClient

Public Sub setMode(Optional cursorType As CursorTypeEnum = adOpenStatic, _
                        Optional lockType As LockTypeEnum = adLockOptimistic, _
                        Optional cursorLocation As CursorLocationEnum = adUseClient, _
                        Optional autoCommit As Boolean = True)
    myCursorType = cursorType
    myLockType = lockType
    myCursorLocation = cursorLocation
    myAutoCommit = autoCommit
End Sub

' ADO-Specific Methods. Same as get Property
' For all others, one should use SQL stmt
Public Function getRecordSet() As ADODB.Recordset
    Set getRecordSet = myTable
End Function

Public Property Get ADODB_Recordset() As ADODB.Recordset
    Set ADODB_Recordset = myTable
End Property



Public Function findRecord(ByVal fieldName As String, _
                            ByVal operatorText As String, _
                            ByVal value As Variant) As Boolean
    ' Strip out "#" since # is delimiter in RecSet.Find()
    value = stripIllegalChar(value, "#")
    fieldName = stripIllegalChar(fieldName, "#")
    
    ' Find
    gotoFirst
    Dim findText As String
    If IsNumeric(value) Then
        findText = fieldName & operatorText & " " & value & " "
    Else
        findText = fieldName & operatorText & "#" & value & "#"
    End If
    myTable.find findText
    If isEof() Then
        findRecord = False 'Not found
     Else
        findRecord = True
    End If
End Function

Public Function commit() As Boolean
'On Error GoTo COMMIT_ERROR
    myTable.Update
    'If myTable.Status = adRecOK Then
        commit = True
        Exit Function
    'End If
    ' if status <> adRecOK, just fall through to error.
COMMIT_ERROR:
    commit = False
    eprint "AdoTable::commit (aka update()) " & mySqlStatement & " failed"
End Function

'batch not supported on Access
Public Function commitBatch() As Boolean
On Error GoTo COMMIT_ERROR
    myTable.UpdateBatch
    commitBatch = True
COMMIT_ERROR:
    Exit Function
    commitBatch = False
    eprint "AdoTable::commitBatch(" & mySqlStatement & ") failed"
End Function

Public Sub refresh()
    myTable.Requery 'Warning: Requery() might not be accurate after adding NEW data.
End Sub

Public Function gotoFirst() As Boolean
    If getSize() = 0 Then
        gotoFirst = False
    Else
        myTable.MoveFirst
        gotoFirst = IIf(myTable.BOF Or myTable.EOF, False, True)
    End If
End Function


Public Function gotoNext()
    myTable.MoveNext
    
End Function

Public Function gotoLast()
    myTable.MoveLast
    
End Function


'get # of records. However, it will only work in certain modes
' Side effect: will reset the cursor position!
Public Function getSize() As Long
    If myTable.cursorType And adOpenStatic Or myTable.cursorType And adOpenKeyset Then
        getSize = myTable.RecordCount
    Else    '-- very slow in this case. UNTESTED
        eprint "Warning: TheAdoTable::getSize() using UNTESTED implementation"
        Dim size As Long
        gotoFirst
        While Not isEof()
            gotoNext
            size = size + 1 'no size++ in VB classic
        Wend
        getSize = size
    End If
End Function

Public Function getColumnSize() As Long
    getColumnSize = myTable.Fields.Count()
End Function

Public Function getSum(ByVal fieldName As String) As Long
                        'Optional cache As Boolean = True) As Long
    'Caches the result and requery when needed. By default.
    'WARNING: Requery() was inaccurate when tested with table that
    '       recently added new data without Update(). It returned
    '       table with old result. So the Cache part was removed.
    'Static LastField As String
    Debug.Assert (fieldName <> "")
    Static SumTable As TheAdoTable2
    'If fieldName <> LastField Then  'NEW QUERY
        'LastField = fieldName
        Set SumTable = New TheAdoTable2
        SumTable.openFile "SELECT SUM([" & fieldName & "])" & _
                            " AS mySum FROM " & myTableName, myDb
    'Else                            'REQUERY
    '    SumTable.refresh
    'End If
    getSum = SumTable.getValue("mySum", 0)
    'If Not cache Then
        SumTable.closeFile
    '    LastField = ""
    'End If
End Function

Public Function getValue(Optional fieldName, Optional defaultValue) As Variant
    Dim returnList As Collection
    If IsMissing(fieldName) Then ' GET the whole row
        Set returnList = New Collection
        Dim i
        For Each i In myTable.Fields
            returnList.add i
        Next
        Set getValue = returnList
    Else
        getValue = myTable.Fields(fieldName).value
        If IsNull(getValue) And Not IsMissing(defaultValue) Then
            getValue = defaultValue
        End If
    End If
End Function

Public Function getRow() As Collection
    'Dim returnValue As Collection
    'Set returnValue = Me.getValue()
    'get getRow = Me.getValue()
    Dim returnList As New Collection
        Dim i
        For Each i In myTable.Fields
            returnList.add i
        Next
    Set getRow = returnList
End Function

Public Function isEof() As Boolean
    isEof = myTable.EOF
End Function

' sets current row's data given the columnNumber name
Public Sub setValue(ByVal fieldName As String, ByVal value)
    On Error GoTo SET_VALUE_ERROR
    
'    Dim testval1, testname1, testname2
'    testname1 = myTable.Fields(2).name
'    testname2 = myTable.Fields(2)
'    testval1 = myTable.Fields(fieldname)
    If IsObject(value) Then
        eprint "AdoTable::setValue() of this type, Not implemented yet"
        'TODO: .value=value.getString()
    ElseIf IsArray(value) Then
        eprint "AdoTable::setValue() of this type, Not implemented yet"
        'TODO: foreach i in value ... value=value+i???
    End If
    myTable.Fields(fieldName).value = value
        ' not sure; myRecordset(columnNumber).value = value ???
    'commit    'automatically call .update() if needed
    'myTable.Update fieldname, value
    
    Exit Sub
SET_VALUE_ERROR:
    eprint "TheAdoTable::setValue(" & fieldName & "," & CStr(value) & _
        " Failed."
End Sub


Public Function addRow(Optional keys, Optional values) As Boolean
    On Error GoTo ADD_FAILED
    'If IsArray(keys) Then
    '    myTable.AddNew keys(), values()
    If IsMissing(keys) Then
        myTable.AddNew
        'commit : Kept getting errors.
    ElseIf isString(keys) Then
        Debug.Assert Not IsMissing(values)
        myTable.AddNew Array(keys), Array(values)
        commit
    Else
        Debug.Assert (IsArray(keys))    'must be arrays
        Debug.Assert Not IsMissing(values)
        myTable.AddNew keys, values '
        commit
        'commit is not needed bc addNew x,y will update immediately
    End If
    addRow = True
    Exit Function
ADD_FAILED:
    addRow = False
    eprint "TheAdoTable::add() failed. " & mySqlStatement
    If Not IsMissing(keys) Then
        eprint the.getString(keys) & the.getString(values)
    End If
End Function


' goes to the 1st record and returns the result in Collection.
' In Perl's DBI, it is same as doing fetch() right after executing.
' In JDBC, it is same as executing, and then doing next() and then fetch().
Public Function getFirst(Optional fieldName As String)
    gotoFirst
    If IsMissing(fieldName) Then 'return the whole row
        Set getFirst = getValue()
    Else        'return only the column
        getFirst = getValue(fieldName)
    End If
End Function

Public Function getNext(Optional fieldName As String)
    gotoNext
    If IsMissing(fieldName) Then 'return the whole row
        Set getFirst = getValue(fieldName)
    Else        'return only the column
        getFirst = getValue(fieldName)
    End If
End Function

Public Sub clear()
    ' Warning: Erases the entire Table.
    'Method 1
    'moveFirst
    'myRecordset.Filter = "id=*"
    'myRecordset.Delete adAffectGroup
    'Method 2
    If getSize() = 0 Then
        Exit Sub    ' nothing do delete
    End If
    the.dprint "Warning--TheAdoTable()::performing clear()/eraseTable()"
    gotoFirst
    While Not isEof()
        remove
        'gotoNext 'remove() automatically does gotoNext()
    Wend
    commitBatch
End Sub

Public Function remove()
    remove = False
    'removes current record.
    If getSize() = 0 Then
        Exit Function
    End If
    myTable.Delete adAffectCurrent
    gotoNext    'Cursor still remains on the deleted rec until we move it.
    remove = True
End Function


Public Function merge(ByRef foreignTable As TheAdoTable2, _
                        ByVal primaryKey As String, ByVal foreignKey As String, _
                        ByVal myField As String, ByVal foreignField As String, _
                        ByVal operationMode As MERGE_OPERATION) As Long
    Dim recordSize As Long
            '# of rec that were affected & had the "operation" performed.
    'Dim balance As Long
            'the balance in the field in myField as the result of this operation
    Dim myValue
    Dim key     'current key
    Dim foreignValue
    Dim found As Boolean
    
    gotoFirst
    foreignTable.gotoFirst

    While Not foreignTable.isEof()
        key = foreignTable.getValue(foreignKey)
        foreignValue = foreignTable.getValue(foreignField)
        found = findRecord(primaryKey, "=", key)
        If found Then   ' ITEM ALREADY EXISTS. GOTO THAT RECORD
            myValue = getValue(myField) 'save old value.Important or else
                                        ' we lose the previous value.
            copyCurrentRow sourceTable:=foreignTable, ignoreEmptyText:=True
                                        ' overwrite previous text field
                                        ' only if new field is not empty.
        Else    ' NEW ITEM is FOUND. ADD A NEW ROW TO MY TABLE
            'add primaryKey, Arraykey
            addRow                 'add an empty record
            copyCurrentRow sourceTable:=foreignTable
            
            'refresh 'for optimization, use SQL. Apparently after adding array of fields,
                    'it loses the cursor. It must then do .Requery() in order to be proper.
                    ' See header for some idea.
            'findRecord primaryKey, "=", key  'delete: cursor is already pointing it.
            myValue = getEmptyTypeBasedOn(foreignValue)
        End If
        Select Case operationMode
            Case MERGE_ADD
               'Dim temp
                'temp = getValue(primaryKey)
                setValue myField, myValue + foreignValue
                recordSize = recordSize + 1 'recordSize++
            Case MERGE_SUBTRACT
                setValue myField, myValue - foreignValue
                recordSize = recordSize + 1 'recordSize++
                If Not found Then
                    eprint "TheAdoTable::merge() Warning: trying to subtract from non-existent record."
                End If
            Case Else
                eprint "TheAdoTable::merge(), this operation not implemented."
        End Select
        commit
        foreignTable.gotoNext
    Wend
    merge = recordSize  'returnValue
    'merge = balance
End Function



Public Function copy(srcTable As TheAdoTable2)
    If Me.getSize() > 0 Then
        warn "TheAdoTable::copy() deleted all the records in " & mySqlStatement
    End If
    Me.clear
    append (srcTable)
End Function

Public Function append(srcTable As TheAdoTable2)
    On Error GoTo APPEND_ERROR
    srcTable.gotoFirst
    While Not srcTable.isEof()
        addRow
        copyCurrentRow (srcTable)
    Wend
    append = True
    Exit Function
APPEND_ERROR:
    append = False
    eprint "TheAdoTable2()::append() Failed"
End Function

Public Function createDuplicate(ByVal tableName As String, _
                            srcTableName As Variant, _
                            Optional db As TheAdoDb, _
                            Optional copyData = False) As Boolean
    On Error GoTo CREATE_DUPLICATE_FAILED
    Me.closeFile    ' close any opened file.
    'Dim newTable As New TheAdoTable2
    Dim sqlText As String
    Dim srcTable As TheAdoTable2
    
    If IsMissing(db) And Not isString(srcTableName) Then
        Set db = srcTableName.getRecordSet().ActiveConnection
    End If
    
    Debug.Assert Not IsEmpty(db)
    
    Dim fieldDict As TheDictionary
    Set fieldDict = db.getMetaData(srcTableName)
    sqlText = TheSQLTool.createTable(tableName, fieldDict)
'    If IsEmpty(db) Then
'        Debug.Assert 1
'    End If
    ' if srcTableName is string then we open a temporary table
    ' and copy all the information we need. Then we close it at the end.
    ' however, if srcTableName is an actual reference to another table,
    ' we just use that.
'    If isString(srcTableName) Then
'        Set srcTable = New TheAdoTable2
'        srcTable.openFile "select * from " & srcTableName, db
'    Else
'        Set srcTable = srcTableName
'    End If
    
'    sqlText = "CREATE TABLE " & tableName & " (temp int, tname varchar(30))"
    'sqlText = "CREATE TABLE " & tableName
    db.execute sqlText
'    Me.closeFile    'close "CREATE" SQL action.
    'copy all the field and its attributes
'    Dim i As Field   'field
'    Me.openFile "select * from " & tableName, db
'    For Each i In srcTable.getRecordSet().Fields
'        myTable.Fields.append i.name, i.Type, i.DefinedSize, i.Attributes
        
'    Next
'    myTable.Fields.Update   'Don't know why it gets error. save it.
'    If isString(srcTableName) Then
'        srcTable.closeFile
'    End If
    Me.openFile "SELECT * FROM " & tableName, db
    
    createDuplicate = True
    Exit Function
CREATE_DUPLICATE_FAILED:
    createDuplicate = False
    eprint "ADOTable::CreateDuplicate() Failed"
End Function

' Internal use only. for debugging use only. used internally to test the functionality
'   of Attribute. However, it didn't seem to work with Access or atleast
'   the KeyColumn, and other vital attributes were not being detected.
'   Instead, Property seemed to be working.
Private Function getFieldAttributeEnumString(ByVal val As FieldAttributeEnum) As String
    Dim str As String
    'Select Case val
     If val And adFldIsNullable Then str = str & "Nullable "
     If val And adFldKeyColumn Then str = str & "keyColumn "
     If val And adFldRowID Then str = str & "RowID "
     If val And adFldUpdatable Then str = str & "Updatable "
     If val And adFldFixed Then str = str & "Fixed "
     If val And adFldLong Then str = str & "long "
'    If val And  adFldFixed Then str = str & "Fixed"

    getFieldAttributeEnumString = str
End Function

Public Function copyCurrentRow(ByRef sourceTable As TheAdoTable2, _
                                Optional ignoreEmptyText = True)

    On Error GoTo COPY_CURRENT_ROW_FAILED
    Dim i As ADODB.Field
    'Dim name As String      'name of the field
    'Dim value               'data
    Dim autoincrementField As String
    autoincrementField = getAutoincrementField()
    
    For Each i In sourceTable.getRecordSet().Fields
        'value = i.value
        'name = i.name
        ' write only on field that is NoT autoincrement and
        ' make sure that i only write text that are not empty
        ' unless i override it in ignoreEmptyText.
        
        ' TODO: check if field exists
        
        If i.name <> autoincrementField And _
                (isEmptyOrNullString(i.value) = False Or _
                     ignoreEmptyText = False) Then
            setValue i.name, i.value
        End If
    Next
    commit 'must save changes here. setValue won't do it.
    
    copyCurrentRow = True
    Exit Function
COPY_CURRENT_ROW_FAILED:
    copyCurrentRow = False
    eprint "TheAdoTable::copyCurrentRow() Failed."
End Function
                                
                                
Public Function isAutoIncrement(ByVal fieldName) As Boolean
    isAutoIncrement = getFieldProperty(fieldName, "ISAUTOINCREMENT")
End Function

Public Function isPrimaryKey(ByVal fieldName) As Boolean
    isPrimaryKey = getFieldProperty(fieldName, "KEYCOLUMN")
End Function

Public Function getPrimaryKey() As String
    getPrimaryKey = getFieldByProperty("KEYCOLUMN")
End Function

Public Function getAutoincrementField() As String
    getAutoincrementField = getFieldByProperty("ISAUTOINCREMENT")
'    Dim i As Field
'    getAutoincrementField = ""  'if no primaryKey exists, returns ""
'    For Each i In myTable.Fields
'        If isAutoIncrement(i.name) Then
'            getAutoincrementField = i.name
'        End If
'    Next
End Function

'see fieldExists()
Public Function keyExists(ByVal key As String) As Boolean
    On Error GoTo KEY_DOESNT_EXIST
    Dim value
    Set value = myTable.Fields(key)
    keyExists = True
    Exit Function
KEY_DOESNT_EXIST:
    keyExists = False
End Function

'alias to keyExists()
Public Function fieldExists(ByVal key As String) As Boolean
    fieldExists = keyExists(key)
End Function

' no need to make it public yet since it isn't used at high-level.
' instead, use high level function like getPrimaryKey(),...
Private Function getFieldProperty(ByVal fieldName As String, _
                                ByVal propertyName As String)
    On Error GoTo GET_FIELD_NO_PROPERTY
    getFieldProperty = False    'default value.
    Dim tempField As Field
    Set tempField = myTable.Fields(fieldName)
'    Dim i As Property
'    For Each i In tempField.Properties
'        the.dprint "Properties = ", i.name
'    Next
   'if tempField.Properties(0).
    Dim result
    result = tempField.Properties(propertyName)
    'If tempField.Properties(propertyName) <> "" Then
        getFieldProperty = True
    'Else
    '    getFieldProperty = False    'probably unreachable condition.
    'End If
    Exit Function
GET_FIELD_NO_PROPERTY:
    'if the property doesn't exist, it means that it doesn't have that
    ' property. It is not an error. It just doesn't have this property.
    getFieldProperty = False
End Function

Private Function getFieldByProperty(ByVal propertyName) As String
    Dim i As Field
    getFieldByProperty = ""  'if no primaryKey exists, returns ""
    
    For Each i In myTable.Fields
        If getFieldProperty(i.name, propertyName) Then
            getFieldByProperty = i.name
            Exit Function
        End If
    Next
End Function

Public Sub saveXML(ByVal path As String)
    myTable.Save path, adPersistXML
End Sub


'=============================================================================
' HELPER METHODS -- TODO: Move these to the TheTableBase Class in VB.NET.
'=============================================================================
Public Function findAndGet(ByVal fieldName, ByVal match, ByVal fieldToReturn)
    If findRecord(fieldName, "=", match) Then
        findAndGet = getValue(fieldToReturn)
    Else
        findAndGet = Null
    End If
End Function


Public Sub findAndSet(ByVal fieldName, ByVal match, ByVal fieldToSet, ByVal value)
    If findRecord(fieldName, "=", match) Then
        setValue fieldToSet, value
    End If
End Sub




'=============================================================================
' CONFIG Table based on ADO TABLE
'TODO: create TheAdoConfigTable and rename them get(), set()
'=============================================================================


Public Function getConfig(ByVal key)
    getConfig = findAndGet(CONFIG_KEY_DEFAULT, key, CONFIG_VALUE_DEFAULT)
End Function

Public Sub setConfig(ByVal key, ByVal value)
    findAndSet CONFIG_KEY_DEFAULT, key, CONFIG_VALUE_DEFAULT, value
End Sub
'WISH:
' public function rollback()
'




Private Sub myTable_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    RaiseEvent cursorMoved
End Sub

'=============================================================================
'
'=============================================================================
#If OLD_CODE Then
''  ("name=John") => current Row's field("Name") = "John"
'' Low-level func, used only by addRow(). Sets the field of currently selected row
'Private Sub addRowItem(ByVal item As String)
'    Dim keyValue As Collection
'    Set keyValue = parseKeyValueString(item)
'    Dim key As String
'    Dim value
'    key = keyValue(1)
'    value = keyValue(2)
'    'myRecordset(key).value = value
'    myRecordset.Fields.item(key).value = value
'End Sub
#End If



