Attribute VB_Name = "TheSQLTool"
Option Explicit

'selectFrom("table", fieldStr="*")
    'Simple SELECT * FROM singleTable
    '? selectFrom ("Employee")
    '=> "SELECT * FROM Employee"


'sqlField("field1","field2",,... OR [fieldsArray] or [fieldsList])
    'Encloses each field in [ ] due to JET DB Engine.
    '? ("SELECT",sqlField("id","name"),"FROM",MyTable)
    '=> "SELECT [id],[name] FROM Employee"


'WISH:
'select(fields, tables, whereClause, groupClause,...) 'Need VB.NET
    'selects multiple-table

Public Function sqlField(ParamArray vararg()) As String
    Dim item, returnString As String
    Dim fieldList
    'in Access, field name should be enclosed in [] in certain cases.
    If isString(vararg(0)) Then     'STR1,STR2,...
        fieldList = vararg
    ElseIf IsArray(vararg(0)) Then  'ARRAY
        fieldList = vararg(0)
    Else                            'TheList, Collection
        Set fieldList = vararg(0)
    End If
        For Each item In fieldList
            If item = "*" Then
                sqlField = "*"
                Exit Function
            End If
            returnString = returnString + "[" & item & "],"
        Next
    'End If
    chop returnString   'remove last ","
    sqlField = returnString
End Function


'makeCreateTable("movie",movieFieldsAndTypes) =>
'   "CREATE TABLE movie (id INTEGER, name VARCHAR(30),..."

Public Function createTable(ByVal tableName As String, _
                                fieldDict As TheDictionary) As String
    Dim sqlText As String
    sqlText = "CREATE TABLE " & tableName & " ("
    Dim i
    For Each i In fieldDict.getKeys()
        sqlText = sqlText & i & " " & fieldDict.getValue(i) & ", "
    Next
    chop sqlText    'remove last comma
    chop sqlText
    sqlText = sqlText & ");"
    'sqlText = "CREATE TABLE (ID INTEGER, NAME INTEGER)"
    createTable = sqlText
End Function

Public Function selectFrom(ByVal tableName As String, _
                            Optional fieldName As String = "*") _
                            As String
    selectFrom = "SELECT " & fieldName & " FROM " & tableName & ";"
End Function
' works with very simple sql statement. Doesn't handle complex sql statement
    'tries to extract table name from the given SQL statement. Doesn't work with mutliple table
    'getTable("SELECT .... FROM ab ....") => Returns "ab"
    'getTable("SELECT * FROM Tbl1,Tbl2") => Returns "Tbl1,tbl2", may not be appropriate value
Public Function getTableName(ByVal sqlText As String) As String
    Dim startPosition As Long, endPosition As Long
    Dim returnValue As String
    If InStr(1, sqlText, "SELECT", vbTextCompare) Then
        sqlText = getTextAfterWord("FROM ", sqlText)
        returnValue = the.getFirstWord(sqlText)
    ElseIf InStr(1, sqlText, "CREATE", vbTextCompare) Then
        sqlText = getTextAfterWord("TABLE ", sqlText)
        returnValue = the.getFirstWord(sqlText)
        
    ElseIf InStr(1, sqlText, "DROP", vbTextCompare) Then
        sqlText = getTextAfterWord("TABLE ", sqlText)
        returnValue = the.getFirstWord(sqlText)
    Else
        eprint "TheSQLTool::gettable()=Unhandled SQL Statement:" & sqlText
        
    End If
    getTableName = returnValue
End Function

'End Function

'Public Function makeDropTable(ByVal tablename As String) As String

'End Function
