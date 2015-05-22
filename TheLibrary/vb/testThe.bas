Attribute VB_Name = "testThe"
Option Explicit

Public Sub test1()
    Dim i As Long
    Dim str As String
    str = "hello"

 
    Dim result As Collection
    the.dprint "hello", getNumberOrString("hello")
    the.dprint "123", getNumberOrString("123")
    the.dprint "34.5", getNumberOrString("34.5")
    str = "abc=hello"
    
    ' getList
    Dim arr(0 To 3)
    arr(0) = "aaa=77"
    arr(1) = "bbb=nissan"
    arr(2) = "ccc=truck"
    arr(3) = "ddd=999.99"
    Dim dict As TheDictionary
    the.dprint "1,2,3", getListFromParamArray(1, 2, 3)
    the.dprint "a,b,c", getListFromParamArray("a", "b", "c")
    
    the.dprint "array(aaa,bbb,ccc,ddd)", getListFromParamArray(arr)
    
    the.dprint "getDict(a=1,b=2,c=3)", getDictFromParamArray("a=1", "b=2", "c=3")
    Set dict = getDictFromParamArray(arr)
    the.dprint "getDictFromParamArray(array(aaa,bbb,ccc,ddd))", dict
    
    Dim var1
    Set var1 = New Collection
    
    the.dprint "isCollection(Coll)", isCollection(var1)
    the.dprint "isString(Coll)", isString(var1)
    var1 = "yo go"
    the.dprint "isCollection(yogo)", isCollection(var1)
    
    the.dprint "isString(yogo)", isString(var1)
    
    'TheDialog.getFont
    
    var1 = FormatCurrency(45.3)
    the.dprint "Currency = ", var1
    
    Dim fs As Object
    Set fs = CreateObject("Scripting.FileSystemObject")
    the.dprint "GetAbsolutePathName()", fs.GetAbsolutePathName("..")
    the.dprint "GetBaseName()", fs.GetBaseName("c:\folder1\folder2\hello.ext")
    the.dprint "GetDriveName()", fs.GetDriveName("c:\folder1\folder2\hello.ext")
    
    the.dprint "GetExtensionName()", fs.GetExtensionName("c:\folder1\folder2\hello.ext")
    the.dprint "GetFileName()", fs.GetFileName("c:\folder1\folder2\hello.ext")
    the.dprint "GetParentFolderName()", fs.GetParentFolderName("c:\folder1\folder2\folder3\")
    Dim meg As Long
    meg = 1024& * 1024&
    the.dprint "AvailableSpace()", (fs.GetDrive("d:").AvailableSpace) / meg
    the.dprint "TotalSize()", fs.GetDrive("d:").TotalSize / meg
    
    'var1 = "abc123"
    'the.dprint "cint(abc123)", (CInt("123"))
    Dim r As New RegExp
    r.Pattern = "[^\d]"
    r.Global = True
    the.dprint "regexp", r.Replace("abc123def", "X")
    
    the.dprint "isNumeric(123 456 )", IsNumeric("123 456")
    
    the.dprint "regExp.getNumber(123abcdef)", TheRegExpTool.getNumber("123abcdef")
    the.dprint "regExp.getNumber(rg1234abc)", TheRegExpTool.getNumber("rg1234abc")
    the.dprint "regExp.getNumber(1234)", TheRegExpTool.getNumber("1234")
    the.dprint "getPrimaryKey()", inventory.productTable.getPrimaryKey()
    the.dprint "getAutoInc()", inventory.productTable.getAutoincrementField()
    the.dprint "isPrimaryKey(id)", inventory.productTable.isPrimaryKey("id")
    
    ' Drop table
    'dropTable "tempTable"
    'Dim fieldDict As New TheDictionary
    'fieldDict.add "[id]", "INTEGER PRIMARY KEY"
    'fieldDict.add "[name]", "VARCHAR(32)"
    'samsDb.createTable "tempTable", fieldDict
    
    'testing static var
    Dim c1 As New TestClass
    Dim c2 As New TestClass
    c1.testStatic (5)
    c2.testStatic (25)
    the.dprint "static1(should be 5)", c1.testStatic()
    the.dprint "static2(should be 25)", c2.testStatic()
    the.dprint "static1(should be 5)", c1.testStatic()
    
    '[]
    
    'sqltool
    Dim fieldArray
    fieldArray = Array("123", "456", "789")
    Dim fieldList As New TheList
    fieldList.add "f1"
    fieldList.add "f2"
    fieldList.add "f3"
    the.dprint "sqlfield(paramarray)", TheSQLTool.sqlField("abc", "def")
    the.dprint "arrays", TheSQLTool.sqlField(fieldArray)
    the.dprint "lists", TheSQLTool.sqlField(fieldList)
    
    'test array.
    Dim myArr(0 To 1) As Long
    myArr(0) = 10
    myArr(1) = 20
    the.dprint "array before", myArr
    
    testArray myArr
    the.dprint "array after", myArr
    
    the.dprint "list before", fieldList
    changeList fieldList
    the.dprint "list after", fieldList
    
    the.dprint "testing getTextAfterWord()", the.getTextAfterWord("FROM", "SELECT * FROM ABC and 123")
    the.dprint "testing getTextAfterWord(NOT FOUND)", the.getTextAfterWord("YIKE", "SELECT * FROM ABC and 123")

    the.dprint "testing getTableName()", TheSQLTool.getTableName("SELECT * FROM myTable WHERE 123")

End Sub

Sub testArray(arr)
    arr(0) = 35
End Sub

Sub changeList(ByVal l As TheList)  'doesn't matter.
    l.add "yoyoFinale"
End Sub
