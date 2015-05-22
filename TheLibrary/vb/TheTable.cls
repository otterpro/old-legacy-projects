VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TheTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'=============================================================================
'   TheTable
'   aka TheGrid
'   Do NOT USE: Broken array mechanism. Cannot resize correctly.
'   Instead, use TheTreeTable(?)
'   TODO: Incorporate this into TheTreeTable class
'   simpler version of table. For more powerful table, use
'   TheTreeTable and its derivatives.
'
'   Dependency: None
'=============================================================================
''
'   addRow(col0,col1,col2,...)
'   addRow(array)
'   setCell(col,row,data)
'   getCell(col,row)
'   getRowSize()
'   getColumnSize()
'   clear()
'       erase table
'   copy()
'   copyTo()
'   getRow()
'   dprint()
'   Example:
'   dim table as New TheTable
'   table.addRow (15,23,"cherry",550."banana")
'   table.addRow (myArray)
'   table.setCell(0,1,"blank")
'   i = table.getCell(0,2)
'
'=============================================================================
Option Explicit
'Implements TheTableInterface

Private myColumnSize As Long, myRowSize As Long
Private myMaxColumnSize As Long, myMaxRowSize As Long
    'actual size in the memory
Const COLUMN_BLOCK_SIZE = 4
Const ROW_BLOCK_SIZE = 4
Const DEFAULT_COLUMN_SIZE = 2
    'if setColumnSize() is not called after class is
    ' initialized, it creates a table with 4 columns.
    
'local variable(s) to hold property value(s)

Private myTable() As Variant 'local copy

Public Property Let table(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.table = 5
    myTable = vData
End Property


Public Property Set table(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.table = Form1
    'Set myTable = vData
End Property


Public Property Get table() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.table
    If IsObject(myTable) Then
        'Set table = myTable
    Else
        table = myTable
    End If
End Property


Private Sub Class_Initialize()
    clear
End Sub

'Public Sub copyArray(data() As Variant)
'    eprint "thetable::copyArray() not implemented yet. "
'End Sub

Public Sub copy(Data As Variant)
    clear
    add Data
End Sub

Public Function getColumnSize() As Long
    getColumnSize = myColumnSize
End Function

Public Function getRowSize() As Long
    getRowSize = myRowSize
End Function

Public Sub clear()
    myColumnSize = DEFAULT_COLUMN_SIZE
    myRowSize = 0
    myMaxColumnSize = COLUMN_BLOCK_SIZE
    myMaxRowSize = ROW_BLOCK_SIZE
    ReDim myTable(0 To myMaxColumnSize, 0 To myMaxRowSize)
End Sub

' create a single row and add the row of data, given the list
' can pass either the array or vararg, but not combination of both.
' If the item exceeds the myColumnSize, it increases myColumnSize.
Public Sub addRow(ParamArray Data())
    Dim colSize As Long
    
    myRowSize = myRowSize + 1 'myRowSize++
    Dim col     'temp column index
    col = 0
    Dim item      ' individual cell's value
    
    If IsArray(Data(0)) Then
        colSize = the.getSize(Data(0))
        myColumnSize = the.Max(colSize, myColumnSize)
        resizeIfNeeded
        'addColumnIfNeeded colSize
        For Each item In Data(0)
            setCell col, myRowSize - 1, item
            col = col + 1  'col++
        Next
 '   ElseIf isCollection(Data(0)) Then

    Else
        colSize = UBound(Data) - LBound(Data) + 1
        myColumnSize = the.Max(colSize, myColumnSize)
        resizeIfNeeded
        For Each item In Data
            setCell col, myRowSize - 1, item
            col = col + 1  'col++
        Next
    End If
    
End Sub

'same as addRow() except that it doesn't use paramarray.
Public Sub addRowOfList(row)
    Dim colSize As Long
    
    myRowSize = myRowSize + 1 'myRowSize++
    Dim col     'temp column index
    col = 0
    Dim item      ' individual cell's value
    
    If IsArray(row) Or isCollection(row) Then
        colSize = the.getSize(row)
        myColumnSize = the.Max(colSize, myColumnSize)
        resizeIfNeeded
        'addColumnIfNeeded colSize
        For Each item In row
            setCell col, myRowSize - 1, item
            col = col + 1  'col++
        Next
 '   ElseIf isCollection(Data(0)) Then
    End If
End Sub

' append multiple rows of data.
' Typically, this shouldn't be needed. Instead, use TheTableBase::CopyTo().
Public Sub add(ByRef Data)
    Dim n
    If TypeName(Data) = "MSFlexGrid" Then
        Dim tempArray()
        ReDim tempArray(0 To Data.Cols() - 1)
        For n = 0 To Data.Rows() - 1 '0-based
            For m = 0 To Data.Cols() - 1 '0-based
                tempArray(m) = Data.TextMatrix(n, m)
            Next
            addRow tempArray
        Next
    ElseIf TypeOf Data Is TheAdoTable2 Then 'ADO-Recordset
        Dim tablePtr As TheAdoTable2
        Set tablePtr = Data
        'Dim tempArray()
        'ReDim tempArray(0 To grid.Columns.Count - 1)
        'Dim tempDataSource As ADODB.Recordset
        'Set tempDataSource = grid.DataSource
        tablePtr.gotoFirst
        While Not tablePtr.isEof()
            addRowOfList tablePtr.getRow()
            tablePtr.gotoNext
        Wend
    ElseIf TypeName(Data) = "TheTable" Then
        For n = 0 To Data.getRowSize() - 1 '0-based
            'addRow Data.getRowStringArray(n)    'convert to array and add
            addRow TheTableBase.getRow(Data, n)   'convert to array and add
        Next
    Else
        MsgBox "TheTable::add() for this data type is not implemented."
    End If
End Sub


Public Sub setCell(ByVal col As Long, ByVal row As Long, ByVal cellData)
    myTable(col, row) = cellData
End Sub

'set the size of the table
' Currently, it only increases the size as needed. It doesn't reduce the size if it shrinks.
Public Sub resize(Optional rowIncrement = 0, Optional columnIncrement = 0)
    If myRowSize + rowIncrement >= myMaxRowSize Then
        myMaxRowSize = the.Max(myMaxRowSize + ROW_BLOCK_SIZE, myMaxRowSize + rowIncrement)
    End If
    If myColumnSize + columnIncrement >= myMaxColumnSize Then
        myMaxColumnSize = the.Max(myMaxColumnSize + COLUMN_BLOCK_SIZE, myMaxColumnSize + columnIncrement)
    End If
    ReDim Preserve myTable(myMaxColumnSize, myMaxRowSize)
End Sub

Public Sub resizeIfNeeded()
    If myRowSize >= myMaxRowSize Or myColumnSize >= myMaxColumnSize Then
        resize 'columnIncrement = myColumn
    End If
End Sub

' set the size of column and the header name, if needed.
' All data will be erased. Call this after instantiating theTable
' if you are creating a new table. If copying a table, don't need
' to call this.
' Instead, use addrow() to add column header in the beginning
'Public Sub setHeader(ParamArray header())

'End Sub


Public Function getCell(ByVal col As Long, ByVal row As Long) As Variant
    rangeIsValid col, row   'range check
    getCell = myTable(col, row)
End Function

' Given the column and row value, returns true if the coordinate is in the valid range.
' Returns false if it is outside the range.
Public Function rangeIsValid(ByVal col, ByVal row) As Boolean
    rangeIsValid = IIf(col >= myColumnSize Or row >= myRowSize, False, True)
   ' If col > myColumnSize Or row > myRowSize Then
        'eprint "TheTable::getCell() out-of-range" ' no need to warn user since this only checks the range.
        '   Warning should be generated by caller of this function.
   '     rangeIsValid = False
   ' Else
   '     rangeIsValid = True
   ' End If
End Function













'=============================================================================
'   OLD CODE
'=============================================================================
#If OLD_CODE Then
'' Use resize(col,row) instead.
''set the size of column. By Default, it is 4 (0..3).
'Public Sub setColumnSize(ByVal col As Long)
'    myColumnSize = col
'    If myColumnSize >= maxColSize Then
'        myMaxColumnSize = myMaxColumnSize + COLUMN_BLOCK_SIZE
'        ReDim Preserve myTable(myMaxColumnSize, myMaxRowSize)
'    End If
'End Sub
'
'
'' Do not use this. Instead, use the.getString(this.getRow())
'Public Function getRowString(ByVal row, Optional separatorChar = ",") As String
'    rangeIsValid 0, row
'    getRowString = ""
'    Dim i
'    For i = 0 To myColumnSize - 1
'        getRowString = getRowString & getCell(i, row) & separatorChar
'    Next
'End Function
'
'' get an array of string
'Public Function getRowStringArray(ByVal row) As String()
'    rangeIsValid 0, row
'
'    Dim i
'    Dim returnArray
'    ReDim returnArray(0 To myColumnSize - 1) As String
'    For i = 0 To myColumnSize - 1
'        returnArray(i) = getCell(i, row)
'    Next
'    getRowStringArray = returnArray
'End Function
'
'' Use resizeIfNeeded() instead.
'' set the size of column if needed
'Public Sub addColumnIfNeeded(ByVal colSize As Long)
'    If colSize > myColumnSize Then
'        setColumnSize (myColumnSize)
'    End If
'End Sub
'
'' create a single row and add the row of data, given the list
'' can pass either the array or vararg, but not combination of both.
'' If the item exceeds the myColumnSize, it increases myColumnSize.
'' The reason it was changed was because it was too inefficient.  Calling
'' resizeIfNeeded() for every row of data that is copied was just too slow.
'Public Sub addRowOLD(ParamArray Data())
'    Dim colSize As Long
'
'    myRowSize = myRowSize + 1 'myRowSize++
'    Dim col     'temp column index
'    col = 0
'    Dim item      ' individual cell's value
'
'    If IsArray(Data(0)) Then
'        colSize = the.getSize(Data(0))
'        myColumnSize = the.Max(colSize, myColumnSize)
'        resizeIfNeeded
'        'addColumnIfNeeded colSize
'        For Each item In Data(0)
'            setCell col, myRowSize - 1, item
'            col = col + 1  'col++
'        Next
'    Else
'        colSize = UBound(Data) - LBound(Data) + 1
'        myColumnSize = the.Max(colSize, myColumnSize)
'        resizeIfNeeded
'        For Each item In Data
'            setCell col, myRowSize - 1, item
'            col = col + 1  'col++
'        Next
'    End If
'
'End Sub
'
'' It was changed to modify the AddRow(). See above.
'' append multiple rows of data.
'' Typically, this shouldn't be needed. Instead, use TheTableBase::CopyTo().
'Public Sub add(Data)
'    If TypeName(Data) = "MSFlexGrid" Then
'        Dim tempArray()
'        ReDim tempArray(0 To Data.Cols() - 1)
'        For n = 0 To Data.Rows() - 1 '0-based
'            For m = 0 To Data.Cols() - 1 '0-based
'                tempArray(m) = Data.TextMatrix(n, m)
'            Next
'            addRow tempArray
'        Next
'    ElseIf TypeName(Data) = "TheTable" Then
'        For n = 0 To Data.getRowSize() - 1 '0-based
'            addRow Data.getRowStringArray(n)    'convert to array and add
'        Next
'    Else
'        MsgBox "TheTable::add() for this data type is not implemented."
'    End If
'End Sub

#End If
