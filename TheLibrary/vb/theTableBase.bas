Attribute VB_Name = "TheTableBase"
'=============================================================================
'   TheTableBase
'   TODO: Incorporate this into TheTreeTable class.
'
'   Base class for supporting common functionality for all table-based data structure.
'   They include TheTable, TheExcelApp, etc.
'   It is very generic and slow in many cases. Override them using the subclass if needed.
'
'   Dependency: The.bas, TheTableInterface.cls
'
'   The derived (sub) class requires the following method to be defined:
'   getCell(x,y), getColumnSize(), addRow()
'
'=============================================================================
Option Explicit

' get a row and return an array containing the entire row. Slow. Override it with subclass.
Public Function getRow(ByRef self, ByVal row As Long) As Variant
    'rangeIsValid 0, row    'no range checking done. Could trap error as an alt?
    Dim i
    Dim returnArray
    ReDim returnArray(0 To self.getColumnSize() - 1)
    For i = 0 To self.getColumnSize() - 1
        returnArray(i) = self.getCell(i, row)
    Next
    getRow = returnArray
End Function

' Export table data to another type. It should only support a universal type, which is TheTable.
' From TheTable, it can be converted to just about any other table format using the object's
' .Copy() method.
'
' Warning: Make sure to Clear() the destination before copying. Or else, it will append to its existing data.
'
' startRow := 0-based, starting row index #. For example, to skip the 1st row, which contains headers,
'       set startRow:=1.
Public Sub copyTo(source As Object, destination As Object, Optional startRow = 0)
    Dim row, tempRow
    
    'If TypeName(destination) = "MSFlexGrid" Then
    '    Dim tempStr
    '    For row = startRow To source.getRowSize()
    '        tempRow = getRow(source, row)
    '        tempStr = the.getStringFromArrayOrCollection(tempRow, vbTab)
    '        destination.addItem tempStr
    '    Next
    'Else
    If TypeName(destination) = "TheTable" Then
        For row = startRow To source.getRowSize() - 1
            tempRow = getRow(source, row)
            destination.addRow tempRow
        Next
    Else
        eprint "TheTableBase::copyTo(), datatype not supported. It only supports TheTable."
    End If
End Sub

Public Function getRowString(source As Object, ByVal row As Long) As String
    getRowString = the.getString(TheTableBase.getRow(source, row))
End Function

' append multiple rows of data from the given table to this table.
' same as .copyFrom() except that it doesn't do .clear() first.
' UNTESTED
Public Sub add(self As Object, table As Object)
    ' Redundant codes were commented out since the core functions are handled in addRow().
    'If TypeName(Data) = "MSFlexGrid" Then
    '    Dim tempArray()
    '    ReDim tempArray(0 To Data.Cols() - 1)
    '    For n = 0 To Data.Rows() - 1 '0-based
    '        For m = 0 To Data.Cols() - 1 '0-based
    '            tempArray(m) = Data.TextMatrix(n, m)
    '        Next
    '        addRow tempArray
    '    Next
    'If TypeName(Data) = "TheTable" Then
    Dim i
    For i = 0 To table.getRowSize() - 1 '0-based
        self.addRow TheTableBase.getRow(table, i)   'get array and add
    Next
    'Else
    '    MsgBox "TheTable::add() for this data type is not implemented."
    'End If
End Sub

Public Sub dprint(source As Object, Optional title = "TheTable::", Optional col = -1, Optional row = -1)
    If row = -1 Then 'Or col =-1 Then 'no need to check for row
        'print all
        Dim columnSize, rowSize
        columnSize = source.getColumnSize()
        rowSize = source.getRowSize()
        the.dprint title, "colSize=" & CStr(columnSize) & _
            ", rowSize=" & CStr(rowSize)
        Dim i, j ', tempString
        For i = 0 To rowSize - 1
            the.dprint "(Row" & str(i) & ")", TheTableBase.getRowString(source, i)
        Next
    Else
        'print only the given columnNumber,row
        the.dprint title & ":(" & CStr(col) & "," & CStr(row) & ")", _
            source.getCell(col, row)
    End If
End Sub





'=============================================================================
'   OLD CODE
'=============================================================================
#If OLD_CODE Then
Public Sub dprintOld(source As Object, Optional title = "TheTable::", Optional col = -1, Optional row = -1)
    If row = -1 Then 'Or col =-1 Then 'no need to check for row
        'print all
        Dim columnSize, rowSize
        columnSize = source.getColumnSize()
        rowSize = source.getRowSize()
        the.dprint title, "colSize=" & CStr(columnSize) & _
            ", rowSize=" & CStr(rowSize)
        Dim i, j ', tempString
        For i = 0 To rowSize - 1
        
            For j = 0 To columnSize - 1
                'tempString =
                the.dprint "(" & str(j) & "," & str(i) & ")", _
                    source.getCell(j, i)
            Next
        Next
    Else
        'print only the given columnNumber,row
        the.dprint title & ":(" & CStr(col) & "," & CStr(row) & ")", _
            source.getCell(col, row)
            
    End If
    
End Sub
#End If
