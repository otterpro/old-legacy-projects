VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TheExcelTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'=============================================================================
'   TheExcelApp
'   Single-sheet Excel
'
'   Dependency: TheTable/TheGrid (used in .copyTo())
'                   MSFlexGrid (Optional)
'
'   openFile(optional path="")
'       .openFile("d:\mypath\my.xls")   'opens this file
'       .openFile   'opens New sheet
'       if path = "" (empty), then it creates a new empty sheet.
'       TODO: If path doesn't exist, create the file first.
'   saveFile("d:\mypath\my.xls")
'   copyTo(TheTable, optional startRow=0)
'       copies the data to table.
'   Original Excel Cells are 1-based.
'   However, this converts to 0-based sheet.
'=============================================================================
Option Explicit
'Implements TheTableInterface

Dim excelApp As excel.Application, _
    workBook As excel.workBook, _
    sheet As excel.Worksheet, _
    range As excel.range

Dim myAutoFit As Boolean
Public Function getCell(ByVal columnNumber As Long, ByVal rowNumber As Long) As Variant
    getCell = sheet.Cells(rowNumber + 1, columnNumber + 1).value
        ' convert 1-based to 0-based
End Function

Public Sub setCell(ByVal columnNumber As Long, ByVal row As Long, ByVal cellData)
    sheet.Cells(row + 1, columnNumber + 1).value = cellData
    If VarType(cellData) = vbDate Then
        sheet.Cells(row + 1, columnNumber + 1).NumberFormat = "m/d/yy h:mm AM/PM"
        'sheet.Cells(row + 1, columnNumber + 1).Width = sheet.Columns(row+1,columnNumber+1).Width
        
    End If
        ' convert 1-based to 0-based
End Sub

' Heavily Modified from code "GetRealLastCell() found in MSDN Book Excerpt
' "Excel 2002 VBA Programmer's Reference" published by Wrox Press Limited; ISBN 1861005709
' This seems to be the only way to find the Column Size and Row Size of excel spreadsheet.
' Sheet.Rows.Count  returns 65535 (ie maximum rows available)
' Sheet.Cells.End... returns all rows that are formatted so it is not accurate.
Public Function getRowSize() As Long
  On Error Resume Next
  getRowSize = sheet.Cells.find("*", sheet.range("A1"), xlFormulas, , xlByRows, _
                                               xlPrevious).row
End Function
Public Function getColumnSize() As Long
  On Error Resume Next
  getColumnSize = sheet.Cells.find("*", sheet.range("A1"), xlFormulas, , _
                                    xlByColumns, xlPrevious).Column
End Function

Public Sub Class_Terminate()
    'close workbook
    closeFile
End Sub



Public Sub closeFile(Optional quit = False)
    If quit Then
        workBook.Close SaveChanges:=False
        excelApp.quit
    End If
    ' quit= false, Do not Quit or close workbook. Leave them open so that user can decide if he/she wants to
    ' close/save them or not.
    Set excelApp = Nothing
    Set workBook = Nothing
    Set sheet = Nothing
    Set range = Nothing

End Sub

' not sure what happens when one tries to open a file that doesn't exist.
Public Sub openFile(Optional path = "", Optional readOnly = False, _
                    Optional visible = True, Optional AutoFit = True)
    If path = "" Then   '-- NEW FILE
        Set excelApp = CreateObject("Excel.Application")
        setVisible (visible)
        ' Get a new workbook.
        Set workBook = excelApp.Workbooks.add
        Set sheet = workBook.ActiveSheet
    Else    '  -- OPEN FILE
        Set excelApp = CreateObject("Excel.Application")
        excelApp.visible = visible
        excelApp.UserControl = visible
        Set workBook = excelApp.Workbooks.Open(path)
        Set sheet = workBook.ActiveSheet
        'excelApp.Visible = True
        'excelApp.UserControl = True
        ' Get a new workbook.
        'Set workBook = excelApp.Workbooks.Add
        'Set sheet = workBook.ActiveSheet
    End If
        myAutoFit = AutoFit
        'sheet.Columns.AutoFit    'select all and autofit
End Sub

' open a sheet from an embedded OLE Excel Sheet.
Public Sub openOle(ByVal path As String, ByRef oleControl As OLE)
    'Set excelApp = oleControl.object.Application '<- no need to set excelApp since it will get closed at the dtor() if we set it here.
    'Set workBook = excelApp.ActiveWorkbook  ' <- doesn't work. OLE doesn't have ActiveWorkbook, only sheets, I think.
    'Set sheet = excelApp.Workbooks(1).ActiveSheet
    'Set sheet = oleControl.object.Application.Workbooks(1).ActiveSheet
    oleControl.CreateEmbed path
    Set excelApp = oleControl.object.Application
    Set workBook = excelApp.Workbooks(1)
    Set sheet = workBook.ActiveSheet
    'Set sheet = oleControl.object.Application.ActiveWorkbook.ActiveSheet
    
End Sub


' accepts  msflexgrid and thetable. Try not to use msflexgrid.
' Instead, use theTable as msflexgrid will be phased out unless
' it is useful for optimizing ( ie no translation overhead from
' grid->table->excel. Instead grid->excel.
'
'Public Sub copy(ByRef data)
Public Sub add(ByRef Data)
    If TypeName(Data) = "MSFlexGrid" Then   ' Deprecated.
        For n = 0 To Data.Rows - 1 '0-based
            For m = 0 To Data.Cols - 1 '0-based
                setCell m, n, Data.TextMatrix(n, m)
                ' excel is 1-based. grid is 0-based
                ' Debug.Print data.TextMatrix(n, m)
            Next
        Next
    ElseIf TypeName(Data) = "TheTable" Then
        warn " TheTable Not supported. Use TheTreeTable instead"
        Dim r, c
        For r = 0 To Data.getRowSize() - 1
            For c = 0 To Data.getColumnSize() - 1
            
                setCell c, r, Data.getCell(c, r)
            Next
        Next
    ElseIf TypeOf Data Is TheAdoTable2 Then
        Dim tablePtr As TheAdoTable2
        Set tablePtr = Data
        tablePtr.gotoFirst
        While Not tablePtr.isEof()
            addRowOfList tablePtr.getRow()
            tablePtr.gotoNext
        Wend
    Else
        eprint "TheExcelApp:copy() for this data type is not implemented."
    End If
    If myAutoFit Then sheet.Columns.AutoFit
End Sub

Public Sub setVisible(Optional visible = True)
        excelApp.visible = visible
        excelApp.UserControl = visible
End Sub

' create a single row and add the row of data, given the list
' can pass either the array or vararg, but not combination of both.
' If the item exceeds the myColumnSize, it increases myColumnSize.
Public Sub addRow(ParamArray Data())
'    Dim colSize As Long
    
    Dim newRow As Long
    newRow = getRowSize()    'RowSize++
    Dim col     'temp column index
    col = 0
    Dim item      ' individual cell's value
    
    If IsArray(Data(0)) Then
        For Each item In Data(0)
            setCell col, newRow, item
            col = col + 1  'col++
        Next
    Else
        For Each item In Data

            setCell col, newRow, item
            col = col + 1  'col++
        Next
    End If
    
End Sub


'same as addRow() except that it doesn't use paramarray.
Public Sub addRowOfList(row)
    Dim newRow As Long
    newRow = getRowSize()    'RowSize++
    
    'myRowSize = myRowSize + 1 'myRowSize++
    Dim col     'temp column index
    col = 0
    Dim item      ' individual cell's value
    
    If IsArray(row) Or isCollection(row) Then
        For Each item In row
            setCell col, newRow, item
            col = col + 1  'col++
        Next
 '   ElseIf isCollection(Data(0)) Then
    End If
End Sub

Public Sub printOut()
    sheet.printOut Copies:=1, Collate:=True
End Sub

