VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TheDataGridTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' UNTESTED, DANGEROUS. DO NOT USE YET.

Option Explicit
Private myGrid As DataGrid

    ' ref to the datagrid


Public Sub openFile(grid As DataGrid)
    Set myGrid = grid
End Sub

Public Function getRow(ByVal row As Long) As Collection
    Dim list As New Collection
    
End Function


Public Function getRowSize() As Long
    Dim recSet As ADODB.recordset
    Set recSet = myGrid.DataSource
    getRowSize = recSet.RecordCount
        'warning: .RecordCount doesn't work with certain ADO CursorType
End Function


#If NOT_IMPLEMENTED_YET Then
Public Sub copy(ByRef source As Object, Optional startRow = 0)
    Dim row, tempRow
    
    If TypeName(source) = "TheTable" Then
        Dim tempStr
        For row = startRow To source.getRowSize() - 1
            tempRow = TheTableBase.getRow(source, row)
            tempStr = the.getStringFromArrayOrCollection(tempRow, vbTab)
            myGrid.AddItem tempStr
        Next
    Else
        eprint "TheFlexGrid::copy() only supports TheTable"
    End If
End Sub

#End If

