VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TheFlexGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private myGrid As MSFlexGrid
    ' points to the msflexgrid
Public Sub openFile(flexGrid As MSFlexGrid)
    Set myGrid = flexGrid
End Sub
Public Sub copy(source As Object, Optional startRow = 0)
    Dim row, tempRow
    
    If TypeName(source) = "TheTable" Then
        Dim tempStr
        For row = startRow To source.getRowSize() - 1
            tempRow = TheTableBase.getRow(source, row)
            tempStr = the.getStringFromArrayOrCollection(tempRow, vbTab)
            myGrid.addItem tempStr
        Next
    Else
        eprint "TheFlexGrid::copy() only supports TheTable"
    End If
End Sub
