










Private Sub Class_Terminate()
    closeTable
End Sub




Public Sub setCell(ByVal col As Long, ByVal row As Long, ByVal cellData)
    'rangeIsValid col, row   'range check not implemented yet
    myRecordset.AbsolutePosition = row + 1  ' convert 1-based to 0-based
    myRecordset(col) = cellData     'col is 0-based
    saveRecord      'save change to disk if needed
End Sub

Public Function getCell(ByVal col As Long, ByVal row As Long) As Variant
    'rangeIsValid col, row   'range check not implemented yet
    myRecordset.AbsolutePosition = row + 1  ' 1-based to 0-based
    getCell = myRecordset(col)  ' col is 0-based
End Function










    







