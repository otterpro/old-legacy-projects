VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TheArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
#Const DEBUG_MODE = 1

'=============================================================================
'TheArray
'
'=============================================================================
Option Explicit

Dim myArray     'array

'public property let array() <- array is reserved keyword. Use getArray()

Public Function getArray() As Variant
    getArray = myArray
End Function

Public Sub setArray(arrays)
    If VarType(arrays) < vbArray Then
        eprint "TheArray::setArray() expecting an array, got something else."
        Exit Sub
    End If
    myArray = arrays
End Sub

'Recursive quicksort implementation. Only supports ascending order.
'It is used by sort() which uses this func to add both asc and desc sort.
'VBTK p7
Private Sub qsort(startElement, stopElement)
    Dim workStart As Long, workStop As Long
    Dim x As Long, y As Long, i As Long
    workStart = startElement
    workStop = stopElement
    'get the halfway point and assign it to x
    x = myArray((startElement + stopElement) / 2)
    'x now holds value of array item that is halfway between start and stopElement.
    'The next stop is to compare the rank of workStart and workStop to determine
    'when workstart (lbound value) equal or exceed workstop(ubound value)
    While (workStart <= workStop)
        'while workstart remains less than workstop, compare the value of the
        'array() element at this position against x (set above). In addition,
        'also make sure the workstart remains less than the stopElement bound value.
        While (myArray(workStart) < x And workStart < stopElement)
            workStart = workStart + 1
        Wend
        'repeat the same for the array value at position workStop, also making
        'sure that workstop remains greater than the startElement bound value.
        While (x < myArray(workStop) And workStop > startElement)
            workStop = workStop - 1
        Wend
        'determine the element to swap based on the final workstart & workstop
        'relative positions.
        If (workStart <= workStop) Then
            'a. assign tmp var y to value at anarray(workStart)
            'b. swap anarray(workStop) for myArray(workStart)
            'c. reassing the value of temp to the type myArray(workstop)
            y = myArray(workStart)
            myArray(workStart) = myArray(workStop)
            myArray(workStop) = y
            'change the start & stop items
            workStart = workStart + 1
            workStop = workStop - 1
        End If
    Wend
    'if the original startelement is still less than workstop, then call the
    'sub again with startelement & workstop as the start & stop points
    If (startElement < workStop) Then qsort startElement, workStop
    'or if the original startelement is still less than workstop, then
    'call the sub again with workstart and stopelement as the start &
    'stop points
    If (workStart < stopElement) Then qsort workStart, stopElement
End Sub

Public Sub sort(Optional startElement As Long, _
                Optional stopElement As Long, _
                Optional sortOrder As SORT_ORDER = ASCENDING)
    Dim x As Variant, y As Variant, i As Long
    If VarType(myArray) < vbArray Then Exit Sub
    
    'assign to the working var the value pased to the sub as startElement &
    'stopElement. If none passed then assign dimensions of array
    If startElement = 0 And stopElement = 0 Then
        startElement = LBound(myArray)
        stopElement = UBound(myArray)
    Else
        'fix up
        If stopElement > UBound(myArray) Then stopElement = UBound(myArray)
        If startElement < LBound(myArray) Then startElement = LBound(myArray)
    End If
    'let our internal qsort array to the work
    qsort startElement, stopElement
    'if descend then inverse element
    If sortOrder = DESCENDING Then
        y = UBound(myArray)
        For i = LBound(myArray) To y \ 2
            x = myArray(i)
            myArray(i) = myArray(y)
            myArray(y) = x
            y = y - 1
        Next
    End If
End Sub
            
'finds an item in a sorted array and returns the index of the match.
'Returns -1 if not found.
Public Function binarySearch(item As Variant) As Long
    Dim i As Long, p As Long, lowerBound As Long, upperBound As Long
    If VarType(myArray) < vbArray Then Exit Function
    lowerBound = LBound(myArray)
    upperBound = UBound(myArray)
    p = upperBound
    'assume not in the list
    binarySearch = -1
    If myArray(lowerBound) < myArray(upperBound) Then
        While item < myArray(p)
            p = p \ 2 'move ptr back 1/2 oflist
            If p <= lowerBound Then Exit Function
        Wend
    Else
        While item > myArray(p)
            p = p \ 2 'move ptr ahead 1/2 oflist
            If p = 0 Then Exit Function
        Wend
    End If
    For i = p To upperBound
        If myArray(i) = item Then
            binarySearch = i
            Exit Function
        End If
    Next
End Function


' VBTK
'slow insert. Works with both ascending and descending sorted array
'Uses brute force to find theinsertion point for new item.
'For faster insert, use it & call binarySearch() instead.
Public Sub insertSorted(item As Variant)
    Dim i As Long, j As Long, p As Long
    Dim startElement As Long, stopElement As Long
    If VarType(myArray) < vbArray Then Exit Sub
    'resize array (make it bigger)
    ReDim Preserve myArray(LBound(myArray) To UBound(myArray) + 1)
    'get start/stop elements
    startElement = LBound(myArray)
    stopElement = UBound(myArray)
    'fixup for inverse sorting. Inverse order.
    If startElement > stopElement Then
        startElement = UBound(myArray)
        stopElement = LBound(myArray)
    End If
    'verify
    If startElement < LBound(myArray) Or stopElement > UBound(myArray) Then Exit Sub
    'find where to insert
    If myArray(startElement) < myArray(stopElement - 1) Then
        For j = startElement To stopElement
            If item < myArray(j) Then Exit For
        Next
    Else
        For j = startElement To stopElement
            If item > myArray(j) Then Exit For
        Next
    End If
    'fix up for end or start of array
    If j > stopElement Then j = stopElement
    'now move elements
    For i = stopElement To j + 1 Step -1
        myArray(i) = myArray(i - 1)
    Next
    'insert item
    myArray(j) = item
End Sub

'find and delete the element. Then move up all the elements.
'TODO: add additional param to be able to remove a range of item
    'from the array.  Also be able to pass an array of items to
    'delete instead of just a single item.
Public Sub remove(item As Variant)
    Dim p As Long, i As Long
    Dim startElement As Long, stopElement As Long
    If VarType(myArray) < vbArray Then Exit Sub
    'get start/stop element
    stopElement = UBound(myArray)
    'find element
    p = binarySearch(item)
    If p > 0 Then
        'now move elements
        For i = p To stopElement - 1
            myArray(i) = myArray(i + 1)
        Next
        'resize array
        ReDim Preserve myArray(LBound(myArray) To UBound(myArray) - 1)
    End If
End Sub







'=============================================================================
'=============================================================================
#If DEBUG_MODE Then
Public Sub testSort()
    Dim arr
    arr = Array(22, 5, 1, 6, 2, 3, 5, 45, 72, 9)
    the.dprint "arr before sort", arr
    Me.setArray (arr)
    Me.sort
    the.dprint "arr after sort", Me.getArray()
    
    Me.insertSorted 33
    the.dprint "arr after inserting 33=", Me.getArray()
    
    Me.sort , , DESCENDING
    the.dprint "arr desc sort", Me.getArray()
    
    the.dprint "find 22=", myArray(Me.binarySearch(22))
    remove 6
    the.dprint "after removing 6", Me.getArray()
End Sub
#End If
