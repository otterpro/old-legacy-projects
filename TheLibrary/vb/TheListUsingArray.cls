VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TheListUsingArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Const CHUNK_SIZE = 100
    ' Array "list" is allocated chunk_size at a time.

'local variable(s) to hold property value(s)
Private mvarlist() As Variant 'local copy
Private size As Long    'size of list, not necessarily amt allocated
Private actualSize As Long  'actual allocated size of list


Public Sub addList(ByVal Data As Variant)
    Dim i
    If TypeName(Data) = "String" Then
        Dim numList As Collection
        Set numList = getNumbersFromString(Data)
        addList numList
    ElseIf TypeName(Data) = "Collection" Or VarType(Data) = vbArray Then
        For Each i In Data
            Add i
        Next
    Else
        'ERROR: unknown type
    End If
    
End Sub

Public Function getSize() As Long
    getSize = size
End Function



Public Property Let List(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.list = 5
    mvarlist = vData
End Property


Public Property Set List(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.list = Form1
    'Set mvarlist = vData
    'can't assign to array
End Property


Public Property Get List() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.list
    'If IsObject(mvarlist) Then
    '    Set list = mvarlist
    'Else
    '    list = mvarlist
    'End If
End Property

Public Sub Class_Initialize()
    size = 0
    actualSize = 0
End Sub
Public Sub Class_Terminate()

End Sub
Public Function getNumberList() As Collection
    Dim i
    Dim numList As New Collection
    For i = 0 To size
        numList.Add (mvarlist(i))
    Next

End Function

Public Function getStringList() As String
    Dim i
    Dim tempString As String
    tempString = ""
    If size = 0 Then
        getStringList = ""
        Exit Function
    End If
    For i = 0 To size - 1
        tempString = tempString & Str(mvarlist(i)) & ","
    Next    ' Remove last ','
    chop tempString
    getStringList = tempString

    'Strange Problem: Suddenly Left() doesn't work.
    ' I get error message saying that It is expecting Array
    'Debug.Print Mid("Hello", 1, 4)

End Function

Public Sub dprint()
    Dim tempString As String
    tempString = getStringList()
    Debug.Print "[ " & tempString & " ]"

End Sub

Public Function getItem(ByVal index As Long) As Variant
    If index >= size Then
        ' ERROR: index out of range
            MsgBox ("Error: list.getItem() index out of range")
        Exit Function
    End If
    getItem = mvarlist(index)
End Function

Public Sub Add(ByVal Data As Variant)
    size = size + 1
    If (size >= actualSize) Then
        actualSize = actualSize + CHUNK_SIZE
        ReDim Preserve mvarlist(actualSize)
    End If
    mvarlist(size - 1) = Data
End Sub

