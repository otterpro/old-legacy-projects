VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TheList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'#Const DEBUG_MODE = 1
'===============================================================================
' TheList
'0-based, not 1-based like the Collection
'
' add(value)
' add(value, "string") like assoc/array hasn't been tested
' add(Object) hasn't been tested.
' add(value, position)
'   adds value at position.
'   warning: add(x,n), if index "n" exceeds the current size of the list, it adds the item to the
'   last position, not at the nth position. Don't rely on add() to maintain the same position as "n".

' TODO:
'   TheList::copy(Collection): ,not tested.
'   does a reference copy,  myList=Collection, it is faster, better.
'===============================================================================


Option Explicit

'local variable(s) to hold property value(s)
Private myList As Collection 'local copy

Public Function getSize() As Long
    getSize = myList.Count
End Function



Public Property Let list(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.list = 5
    ' myList = vData
End Property


Public Property Set list(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.list = Form1
    Set myList = vData
    'can 't assign to array
End Property


Public Property Get list() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.list
    If IsObject(myList) Then
        Set list = myList
    Else
        list = myList
    End If
End Property

Public Sub Class_Initialize()
   Set myList = New Collection
   'the.dprint "TheList::Initialize()", ""
End Sub
Public Sub Class_Terminate()
    'myList = Null
End Sub

' USE get List() instead.
'Public Function getNumberList() As Collection
'    Set getNumberList = myList
'End Function

'getString( ) -> returns "12,34,56,78"
'getString() -> returns "apple,banana,cherry"
Public Function getString() As String
    getString = the.getString(myList)
End Function

Public Sub dprint()
    the.dprint "TheList", "[ " & getString() & " ]"
End Sub

Public Function getItem(ByVal index As Long) As Variant
    'If Not rangeIsValid(index) Then  'Removed for efficiency.
    'End If
    getItem = myList.item(index + 1)
        ' add 1 to convert from 0-based to 1-based (collection=1)
End Function

' Insert operation. If item exists, it removes it first and replaces the item.
' if the item doesn't exist already, it adds to the end of the list.
' Collection cannot set data using .Item(x). So delete the item, and then re-insert it.
Public Sub setItem(key, value)
    'If Not rangeIsValid(index) Then  'Removed for efficiency.
    'End If
    If Not Me.removeKey(key) Then
        add value, key 'add new item
    Else
        add value, key
    End If
End Sub

' returns true if remove is ok. False if remove fails.
Public Function removeKey(key) As Boolean
    On Error GoTo removeError  'if item doesn't exist, add it as new item.
    If VarType(key) = vbString Then
        myList.remove key 'remove the item first.
    Else
        myList.remove key + 1 ' collection=1 based, list=0 based
    End If
    removeKey = True
    Exit Function
removeError:
    removeKey = False
    'the.dprint "theList::remove(key)", "Key does not exist"
End Function

Public Function remove(value) As Boolean
    'find it first.
    Dim i As Long
    
    For i = 1 To myList.Count
        If IsObject(myList.item(i)) Then
            If myList.item(i) Is value Then
                remove = True
                myList.remove i
                Exit Function
            End If
        ElseIf myList.item(i) = value Then
            remove = True
            myList.remove i
            Exit Function
        Else
            'nothing to do.
        End If
    Next
    remove = False
End Function

' add(5)
' add(3.4)
'add ("Apple")
'add("Banana")
' list = [5, 3.4, "Apple", "Banana"]

Public Sub add(ByVal Data As Variant, Optional key)
    If IsMissing(key) Then
        myList.add Data   'normal add
    Else    '
        If VarType(key) = vbString Then
            myList.add Data, key  ' add with key, ie like hash table/assoc array
        Else
            If key + 1 >= myList.Count Then    ' last item
                myList.add Data, after:=Min(key + 1, myList.Count) ' add at last position
                    ' to avoid exceeding the range, use Min()
            Else
                myList.add Data, before:=key + 1 ' add at x position
            End If  'last item
        End If  ' vbString
    End If  'IsMissing()
End Sub

' add list of number in string format or a collection of any data.
' addList("12,45,423,45")
' addList (myCollectionOfCar)
Public Sub addNumberList(ByVal Data As Variant)
    Dim i
    If TypeName(Data) = "String" Then
        Dim numList As Collection
        Set numList = getNumbersFromString(Data)
        add numList
    ElseIf TypeName(Data) = "Collection" Or VarType(Data) = vbArray Then
        For Each i In Data
            add i
        Next
    Else
        'ERROR: unknown type
    End If
    
End Sub

' Do a quick shallow copy of Collection. Other objects are not implemented yet.
Public Sub copy(ByRef Data)
    If TypeName(Data) = "Collection" Then
       Set myList = Data      'quick & easy. Just copy ref
    Else
        eprint "TheList::copy(data), data type not recognized or implemented."
    End If
End Sub

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
   Set NewEnum = myList.[_NewEnum]
End Function

' Text taken from  House of Bricks (Custom Control) example from MSDN.
'Enabling For Each ・ Next
'Along with robustness, you get For Each ・ Next back. Once again you can delegate all the work to the Collection object, by adding the following method:
'
'' NewEnum must return the IUnknown interface of a
'' collection's enumerator.
'Public Function NewEnum() As IUnknown
'   Set NewEnum = mcolEmployees.[_NewEnum]
'End Function
'
'The important thing you're delegating to the Collection object is its enumerator. An enumerator is a small object that knows how to iterate through the items in a collection. You can't write an enumerator object with Visual Basic, but because the Employees class is based on a Collection object, you can return the Collection object's enumerator ? which naturally enough knows how to enumerate the items the Collection object is holding.
'
'The square brackets around the Collection object's _NewEnum method are necessary because of the leading underscore in the method name. This leading underscore is a convention indicating that the method is hidden in the type library. You can't name your method _NewEnum, but you can hide it in the type library and give it the procedure ID that For Each ・ Next requires.
'
'To hide the NewEnum method and give it the necessary procedure ID
'
'On the Tools menu, click Procedure Attributes to open the Procedure Attributes dialog box. In Name box, select the NewEnum method.
'
'
'Click Advanced to show the advanced features. Check Hide this member to make NewEnum hidden in the type library.
'
'
'In the Procedure ID box, type ?4 (minus four) to give NewEnum the procedure ID required by For Each ・ Next. Click OK.
'Important   In order for your collection classes to work with For Each ・ Next, you must provide a hidden NewEnum method with the correct procedure ID.

#If DEBUG_MODE Then
Public Sub test()
    the.dprint "===="
    add "alpha"
    add "Beta"
    add 1
    add 2
    add 3
    dprint
    the.dprint "----"
    'remove 2
    setItem 2, 500
    add "quest"
    setItem 0, "alpha_new"
    setItem 8, "last"
    setItem 6, "ultimateLast"
    dprint
    
End Sub
#Else
    Public Sub test()
    ' do nothing
    End Sub
#End If

