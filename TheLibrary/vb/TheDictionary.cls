VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TheDictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'===============================================================================
' TheDictionary
' aka assoc array, hash table, etc.
'
' add(key, value)
'       Keys are Type-Sensitive.  Dict("ABC") <> Dict("abc"). Key is always String
'       Values can be object as well as built-in type. Tested ok with TheList.
'       also Identical to setValue(key,value). add() is preferred.
'       if key is empty ("") then generate a numeric key ("0".."n"). However, this
'           is treated as an error.
' remove(key)
'
' getValue(key)
'
' getString()
' clear()
' keyExists()
' getSize()

'===============================================================================

' Example Code
'     Dim dict
'    Set dict = New TheDictionary
'    dict.add "apple", 3
'    dict.add "banana", 4
'    dict.add "apple", 5
'    val = dict.getValue("candy")
'    dict.dprint
' see also Test()


'local variable(s) to hold property value(s)
Option Explicit
Private myDict As Object 'local copy

Public Sub clear()
Attribute clear.VB_Description = "Remove all item"
    myDict.RemoveAll
End Sub

'see hasKey().
Public Function keyExists(ByVal key As String) As Boolean
    keyExists = CBool(myDict.Exists(key))
End Function

'alias to keyExists()
Public Function hasKey(ByVal key As String) As Boolean
    hasKey = keyExists(key)
End Function

Public Function getSize() As Long
    getSize = myDict.Count()
End Function

Public Function getValue(ByVal key As String) As Variant
    ' getValue = IIf(keyExists(Key), myDict.Item(Key), "")
    'can't use this because myDict() is evaluated when it shouldn't
    If keyExists(key) Then
        If VarType(myDict.item(key)) = vbObject Then
            Set getValue = myDict.item(key)
        Else
            getValue = myDict.item(key)
        End If
    Else
        getValue = ""
    End If
End Function

Public Function getKeys()
    'Dim returnValue
    getKeys = myDict.keys()
    
    'returnValue = myDict.keys()
    'getKeys = returnValue
End Function

' Identical to Add() since add() handles the duplicate key
Public Sub setValue(ByVal key As String, value)
    add key, value
End Sub

Public Sub Class_Initialize()
    Set myDict = CreateObject("Scripting.Dictionary")
End Sub

'if key is empty (""), then create a numeric key. However, this should be treated as an error case.
Public Sub add(ByVal key As String, value)
    Static index As Long
        'if key is empty (""), then create a numeric key. However, this should be treated as an error case.
    If key = "" Then
        key = CStr(index)   'generate a key (ie "0"..."n" since key is missing)
        eprint "TheDictionary::add(): key is empty string.The dict should not have empty key. Value=" & _
            the.getString(value) & " Generated Key=" & key
        index = index + 1 'no ++ in VB although exists in VB.Net
    End If
    remove key  'removes if the key exists. Else, it ignores it.
    myDict.add key, value
End Sub

Public Function remove(ByVal key) As Boolean
    If myDict.Exists(key) Then
        myDict.remove key
        remove = True
    Else
        remove = False
    End If
End Function

Public Function getString() As String
    Dim key
    getString = ""
    For Each key In myDict
        getString = getString & key & "=>" & the.getString(myDict.item(key)) & ","
    Next
    chop getString  ' get rid of last comma
End Function

Public Sub dprint()
Attribute dprint.VB_Description = "print all dict item to debug "
    the.dprint "TheDictionary", "[" & getString() & "]"
End Sub

Public Property Let dict(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dict = 5
    myDict = vData
End Property


Public Property Set dict(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.dict = Form1
    Set myDict = vData
End Property


Public Property Get dict() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dict
    If IsObject(myDict) Then
        Set dict = myDict
    Else
        dict = myDict
    End If
End Property

' UNTESTED!
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
   Set NewEnum = myDict.[_NewEnum]
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

'=============================================================================
'=============================================================================
#If DEBUG_MODE Then
'    Public Sub selftest()
'
'        add "apple", 3
'        add "banana", 4
'        add "cherry", 100
'        add "apple", 5
'
'        Dim l As New TheList
'        l.add "L1"
'        l.add "L2"
'        l.add "L3"
'
'        add "List", l
'        dprint
'        the.dprint "getValue(candy)", getValue("candy")
'        the.dprint "getValue(banana)", getValue("banana")
'        the.dprint "getValue(List)", getValue("List")
'    End Sub
#Else
'    Public Sub test()
'    End Sub
#End If

'=============================================================================
' OLD
'=============================================================================
#If OLD Then
' Counter Broken
' counter methods - used when dictionary is used as a counter
'   for counting the # of words, etc.
' .inc (key) increase value by 1
' .dec (key) decreases its value by 1

'   dec(), inc() increases counter value if it is used
'   for counting. IT also keeps track of the index #.
'   Value format = "index #, count #"
'   Negative #s are NOT ALLOWED for simplicity.
'   .getCount("key") returns only the counter portion of the value
'   same as .getSize("")
'Public Sub inc(ByVal key As Variant)
'    Dim numList As Collection, countValue, _
'        newList As New Collection
'
'    Dim text As String, valueStr As String
'
'    If keyExists(key) Then
'        valueStr = myDict.item(key)
'        Set numList = getNumbersFromString(valueStr)
'        'dprint "numList", numList
'        countValue = numList(2)
'        'dprint "numList", numList
'        If IsNumeric(countValue) Then
'            newList.add (numList(1))
'            newList.add countValue + 1
'            text = getStringFromNumbers(newList)
'            add key, text
'        Else
'            ' add key, 1
'            ' Error. invalid format
'        End If
'
'    Else    ' NEW KEY
'        Set numList = New Collection
'        numList.add getSize()
'        numList.add 1
'        text = getStringFromNumbers(numList)
'        'Debug.Print "text="; text
'        add key, text
'    End If
'End Sub
'
'
'Public Function getCount(ByVal key As String) As Long
'    Dim numList As Collection, countValue, valueStr
'
'    If keyExists(key) Then
'        valueStr = myDict.item(key)
'        Set numList = getNumbersFromString(valueStr)
'        'dprint "numList", numList
'        countValue = numList(2)
'        'dprint "numList", numList
'        If IsNumeric(countValue) Then
'            getCount = countValue
'        Else
'            getCount = 0
'            ' not a valid count format
'        End If
'    Else
'        getCount = 0
'        ' key doesn't exist. count=0
'    End If
'End Function
'
'Public Sub dec(ByVal key As Variant)
'    Dim numList As Collection, countValue, _
'        newList As New Collection
'
'    Dim text As String, valueStr As String
'
'    If keyExists(key) Then
'        valueStr = myDict.item(key)
'        Set numList = getNumbersFromString(valueStr)
'        'dprint "numList", numList
'        countValue = numList(2)
'        'dprint "numList", numList
'        If IsNumeric(countValue) Then
'            If countValue <= 0 Then
'                'ERROR: Cannot be a negative #
'                Exit Sub
'            End If
'            newList.add (numList(1))
'            newList.add countValue - 1
'            text = getStringFromNumbers(newList)
'            add key, text
'        Else
'            ' add key, 1
'            ' Error. invalid format
'        End If
'    Else
'        ' ERROR. key doesn't exist. Impossible.
'    End If
'
'End Sub
#End If

