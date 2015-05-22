Attribute VB_Name = "the"
#Const DEBUG_MODE = 1

'=============================================================================
'The - generic library of functions
'=============================================================================
'
'STRING
    'getNumbersFromString("12,45") -> returns collection [12,45]
    'getStringFromNumbers([12,45]) -> returns "12,45" ' Not tested?
    'getNumberOrString(str): returns either number or string depending on string type
    '       ("45") returns 45(int). ("abc") returns "abc". ("2.45") returns 2.45(float)
    'getCharAt(str,index) -> like str[i] which VB lacks. str Index=0-based
    'stripIllegalChar(str, findChars)
    '
'    isEmptyOrNullString(var) = True if var is string and is "".
'       also True if var is NULL (variant)
'   isNotNull()
'   getSize() returns size of an array,paramarray, string, Collection, The
'   isAlmost(float1, float2)
'       compare 2 real values, and returns true if they are almost same
'       ie 9.999 = 10.0000
'       useful for comparing real # since they can vary a little bit.
'       isAlmost(9.99, 10.0)=> returns TRUE
'Var/Class Type
'   setEmptyTypeBasedOn(exampleVar)
'       if exampleVar is string, it returns "". If it is number, it returns 0.
'       good for setting empty variant type which needs specific type.
'
' UI
    'moveLine(lineControl, x1,y1, x2,y2)
    'NOT IMPL: 'moveLineBox (lineControl
'Debugging Functions
'   dprint "title", var1,var2,...
'       prints debugging message. SImilar to MFC's TRACE()
'=============================================================================

Option Explicit

Public Const QUOTE_CHARS = "'"""    ' Single and Double quotes ['"]
Public Const QUOTE_CHAR = """"

Public Enum SORT_ORDER  'used by sort()
    ASCENDING
    DESCENDING
End Enum


    'remember, #Const are always Private. I wish for #Include.

Public m, n   ' fast temp reusable var, for quick usage. Don't use it to call sub since this
            ' value may change. similar to $_ in perl. Strongly recommended that it is not used.
            
'   Dim numList, i
'   Set numList = getNumbersFromString("1 ,3 , 5")
'   dprint "type", TypeName(numList)
'   dprint "numList", numList


Private Const EPRINT_ERROR_CODE = 1001

' returns a collection given the embedded number in string
' getNumbersFromString("12,45") -> returns collection [12,45]
Public Function getNumbersFromString(ByVal text As String) As Collection
' Must be comma delimited
    Dim strList, i, numList As New Collection
    strList = Split(text, ",")
    For Each i In strList
        'Debug.Print i & "-"
        numList.add val(i)
    Next
    Set getNumbersFromString = numList
End Function


' get string version of any built-in type or user-type
' my extension of CStr(x) and Str(x)
Public Function getString(Data, Optional addQuoteToString = False) As String
    'Dim typeCode
   ' dprint "vartype", CStr(vartype(data))
   ' Exit Function
    If IsArray(Data) Then
        getString = getStringFromArrayOrCollection(Data)
        'getString = "array not impl"
    Else
        Select Case VarType(Data)
            Case vbString
                If addQuoteToString Then
                    getString = Chr(34) & Data & Chr(34)
                Else
                    getString = Data
                End If
            Case vbObject
                If TypeOf Data Is Collection Then
                    getString = getStringFromArrayOrCollection(Data)
                Else
                    getString = Data.getString()    'assuming that object has getString() method.
                End If
            Case Else
                getString = CStr(Data)
        End Select
    'getString = IIf(VarType(data) Or vbString, data, Str(data))
    End If
End Function

Public Function getNumberOrString(ByVal str As String) As Variant
    If IsNumeric(str) Then  ' number. find if float or int.
        If InStr(str, ".") > 0 Then 'must be floating point
            getNumberOrString = CSng(str)
        Else    ' must be integer
            getNumberOrString = CLng(str)
        End If
    Else
        getNumberOrString = str 'return itself. it is a string.
    End If
End Function

' separatorChar := char that separates each item. By default, it is ",". But vbTab may be needed
'   especially for entering data for msflexgrid's addItem()
' addQuoteToString := surrounds each string with " ". False by default.
Public Function getStringFromArrayOrCollection(Data, Optional separatorChar = ",", _
    Optional addQuoteToString = False) As String
    Dim i
    getStringFromArrayOrCollection = ""
    For Each i In Data
        getStringFromArrayOrCollection = getStringFromArrayOrCollection & getString(i, addQuoteToString:=addQuoteToString) & separatorChar
    Next
    the.chop getStringFromArrayOrCollection     'remove the last ","
End Function



' Returns a string enclosed by quote if there is any space in between the string.
' [As times go By] -> "[As times go by]"  ' quotes added
' Cars -> Cars  'no change
' useful for filenames with spaces.
' Warning: Does not account for the leading/trailing spaces.
' Don't know which module uses this function.
Public Function addQuoteIfNeeded(ByVal text As String) As String
    addQuoteIfNeeded = IIf(InStr(1, text, " "), QUOTE_CHAR & text & QUOTE_CHAR, text)
    ' --- Shortcut to codes below  -----
    'If InStr(1, text, " ") Then
    '    addQuoteIfNeeded = Chr(34) & text & Chr(34)
    'Else
    '    addQuoteIfNeeded = text
    'End If
End Function


Public Function getCharAt(ByVal text, ByVal index) As String
    getCharAt = Mid$(text, index + 1, 1)
        'index+1: convert 0-based to 1-based index.
End Function

' Strips the surrounding quote marks of a string and returns it.
' Also removes leading/trailing whitespace chars if quote was also removed.
Public Function removeLeadingAndTrailingQuote(ByVal text As String) As String
    Dim returnString As String
    returnString = lstrip(text)
    returnString = rstrip(returnString)
    Dim removedQuote As Boolean
    removedQuote = False
    If Mid(returnString, 1, 1) = Chr(34) Then
        returnString = Mid(returnString, 2, Len(returnString) - 1)
        removedQuote = True
    End If
    If Mid(returnString, Len(returnString), 1) = Chr(34) Then
        returnString = Mid(returnString, 1, Len(returnString) - 1)
        removedQuote = True
    End If
   removeLeadingAndTrailingQuote = IIf(removedQuote, returnString, text)
End Function


' get size of array or collection or lists (paramArray)
' Not tested on Collection
Public Function getSize(ByRef Data As Variant) As Long
    If IsArray(Data) Then
        getSize = UBound(Data) - LBound(Data) + 1
    ElseIf VarType(Data) = vbString Then
        getSize = Len(Data)
    ElseIf IsNumeric(Data) Then
        getSize = LenB(Data)
    ElseIf TypeOf Data Is Collection Then
        getSize = Data.Count()
    ElseIf TypeOf Data Is TheDictionary Then
        getSize = Data.getSize()
    ElseIf TypeOf Data Is TheList Then
        getSize = Data.getSize()
    Else
        getSize = 0 'unable to determine size yet.
    End If
End Function

' given a string, parse the key and value using = as delimiter.
' ("a=125") ->'   returns Collection with 1st element ="a", 2nd=125
'       we return Collection and not Array because we may implement
'       a dictionary with non-unique key
'       ie ("name=Doe,John,Anderson") -> (name=Doe,name=John,name=Anderson)
' previous: returns Array with 1st element ="a", 2nd=125
' If it is missing "=" then key is empty "" and only value is filled.
Public Function parseKeyValueString(ByVal keyValueStr As String) As Collection
    Dim keyValue As New Collection
    'Dim keyValue(0 To 1)   'used arrays before
    Dim key As String, value
    key = ""
    keyValueStr = Trim(keyValueStr)
    Dim index As Long
    index = InStr(keyValueStr, "=")
    value = getNumberOrString(Mid$(keyValueStr, index + 1))
    If index = 1 Then   'only case: if "=abc", missing key, but has leading "=".
        eprint "parseKeyValueString(): Mya be Illegal keyValueStr. <" & keyValueStr & ">"
    ElseIf index > 1 Then   ' key exists.
        key = Mid$(keyValueStr, 1, index - 1)
    End If
    keyValue.add key
    keyValue.add value
    Set parseKeyValueString = keyValue
End Function

' given a list of paramArray, convert it to Collection
Public Function getListFromParamArray(ParamArray vararg()) As Collection
    Dim list As New Collection
    Dim item
    If IsArray(vararg(0)) Then  ' func(myArray)
        For Each item In vararg(0)
            list.add item
        Next
    ElseIf TypeOf vararg(0) Is Collection Then  'func(myCollection)
        ' already a collection. Do nothing.
    Else    'a normal paramarray. func(1,3,5,7,9)
        For Each item In vararg
            list.add item
        Next
    End If
    Set getListFromParamArray = list
End Function

'given a list of paramArray, convert it to Dictionary.
' If key is missing ie "", then we generate a numeric key.
' However, don't rely on this feature since another key with same number will
' overwrite it.
'
Public Function getDictFromParamArray(ParamArray vararg()) As TheDictionary
    Dim dict As New TheDictionary
    Dim list As Collection, item
    ' first, convert paramArray to List
    If IsArray(vararg(0)) Then  ' func(myArray)
        Set list = getListFromParamArray(vararg(0))
    ElseIf TypeOf vararg(0) Is Collection Then  'func(myCollection)
        Set list = vararg(0)
        ' already a collection. Do nothing.
    Else    'a normal paramarray. func(1,3,5,7,9)
        Set list = New Collection
        For Each item In vararg
            list.add item
        Next
    End If
    
    Dim keyValue As Collection
    Dim key As String, value
    For Each item In list
        Set keyValue = parseKeyValueString(CStr(item))
        key = keyValue.item(1)
        value = getNumberOrString(keyValue.item(2)) 'value could be number or string
        dict.add key, value
    Next
    Set getDictFromParamArray = dict
End Function

Public Sub eprint(text)
    'print error message to user or hide it if necessary.
    Dim answer
    answer = MsgBox(getString(text) & vbCrLf & vbCrLf & Err.Description _
                , vbCritical + vbYesNo, _
                "Program Error: Click Yes to continue, No to quit.")
    If answer = vbNo Then
        Err.Raise vbObjectError + EPRINT_ERROR_CODE, "error", text
        End    ' halt the program here. Data are not saved!
    End If
End Sub

Public Sub chop(ByRef text As String)
    'text = Left(text, Len(text) - 1)
    ' alternate
    text = Mid(text, 1, Len(text) - 1)
    ' For odd reason, Left() isn't working.
End Sub

' strip any beginning whitespace including RET,TAB,etc. and any
' char <=32 (SPACE).
Public Function lstrip(ByVal text As String) As String

    For n = 1 To Len(text)  'n=char position, 1-based
        If Asc(Mid(text, n, 1)) > vbKeySpace Then
            Exit For
        End If
    Next
    lstrip = Right(text, Len(text) - n + 1)
End Function

' strip any ending whitespace including RET,TAB,etc. and any
' char <=32 (SPACE).
Public Function rstrip(ByVal text As String) As String
    For n = Len(text) To 1 Step -1  'n=char position, 1-based
        If Asc(Mid(text, n, 1)) > vbKeySpace Then
            Exit For
        End If
    Next
    rstrip = Left(text, n)
End Function

Public Function strip(ByVal text As String) As String
    text = lstrip(text)
    text = rstrip(text)
    strip = text
End Function

' Returns 1st word in a string, delimited by whitespace.
' Can use explode()/split() but decided to simply look for the first
' word that ends with ascii <=32.
' Does not strip the leading whitespace / no ltrim()
Public Function getFirstWord(ByVal text As String) As String
    'text = LTrim(text)
    ' LTRIM() doesn't remove CRLF So use
    text = lstrip(text)
    'skip any whitespace first.
    Dim test As String
    'look for whitespace indicating the word
    For n = 1 To Len(text)  '
        test = Mid$(text, n, 1)
        If Asc(Mid$(text, n, 1)) <= vbKeySpace Then
            n = n - 1   ' account for the extra "space"
            Exit For    ' whitespace is found. Get word.
        End If
    Next
        getFirstWord = Left$(text, n)
End Function

'
' skipOneWord("ABC 123 DEF GHI)" => " 123 DEF GHI"
Public Function skipOneWord(ByVal text As String, _
                            Optional skipWhiteSpace As Boolean = True) As String
    text = LTrim(text)
    'skip any whitespace first.
    'Dim position As Long
    Dim n As Long
    
    'skip the first word.
    For n = 1 To Len(text)  '
        If Asc(Mid(text, n, 1)) <= vbKeySpace Then
            'n = n - 1   ' account for the extra "space"
            Exit For    ' whitespace is found. Get word.
        End If
    Next
    
    text = Mid$(text, n)
    If skipWhiteSpace Then
        text = LTrim(text)      'skip whitespace
    End If
    'look for whitespace indicating the word
        skipOneWord = text
    
End Function

'End Function
'skips all words until it sees the keyWord. By default, it ignores case.'
'getTextAfterWord("FROM","SELECT * FROM abc WHERE i>0")
'=> " abc WHERE i>0",    skips "SELECT * FROM"
Public Function getTextAfterWord(ByVal keyWord As String, ByVal text As String, _
                            Optional ignoreCase As Boolean = True) As String
    Dim compareMode As VbCompareMethod
    compareMode = IIf(ignoreCase, vbTextCompare, vbBinaryCompare)
    Dim position As Long
    position = InStr(1, text, keyWord, compareMode)
    If position <= 0 Then
        Exit Function    ' word not found
    End If
    text = Mid$(text, position)
    getTextAfterWord = skipOneWord(text)
End Function

Public Function Max(data1, data2)
    Max = IIf(data1 > data2, data1, data2)
End Function
Public Function Min(data1, data2)
    Min = IIf(data1 < data2, data1, data2)
End Function

' remove all character in the string except those in the exceptText string.
' removeAllCharExcept("a") ' remove all char except a. Will end up with "a", or "aa...a"
' removeAllCharExcept("0123456789+-.")  ' keep only the number
' TODO: use filter() instead or use regexp.
'Public Function removeAllCharExcept(exceptText) As String
'    Dim i
'
'End Function


Public Function isNotNull(var) As Boolean
    isNotNull = IIf(IsNull(var), False, True)
End Function

Public Function isString(var) As Boolean
    isString = IIf(VarType(var) = vbString, True, False)
End Function

Public Function isEmptyOrNullString(var) As Boolean
    isEmptyOrNullString = False   'default
    If isString(var) Then
        Dim text As String
        text = Trim(var)
        If text = "" Then
            isEmptyOrNullString = True
        End If
    ElseIf IsNull(var) Then
            isEmptyOrNullString = True
        
    End If
End Function

' untested.
Public Function isCollection(var) As Boolean
    isCollection = IIf(TypeOf var Is Collection, True, False)
End Function

Public Function isAlmost(ByVal num1 As Double, ByVal num2 As Double, _
    Optional threshold As Double = 0.01) As Boolean
    'change the threshold for more precision.
    isAlmost = IIf(Abs(num1 - num2) < threshold, True, False)
End Function

' same as join() except it uses paramarray instead of array.
' aka join() PHP,Python, Implode() PHP
' Also, uses SPACE " " as delim. See joinField
Public Function joinField(ParamArray vararg()) As String
    Dim item, returnString As String
    'For Each item In vararg
    '    returnString = returnString & item & " "
    'Next
    'chop returnString   ' remove last space " ".
    ' joinField = returnString
    joinField = Join(vararg, " ")
End Function

'Useful for avoiding lots of checks for NULL value since assigning a null value
'to a variable will result in Error.
Public Function getEmptyTypeBasedOn(exampleVar)
    If IsNumeric(exampleVar) Then
        getEmptyTypeBasedOn = 0
    ElseIf isString(exampleVar) Then
        getEmptyTypeBasedOn = ""
    Else
        getEmptyTypeBasedOn = Null
        eprint "the.getEmptyTypeBasedOn() tried illegal var type"
        'could not determine type.
    End If
End Function

'Get bit value at nth position of a number (works only with 16bit value).
'Returns true (-1) if bit is on, 0 if false. Warning: it returns TRUE(-1),not 1.
'VBTK
Public Function getBit(ByVal value As Integer, ByVal bitPosition As Integer) As Boolean
    Dim temp As Long    'holds a long value to catch overflow
    If bitPosition < 0 Or bitPosition > 15 Then
        eprint "getBit() failed. Position # is invalid. Works only on 16bit. not 32."
        Exit Function
    End If
    temp = value And &HFFFF&
    getBit = (temp And 2 ^ bitPosition)
End Function
        
'sets the bit at nth position and returns the new value. Only works on 16bit Int.
'VBTK
Public Function setBit(ByVal value As Integer, ByVal bitPosition As Integer) As Integer
    Dim temp As Long
    If bitPosition < 0 Or bitPosition > 15 Then
        eprint "setBit() failed. Position # is invalid. Works only on 16bit. not 32."
        Exit Function
    End If
    temp = value And &HFFFF&
    temp = (temp Or 2 ^ bitPosition)
    setBit = (temp And &H7FFF&) - (temp And &H8000&)
End Function

'return HIGH byte (bit 8 to 15) from int/word
'VBTK. Not tested.
Public Function getHighByte(ByVal value As Integer) As Integer
    Dim temp As Long
    temp = value And &HFFFF&    'FFFF
    temp = temp \ (2 ^ 8) ' chop out high byte
    temp = temp * (2 ^ 8) ' right shift 8 bit
    getHighByte = temp
End Function

'return LOW byte (bit 0-7) from int/word
'VBTK. Not tested.
Public Function getLowByte(ByVal value As Integer) As Integer
    Dim temp As Long
    temp = value And &HFFFF& 'FFFF
    temp = temp And &HFF&   '00FF
    getLowByte = temp
End Function

'should be making Integer(16bit word) instead of long but not possible because
'Byte higher than 127 will cause overflow due to it being SIGNED.
'VBTK.
Public Function makeLong(ByVal highByte As Byte, ByVal lowByte As Byte) As Long
    makeLong = (highByte * 256) + lowByte
End Function

Public Function getLowWord(ByVal value As Long) As Long
    getLowWord = value And &HFFFF&
End Function


Public Function isLeapYear(ByVal dateValue) As Boolean
    Dim yearValue
    yearValue = Year(dateValue)
    If (yearValue Mod 400) = 0 Then
        isLeapYear = True
        Exit Function
    End If
    If (yearValue Mod 100) = 0 Then
        Exit Function   'not leap year
    End If
    isLeapYear = (yearValue Mod 4) = 0
End Function

'given date, return # of days in that month.
'based on VBTK, modified to make it work for any year.
Public Function getDaysInMonth(ByVal dateValue)
    Dim monthValue
    monthValue = Month(dateValue)
    Select Case monthValue
        Case 2
            If isLeapYear(dateValue) Then
                getDaysInMonth = 29
            Else
                getDaysInMonth = 28
            End If
        Case 1, 3, 5, 7, 8, 10, 12
            getDaysInMonth = 31
        Case Else
            getDaysInMonth = 30
        End Select
End Function

Public Function getDaysInYear(ByVal dateValue As Date) As Integer
    getDaysInYear = IIf(isLeapYear(dateValue), 366, 365)
End Function

' 1-365 (or 366 in leap year), get day#
Public Function getDayOfYear(ByVal dateValue As Date) As Integer
    getDayOfYear = DatePart("y", dateValue)
End Function

' 0-365, how many days left in the given year
Public Function getDaysLeftInYear(ByVal dateValue As Date) As Integer
    getDaysLeftInYear = getDaysInYear(dateValue) - getDayOfYear(dateValue)
End Function

'Get printable Locale-Independent date
'When representing a date as universal as possible in the string format,
'use the YYYYMMDD format. It can be saved to a file without need for locale conversion.
Public Function getUniversalDateString(ByVal dateValue As Date) As String
    getUniversalDateString = Format(dateValue, "yyyymmdd")
End Function
Public Function universalDateString2Date(ByVal dateString As String) As Date
    universalDateString2Date = DateSerial(Left$(dateString, 4), _
        Mid$(dateString, 5, 2), Mid$(dateString, 7, 2))
End Function


' TODO: Convert lists, paramarrays, collections, single var into array.
' More flexible than Array() function.
'Public Function makeArray(ParamArray args())
    'see makeList()
'End Function


'TODO: convert Collection, arrays, paramarrays, single variable into Collection
'Similar to VB's Array() function and my makeArray().
'can be very slow since it has to copy every items in the array/collection/etc.
'Dictionary should use dict.getKeys() and .getValues() instead.
'Public Function makeList(ParamArray args())
'    Dim item, returnString As String
'    Dim fieldList
'    'in Access, field name should be enclosed in [] in certain cases.
'    If isString(vararg(0)) Then     'STR1,STR2,...
'        fieldList = vararg
'    ElseIf IsArray(vararg(0)) Then  'ARRAY
'        fieldList = vararg(0)
'    Else                            'TheList, Collection
'        Set fieldList = vararg(0)
'    End If
'        For Each item In fieldList
'            If item = "*" Then
'                sqlField = "*"
'                Exit Function
'            End If
'            returnString = returnString + "[" & item & "],"
'        Next
' Remove unwanted/illegal char and issue warning if it does.
'End Function

Public Function stripIllegalChar(ByVal text, ByVal charsToRemove)
    Dim c, i                'index, char
    Dim oldText, newText    'str
    For i = 0 To Len(charsToRemove) - 1
        c = getCharAt(charsToRemove, i)
        oldText = text
        newText = Replace(text, c, "")
        If Len(oldText) <> Len(newText) Then
            the.dprint "StripIllegalChar() " & _
                        "Warning: Removed chars. OldText=" & text & _
                        "newText=" & stripIllegalChar
        End If
        text = newText
    Next
    stripIllegalChar = text
End Function

Public Sub dprint(ParamArray vararg())
    On Error GoTo IN_DEVELOPMENT_ENVIRONMENT
    Debug.Print 1 / 0
    Exit Sub
IN_DEVELOPMENT_ENVIRONMENT:
    Dim item
    Dim i As Long
    'For Each item In vararg
    Debug.Print vararg(0);   '
    If UBound(vararg) - LBound(vararg) + 1 = 1 Then
        Debug.Print     'println since it doesn't have trailing vars.
    Else
        Debug.Print "=";
    End If
    For i = LBound(vararg) + 1 To UBound(vararg)
        If IsObject(vararg(i)) Then
            Set item = vararg(i)
        Else
            item = vararg(i)
        End If
        Debug.Print TheDebug.getString(item)
    Next
End Sub

Public Sub warn(ByVal text As String)
    dprint "WARNING: " & text
End Sub
'=============================================================================
'   UI-Related
'=============================================================================
Public Sub moveLine(lineControl As Line, ByVal x1, ByVal y1, ByVal x2, ByVal y2)
    lineControl.x1 = x1
    lineControl.x2 = x2
    lineControl.y1 = y1
    lineControl.y2 = y2
End Sub

'=============================================================================
'   DEPRECATED CODES
'=============================================================================
' Deprecated function
' VB-Specific.  Since there is no RemoveAll method, this explicitly removes the
' collection. However, this isn't really needed since VB has its own garbage-collection.
' Just pointing the collection to Nothing/Null will satisfy it.
Public Sub clearCollection(collectionName As Collection)
    Dim i
    For i = collectionName.Count To 1 Step -1
        collectionName.remove i
        
    Next
End Sub
'=============================================================================
'=============================================================================
#If DEBUG_MODE Then
    Public Sub test(text)
        ' "getString", "all",
        If text = "getString" Then  '-- test getString()
            Dim a, b, c, d
            a = 34.5
            b = 44
            c = True
            d = "hello"
            Dim e(1 To 3)
            e(1) = 50
            e(2) = 51
            e(3) = 52
            
            Dim f As New Collection
            f.add "c1"
            f.add "c2"
            
            the.dprint "a:", the.getString(a)
            the.dprint "b:", the.getString(b)
            the.dprint "c:", getString(c)
            the.dprint "d:", getString(d)
            the.dprint "e(array):", getString(e)
            the.dprint "f(collection):", getString(f)

        Else
            the.dprint "No test defined"
        End If
    End Sub
    

    
#Else
    Public Sub test(text)
    End Sub
#End If





'=============================================================================
'   OLD CODE
'=============================================================================
#If OLD_CODE Then
'#Const DEBUG_MODE = 1  'deprecated. Use TheDebug. No TheDebug yet.
'' deprecated. Use TheDebug.dprint() instead.
'Public Sub dprint(title, Optional Data)
'
'
'    TheDebug.dprint title, Data
''    Debug.Print getString(title);
' '   If IsMissing(Data) Then
' '       Debug.Print
'  '      Exit Sub
''    End If
''    Debug.Print ":=" & getString(Data) & " <type=" & TypeName(Data) & "size=" & getSize(Data) & ">"
'
'End Sub
'
'
'' Deprecated. Use getString() instead.
'Public Function getStringFromNumbers(Data) As String
'    Dim i
'    For Each i In Data
'        getStringFromNumbers = getStringFromNumbers & i & ","
'    Next
'    If IsEmpty(Data) Then
'        getStringFromNumbers = ""
'        Exit Function
'    End If
'    ' Remove last ','
'    chop getStringFromNumbers
'
'End Function
'
'' VB lacks ++, -- post/pre-inc/dec. although VB.net has it.
'' Dangerous. If you call sub with ( ), then it will default the var as ByVal, not byRef!
'Public Sub inc(ByRef var As Long)
'    var = var + 1
'End Sub
'
'Public Sub dec(ByRef var)
'    var = var - 1
'End Sub
'
'' Same as x+=y. To subtract, call add(var, -value).
'Public Sub incBy(ByRef var, value)
'
'    var = var + value
'End Sub
'
'Public Sub decBy(ByRef var, value)
'    var = var - value
'End Sub
'' Same as x-=y
''Public Sub subtractSelf(ByRef var, value)
''    var = var - value
''End Sub

#End If

