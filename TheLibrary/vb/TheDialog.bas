Attribute VB_Name = "TheDialog"
'=============================================================================
' TheDialog:
'   Simple dialogs can be called from here and its return value can be retrieved from
'   those dialog forms.
'
'   get___(byref returnVariable, optional title, optional defaultValue)
'       returnVariable:= sets this variable to the value that user selects.
'       title:= Caption of dialog box
'       defaultValue:= default value to select.
'       Returns true if successful, false if user presses cancel or error occurs.
'   getInteger()
'   getText()
'   getFileName()
'   getFolder()
'
'   Dialog forms must retrieve the TheDialog.myTitle, TheDialog.myDefaultValue, and
'       return using TheDialog.myReturnValueStack.push retVal
'=============================================================================
Option Explicit
'Table of Return Values from forms. This stores all the returned values from any TheForms that
' needs to return any values back. Since forms lose its data once it is unloaded, this mechanism was
' implemented to store the return value temporarily until it was needed.

'Public myReturnValueStack As New TheStack
'   Removed the stack. Instead, using a single but simple var to hold the return value.
Dim myReturnValue
Dim myDefaultValue As String, myTitle As String
' Used by dialog forms to simplify setting the initial variables.

Private Sub setDialogInfo(ByVal title As String, defaultValue)
        myDefaultValue = defaultValue
        myTitle = title
End Sub

' alias for push. Used by the dialog forms, makes it simpler than pushing().
Public Sub setReturnValue(returnValue)
    'myReturnValueStack.push (returnValue)
    myReturnValue = returnValue
End Sub

Public Function getInteger(Optional title = "Enter Number", Optional defaultValue = 0)
    setDialogInfo title, defaultValue
    TheGetIntegerDialog.Show vbModal
    getInteger = getReturnValue()
End Function


Public Function getText(Optional title = "Enter Text", Optional defaultValue = "")
    setDialogInfo title, defaultValue
    TheGetTextDialog.Show vbModal
    getText = getReturnValue()
End Function

Private Function getReturnValue()
    getReturnValue = myReturnValue
End Function

Public Function getDefaultValue()
    getDefaultValue = myDefaultValue
End Function

Public Function getTitle()
    getTitle = myTitle
End Function

Public Function getFont()
    eprint "Not implemented yet"
    Dim dialog As Object
    'Late-binding
    'Set dialog = comdlg
    
    Set dialog = CreateObject("CommonDialog.CommonDialog")  'could be ado2.1 or ado2.6. Whatever the user has...
    

End Function

'=============================================================================
#If OLD_CODE Then
Public Enum DialogType
    getInteger
    getText
    getFile
End Enum
Public Function getDialog(dialogName As DialogType, Optional title = "", Optional defaultValue = 0) As Variant
        myDefaultValue = defaultValue
        myTitle = title
        Select Case dialogName = getInteger
            TheGetIntegerDialog.Show vbModal, Me
        Case dialogName = getText
            TheGetTextDialog.Show vbModal, Me
        Case Else
            eprint "getDialog: invalid dialog name"
        End Select
        getDialog = myReturnValueStack.pop()
End Function
'Public Function getReturnValue(ByVal name As String)
    'getReturnValue = myReturnValueDict.getValue(name)
'End Function
Public Function getText(Optional title = "", Optional defaultText = "") As String
    myDefaultValue = defaultText
    TheGetTextDialog.Show vbModal, Me
    getText = myReturnValueStack.pop()
End Function

Public Function getInteger(Optional title = "", Optional defaultValue = 0) As Long
        myDefaultValue = defaultValue
        TheGetTextDialog.Show vbModal, Me
        getInteger = myReturnValueStack.pop()
End Function
Public Function getDefaultValue()
    getDefaultValue = myDefaultValue
End Function

Public Function getTitle()
    getTitle = myTitle
End Function

Private Function getReturnValue(ByRef returnVar)
    Dim returnValue
    'returnValue = myReturnValueStack.pop()
    returnValue = myReturnValue
    If IsNull(returnValue) Then
        'returnVar = 77 'returnValue
        getReturnValue = False
    Else
            returnVar = returnValue
            getReturnValue = True
    End If
End Function

#End If


