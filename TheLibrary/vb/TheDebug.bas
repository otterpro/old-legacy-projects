Attribute VB_Name = "TheDebug"
Option Explicit

' dprint. Conditionally compiled. If compiled, the dprint is empty. Else, dprint works.
' isDebug
'   Returns TRUE if the VB is running in debug mode (ie in development environment)
'       False if it is compiled.
'   See Bug Proofing Visual Basic, Rod Stephens, p76 for more info.

'Private myDebugMode As Boolean


Public Function isDebug() As Boolean
    'isDebug = myDebugMode
    On Error GoTo IN_DEVELOPMENT_ENVIRONMENT
    'if the program is running in the development envirnoment, it will try to execute
    ' the following line resulting in an error, indicating that we are in debug mode
    Debug.Print 1 / 0
    isDebug = False ' we know that it is compiled bc 1/0 didn't compile
    Exit Function
IN_DEVELOPMENT_ENVIRONMENT:
    isDebug = True
End Function

' used exclusively by the.dprint()
Public Function getString(ByRef item) As String
    getString = the.getString(item) & " <type=" & TypeName(item) & " size=" & getSize(item) & ">"
End Function


'=============================================================================
' OLD
'=============================================================================
#If OLD_CODE Then
Public Sub dprint(ParamArray vararg())
    On Error GoTo IN_DEVELOPMENT_ENVIRONMENT
    Debug.Print 1 / 0
    Exit Sub
IN_DEVELOPMENT_ENVIRONMENT:
    Debug.Print getString(title);
    If IsMissing(Data) Then
        Debug.Print
        Exit Sub
    End If
    Debug.Print ":=" & getString(Data) & " <type=" & TypeName(Data) & " size=" & getSize(Data) & ">"
End Sub

#End If







