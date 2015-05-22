Attribute VB_Name = "TheRegExpTool"
' TODO: Move these to The or TheStringTool since regex is considered an
' intrinsic object (although very hidden)

' getNumber(str)
'    Extracts # from string and returns the first # as int.
'    getNumber ("abc123def") ' 123
'    getNumber ("123qrs456") ' 123

Option Explicit

Private myRegExp As New RegExp

Public Function getNumber(ByVal text As String)
    setPattern "[^\d]", isGlobal:=True
    text = myRegExp.Replace(text, " ")
    text = getFirstWord(text)
    text = Trim(text)
    If InStr(text, ".") Then
        getNumber = CSng(text)
    ElseIf text = "" Then
        getNumber = CLng(0)
        'warning: no # found!
    Else
        getNumber = CLng(text)
    End If
End Function

' sets property for RegExp Object. Used to give consistent way to set the property.
Private Function setPattern(ByVal patternText As String, _
                            Optional isGlobal = True, Optional ignoreCase = False)
    myRegExp.Global = isGlobal
    myRegExp.Pattern = patternText
End Function


