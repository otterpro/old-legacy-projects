Attribute VB_Name = "TheCGITool"
'Tool for HTML and XML and CGI stuff. Too broad
Option Explicit

'Properly escapes the given string
'by translating "&", "<" and ">" char. It is supposed to be similar to
'Python's CGI module's escape(). VERY SLOW. Replace with RegExp Later.
'Very limited functionality. Only handles simple string and it assumes
'that it is not in a literal block. Also it expects that it is NOT escaped yet.
Public Function escape(ByVal text As String) As String
    'Dim returnValue As String
    'escape = text
    'HTML can ends these char with ";" semi-colon. Omitting ";" leads to confusion.
    If InStr(1, text, "&amp", vbTextCompare) Then
        the.eprint "CGI::escape() does not handle string that contains escaped chars: " & text
        Exit Function
    End If
    If InStr(text, "&") Then
        text = Replace(text, "&", "&amp;", Compare:=vbTextCompare)
        'kinda tricky because you don't want to replace "&amp"
    End If
    If InStr(text, the.QUOTE_CHAR) Then
        text = Replace(text, the.QUOTE_CHAR, "&quot;")
    End If
    If InStr(text, "<") Then
        text = Replace(text, "<", "&lt;")
    End If
    If InStr(text, ">") Then
        text = Replace(text, ">", "&gt;")
    End If
        
    escape = text
End Function
