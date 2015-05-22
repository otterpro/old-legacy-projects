VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TheIni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'TODO: CONVERT TO TheWin32IniTable class (TheWinIni for short)
'
'   dim ini as New TheIni  'by default, opens app.ExeName().ini (minus ".exe")
'   name= ini.getValue("kid","name")
'
'
'   to open different ini
'   ini.open("d:\path\myini.ini")
Option Explicit
Private Const MAX_STRING_SIZE = 255

Private misc As New TheMisc

Private Declare Function WritePrivateProfileString Lib "kernel32" Alias _
    "WritePrivateProfileStringA" (ByVal section As String, _
    ByVal key As Any, _
    ByVal value, _
    ByVal filename As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias _
    "GetPrivateProfileStringA" (ByVal applicationName As String, _
    ByVal lpKeyName As String, ByVal lpDefault As String, _
    ByVal lpReturnedString As String, ByVal nSize As Long, _
    ByVal lpFileName As String) As Long

Private IniPath As String

Public Function openFile(filename As String) As Boolean
    'search for INI file. If it exists, returns TRUE. Else returns false.
    IniPath = filename
    If Dir(IniPath) = "" Then
        openFile = False 'file not found
    Else
        openFile = True
    End If
End Function

' Open App's INI file.  FIlename must match app's name.
Public Sub openConfig()
    Dim returnValue As Boolean
    Dim newPath As String
    Dim filename As String
    Dim fileSystem As New TheFileSystem
    filename = App.path & "\" & App.EXEName
    filename = fileSystem.replaceFileExtension(IniPath, ".ini")
    returnValue = openFile(filename)

End Sub

Public Sub setValue(ByVal section As String, ByVal key As String, Data)
    Dim result
    result = WritePrivateProfileString(section, key, Data, IniPath)
End Sub

Public Function getValue(ByVal section As String, ByVal key As String, _
    Optional ByVal defaultValue = "") As Variant
    Dim defaultValueIfAny As String
    'If IsMissing(defaultValue) Then
     '   defaultValueIfAny = ""
    Dim returnString As String
    returnString = String(MAX_STRING_SIZE, Chr(0))

    Dim result
    result = GetPrivateProfileString(section, key, defaultValue, returnString, _
        MAX_STRING_SIZE, IniPath)
    getValue = returnString
    
End Function

