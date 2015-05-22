VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TheSimpleRegistry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' TODO: Convert to TheWin32RegistryTable (TheWinReg for short)
'
Option Explicit
Private myFilename As String

' Open App's INI file.  FIlename must match app's name.
Public Sub openConfig()
    myFilename = App.EXEName
End Sub

Public Sub setValue(ByVal section As String, ByVal key As String, Data)
    Debug.Assert (myFilename)
    SaveSetting myFilename, section, key, Data
End Sub

Public Function getValue(ByVal section As String, ByVal key As String, _
    Optional ByVal defaultValue = "") As Variant
    Debug.Assert (myFilename)
    getValue = GetSetting(myFilename, section, key, defaultValue)
End Function




