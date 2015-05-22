VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TheGuiList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Dim frameList As Collection
Dim formList As Collection
Public Sub dprint()
    Dim i As Object
    Debug.Print "Form List"
    For Each i In formList
        Debug.Print i.Name
    Next
    Debug.Print "Frame List:"
    For Each i In frameList
        Debug.Print i.Name
    Next
    
End Sub



Public Sub hideAllFramesExcept(frameRef As Frame)
    ' ParamArray frameLists())
    hideAllFrames
    frameRef.Visible = True
    'Dim i
    'For Each i In frameLists
    '    i.Visible = True
    'Next
    
End Sub

Public Sub hideAllFrames()
    Dim i As Object
    For Each i In frameList
        i.Visible = False
    Next
End Sub

Public Sub Class_Terminate()
End Sub

Public Sub Class_Initialize()
    Set frameList = New Collection
    Set formList = New Collection
    
    Dim f, c 'form and control
    For Each f In Forms
        formList.Add f
        'Debug.Print "Form:" & f.Name
        For Each c In f
            
        Select Case TypeName(c)
            Case "Frame"
                frameList.Add c
                'Debug.Print "Frame:" & c.Name
            Case Else
        End Select
        Next
    Next
    
End Sub

