VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TheOle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' see MS VB Example from MSDN "OleCnt"
' TODO: See TheExcelTable's openAsEmbedded(). Not working.
'
Option Explicit
Dim myOleControl As OLE
Dim myFilename As String
Dim myFileIsSavedAutomaticallyWhenUnloaded As Boolean
    ' aka autoSave. by default, FALSE. Currently doesn't work unless this obj is instantiated on the form and not globally on the module.
    ' since Control will go out of scope if it is on the module.
    
Public Function openFile(ByVal filename As String, ByRef oleControl As OLE, Optional autoSave As Boolean = False) As Boolean
    Dim fileNumber As Long
    myFileIsSavedAutomaticallyWhenUnloaded = autoSave
    openFile = False
    Set myOleControl = oleControl
    myFilename = filename
    fileNumber = FreeFile()
    Open filename For Binary As fileNumber
    If Err Then ' Error: File Not found or something else
        the.eprint "TheOle::openFile(), Error loading file " & filename
        Exit Function
    End If
        ' Display the hourglass mouse pointer.
    Screen.MousePointer = 11
    oleControl.ReadFromFile fileNumber
    If (Err) Then
        If Err = 30015 Then
            eprint "TheOle::openFile(),Not a valid object."
        Else
            eprint "TheOle::openFile()," & Error$
        End If
    End If
    
    ' Set the form properties now that the OLE container control contains an object.
    'UpdateCaption
    ' Restore the mouse pointer.
    Screen.MousePointer = 0
    openFile = True
    Close fileNumber
End Function

Public Function saveFile(Optional filename As String) As Boolean
    Dim fileNumber As Long
    saveFile = False
    fileNumber = FreeFile()
    Dim saveFilename As String      ' filename to save. By default, saves using opened filename
    If IsMissing(filename) Then
        saveFilename = myFilename
    Else
        saveFilename = filename
    End If
    
    Open myFilename For Binary As fileNumber
    If Err Then
        the.eprint (Error)
            Exit Function
    End If
    myOleControl.SaveToFile fileNumber
    If Err Then
        the.eprint (Error)
        Close fileNumber
        Exit Function
    End If
    saveFile = True
    Close fileNumber
End Function

'enable the control in edit mode. ie activate it.
Public Sub activate()
    'myOleControl.DoVerb -1
    myOleControl.DoVerb vbOLEUIActivate
        'not sure if this constant is defined.
End Sub

Private Sub Class_Terminate()
    If myFileIsSavedAutomaticallyWhenUnloaded = True Then
        saveFile
    End If
End Sub

Private Function oleIsValid() As Boolean
    If myOleControl.OLEType = vbOLENone Then
        oleIsValid = False
    Else
        oleIsValid = True
    End If
End Function

