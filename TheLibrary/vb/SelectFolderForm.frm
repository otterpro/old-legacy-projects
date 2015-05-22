VERSION 5.00
Begin VB.Form SelectFolderForm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Folder"
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6090
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   6090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.FileListBox fileList 
      Height          =   4185
      Left            =   2400
      System          =   -1  'True
      TabIndex        =   4
      Top             =   840
      Width           =   3495
   End
   Begin VB.CommandButton cancelButton 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   3120
      TabIndex        =   3
      Top             =   5160
      Width           =   1695
   End
   Begin VB.CommandButton okButton 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   1200
      TabIndex        =   2
      Top             =   5160
      Width           =   1815
   End
   Begin VB.DriveListBox driveList 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   3615
   End
   Begin VB.DirListBox dirList 
      Height          =   4140
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   2175
   End
End
Attribute VB_Name = "SelectFolderForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'======================================================================
' SelectFolderForm.getPath(optional filePattern="*.*")
' loads and shows the Select Folder Window
' returns the path selected by the user.
' Returns "" if user presses Cancel
'======================================================================


Private path As String ' full path of folder
' Private pressedCancel As Boolean 'if user pressed the Cancel button'
'----------------------------------------------------------------------
Public Function getPath(Optional ByVal initialPath, Optional ByVal filePattern)

    SelectFolderForm.Show vbModal
    If Not IsMissing(initialPath) Then
        'ChDir (initialPath)
        driveList.Drive = Left(initialPath, 2)
        dirList.path = initialPath
        fileList.path = initialPath
    End If
    getPath = path
End Function

Private Sub CancelButton_Click()
    path = ""
    Unload Me
End Sub


Private Sub dirList_Change()
    fileList.path = dirList.path()
End Sub

Private Sub driveList_Change()
    dirList.path = Left(driveList.Drive, 2)
End Sub

Private Sub OKButton_Click()
    path = dirList.path
    Unload Me
End Sub
