VERSION 5.00
Begin VB.Form configWindow 
   Caption         =   "Configuration"
   ClientHeight    =   5265
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8655
   LinkTopic       =   "Form1"
   ScaleHeight     =   5265
   ScaleWidth      =   8655
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cancelButton 
      Caption         =   "Cancel "
      Height          =   735
      Left            =   4440
      TabIndex        =   2
      Top             =   4440
      Width           =   4095
   End
   Begin VB.CommandButton okButton 
      Caption         =   "Save Changes"
      Default         =   -1  'True
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   4440
      Width           =   4215
   End
   Begin VB.Frame Frame1 
      Caption         =   "WARNING: Advanced User Only. Making Changes May Render This Software Inoperable."
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8415
      Begin VB.TextBox dbPathText 
         Height          =   375
         Left            =   1440
         TabIndex        =   3
         Text            =   "dbPath"
         Top             =   480
         Width           =   6495
      End
      Begin VB.Label Label1 
         Caption         =   "Database Path:"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   1215
      End
   End
End
Attribute VB_Name = "configWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub cancelButton_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    dbPathText.text = App.path() & "\" & DB_FILENAME
End Sub

Private Sub OKButton_Click()
    ' TODO: set config
    ' dbPath = dbPathText.text
End Sub
