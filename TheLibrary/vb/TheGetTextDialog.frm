VERSION 5.00
Begin VB.Form TheGetTextDialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Enter Name"
   ClientHeight    =   2940
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox text 
      Height          =   1335
      Left            =   360
      TabIndex        =   2
      Text            =   "Text"
      Top             =   360
      Width           =   4095
   End
   Begin VB.CommandButton cancelButton 
      Caption         =   "Cancel"
      Height          =   1215
      Left            =   4680
      MaskColor       =   &H00000000&
      Picture         =   "TheGetTextDialog.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1440
      UseMaskColor    =   -1  'True
      Width           =   1215
   End
   Begin VB.CommandButton okButton 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   1215
      Left            =   4680
      MaskColor       =   &H00000000&
      Picture         =   "TheGetTextDialog.frx":1B42
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   1215
   End
End
Attribute VB_Name = "TheGetTextDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim returnValue

Private Sub cancelButton_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    Me.text.SetFocus
End Sub

Private Sub Form_Load()
    returnValue = Null
    Me.text = CStr(TheDialog.getDefaultValue())
    Me.Caption = TheDialog.getTitle()
End Sub



Private Sub Form_Unload(Cancel As Integer)
    TheDialog.setReturnValue returnValue
End Sub

Private Sub okButton_Click()
    returnValue = Me.text
    Unload Me
End Sub


