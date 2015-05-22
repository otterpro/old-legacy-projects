VERSION 5.00
Begin VB.Form changeValueDialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change Value"
   ClientHeight    =   2445
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   4215
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton subtractButton 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2040
      TabIndex        =   4
      Top             =   1080
      Width           =   495
   End
   Begin VB.CommandButton addButton 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2040
      TabIndex        =   3
      Top             =   360
      Width           =   495
   End
   Begin VB.TextBox valueText 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Text            =   "valueText"
      Top             =   600
      Width           =   1815
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   1095
      Left            =   2760
      MaskColor       =   &H00000000&
      Picture         =   "changeValueDialog.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1200
      UseMaskColor    =   -1  'True
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   975
      Left            =   2760
      MaskColor       =   &H00000000&
      Picture         =   "changeValueDialog.frx":1B42
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   1215
   End
End
Attribute VB_Name = "changeValueDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub addButton_Click()
    Me.valueText.text = Val(Me.valueText.text) + 1
End Sub

Private Sub cancelButton_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.valueText.text = valueToChange
End Sub


Private Sub okButton_Click()
    valueToChange = CStr(Me.valueText.text)
    'changeValueDialogValue = valueToChange
    Unload Me
End Sub

Private Sub subtractButton_Click()
    Me.valueText.text = Max(Val(Me.valueText.text) - 1, 0)
End Sub

Private Sub valueText_Change()
    Me.valueText.text = CStr(Me.valueText.text)
End Sub

'Private Sub multiplierText_Validate(Cancel As Boolean)
'    If Not IsNumeric(Me.multiplierText.text) Then
'        removeAllCharExcept ("0123456789")
'    End If
'    multiplier = Val(Me.multiplierText.text)
'End Sub

Private Sub valueText_KeyPress(KeyAscii As Integer)
    If KeyAscii < 48 Or KeyAscii > 57 Then
        ' illegal entry.
        KeyAscii = 0
    End If
End Sub

