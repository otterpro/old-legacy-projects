VERSION 5.00
Begin VB.Form dataEntryForm 
   Caption         =   "Detail Data Entry"
   ClientHeight    =   4560
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4560
   ScaleWidth      =   6000
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cancelButton 
      Cancel          =   -1  'True
      Caption         =   "Ignore this change and return to the previous screen"
      Height          =   1215
      Left            =   3600
      TabIndex        =   4
      Top             =   3240
      Width           =   2295
   End
   Begin VB.CommandButton okButton 
      Caption         =   "Save this result and return to the previous screen"
      Default         =   -1  'True
      Height          =   1215
      Left            =   120
      MaskColor       =   &H00000000&
      Picture         =   "dataEntryForm.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3240
      UseMaskColor    =   -1  'True
      Width           =   3375
   End
   Begin VB.CommandButton minusButton 
      Caption         =   "-"
      Height          =   375
      Left            =   2760
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1080
      Width           =   615
   End
   Begin VB.CommandButton plusButton 
      Caption         =   "+"
      Height          =   375
      Left            =   2760
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   600
      Width           =   615
   End
   Begin VB.TextBox quantity 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1320
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox note 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   1320
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1560
      Width           =   3255
   End
   Begin VB.Label Label2 
      Caption         =   "Note"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Quantity"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   855
   End
End
Attribute VB_Name = "dataEntryForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cancelButton_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    quantity.text = CStr(dataEntryParam.quantity)
    note.text = dataEntryParam.note
End Sub

Private Sub minusButton_Click()
    quantity.text = Max(val(quantity.text) - 1, 0)
End Sub

Private Sub okButton_Click()
    dataEntryParam.quantity = val(quantity.text)
    dataEntryParam.note = note.text
    Unload Me
End Sub

Private Sub plusButton_Click()
    quantity.text = CStr(val(quantity.text) + 1)
End Sub

Private Sub quantity_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyDelete Or KeyAscii = vbKeyBack Then
        Exit Sub
    ElseIf KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
        KeyAscii = 0
    End If
End Sub
