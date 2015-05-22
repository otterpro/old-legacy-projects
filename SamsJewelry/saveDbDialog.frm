VERSION 5.00
Begin VB.Form saveDbDialog 
   BorderStyle     =   0  'None
   Caption         =   "Please Wait"
   ClientHeight    =   3195
   ClientLeft      =   2715
   ClientTop       =   3315
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.Label Label1 
      Caption         =   "Please wait until Excel is finished loading."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   5055
   End
End
Attribute VB_Name = "saveDbDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

