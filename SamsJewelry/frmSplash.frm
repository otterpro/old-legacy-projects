VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4785
   ClientLeft      =   30
   ClientTop       =   30
   ClientWidth     =   7485
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   7485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame fraMainFrame 
      Height          =   4590
      Left            =   45
      TabIndex        =   0
      Top             =   -15
      Width           =   7380
      Begin VB.PictureBox Picture1 
         Height          =   1455
         Left            =   480
         Picture         =   "frmSplash.frx":0000
         ScaleHeight     =   1395
         ScaleWidth      =   915
         TabIndex        =   8
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lblLicenseTo 
         Alignment       =   1  'Right Justify
         Caption         =   "LicenseTo"
         Height          =   255
         Left            =   270
         TabIndex        =   1
         Tag             =   "1017"
         Top             =   300
         Width           =   6855
      End
      Begin VB.Label lblProductName 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "A-Team and Coconut Studio Presents Inventory Control"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   1800
         TabIndex        =   7
         Tag             =   "1016"
         Top             =   1200
         Width           =   5355
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         Caption         =   "Sam's International"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2880
         TabIndex        =   6
         Tag             =   "1015"
         Top             =   720
         Width           =   3330
      End
      Begin VB.Label lblPlatform 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Platform"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   384
         Left            =   5724
         TabIndex        =   5
         Tag             =   "1014"
         Top             =   2400
         Width           =   1284
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "1.0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6615
         TabIndex        =   4
         Tag             =   "1013"
         Top             =   2760
         Width           =   390
      End
      Begin VB.Label lblCompany 
         Caption         =   "Dan Kim, A-Team and Coconut Studio"
         Height          =   615
         Left            =   4710
         TabIndex        =   3
         Tag             =   "1011"
         Top             =   3330
         Width           =   2415
      End
      Begin VB.Label lblCopyright 
         Caption         =   "Copyright 2003"
         Height          =   255
         Left            =   4710
         TabIndex        =   2
         Tag             =   "1010"
         Top             =   3120
         Width           =   2415
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    'LoadResStrings Me
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    'lblProductName.Caption = App.title
End Sub



