VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form viewSalespersonForm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "View Salesperson"
   ClientHeight    =   5790
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6945
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   6945
   ShowInTaskbar   =   0   'False
   Begin MSDataGridLib.DataGrid salespersonGrid 
      Height          =   4335
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   7646
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid salespersonInventoryGrid 
      Height          =   4335
      Left            =   2280
      TabIndex        =   1
      Top             =   120
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   7646
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "Finish"
      Height          =   975
      Left            =   120
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4560
      UseMaskColor    =   -1  'True
      Width           =   6495
   End
End
Attribute VB_Name = "viewSalespersonForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim mySalespersonTable As New TheAdoTable2
Dim myInventoryTable As New TheAdoTable2
Dim myInventoryTableIsOpen As Boolean

Private Sub salespersonInventoryGrid_Click()
    'nothing happens if user tries to edit this field.
    Exit Sub
End Sub

Private Sub Form_Load()
    'mySalespersonTable.openTable "select [name],[out] from salesperson", inventory..getConnection()
    'mySalespersonTable.moveFirst
    'Set salespersonGrid.DataSource = mySalespersonTable.getRecordSet()
    'myInventoryTableIsOpen = False

End Sub

Private Sub okButton_Click()
    Unload Me
End Sub

Private Sub salespersonGrid_Click()
    If (myInventoryTableIsOpen) Then
        myInventoryTable.closeFile
    End If
    If mySalespersonTable.isEof() Then
        ' Nothing to do since there is no salesperson in the db
        Exit Sub
    End If
    myInventoryTableIsOpen = True
    'myInventoryTable.openTable "select * from " & _
        mySalespersonTable.getValue("tableName"), _
        inventory..getConnection()
    Set salespersonInventoryGrid.DataSource = myInventoryTable.getRecordSet()
    salespersonInventoryGrid.refresh
End Sub


