VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form chooseSalespersonDialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Choose Salesperson"
   ClientHeight    =   7470
   ClientLeft      =   1305
   ClientTop       =   1395
   ClientWidth     =   9645
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7470
   ScaleWidth      =   9645
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton okButton 
      Caption         =   "Choose "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   3960
      TabIndex        =   7
      Top             =   120
      Width           =   3975
   End
   Begin MSDataGridLib.DataGrid inventoryGrid 
      Height          =   4335
      Left            =   3960
      TabIndex        =   5
      Top             =   3000
      Width           =   5535
      _ExtentX        =   9763
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
   Begin VB.Frame advancedOptionFrame 
      Caption         =   "Advanced Options"
      Height          =   1815
      Left            =   120
      TabIndex        =   2
      Top             =   5520
      Width           =   3735
      Begin VB.CommandButton changePictureButton 
         Caption         =   "Change Picture of this Salesperson"
         Enabled         =   0   'False
         Height          =   1455
         Left            =   2400
         TabIndex        =   6
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton deleteSalespersonButton 
         Caption         =   "Delete This Salesperson"
         Height          =   1455
         Left            =   1200
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton addNewSalespersonButton 
         Caption         =   "Add New Salesperson"
         Height          =   1455
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
   End
   Begin MSDataGridLib.DataGrid salespersonGrid 
      Height          =   5295
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   9340
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   29
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
         Size            =   13.5
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
   Begin VB.CommandButton cancelButton 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   2775
      Left            =   8040
      MaskColor       =   &H00000000&
      Picture         =   "chooseSalespersonDialog.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   1455
   End
End
Attribute VB_Name = "chooseSalespersonDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=============================================================================
'
'=============================================================================

Option Explicit

Private Const DEFAULT_NAME_WIDTH_PERCENTAGE = 0.7
Private Const DEFAULT_INDEX_WIDTH = 100 'twips
Private Sub addNewSalespersonButton_Click()
    Dim name
    name = TheDialog.getText("Enter Salesperson's name")
    name = Trim(UCase(name))
    If name <> "" Then
        inventory.addNewSalesperson name
        salespersonGrid.refresh
        refreshInventoryGrid
    End If
End Sub

Private Sub cancelButton_Click()
    ChooseSalesperson = False
    Unload Me
End Sub

Private Sub deleteSalespersonButton_Click()
    Dim userInput
    userInput = MsgBox("Are you sure you want to remove this salesperson? " & _
            "Click 'YES' to remove this person.", vbQuestion + vbYesNo)
    If userInput <> vbYes Then
        Exit Sub
    End If
    If Not inventory.deleteSalesperson() Then
        MsgBox "Cannot remove this salesperson because he/she " & _
                " still has items in their inventory. Please do an " & _
                " 'Sale-In' for this person before deleting this person."
    Else
        
        refreshInventoryGrid
    End If
End Sub

Private Sub Form_Load()
    
    Set inventoryGrid.DataSource = Nothing
    Set salespersonGrid.DataSource = _
                            inventory.salespersonTable.getRecordSet()
    refreshInventoryGrid
    salespersonGrid.Columns(0).Width = DEFAULT_NAME_WIDTH_PERCENTAGE _
                    * salespersonGrid.Width
    inventoryGrid.Columns(0).Width = DEFAULT_INDEX_WIDTH
    showSalespersonName
End Sub



Private Sub Form_Unload(Cancel As Integer)
    Set salespersonGrid.DataSource = Nothing
    Set inventoryGrid.DataSource = Nothing
End Sub

Private Sub okButton_Click()
    
    If inventory.salespersonTable.getSize() < 1 Then
        MsgBox "You must add a salesperson before continuing."
        Exit Sub
    End If
    ChooseSalesperson = True
    Unload Me
End Sub



Private Sub salespersonGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Not IsNull(LastRow) Then
'        Call changeSalesperson
        refreshInventoryGrid
        showSalespersonName
    End If
End Sub

Private Sub refreshInventoryGrid()
    
    Set inventoryGrid.DataSource = Nothing
    Set inventoryGrid.DataSource = inventory.salespersonInventoryTable.getRecordSet()
    
End Sub

Private Sub showSalespersonName()
    okButton.Caption = "Choose " & inventory.salespersonName
End Sub
