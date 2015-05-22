VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sam's International"
   ClientHeight    =   5970
   ClientLeft      =   465
   ClientTop       =   1155
   ClientWidth     =   8250
   LinkTopic       =   "Form1"
   ScaleHeight     =   502.525
   ScaleMode       =   0  'User
   ScaleWidth      =   700
   Begin VB.Frame introFrame 
      Height          =   255
      Left            =   0
      TabIndex        =   24
      Top             =   360
      Width           =   8175
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   4275
         Left            =   240
         Picture         =   "frmMain.frx":0000
         ScaleHeight     =   4245
         ScaleWidth      =   7710
         TabIndex        =   25
         Top             =   240
         Width           =   7740
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Sam's International"
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
         Left            =   240
         TabIndex        =   26
         Top             =   4560
         Width           =   7695
      End
   End
   Begin VB.Frame advancedFrame 
      Caption         =   "Advanced Options"
      Height          =   495
      Left            =   0
      TabIndex        =   20
      Top             =   360
      Width           =   8175
      Begin VB.CommandButton openAccessButton 
         Caption         =   "open Access database and Quit"
         Height          =   2295
         Left            =   120
         Picture         =   "frmMain.frx":6AB1A
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Quit this application and load the database file using MS Access. If Access is not found, it will issue an error. "
         Top             =   2880
         Width           =   1335
      End
      Begin VB.CommandButton exportDbToExcelButton 
         Caption         =   "Save the master inventory list as an Excel Spreadsheet (for printing, etc)"
         Height          =   2295
         Left            =   1560
         Picture         =   "frmMain.frx":6AE24
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   2880
         Width           =   2055
      End
      Begin VB.Label dbPathLabel 
         Caption         =   "dbPath"
         Height          =   735
         Left            =   240
         TabIndex        =   21
         Top             =   1080
         Width           =   5055
      End
   End
   Begin VB.Frame editDbFrame 
      Height          =   615
      Left            =   120
      TabIndex        =   12
      Top             =   360
      Width           =   7935
      Begin MSDataGridLib.DataGrid grid 
         Height          =   3975
         Left            =   360
         TabIndex        =   13
         ToolTipText     =   "This shows the actual database itself. Modifying the data here will result in immediate update of the data. So becareful. "
         Top             =   240
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   7011
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         AllowAddNew     =   -1  'True
         AllowDelete     =   -1  'True
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
      Begin VB.Label editDbLabel 
         Caption         =   "Becareful when editing the data. Once the data is entered, it is updated immediately. "
         Height          =   495
         Left            =   120
         TabIndex        =   14
         Top             =   4680
         Width           =   4695
      End
   End
   Begin VB.Frame inventoryScanFrame 
      Height          =   3615
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   7935
      Begin VB.TextBox descText 
         DataField       =   "desc"
         DataSource      =   "samsDB"
         Height          =   975
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   18
         Text            =   "desc"
         Top             =   2400
         Width           =   1335
      End
      Begin VB.TextBox quantityText 
         DataField       =   "quantityInStock"
         DataSource      =   "samsDB"
         Height          =   375
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   16
         Text            =   "quantity"
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox productIdText 
         DataField       =   "id"
         DataSource      =   "samsDB"
         Height          =   375
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   15
         Text            =   "productId"
         Top             =   960
         Width           =   1335
      End
      Begin VB.PictureBox productPicture 
         Height          =   2292
         Left            =   120
         ScaleHeight     =   2235
         ScaleWidth      =   2235
         TabIndex        =   6
         ToolTipText     =   "This shows the picture of the product. If there is no picture, it is blank. "
         Top             =   960
         Width           =   2292
      End
      Begin VB.Frame productGridFrame 
         Height          =   4812
         Left            =   3960
         TabIndex        =   4
         Top             =   120
         Width           =   3852
         Begin VB.Frame finishButtonFrame 
            Height          =   4332
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   972
            Begin VB.CommandButton cancelButton 
               Caption         =   "Cancel (Do not Save)"
               Height          =   1212
               Left            =   120
               TabIndex        =   11
               Top             =   3000
               Width           =   732
            End
            Begin VB.CommandButton finishButton 
               Caption         =   "Finish and Save"
               Height          =   1692
               Left            =   120
               Picture         =   "frmMain.frx":6B12E
               Style           =   1  'Graphical
               TabIndex        =   10
               Top             =   1200
               Width           =   732
            End
            Begin VB.CommandButton removeItemButton 
               Caption         =   "<"
               Height          =   852
               Left            =   120
               TabIndex        =   9
               Top             =   240
               Width           =   732
            End
         End
         Begin MSFlexGridLib.MSFlexGrid productGrid 
            Height          =   4332
            Left            =   1200
            TabIndex        =   5
            Top             =   240
            Width           =   2532
            _ExtentX        =   4471
            _ExtentY        =   7646
            _Version        =   393216
            Rows            =   0
            FixedRows       =   0
            FixedCols       =   0
            ScrollTrack     =   -1  'True
            SelectionMode   =   1
         End
      End
      Begin VB.TextBox currentProductIDText 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   672
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   3732
      End
      Begin VB.Label Label2 
         Caption         =   "Description"
         Height          =   255
         Left            =   2520
         TabIndex        =   19
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Quantity In Stock"
         Height          =   255
         Left            =   2520
         TabIndex        =   17
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label inventoryScanHelpText 
         Caption         =   "inventoryScanHelpText"
         Height          =   855
         Left            =   120
         TabIndex        =   7
         Top             =   4080
         Width           =   3735
      End
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   7080
      Top             =   0
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar statusBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   5715
      Width           =   8250
      _ExtentX        =   14552
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14499
            Text            =   "Status"
            TextSave        =   "Status"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TabStrip tabStrip 
      Height          =   4695
      Left            =   0
      TabIndex        =   1
      Top             =   840
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   8281
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   6
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Sam's Int"
            Key             =   "sam"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Inventory View"
            Key             =   "view"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Remove Items (-)"
            Key             =   "-"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Add Items (+)"
            Key             =   "+"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Edit Database"
            Key             =   "edit"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Advanced"
            Key             =   "advanced"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFileExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim currentUserInput As String
Dim currentProductIDTextHasFocus As Boolean



Private Sub Form_Activate()
    tabStrip_Click
End Sub

Private Sub Form_Initialize()
    productGrid.ColWidth(0) = productGrid.Width * 0.6
    productGrid.ColWidth(1) = productGrid.Width * 0.1

    'currentProductIDText.text = "WOW"

    
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If Not currentProductIDTextHasFocus Then
        currentProductIDText.SetFocus
        currentProductIDText.Text = currentProductIDText.Text + Chr(KeyAscii)
    End If
    'currentProductIDText_KeyPress (KeyAscii)
End Sub

Private Sub Form_Load()
    initDb
    Set grid.DataSource = dbCursor
    dbPathLabel.Caption = dbPath

    'LoadResStrings Me
    'Me.Left = GetSetting(App.title, "Settings", "MainLeft", 1000)
    'Me.Top = GetSetting(App.title, "Settings", "MainTop", 1000)
    'Me.Width = GetSetting(App.title, "Settings", "MainWidth", 6500)
    'Me.Height = GetSetting(App.title, "Settings", "MainHeight", 6500)
    
End Sub






Private Sub mnuHelpAbout_Click()
    frmAbout.Show vbModal, Me
End Sub

Private Sub mnuToolsOptions_Click()
    frmOptions.Show vbModal, Me
End Sub

Private Sub mnuFileExit_Click()
    'unload the form
    Unload Me

End Sub

Private Sub mnuFileSend_Click()
    'ToDo: Add 'mnuFileSend_Click' code.
    MsgBox "Add 'mnuFileSend_Click' code."
End Sub

Private Sub mnuFilePrintPreview_Click()
    'ToDo: Add 'mnuFilePrintPreview_Click' code.
    MsgBox "Add 'mnuFilePrintPreview_Click' code."
End Sub

Private Sub mnuFilePageSetup_Click()
    On Error Resume Next
    With dlgCommonDialog
        .DialogTitle = "Page Setup"
        .CancelError = True
        .ShowPrinter
    End With

End Sub

Private Sub mnuFileProperties_Click()
    'ToDo: Add 'mnuFileProperties_Click' code.
    MsgBox "Add 'mnuFileProperties_Click' code."
End Sub





Private Sub productGrid_Click()
    productGrid.col = 0 ' always point to id, not quantity
    showProductData productGrid.Text
End Sub





Private Sub removeItemButton_Click()
    'Debug.Print "Clicked RemoveItemButton"
    Dim index, count, Key As String
    If inventoryCount.getSize() = 0 Then
        Exit Sub
    End If
    productGrid.col = 0
    Key = productGrid.Text
    'Debug.Print "Key =" & key
    inventoryCount.dec (Key)
    index = inventoryCount.getValue(Key, 0)
    count = inventoryCount.getValue(Key, 1)
    productGrid.row = index
    productGrid.col = 1
    productGrid.Text = CStr(count)
    productGrid.TopRow = index
    'Debug.Print productGrid.text
End Sub



Private Sub tabStrip_Click()
    If tabStrip.SelectedItem.Key = previousTabStripKey Then
        Exit Sub
    End If
    currentTabKey = tabStrip.SelectedItem.Key
    clearProductGrid
    Select Case tabStrip.SelectedItem.Key
    Case "sam"
        guiList.hideAllFramesExcept introFrame
        Me.KeyPreview = False
    Case "view"
        guiList.hideAllFramesExcept inventoryScanFrame
        productGridFrame.Visible = True
        inventoryScanHelpText = "Enter the product ID to view the " & _
            "product data."
        currentProductIDText.SetFocus
        Me.KeyPreview = True
    Case "+"
        guiList.hideAllFramesExcept inventoryScanFrame
        finishButtonFrame.Visible = True
        productGridFrame.Visible = True
        currentProductIDText.SetFocus
        inventoryScanHelpText = "Enter the product ID to add the " & _
            "product to the database. Click on Finish to save the" & _
            "database. If you don't want to save the database, click on " & _
            "Cancel button."
         Me.KeyPreview = True
            
    Case "-"
        guiList.hideAllFramesExcept inventoryScanFrame
        finishButtonFrame.Visible = True
        productGridFrame.Visible = True
        currentProductIDText.SetFocus
        inventoryScanHelpText = "Enter the product ID to remove the " & _
            "product to the database. Click on Finish to save the" & _
            "database. If you don't want to save the database, click on " & _
            "Cancel button."
        Me.KeyPreview = True
    Case "edit"
        guiList.hideAllFramesExcept editDbFrame
        Me.KeyPreview = False
    Case "advanced"
        guiList.hideAllFramesExcept advancedFrame
        Me.KeyPreview = False
    Case Else
        MsgBox ("Error: unknown tab selected (" & tabStrip.SelectedItem.Key & ")")
        ' ERROR: unknown tab
    End Select
    previousTabStripKey = tabStrip.SelectedItem.Key
    'guiList.dprint
    
    'Debug.Print "index="; tabStrip.SelectedItem.index
    'Debug.Print "key=" & tabStrip.SelectedItem.key
End Sub

 
Sub clearProductGrid()
    productGrid.clear
    productGrid.Rows = 0
    inventoryCount.clear
End Sub





