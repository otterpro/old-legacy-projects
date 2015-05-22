VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form mainForm 
   Caption         =   "Inventory Control"
   ClientHeight    =   7200
   ClientLeft      =   1110
   ClientTop       =   1680
   ClientWidth     =   9615
   LinkTopic       =   "Form1"
   ScaleHeight     =   7200
   ScaleWidth      =   9615
   Begin VB.Frame salesFrame 
      Caption         =   "Salesperson"
      Height          =   2295
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   3615
      Begin VB.CommandButton outSalesButton 
         Caption         =   "Out (Sales)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1920
         Left            =   1800
         MaskColor       =   &H00FF00FF&
         Picture         =   "mainForm.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   1695
      End
      Begin VB.CommandButton inSalesButton 
         Caption         =   "Sold (Sales)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1920
         Left            =   120
         MaskColor       =   &H00000000&
         Picture         =   "mainForm.frx":1B42
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Report Summary"
      Height          =   4455
      Left            =   120
      TabIndex        =   6
      Top             =   2280
      Width           =   9615
      Begin TabDlg.SSTab SSTab1 
         Height          =   4095
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   9285
         _ExtentX        =   16378
         _ExtentY        =   7223
         _Version        =   393216
         Style           =   1
         TabHeight       =   520
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "General"
         TabPicture(0)   =   "mainForm.frx":3684
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).ControlCount=   0
         TabCaption(1)   =   "Report"
         TabPicture(1)   =   "mainForm.frx":36A0
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "inOutReportGrid"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Master Inventory"
         TabPicture(2)   =   "mainForm.frx":36BC
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "productTableGrid"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).ControlCount=   1
         Begin MSDataGridLib.DataGrid inOutReportGrid 
            CausesValidation=   0   'False
            Height          =   3495
            Left            =   -74880
            TabIndex        =   11
            TabStop         =   0   'False
            ToolTipText     =   "To print this report, select Print Report from the File Menu."
            Top             =   480
            Width           =   9015
            _ExtentX        =   15901
            _ExtentY        =   6165
            _Version        =   393216
            AllowUpdate     =   0   'False
            Enabled         =   -1  'True
            HeadLines       =   1
            RowHeight       =   21
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
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
         Begin MSDataGridLib.DataGrid productTableGrid 
            Height          =   3495
            Left            =   -74880
            TabIndex        =   12
            TabStop         =   0   'False
            ToolTipText     =   "To print this report, select Print Report from the File Menu."
            Top             =   480
            Width           =   9015
            _ExtentX        =   15901
            _ExtentY        =   6165
            _Version        =   393216
            AllowUpdate     =   0   'False
            HeadLines       =   1
            RowHeight       =   21
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
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
      End
      Begin MSComDlg.CommonDialog comdlg 
         Left            =   8760
         Top             =   3840
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin MSComctlLib.StatusBar mainStatusBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   6825
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   661
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   16457
            Text            =   "statusBar"
            TextSave        =   "statusBar"
         EndProperty
      EndProperty
   End
   Begin VB.Frame buttonFrame 
      Caption         =   "Master Inventory"
      Height          =   2295
      Left            =   3720
      TabIndex        =   0
      Top             =   0
      Width           =   5775
      Begin VB.CommandButton testButton 
         Caption         =   "test"
         Height          =   975
         Left            =   5280
         TabIndex        =   9
         Top             =   1320
         Width           =   495
      End
      Begin VB.CommandButton quitButton 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1920
         Left            =   3840
         MaskColor       =   &H00000000&
         Picture         =   "mainForm.frx":36D8
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   1815
      End
      Begin VB.CommandButton outButton 
         Caption         =   "Out"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1920
         Left            =   1920
         MaskColor       =   &H00FF00FF&
         Picture         =   "mainForm.frx":521A
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   1815
      End
      Begin VB.CommandButton inButton 
         Caption         =   "In"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1920
         Left            =   120
         MaskColor       =   &H00FF00FF&
         Picture         =   "mainForm.frx":6D5C
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   1815
      End
   End
   Begin VB.Menu fileMenu 
      Caption         =   "File"
      Begin VB.Menu PrintMenu 
         Caption         =   "Print Report"
      End
      Begin VB.Menu viewSalespersonMenu 
         Caption         =   "View Salesperson"
      End
      Begin VB.Menu viewItemMenu 
         Caption         =   "View Product"
         Enabled         =   0   'False
      End
      Begin VB.Menu editDbMenu 
         Caption         =   "Edit Database (Advanced)"
         Enabled         =   0   'False
      End
      Begin VB.Menu configMenu 
         Caption         =   "Configuration..."
         Enabled         =   0   'False
      End
      Begin VB.Menu aboutMenu 
         Caption         =   "About"
      End
      Begin VB.Menu exitMenu 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "mainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const IO_REPORT_ID_WIDTH = 0#
Private Const IO_REPORT_DATE_WIDTH_PERCENTAGE = 0.3
Private Const IO_REPORT_INOUT_WIDTH_PERCENTAGE = 0.1
Private Const IO_REPORT_NOTE_WIDTH_PERCENTAGE = 0.4

Private Const PRODUCT_TABLE_INDEX = 0
Private Const PRODUCT_TABLE_ID = 0.2
Private Const PRODUCT_TABLE_QUANTITY = 0.1
Private Const PRODUCT_TABLE_NOTE = 0.2
Private Const PRODUCT_TABLE_DATE = 0.2

'Dim inOutReportTable As New TheAdoTable2
'Dim inOutReportDb As New TheAdoDb

' These tables are reference to actual tables. They are used by the
' DataGrids in the report view.
Private myProductTable As New TheAdoTable2
Private myInOutReportTable As New TheAdoTable2

'Dim inOutReportIsChanged As Boolean

Private Sub aboutMenu_Click()
    frmAbout.Show vbModal, Me
End Sub

Private Sub configMenu_Click()
    configWindow.Show vbModal, Me
End Sub

Private Sub editDbMenu_Click()
    inventory.currentOperation = EDIT_INVENTORY
    viewInventoryForm.Show vbModal, Me
End Sub







Private Sub inOutReportGrid_HeadClick(ByVal ColIndex As Integer)
    'TODO: Create new DataGrid control subclassed and use the following....
    'Dim tempGrid As DataGrid
    'Set tempGrid = inOutReportGrid
    Dim header As String
    Static LastHeader As String
    Dim sortText As String
    header = inOutReportGrid.Columns(ColIndex).Caption
    'If header = "note" Then 'cannot sort note field.
    '    Exit Sub
    'End If
    If myInOutReportTable.getRecordSet().Fields(ColIndex).Type = adLongVarWChar Then
        Exit Sub    'ignore Text/Memo/Long strings since it can't Sort.
    End If
    
    If LastHeader = header Then
        LastHeader = ""
        sortText = header & " DESC"
    Else
        sortText = header
        LastHeader = header
    End If
    myInOutReportTable.getRecordSet().sort = sortText
    'the.dprint "Sort order=", sortText
End Sub

Private Sub PrintMenu_Click()
    'MsgBox CStr(SSTab1.Tab)
    Dim src As TheAdoTable2
    Select Case SSTab1.Tab
    Case 0  'general tab
        Exit Sub
    Case 1  'in-out report
        Set src = myInOutReportTable
    Case 2  'master inventory
        Set src = myProductTable
    Case Else
        'unknown tab
        Exit Sub
    End Select

    Dim sheet As New TheExcelTable

    'sheet.openFile visible:=False
    sheet.openFile visible:=True
    sheet.add src
    sheet.printOut
    sheet.closeFile quit:=True
    
End Sub

Private Sub productTableGrid_HeadClick(ByVal ColIndex As Integer)
    Dim header As String
    Static LastHeader As String
    Dim sortText As String
    
    header = productTableGrid.Columns(ColIndex).Caption
    If myProductTable.getRecordSet().Fields(ColIndex).Type = adLongVarWChar Then
        Exit Sub    'ignore Text/Memo/Long strings since it can't Sort.
    End If
    If LastHeader = header Then
        LastHeader = ""
        sortText = header & " DESC"
    Else
        sortText = header
        LastHeader = header
    End If
    myProductTable.getRecordSet().sort = sortText
    
    
End Sub

Private Sub quitButton_Click()
    Unload Me
End Sub







Private Sub testButton_Click()
    testThe.test1
End Sub

'Private Sub inOutReportOle_Updated(Code As Integer)
    'no need to check that vbOLEChanged but we're assuming that any event means it was altered.
'    inOutReportIsChanged = True
'    inOutReportOle.update
'End Sub

Private Sub exitMenu_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    buttonFrame.visible = True
    buttonFrame.Enabled = True
    mainStatusBar.SimpleText = "Total Items: " & CStr(inventory.getInventoryCount())
    'inventory.salespersonTable.refresh    'same as recset.Requery
    
    'set grid width


End Sub

Private Sub Form_Load()
    ' in-out report OLE
    'inOutReport.openFile App.path() & "\" & IN_OUT_REPORT_FILENAME_TEST, inOutReportOle
    'inOutReportOle.CreateEmbed App.path() & "\" & IN_OUT_REPORT_FILENAME
    myInOutReportTable.clone inventory.inOutReportTable
    Set inOutReportGrid.DataSource = myInOutReportTable.getRecordSet()
    inOutReportGrid.Columns(0).Width = IO_REPORT_ID_WIDTH
    inOutReportGrid.Columns(1).Width = IO_REPORT_DATE_WIDTH_PERCENTAGE * inOutReportGrid.Width
    inOutReportGrid.Columns(2).Width = IO_REPORT_INOUT_WIDTH_PERCENTAGE * inOutReportGrid.Width
    inOutReportGrid.Columns(3).Width = IO_REPORT_INOUT_WIDTH_PERCENTAGE * inOutReportGrid.Width
    inOutReportGrid.Columns(4).Width = IO_REPORT_NOTE_WIDTH_PERCENTAGE * inOutReportGrid.Width
    myProductTable.clone inventory.productTable
    Set productTableGrid.DataSource = myProductTable.getRecordSet()
    productTableGrid.Columns(0).Width = PRODUCT_TABLE_INDEX * productTableGrid.Width
    productTableGrid.Columns(1).Width = PRODUCT_TABLE_ID * productTableGrid.Width
    productTableGrid.Columns(2).Width = PRODUCT_TABLE_QUANTITY * productTableGrid.Width
    productTableGrid.Columns(3).Width = PRODUCT_TABLE_DATE * productTableGrid.Width
    'productTaBLEgrid.Columns(4).Width = * productTableGrid.Width
    'productTaBLEgrid.Columns(5).Width = * productTableGrid.Width
    
    
End Sub

Private Sub inButton_Click()
    inventory.currentOperation = INVENTORY_IN
    showDataWindow
End Sub

Sub showDataWindow()
    If inventory.hasUnfinishedWork() Then
        Exit Sub    ' work must match (out=out, in=in, salein=salein,etc)
    End If
    
    If inventory.currentOperation = SALE_IN Or inventory.currentOperation = SALE_OUT Then
        If inventory.workTable.getSize() > 0 Then
            inventory.changeSalesperson (inventory.getPastSalespersonId())  'use previous salesguy
        Else
            chooseSalespersonDialog.Show vbModal, Me    'select salesperson
            If ChooseSalesperson = False Then
                Exit Sub
            End If
        End If
        'special case: You can't do IN when there was no OUT
        If inventory.currentOperation = SALE_IN And inventory.SalespersonOut = 0 Then
            MsgBox "You must do an Out first before doing the In. " & _
            "Currently it is showing that the salesperson's inventory is empty."
            Exit Sub
        End If
    End If
    dataWindow.Show vbModal, Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Dim i As Integer
    'If inOutReportIsChanged Then
    '    inOutReport.saveFile     'Save the inOutReport.XLS (.OLE, actually)
   'End If
   
    'close all sub forms
    For i = Forms.Count - 1 To 1 Step -1
        Unload Forms(i)
    Next
    
End Sub


'Private Sub inFromExcelButton_Click()
Private Sub inFromExcelMenu_Click()
'    comdlg.Filter = "Docs (*.xls)|*.xls"
'    comdlg.Flags = cdlOFNFileMustExist
'    comdlg.CancelError = True 'insure that "cancel" button results in Error Evt
'    On Error GoTo CancelError
'    comdlg.ShowOpen
'    If comdlg.filename = "" Then
'        Exit Sub        ' usually won't occur since OnError will goto CancelError
'    End If
'
'    ' disable form to prevent user from clicking on anything while loading excel sheet
'    mainForm.Enabled = False
'    Dim excelApp As New TheExcelTable
'    excelApp.openFile comdlg.filename, readOnly:=True, visible:=False
'    mainStatusBar.SimpleText = "Loading " & comdlg.filename
'    'dprint "rowSize", CStr(excelApp.getRowSize())
'    'dprint "colSize", CStr(excelApp.getColumnSize())
'    'TheTableBase.dprint importExcelTable, "Before- importExcelTable"
'    TheTableBase.copyTo excelApp, importExcelTable, 1
'    'TheTableBase.dprint excelApp, "excelTable"
'    'TheTableBase.dprint importExcelTable, "After-importExcelTable"
'    excelApp.closeFile quit:=True
'    inButton_Click
'    mainStatusBar.SimpleText = ""
'    mainForm.Enabled = True
'CancelError:    'exit sub
End Sub

Private Sub inSalesButton_Click()
    inventory.currentOperation = SALE_IN
    showDataWindow
End Sub

Private Sub outButton_Click()
    inventory.currentOperation = INVENTORY_OUT
    showDataWindow
End Sub

Private Sub outSalesButton_Click()
    inventory.currentOperation = SALE_OUT
    showDataWindow
End Sub


Private Sub viewItemMenu_Click()
'Private Sub viewButton_Click()
    inventory.currentOperation = VIEW_INVENTORY
    showDataWindow
End Sub

Private Sub viewSalespersonMenu_Click()
    chooseSalespersonDialog.Show vbModal, Me
End Sub



'=============================================================================
'=============================================================================
#If OLD_CODE Then
'Private Sub reportGrid_AfterColEdit(ByVal ColIndex As Integer)
'    dbReportCursor.update
'End Sub

'    Dim previousSalesperson As String
'    'previousWorkTableOperation = inventory.workTable.getValue(CONFIG_WORK_TABLE_OPERATION)
'    If hasUnfinishedWorkTable() Then
'        hasUnfinishedSalespersonWorkTable = True
'    ElseIf SalespersonID <> _
'        ConfigTable.getConfig(CONFIG_WORK_TABLE_SALESPERSON_ID) Then
'         hasUnfinishedSalespersonWorkTable = True
'         MsgBox "You still have unfinished " & _
'            getOperationString(ConfigTable.getConfig(CONFIG_WORK_TABLE_OPERATION)) & _
'            " for " & previousSalesperson
'    Else
'         hasUnfinishedSalespersonWorkTable = False
'    End If


'' returns TRUE if inventory.workTable is not empty and the current operation is not the same.
'' or if the operation is SalesIn/Out, but the salesperson doesn't match.
'Private Function hasUnfinishedWorkTable()
'    Dim previousWorkTableOperation As operationMode
'    'Dim warningMsg As String
'    'OLD:ConfigTable.findRecord CONFIG_KEY, "=", CONFIG_WORK_TABLE_SIZE
'    If ConfigTable.getConfig(CONFIG_WORK_TABLE_SIZE) <= 0 Then
'        hasUnfinishedWorkTable = False
'        Exit Function
'    End If
'    previousWorkTableOperation = ConfigTable.getConfig(CONFIG_WORK_TABLE_OPERATION)
'    If previousWorkTableOperation <> currentOperation Then
'        hasUnfinishedWorkTable = True
'        'Select Case previousWorkTableOperation
'        MsgBox "You still have unfinished " & getOperationString(previousWorkTableOperation)
'    Else
'        hasUnfinishedWorkTable = False
'    End If
'
'End Function

'Private Function hasUnfinishedSalespersonWorkTable()
'    Dim previousSalesperson As String
'    'previousWorkTableOperation = inventory.workTable.getValue(CONFIG_WORK_TABLE_OPERATION)
'    If hasUnfinishedWorkTable() Then
'        hasUnfinishedSalespersonWorkTable = True
'    ElseIf SalespersonID <> _
'        ConfigTable.getConfig(CONFIG_WORK_TABLE_SALESPERSON_ID) Then
'         hasUnfinishedSalespersonWorkTable = True
'         MsgBox "You still have unfinished " & _
'            getOperationString(ConfigTable.getConfig(CONFIG_WORK_TABLE_OPERATION)) & _
'            " for " & previousSalesperson
'    Else
'         hasUnfinishedSalespersonWorkTable = False
'    End If
'
'End Function
'
'Sub showDataWindow()
'    If hasUnfinishedWork() Then
'        Exit Sub    ' work must match (out=out, in=in, salein=salein,etc)
'    End If
'    If currentOperation = SALE_IN Or currentOperation = SALE_OUT Then
'        If hasUnfinishedSalespersonWork() Then
'            chooseSalespersonDialog.Show vbModal, Me
'            If ChooseSalesperson = False Then
'                Exit Sub
'            End If
'            'Exit Sub
'        Else
'            'Use previously selected salesperson.
'        End If
'
'        Dim out As Long
'        out = salespersonTable.getValue(SALESPERSON_OUT)
'        If currentOperation = SALE_IN And out = 0 Then
'            MsgBox "You must do an Out first before doing the In. " & _
'            "Currently it is showing that the salesperson's inventory is empty."
'            Exit Sub
'        ElseIf currentOperation = SALE_OUT And salespersonTable.getValue("out") > 0 Then
'            MsgBox "Currently it is showing that the salesperson " & Salesperson.name & _
'            " has " & CStr(out) & "items Out."
'        End If
'    End If
'    ' hideButton
'    'buttonFrame.Enabled = False
'    'buttonFrame.Visible = False
'    'Me.Visible = False
'    'dataWindow.Show
'    dataWindow.Show vbModal, Me
'    'Me.Visible = True
'End Sub

#End If

