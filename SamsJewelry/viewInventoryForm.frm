VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form viewInventoryForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sams - Data View"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   400
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   640
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton printButton 
      Caption         =   "Print"
      Height          =   1095
      Left            =   5640
      MaskColor       =   &H00000000&
      Picture         =   "viewInventoryForm.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4800
      UseMaskColor    =   -1  'True
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   4575
      Left            =   7680
      TabIndex        =   4
      Top             =   0
      Width           =   1815
   End
   Begin VB.CommandButton openAccessButton 
      Caption         =   "Open Access"
      Enabled         =   0   'False
      Height          =   1095
      Left            =   8400
      MaskColor       =   &H00000000&
      Picture         =   "viewInventoryForm.frx":1B42
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4800
      UseMaskColor    =   -1  'True
      Width           =   1095
   End
   Begin VB.CommandButton exportToExcelButton 
      Caption         =   "Export to Excel"
      Enabled         =   0   'False
      Height          =   1095
      Left            =   7080
      MaskColor       =   &H00000000&
      Picture         =   "viewInventoryForm.frx":3684
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4800
      UseMaskColor    =   -1  'True
      Width           =   1215
   End
   Begin VB.CommandButton saveButton 
      Appearance      =   0  'Flat
      Caption         =   "Finish"
      Default         =   -1  'True
      Height          =   1095
      Left            =   120
      MaskColor       =   &H00000000&
      Picture         =   "viewInventoryForm.frx":51C6
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4800
      UseMaskColor    =   -1  'True
      Width           =   5415
   End
   Begin MSDataGridLib.DataGrid dbGrid 
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   7858
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
End
Attribute VB_Name = "viewInventoryForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim myProductTable As New TheAdoTable

Private myProductTable As New TheAdoTable2

Private Sub dbGrid_AfterColUpdate(ByVal ColIndex As Integer)
    'myProductTable.update
End Sub

Private Sub dbGrid_HeadClick(ByVal ColIndex As Integer)
    If ColIndex >= 2 Then
        Exit Sub
    End If
    
    'sortDb dbGrid.Columns(ColIndex).DataField
     'sortDb PRODUCT_ID
    'Set dbGrid.DataSource = dataViewCursor
    'dbGrid.refresh
End Sub

Private Sub exportToExcelButton_Click()
    ' TODO: use Import/Export()/exportTable()/copyTable() instead.
    Dim sheet As New TheExcelTable
    sheet.openFile
    
    Dim myTable As TheTable
    'TODO: sheet.copyTable (productTable)
    
    'ignore below
    Dim row As Long

    productTable.gotoFirst
    sheet.setCell 0, 0, PRODUCT_ID   'write heading
    sheet.setCell 1, 0, QUANTITY
    row = 1
    
    Do While (Not productTable.EOF())
        sheet.setCell 0, row, productTable.getValue(PRODUCT_ID)
        sheet.setCell 1, row, productTable.getValue(QUANTITY)
        row = row + 1       'row++
        productTable.gotoNext
    Loop
    Unload Me
End Sub

Private Sub Form_Load()
    'sortDb PRODUCT_ID
    'productTable.refresh
    Set dbGrid.DataSource = inventory.ProductEditTable.getRecordSet()
    If currentOperation = EDIT_INVENTORY Then
        ' make it editable
        openAccessButton.Enabled = True
        exportToExcelButton.Enabled = True
        dbGrid.AllowDelete = True
        dbGrid.AllowUpdate = True
        dbGrid.AllowAddNew = True
        
    End If
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'dataViewCursor.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set dbGrid.DataSource = Nothing
    myProductTable.closeFile    'just to be safe.
End Sub

Private Sub openAccessButton_Click()
    fileSystem.launchApp App.path() & "\" & DB_FILENAME
    Unload Me
End Sub



Private Sub printButton_Click()
    MsgBox "Not implemented yet."
End Sub

Private Sub saveButton_Click()
    Unload Me
End Sub

Private Sub sortDb(fieldToSort As String)
    'TODO: See OLD CODE
End Sub







'=============================================================================
'=============================================================================
#If OLD_CODE Then
Private Sub dbGrid_HeadClick(ByVal ColIndex As Integer)
    If ColIndex >= 2 Then
        Exit Sub
    End If
    
    sortDb dbGrid.Columns(ColIndex).DataField
     'sortDb PRODUCT_ID
    Set dbGrid.DataSource = dataViewCursor
    dbGrid.refresh
End Sub

Private Sub exportToExcelButton_Click()
    Dim sheet As New TheExcelTable
    sheet.openFile
    
    Dim row As Long

    dbCursor.MoveFirst
    sheet.setCell 0, 0, "ID"    'write heading
    sheet.setCell 1, 0, "Qty"
    row = 1
    
    Do While (Not dbCursor.EOF)
        sheet.setCell 0, row, dbCursor("id").value
        sheet.setCell 1, row, dbCursor("quantityInStock").value
        row = row + 1       'row++
        dbCursor.MoveNext
    Loop
    Unload Me
End Sub

Private Sub Form_Load()

    sortDb PRODUCT_ID
    Set dbGrid.DataSource = dataViewCursor

    
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    dataViewCursor.Close
    ' Me.dbGrid.ClearFields
End Sub

Private Sub openAccessButton_Click()
    fileSystem.launchApp App.path() & "\" & DB_FILENAME
    Unload Me
End Sub



Private Sub saveButton_Click()
    dataViewCursor.Update
    Unload Me
End Sub

Private Sub sortDb(fieldToSort As String)
    Static sortMode As String
        'either "desc" or "asc"
    Static sortedColumnName As String
        'name of column being sorted currently
        
    If dataViewCursor.state <> adStateClosed Then
        dataViewCursor.Close
    Else    '--------- First time running this func.
        'sortMode = "desc"
        sortedColumnName = fieldToSort
    End If
    If fieldToSort = sortedColumnName Then  'user clicked on same column
        sortMode = IIf(sortMode = "desc", "asc", "desc")
    Else
        sortMode = "desc"
    End If
    sortedColumnName = fieldToSort
    dataViewCursor.cursorLocation = adUseClient
    dataViewCursor.cursorType = adOpenKeyset
    dataViewCursor.lockType = adLockPessimistic
    dataViewCursor.Open "select * from products order by " & fieldToSort & " " & sortMode, samsDb
End Sub

#End If

