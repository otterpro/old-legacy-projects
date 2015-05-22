VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form dataWindow 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6210
   ClientLeft      =   1485
   ClientTop       =   720
   ClientWidth     =   9645
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   414
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   643
   Begin VB.Frame Frame6 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   6000
      TabIndex        =   17
      Top             =   1680
      Width           =   2415
      Begin VB.PictureBox Picture1 
         Height          =   1935
         Left            =   0
         ScaleHeight     =   1875
         ScaleWidth      =   2355
         TabIndex        =   21
         Top             =   120
         Width           =   2415
      End
   End
   Begin VB.Frame Frame5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   6000
      TabIndex        =   16
      Top             =   0
      Width           =   3615
      Begin VB.TextBox outputText 
         Enabled         =   0   'False
         Height          =   1575
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   20
         Text            =   "dataWindow.frx":0000
         Top             =   120
         Width           =   3375
      End
   End
   Begin VB.Frame Frame4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4935
      Left            =   120
      TabIndex        =   14
      Top             =   1200
      Width           =   5895
      Begin MSDataGridLib.DataGrid productGrid 
         Height          =   4455
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   7858
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   21
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
   Begin VB.CommandButton decreaseMultiplierButton 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   11
      Top             =   720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Frame Frame3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   6000
      TabIndex        =   2
      Top             =   4200
      Width           =   3615
      Begin VB.CommandButton finishLaterButton 
         Caption         =   "Finish Later"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   1560
         MaskColor       =   &H00000000&
         Picture         =   "dataWindow.frx":0006
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   975
      End
      Begin VB.CommandButton cancelButton 
         Appearance      =   0  'Flat
         Caption         =   "Empty"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   2520
         MaskColor       =   &H00000000&
         Picture         =   "dataWindow.frx":0D48
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   975
      End
      Begin VB.CommandButton finishButton 
         Caption         =   "Finish"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   120
         MaskColor       =   &H00000000&
         Picture         =   "dataWindow.frx":288A
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   5895
      Begin VB.CommandButton increaseMultiplierButton 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5280
         TabIndex        =   10
         Top             =   360
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox multiplierText 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   4800
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   "1"
         Top             =   480
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.TextBox productIdText 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   8
         Top             =   420
         Width           =   4335
      End
      Begin VB.Label Label2 
         Caption         =   "Product ID: Use Keyboard or Barcode scanner to enter ID Number"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   120
         Width           =   4935
      End
      Begin VB.Label Label1 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4560
         TabIndex        =   12
         Top             =   600
         Visible         =   0   'False
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   8400
      TabIndex        =   0
      Top             =   1680
      Width           =   1215
      Begin VB.Label Label5 
         Caption         =   "Quantity:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   600
         Width           =   615
      End
      Begin VB.Label descText 
         Caption         =   "desc"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   120
         TabIndex        =   7
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label quantityText 
         Caption         =   "quantity"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   615
      End
      Begin VB.Label productNameText 
         Caption         =   "productName"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   975
      End
   End
End
Attribute VB_Name = "dataWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=============================================================================
'   dataWindow:
'       main data-entry form
'=============================================================================
Option Explicit
Private Const INDEX_WIDTH = 50  '50 twips wide
Private Const PRODUCT_ID_WIDTH_PERCENTAGE = 0.3
Private Const QUANTITY_WIDTH_PERCENTAGE = 0.2
Private Const NOTE_WIDTH_PERCENTAGE = 0.4
    'width of columns are 50%, '20%, 30% respectively


Dim productIdTextHasFocus As Boolean

Private Sub cancelButton_Click()
    inventory.clearWorkTable  'erase inventory.inventory.workTable if necessary. Asks user too.
    ' Unload Me
End Sub


Private Sub decreaseMultiplierButton_Click()
    multiplier = Max(multiplier - 1, 1)
    multiplierText.text = CStr(multiplier)
End Sub

Private Sub finishButton_Click()
    '-------- view mode doesn't do anyting.
    If inventory.currentOperation = VIEW_INVENTORY Then
        Unload Me
        Exit Sub    ' exit is needed because sometimes Unload is delayed. In the meantime, it continues with the program!
    End If
    dataWindow.Caption = "Please wait while saving."
    dataWindow.Enabled = False  'better to disable whole form than just button
    inventory.updateDb
    Unload Me
End Sub

Private Sub finishLaterButton_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    productIdText.SetFocus
    productIdTextHasFocus = False
End Sub

Private Sub Form_Load()
    multiplier = 1
    multiplierText.text = str(multiplier)
    'productGrid.ColWidth(0) = 120 * 4
    'productGrid.ColWidth(1) = 120 * 20
    
    'copy field name from inventory class
    fieldNameConstant.id = inventory.getFieldName(PRODUCT_NAME)
    fieldNameConstant.note = inventory.getFieldName(PRODUCT_DESC)
    fieldNameConstant.quantity = inventory.getFieldName(PRODUCT_QUANTITY)
    
    Set productNameText.DataSource = inventory.productTable.getRecordSet()
    productNameText.DataField = fieldNameConstant.id
    'productNameText.DataField = inventory.getFieldName(PRODUCT_NAME)
    
    Set quantityText.DataSource = inventory.productTable.getRecordSet()
    quantityText.DataField = fieldNameConstant.quantity
    'quantityText.DataField = inventory.getFieldName(PRODUCT_QUANTITY)
    Set descText.DataSource = inventory.productTable.getRecordSet()
    descText.DataField = fieldNameConstant.note
    'descText.DataField = inventory.getFieldName(PRODUCT_DESC)
    'Set productPicture.DataSource = productTable.getRecordSet()
    
    ' Show Title and save setting
    inventory.ConfigTable.setConfig inventory.getFieldName(LAST_OPERATION), _
    inventory.currentOperation
    
    'ConfigTable.setValue CONFIG_WORK_TABLE_OPERATION, currentOperation 'OLD way
    If inventory.currentOperation = SALE_IN Then
        dataWindow.Caption = "Sold/In (Sales) for " & inventory.salespersonName & _
            ". Enter the items that are sold. The unsold items are placed back into the inventory."
        inventory.ConfigTable.setConfig inventory.getFieldName(LAST_SALESPERSON_ID), inventory.SalespersonId
    ElseIf inventory.currentOperation = SALE_OUT Then
        dataWindow.Caption = "Out (Sales) for " & inventory.salespersonName
        inventory.ConfigTable.setConfig inventory.getFieldName(LAST_SALESPERSON_ID), inventory.SalespersonId
    Else
        dataWindow.Caption = inventory.getCurrentOperationString()   'title bar reflects the Mode
    End If
    
    
    
    'If importExcelTable.getRowSize() > 0 Then
    '    Dim row, col
    '    Dim grid As New TheFlexGrid
    '    grid.openFile productGrid
    '    grid.copy importExcelTable
    '    importExcelTable.clear
    'End If

    inventory.workTable.refresh
    
    'inventory.inventory.workTable.openFile "select " & WORK_PRODUCT_ID & "," & WORK_QUANTITY & ", " & _
        WORK_NOTE & " from " & WORK_TABLE, samsDb

    Set productGrid.DataSource = inventory.workTable.getRecordSet()
    If productGrid.Columns.Count > 2 Then
        productGrid.Columns(0).Width = INDEX_WIDTH    ' Index Field should not be shown.
        productGrid.Columns(1).Width = productGrid.Width * PRODUCT_ID_WIDTH_PERCENTAGE
        productGrid.Columns(2).Width = productGrid.Width * QUANTITY_WIDTH_PERCENTAGE
        productGrid.Columns(3).Width = productGrid.Width * NOTE_WIDTH_PERCENTAGE
    End If

    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'inventoryCount.clear
    inventory.ConfigTable.setConfig inventory.getFieldName(WORK_TABLE_SIZE), inventory.workTable.getSize()
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If Not productIdTextHasFocus Then
        productIdText.SetFocus
        productIdText.text = productIdText.text + Chr(KeyAscii)
    End If
    'currentProductIDText_KeyPress (KeyAscii)
End Sub


Private Sub increaseMultiplierButton_Click()
    multiplier = multiplier + 1
    multiplierText.text = CStr(multiplier)
End Sub



'Private Sub multiplierText_Change()
    'multiplier = Val(Me.multiplierText.text)
'End Sub


Private Sub multiplierText_Click()
    Dim valueToChange As Long
    valueToChange = val(Me.multiplierText.text)
    'changeValueDialog.Show vbModal, Me
    'Me.multiplierText.text = CStr(valueToChange)
    If (TheDialog.getInteger(multiplier, defaultValue:=valueToChange)) Then
            Me.multiplierText.text = CStr(multiplier)
    End If
End Sub

Private Sub productGrid_DblClick()
    If inventory.workTable.getSize() = 0 Then Exit Sub
    Dim old As ProductType
    
    old.quantity = inventory.workTable.getValue(fieldNameConstant.quantity, 0)
    dataEntryParam.quantity = old.quantity
    old.note = inventory.workTable.getValue(fieldNameConstant.note, "")
    
    dataEntryParam.note = old.note
    dataEntryParam.id = inventory.workTable.getValue(fieldNameConstant.id)
    dataEntryForm.Show vbModal, Me
    If dataEntryParam.quantity <> old.quantity Or _
        dataEntryParam.note <> old.note Then
        'Change value.
        inventory.workTable.setValue fieldNameConstant.quantity, _
                dataEntryParam.quantity
        inventory.workTable.setValue fieldNameConstant.note, _
                dataEntryParam.note
        
    End If
End Sub

'Private Sub multiplierText_GotFocus()
    'dataWindow.KeyPreview = False
'End Sub

'Private Sub multiplierText_LostFocus()
    'Me.KeyPreview = True
'End Sub

Private Sub productGrid_KeyPress(KeyAscii As Integer)
    productGrid.col = 0 ' always point to id, not quantity
    'showProductData
End Sub


Private Sub productIdText_gotfocus()
    productIdText.SelStart = Len(productIdText.text)
    productIdTextHasFocus = True
End Sub

Private Sub productIdText_KeyUp(KeyCode As Integer, Shift As Integer)
    ' clear the CR-LF key. It doesn't have to take care of it since
    ' keyPress() already handled the previous content.
    If KeyCode = vbKeyReturn Then
        productIdText.text = ""
    End If
End Sub

Private Sub productIdText_lostfocus()
    productIdTextHasFocus = False
    productIdText.SelStart = 0
        'currently not working bc it needs focus
    productIdText.SelLength = Len(productIdText.text)
    productIdText.text = ""
End Sub

Private Sub ProductIDText_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        productIdText.text = lstrip(productIdText.text)
        If productIdText.text = "" Then Exit Sub
        productIdText.text = UCase(productIdText.text)
        addItemToWorkTable productIdText.text
        showProductData
        productIdText.text = ""
        
    End If
End Sub


Sub showProductData(Optional productId)
    ' Show product's data from the DB when the user selects
    ' an item from the grid (either via kb or mouse)
    'dprint productGrid.text
    If Not IsMissing(productId) Then
        Dim result As Boolean
        'result = inventory.productTable.findRecord(PRODUCT_ID_FIELD, "=", productId)
        result = inventory.findDb(productId)
    End If
    productNameText.refresh
    'productPicture.refresh
    quantityText.refresh
    descText.refresh
End Sub

Sub addItemToWorkTable(ByVal productId As String)
    On Error GoTo ADD_ITEM_TO_WORK_TABLE_FAILED
    
    'Check for spelling and validation.
    productId = inventory.validateProductId(productId)
    
    showProductData productId    'search and show data
    
    'Make sure the type is correct (ie it is Ring, Earring, etc.)
    'Verify if user enters something that doesn't look like a productId.
    If Not inventory.productTypeIsValid(productId) Then
        Dim userInput
        userInput = MsgBox("The item that you've just entered doesn't belong in any " & _
                       " of the categories. Are you sure it is correct?" & _
                       " Click 'YES' if you are sure. Click 'NO' to cancel this entry.", _
                       vbYesNo + vbQuestion)
        If userInput = vbNo Then
            Exit Sub
        End If
    End If
    
    inventory.addNewWorkItem productId
     
    Exit Sub
ADD_ITEM_TO_WORK_TABLE_FAILED:
    MsgBox Err.Description
    'productGrid.refresh
End Sub

Private Sub showPicture()
    ' FUTURE/WISH: picture not decided yet. Should pic be embedded? or external?
End Sub














'=============================================================================
'   OLD CODE
'=============================================================================
#If OLD_CODE Then
'Private Sub productGrid_Click()
'    Dim productId As String
'    productId = Me.productGrid.TextMatrix(Me.productGrid.row, 0)
'    showProductData productId
'    showPicture
'    'If currentOperation = VIEW_INVENTORY Then
'    '    Exit Sub
'    'End If
'    'dprint "x=", CStr(Me.productGrid.col)
'    'dprint "y=", CStr(Me.productGrid.row)
'    valueToChange = Me.productGrid.text()
'    valueToChange = TheDialog.getInteger(productId, CInt(valueToChange))
'    'dprint "valueToChange=", CStr(valueToChange)
'    'TheGetIntegerDialog.Caption = productId
'
'    'TheGetIntegerDialog.show vbModal, Me
'    Me.productGrid.text = CStr(valueToChange)
'
'End Sub
'
'Private Sub finishButton_Click()
'    '-------- view mode doesn't do anyting.
'    If currentOperation = VIEW_INVENTORY Then
'        Unload Me
'        Exit Sub    ' exit is needed because sometimes Unload is delayed. In the meantime, it continues with the program!
'    End If
'
'    'Dim inventoryTable As New TheAdoTable
'    Dim sqlStmt As String
'
'    'finishButton.Enabled = False
'    'CancelButton.Enabled = False
'    dataWindow.Caption = "Please wait while saving."
'    dataWindow.Enabled = False  'better to disable whole form than just button
'
'    saveExcelSheet  ' Show Excel Spreadsheet
'    If currentOperation = INVENTORY_IN Or currentOperation = INVENTORY_OUT Then
'        reportTableNote = ""
'        '-------- No items in the table.
'        If inventory.inventory.workTable.getSize() = 0 Then
'            MsgBox "Please enter at least one item. " & _
'                "If you don't want to enter anything, press 'Finish Later' button to end this session."
'            Exit Sub
'        End If
'        updateDb
'    ElseIf currentOperation = SALE_IN Then
'        ' copy everything back in, except for the items entered in.
'        'open the salesperson's DB. Append the table to the DB.
'
'        'copy everything back into the "IN" table
'        salespersonTable.Find SALESPERSON_NAME & "='" & Salesperson & "'"
'        Dim inTable As New TheTable
'        sqlStmt = "select * from " & salespersonTable.getValue("tableName")
'        inventoryTable.openFile sqlStmt, samsDb
'        TheTableBase.add inTable, inventoryTable
'        currentOperation = INVENTORY_IN
'        reportTableNote = Salesperson
'        updateDb sourceTable:=inTable, updateInOutReport:=False  '
'        'delete all items in the inventoryTable since it is no longer needed.
'        inventoryTable.deleteAll
'
'
'        inventoryTable.closeTable
'
'        'now do an out of the table
'        currentOperation = INVENTORY_OUT
'        reportTableNote = Salesperson & " (Items sold)"
'        updateDb
'
'        'salesperson's Out is set to 0.
'        salespersonTable.setValue SALESPERSON_OUT, 0
'        salespersonTable.Update
'
'    ElseIf currentOperation = SALE_OUT Then
'        'open the salesperson's DB. Append the table to the DB.
'        Dim table As New TheTable
'        table.add productGrid
'        salespersonTable.Find "name='" & Salesperson & "'"
'
'        sqlStmt = "select * from " & salespersonTable.getValueAtColumn("tableName")   ' ex: table=out23
'        inventoryTable.openTable sqlStmt, samsDb
'        TheTableBase.add inventoryTable, table
'        Dim outCount As Long
'        reportTableNote = Salesperson & " (Out for show/sales)"
'        outCount = updateDb()
'        inventoryTable.Update
'        inventoryTable.closeTable
'
'        Dim previousOutCount
'        previousOutCount = salespersonTable.getValueAtColumn(SALESPERSON_OUT)
'        salespersonTable.setValue SALESPERSON_OUT, previousOutCount + outCount
'        'dbSalespersonCursor.update 'redundant.
'        'MsgBox dbSalespersonCursor.Status
'
'        'salespersonTable.RecordSet.Find("name")
'    Else
'        eprint "dataWindow::finishButton_Click() currentOperation not valid"
'    End If
'    Unload Me
'End Sub

#End If

