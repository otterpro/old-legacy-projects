Attribute VB_Name = "mainModule"
'=============================================================================
'=============================================================================
'TODO:
' _ in DataWindow, Cancel Button doesn't erase the table. Instead, let the user resume from
'   the previous data-entry and add "Clear" button.
'   - save to spreadsheet of entire inventory button (save automatically as "inventory.xls")
' _ add listbox sorted with item type, and its quantity and the total sum in dataWindow
'   _in Report, eliminate the duplicate In part. Just have the outpart.
' _ make the 0'ed item to be deleted from the listbox. Using Excel control will fix this.
' _ warn the user if the OUT item doesn't exist.
'  _ each item in the In is separate and has a price field, is unique.
' _ sort by item #, not by item type. Or separate item type in separate field.
'       [RG] [1234/a], etc.
'  _ Item type: RG,ER,PT,BL,NL,BR,BG
'   _ quick entry using radio buttons, user only enters the #.
'  _ print subtotal in the spreadsheet report.
' _ add price field in "Sold/In Sales"

' INSTALL ISSUES:
' make sure to backup original MDB in the user's PC before overwriting it.
' make folder called "inOutReport" folder? DOn't know if Save in Excel will create the folder.
'

'FIXED:
'_ detect multiple instance of this app and allow only 1 launch of this app.  Uses App.previnstance()
'  _use Excel Control to show report instead of list. Save the DB as Excel Spreadsheet?

' NOTES:
' Sets
' 1x2 = ER+RG, ER+PT
' 1x3 = RG,ER,PT
' not all sets have same item #. ex: set is ER-1000, RG-1002
' also, when the shipment is received, there is no info that tells us that it is part of a set.
' Salesperson always puts the rest of the unsold items back into the master inventory. Always.
' However, salesperson may add (out) items anytime.

'FUTURE/WISH:
' _ Get a copy of Database button so salesperson can get the latest copy of db if needed.
' _ add pictures, add option to get picture in copy of Database.
'


Option Explicit

'GLOBAL

Public mainWindow As mainForm
'Public mainWindow As dataViewWindow
'Public inventoryCount As New TheDictionary
Public fileSystem As New TheFileSystem
' Public inOutReport As New TheOle



Public guiList As New TheGuiList
'Public testNumber As Integer 'deleteme
'Public reportTableNote As String
    ' notes to add to the ReportTable when doing in/out.
'Public valueToChange As Long
    ' value to change in valueChangeDialog
Public multiplier As Long
    ' multiplies the quantity by this value when doing inventory

Public Type ProductType    'generic data structure for product
    quantity As Variant 'can be string in case of field name
    note As String
    id As String
End Type

Public dataEntryParam As ProductType
    ' used to pass var between dataEntryForm and main/data form
Public fieldNameConstant As ProductType
    'used to keep the name of fields (strings).
    
' DB and Tables
Public inventory As InventoryControl
'Public withevents productTable As TheAdoTable2

'Public salespersonTable As New TheAdoTable2
'Public salespersonInventoryTable As New TheAdoTable2
'Public balanceQuery As New TheAdoTable2
'Public inOutReportTable As New TheAdoTable2
'Public inventory.workTable As New TheAdoTable2
    'inventory.workTable is used as a temp table for keeping track of user's data-entry.
    
'external db/xls/ole filename
Public Const DB_FILENAME = "sams.mdb"
'Public Const IN_OUT_REPORT_FILENAME_TEST = "inOutReport.ole"

''names of tables in sams.mdb
'Public Const SALESPERSON_INVENTORY_TABLE_PREFIX = "salesInventory"
'Public Const PRODUCT_TABLE = "product"
'Public Const WORK_TABLE = "inventoryWork"
'    'temp table to hold the In/Out/etc item
''Public Const PRODUCT_TABLE_TEMPLATE = "productTemplate"
'    'just use WORK_TABLE.
'Public Const IN_OUT_REPORT_FILENAME_PREFIX = "inout"
'    'ex: inout1.xls, inout2.xls, inout3.xls, ...
'Public Const IN_OUT_REPORT_TABLE = "inOutReport"
'Public Const IN_OUT_REPORT_FOLDER = "inOutReport"
'    ' all inout*.xls is stored in <appPath>/inOutReport/ folder.
    
'' PRODUCT TABLE
'Public Const INDEX = "index"        'autonumber field
'Public Const QUANTITY = "quantity"
'Public Const PRODUCT_ID = "id"      ' "RG-1000/A"
'Public Const DESCRIPTION = "note"   ' "This is a turquoise ring"
'Public Const LAST_MODIFIED = "lastModified"
'Public Const PRODUCT_NUMBER = "productNumber"   ' 1000
'
'' Currently Not used
'Public Const PRODUCT_TYPE = "type"  '
'
'Public Const REPORT_ID = "id"
'Public Const REPORT_DATE = "date"
'Public Const REPORT_NOTE = "note"
'Public Const REPORT_IN = "in"
'Public Const REPORT_OUT = "out"
'Public Const REPORT_SALESPERSON = "salesperson"
'Public Const REPORT_SHEET_FILENAME = "sheetFilename"
'Public Const DB_PRODUCT_BALANCE = "balance"
'
'' WORK TABLE have same structure as PRODUCT TABLE
'Public Const WORK_INDEX = "index"
'Public Const WORK_PRODUCT_ID = "id"
'Public Const WORK_QUANTITY = "quantity"
'Public Const WORK_NOTE = "note"
'
''SALESPERSON TABLE
'Public Const SALESPERSON_TABLE = "salesperson"
'Public Const SALESPERSON_OUT = "out"
'Public Const SALESPERSON_ID = "id"
'Public Const SALESPERSON_NAME = "name"
'Public Const SALESPERSON_INVENTORY_TABLE_NAME = "tableName"
'
''Public Salesperson As String
'    ' currently selected salesperson, chosen in chooseSalespersonDialog. Used for In/Sales, Out/sales
''Public SalespersonID As Long
'    ' currently selected salesperson's id. Convenient to store it in a var.
''Public SalespersonInventoryTableName As String
'
'Public Type SalespersonType
'    name As String      '
'    id As Long          '
'    tableName As String 'inventory table name
'    out As Long         ' # of items out.
'End Type
'
'Public Salesperson As SalespersonType   'current salesperson
'                    'salesperson can be selected from In-Sales, Out-Sales
'
'' CONFIG TABLE
'Public ConfigTable As New TheAdoTable2
'Public Const CONFIG_TABLE = "config"
'Public Const CONFIG_WORK_TABLE_OPERATION = "workTableOperation"
'    'in,out,saleIn,saleOut (see OperationMode and currentOperation)
'    ' is in numbercode (enumerated)?
'Public Const CONFIG_WORK_TABLE_SALESPERSON_ID = "workTableSalesperson"
'    'id # of salesperson if the inventory.workTable operation is saleIn or saleOut
'Public Const CONFIG_WORK_TABLE_SIZE = "workTableSize"
'Public Const CONFIG_KEY = "key"
'Public Const CONFIG_VALUE = "value"
'
'    ' the size of work table the last time it was being used.
'    ' if 0, there is no need to worry about checking which operation we're doing.
'
'Public importExcelTable As New TheTable
'    ' if the user imports excel sheet for "in", it stores the data here temporarily
'    ' until it is retrieved by the dataWindow form.

''Dialog related
Public ChooseSalesperson As Boolean
'    'set by ChooseSalespersonDialog. If the user successfully chooses
'    'a salesperson, it is true. If not, it is false.
Sub Main()
    the.dprint "================================================"
    the.dprint "Sams Invoice App Started"
    the.dprint "================================================"
    
    If App.PrevInstance = True Then
        'MsgBox CStr(TheMSWin.findWindowHandle("Inventory Control"))
        Exit Sub    ' cannot run more than 1 instance.
    End If
    multiplier = 1
    ' Set inventoryCount = New TheDictionary
    'dbPath = App.path() & "\" & DEFAULT_DB_PATH
    'MsgBox "is it working???"
    frmSplash.Show
    frmSplash.refresh
    
    Dim dbPath As String
    dbPath = fileSystem.joinPath(App.path(), DB_FILENAME)
    If Dir(dbPath) = "" Then
        MsgBox "WARNING: Database file " & dbPath & " was not found. " _
            & "This application will not function properly without it."
        Exit Sub
    End If
    
    Set inventory = New InventoryControl
    inventory.openFile dbPath
    
    Set mainWindow = New mainForm
    'Set mainWindow = New dataViewWindow
    Load mainWindow
    Unload frmSplash
    mainWindow.Show
End Sub

Sub LoadResStrings(Frm As Form)
    On Error Resume Next
End Sub

'Public Function findProduct(ByRef adoTable As TheAdoTable2, ByVal productId As String) As Boolean
'    findProduct = adoTable.findRecord(PRODUCT_ID, "=", productId)
'End Function
'
'Public Function findDb(ByVal productId)
'    findDb = productTable.findRecord(PRODUCT_ID, "=", productId)
'End Function
'
'Public Function findInSalespersonTable(ByVal id) As Boolean
'    findInSalespersonTable = salespersonTable.findRecord(PRODUCT_ID, "=", id)
'End Function
'
'' save as excel sheet and update data
'Public Sub saveExcelSheet()
'    Dim sheet As New TheExcelTable
'    Dim table As New TheTable
'    sheet.openFile visible:=False
'    table.addRow PRODUCT_ID, QUANTITY, REPORT_NOTE
'    'table.add dataWindow.productGrid
'    'sheet.copy table
'    TheTableBase.copyTo inventory.workTable, table
'    sheet.add table
'    sheet.setVisible
'End Sub
'
'
'Public Sub clearWorkTable(Optional askUser As Boolean = True)
'    If inventory.workTable.getSize() > 0 And askUser Then
'        If vbOK <> MsgBox("Click OK to erase everything. Click Cancel to go back.", vbOKCancel + vbQuestion, _
'                        "Are you sure you want to delete the current sheet?") Then
'            Exit Sub    'NO. User doesn't want to erase.
'        End If
'    End If
'    inventory.workTable.clear 'erase the content.
'    ConfigTable.setConfig CONFIG_WORK_TABLE_SIZE, 0 ' save 0 to config file.
'End Sub
'
'' update Products table and Report table.
'' updateInOutReport
'' Returns # of items that were either placed in the "In" or the "Out" box.
'Public Sub updateDb()  'Optional sourceTable, Optional updateInOutReport As Boolean = True)
'   'Dim inventoryTable As New TheAdoTable2
'    Dim sqlStmt As String
'    Dim salespersonInventoryTable As New TheAdoTable2
'    Dim note As String  'additional text to add to the note field in the InOutReportTable.
'
'    'saveExcelSheet  ' TODO:Show Excel Spreadsheet
'
'    If currentOperation = INVENTORY_IN Then
'        'productT += workT
'        productTable.merge inventory.workTable, PRODUCT_ID, PRODUCT_ID, QUANTITY, QUANTITY, MERGE_ADD
'    ElseIf currentOperation = INVENTORY_OUT Then
'        'productT -= workT
'        productTable.merge inventory.workTable, PRODUCT_ID, PRODUCT_ID, QUANTITY, QUANTITY, MERGE_SUBTRACT
'    ElseIf currentOperation = SALE_IN Or currentOperation = SALE_OUT Then
'        'open salesperson's inventory Table
'        sqlStmt = "SELECT * FROM " & salespersonTable.getValue(SALESPERSON_INVENTORY_TABLE_NAME)
'        salespersonInventoryTable.openFile sqlStmt, samsDb
'        If currentOperation = SALE_IN Then  'SOLD
'            ' formula:
'            '1. salespersonT -= inventory.workTable  # of items sold
'            '2. productT += salespersonT   # of returned items to inventory
'            salespersonTable.merge inventory.workTable, PRODUCT_ID, PRODUCT_ID, QUANTITY, QUANTITY, MERGE_SUBTRACT
'            productTable.merge salespersonInventoryTable, PRODUCT_ID, PRODUCT_ID, QUANTITY, QUANTITY, MERGE_ADD
'            salespersonInventoryTable.clear
'        Else    ' SALE_OUT, append to salesperson's out table
'            ' salespersonT += inventory.workTable
'            salespersonInventoryTable.merge inventory.workTable, PRODUCT_ID, PRODUCT_ID, QUANTITY, QUANTITY, MERGE_ADD
'        End If
'    Else
'        eprint "UpdateDB() Undefined operation attempted."
'    End If
'    'updateInOutReportTable inCount, outCount, note  'TODO: impl table::getSum(field)
'    clearWorkTable askUser:=False
'End Sub
'
'Public Sub updateInOutReportTable(ByVal inCount As Long, ByVal outCount As Long, ByVal note As String)
'    'should I just use addRow()?
'    'dbBalance.Update
'    'dbReportCursor(REPORT_DATE) = CStr(Now())
'    'dbReportCursor(REPORT_IN) = inCount
'    'dbReportCursor(REPORT_OUT) = outCount
'    'If (reportTableNote <> "") Then
'    '    dbReportCursor(REPORT_NOTE) = reportTableNote
'    'End If
'    'dbReportCursor(DB_PRODUCT_BALANCE) = getInventoryCount()
'
'    'dbReportCursor.update
'
'End Sub
'
'' get current count size of balance.
'Public Function getInventoryCount()
'    balanceQuery.refresh
'    'getInventoryCount = balanceQuery.getCell(0, 0)  '(DB_PRODUCT_BALANCE)
'    getInventoryCount = balanceQuery.getValue(DB_PRODUCT_BALANCE)
'
'    If IsNull(getInventoryCount) Then
'        getInventoryCount = 0
'    End If
'End Function
'
'Public Function getOperationString(ByVal op As operationMode) As String
'    Dim returnString As String
'    Select Case op
'    Case INVENTORY_IN
'        returnString = "In"
'    Case INVENTORY_OUT
'        returnString = "Out"
'    Case VIEW_INVENTORY
'        returnString = "View Inventory"
'    Case EDIT_INVENTORY
'        returnString = "Edit Inventory"
'    Case SALE_IN
'        returnString = "Sale In"
'    Case SALE_OUT
'        returnString = "Sale Out"
'    Case Else
'        eprint "getOperationEnumString(): this operationMode not implemented"
'    End Select
'    getOperationString = returnString
'End Function
'
'
'Public Function getCurrentOperationString() As String
'    getCurrentOperationString = getOperationString(currentOperation)
'End Function
'
'Public Sub addNewSalesperson(ByVal name As String)
'        'creating a new salesperson
'    salespersonTable.addRow SALESPERSON_NAME, name
'    Dim id, tableName
'
'    id = salespersonTable.getValue(SALESPERSON_ID)
'    tableName = SALESPERSON_INVENTORY_TABLE_PREFIX & CStr(id)
'                        ' tableName can be out23,out45,etc.
'    salespersonTable.setValue SALESPERSON_INVENTORY_TABLE_NAME, _
'                                tableName
'    salespersonTable.commit
'
'
'
'    ' Create Inventory table for this salesperson
'    Dim tempTable As New TheAdoTable2
'    tempTable.createDuplicate tableName, _
'                                PRODUCT_TABLE, samsDb
'    tempTable.closeFile
'
'End Sub
'' should be called only the _Change Evt() in ChooseSalesperson Form
'Public Sub changeSalesperson(Optional id As Long = -1)
'    Dim found As Boolean
'    found = True
'
'    salespersonInventoryTable.closeFile
'
'    If id >= 0 Then 'FIND THE RECORD
'        found = salespersonTable.findRecord(SALESPERSON_ID, "=", id)
'    End If
'    If salespersonTable.getSize() < 1 Or found = False Then
'        Salesperson.name = ""
'        Salesperson.id = -1
'        Salesperson.out = 0
'    Else    'GET CURRENT SALESPERSON
'        'If salespersonTable.getSize() = 1 Then
'        '    salespersonTable.gotoFirst
'        'End If
'        Salesperson.name = salespersonTable.getValue(SALESPERSON_NAME)
'        Salesperson.id = salespersonTable.getValue(SALESPERSON_ID)
'        Salesperson.out = salespersonTable.getValue(SALESPERSON_OUT, 0)
'        Dim tempTableName
'        tempTableName = salespersonTable.getValue( _
'                                    SALESPERSON_INVENTORY_TABLE_NAME)
'        If Not IsNull(tempTableName) Then
'            Salesperson.tableName = tempTableName
'            'salespersonInventoryTable.closeFile 'close old file
'            salespersonInventoryTable.openFile "SELECT * FROM " & _
'                        Salesperson.tableName, samsDb
'        End If
'
'    End If
'End Sub
'
''True if delete operation is possible. It isn't an indication that delete operation
''   was successful.
'' False if delete operation can't be done because user has inventory items.
'Public Function deleteSalesperson() As Boolean
'    If salespersonTable.getSize() <= 0 Then
'        deleteSalesperson = True
'            'it is T because deleting an empty rec is always true.
'            ' ie 0 - 0 = 0. It won't let the user of this func flip-out
'            ' by putting out error message either.
'        Exit Function   ' no record to delete
'    End If
'
'    If Salesperson.out > 0 Then
'        deleteSalesperson = False
'        ' CAN'T delete a salesperson when he/she has some items in the inventory.
'        Exit Function
'    End If
'    Dim oldTable As String
'    oldTable = Salesperson.tableName
'    salespersonTable.remove
'    salespersonTable.gotoFirst
'    changeSalesperson
'    If oldTable <> "" Then
'        samsDb.dropTable oldTable
'    End If
'    deleteSalesperson = True
'End Function
'
'
'Public Function hasUnfinishedWork() As Boolean
'    'table is already empty.
'    If ConfigTable.getConfig(CONFIG_WORK_TABLE_SIZE) <= 0 Then
'        hasUnfinishedWork = False
'        Exit Function
'    End If
'
'    'table is active. What was the last activity?
'    Dim previousOperation As operationMode
'    previousOperation = ConfigTable.getConfig(CONFIG_WORK_TABLE_OPERATION)
'    If previousOperation <> currentOperation Then
'        hasUnfinishedWork = True
'        MsgBox "You still have unfinished " & getOperationString(previousOperation)
'    Else
'        hasUnfinishedWork = False   'same operation. Resume.
'    End If
'
'End Function
'
'Public Function getPastSalespersonId() As Long
'    getPastSalespersonId = ConfigTable.getConfig(CONFIG_WORK_TABLE_SALESPERSON_ID)
'End Function


















'=============================================================================
'   OLD CODE
'=============================================================================
#If OLD_CODE Then

'Public Const INVENTORY_OUT_WORK_TABLE = "inventoryOutWork"
    'temp table to hold the Out items.
'Public Const SALES_OUT_WORK_TABLE_PREFIX = "salesOutWork"
'Public Const SALES_IN_WORK_TABLE_PREFIX = "salesInWork"
    ' temp Table to hold each salesperson's In/Out Temporary tables
    
'Public Const IN_OUT_REPORT_FILENAME = "inOutReport.xls"

'Public dbPath As String 'relative path to MDB file
    ' can be changed from configMenu

'Public inOutReportReadOnlyTable As New TheAdoTable2
    ' for grid. Simplified, only 4 fields are visible.

'Public dbCursor As New adodb.RecordSet          'DELETEME: products table
'Public dbReportCursor As New adodb.RecordSet    ' report table.
'Public dbSalespersonCursor As New adodb.RecordSet
'Public dbBalance As New adodb.RecordSet     ' Count of all items in db
'Public salespersonTable As New adodb.RecordSet

'Public Const DEFAULT_DB_PATH = "sams.mdb"   'DELETEME:
'
'
'Public Sub setConfig(ByVal key, ByVal value)
'    ConfigTable.findRecord CONFIG_KEY, "=", key
'    ConfigTable.findRecord CONFIG_VALUE, "=", value
'End Sub
'
'
''in initDb()
'    dbReportCursor.cursorLocation = adUseClient
'    dbReportCursor.cursorType = adOpenKeyset
'    dbReportCursor.lockType = adLockPessimistic
'    dbReportCursor.Open "select [date],[in],[out],[balance],[note] from report", samsDb
'
'
'Sub tempsub()
'      Dim keys, i
'    keys = inventoryCount.dict.keys()
'    For Each i In keys
'        outputText = ""
'        found = findDb(i)
'        count = inventoryCount.getCount(i)
'        If currentOperation = INVENTORY_IN Then 'Add'
'            'outputText = "+"
'            If found Then
'                dbCursor(QUANTITY_IN_STOCK).value = _
'                    dbCursor(QUANTITY_IN_STOCK).value + _
'                    count
'
'            Else
'                ' Create a new item.
'                dbCursor.AddNew
'                dbCursor(PRODUCT_ID).value = i
'                dbCursor(QUANTITY_IN_STOCK).value = _
'                    count
'                outputText = outputText & "Adding New Product "
'                'excelSheet.setCell 3, excelCurrentRow, "Added new product"
'            End If
'            dbCursor.Update
'        Else    '- sub'
'            'outputText = "-"
'            If found Then
'                Dim tempNumber
'                tempNumber = dbCursor(QUANTITY_IN_STOCK).value - _
'                    count
'                If tempNumber >= 0 Then
'                    dbCursor(QUANTITY_IN_STOCK).value = tempNumber
'                Else
'                    dbCursor(QUANTITY_IN_STOCK).value = 0
'                    outputText = outputText & "Error: Negative Number"
'                    ' ERROR: can't go below 0. TODO: Warn the user?
'                End If
'                dbCursor.Update
'            Else
'                ' TODO: item not found. Should I warn the user?
'            End If
'        End If
'        outputText = outputText & i & str(count)
'        'excelSheet.setCell 1, excelCurrentRow, i
'        'excelSheet.setCell 2, excelCurrentRow, count
'        'excelCurrentRow = excelCurrentRow + 1
'        'statusList.addItem outputText
'        'statusList.TopIndex = statusList.NewIndex
'    Next
'
'    closeStatusForm.Enabled = True
'
'End Sub
'
'Public Sub updateReportTable_USING_MDBTABLE_NOT_OLEXLS(ByVal inCount As Long, ByVal outCount As Long)
'    'dbBalance.Update
'
'    dbReportCursor.AddNew
'    dbReportCursor(REPORT_DATE) = CStr(Now())
'    dbReportCursor(REPORT_IN) = inCount
'    dbReportCursor(REPORT_OUT) = outCount
'    'If (currentOperation = "saleOut" Or currentOperation = "saleIn") Then
'    If (reportTableNote <> "") Then
'        dbReportCursor(REPORT_NOTE) = reportTableNote
'    End If
'    dbReportCursor(DB_PRODUCT_BALANCE) = getInventoryCount()
'
'    dbReportCursor.Update
'
'End Sub
'Sub initDb_OLD()
'
'    If Dir(dbPath) = "" Then
'        MsgBox "WARNING: Database file " & dbPath & " was not found. " _
'            & "This application will not function properly without it."
'        Exit Sub
'        ' Allow user to fix the dbpath via Config menu.
'    End If
'
'    samsDb.openDb App.path() & "\" & DB_FILENAME
'    samsDb.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" _
'            & "Data Source=" & dbPath & ";Persist Security Info=False;" _
'            & "Mode=Read|Write"
'    samsDb.Open
'    ' MsgBox samsDb.State
'
'    dbCursor.cursorLocation = adUseClient
'    dbCursor.cursorType = adOpenKeyset
'    dbCursor.lockType = adLockPessimistic
'    dbCursor.Open "select * from products", samsDb
'
'
'    dbSalespersonCursor.cursorLocation = adUseClient
'    dbSalespersonCursor.cursorType = adOpenKeyset
'    dbSalespersonCursor.lockType = adLockPessimistic
'    dbSalespersonCursor.Open "select [name],[out],[tableName],[id] from salesperson", samsDb
'
'    'Debug.Print dbCursor(0)
'    'Debug.Print dbCursor(1)
'    'TODO: dbCursor.MoveFirst
'
'    dbBalance.Open "select sum(" & QUANTITY_IN_STOCK & ") as " & DB_PRODUCT_BALANCE & " from products", samsDb
'
'    'salespersonTable.openTable "select [name],[out],[tableName],[id] from salesperson", samsDb
'End Sub
'' update Products table and Report table.
'' updateInOutReport
'' Returns # of items that were either placed in the "In" or the "Out" box.
'Public Function updateDb(Optional sourceTable, Optional updateInOutReport As Boolean = True)
'    Dim table As New TheTable
'    If IsMissing(sourceTable) Then
'        table.add dataWindow.productGrid
'    Else
'        table.add sourceTable
'    End If
'    Dim inCount, outCount   'in,out count for this report
'
'    'update db
'    Dim i, found As Boolean, count As Long, productId As String
'    'Dim keys,
'    'keys = inventoryCount.dict.keys()
'    For i = 0 To table.getRowSize() - 1
'        productId = table.getCell(0, i)
'        count = table.getCell(1, i)     ' count
'        found = findDb(productId)  'look productID
'        If currentOperation = INVENTORY_IN Then 'Add'
'            If found Then
'                Dim newValue As Long
'                newValue = productTable.getValueAtColumn(QUANTITY) + count
'                productTable.setValueAtColumn QUANTITY, newValue
'                ' productTable.update ' automatically done.
'            Else
'                ' Create a new item.
'                productTable.addRow productId, count
'            End If
'            inCount = inCount + count
'        ElseIf currentOperation = INVENTORY_OUT Or currentOperation = SALE_OUT Then 'sub'
'            If found Then
'                Dim tempNumber
'                tempNumber = productTable.getValueAtColumn(QUANTITY) - count
'                productTable.setValueAtColumn QUANTITY, Max(tempNumber, 0)
'                    ' ERROR: can't go below 0. TODO: Warn the user?
'                'End If
'                'dbCursor.update 'redundant using setValueAtColumn()
'            Else
'                ' TODO: item not found. Should I warn the user?
'            End If
'            outCount = outCount + count
'        Else 'View mode
'            ' Do nothing
'        End If
'        ' outputText = outputText & i & Str(count)
'
'    Next
'    ' generate Report Table
'    updateReportTable inCount, outCount
'    updateDb = Max(inCount, outCount)
'End Function
'Private Function hasUnfinishedSalespersonWork()
'    If Salesperson.id <> pastSalesperson.id Then
'         hasUnfinishedSalespersonWork = True
'         MsgBox "You still have unfinished " & _
'            getOperationString(ConfigTable.getConfig(CONFIG_WORK_TABLE_OPERATION)) & _
'            " for " & pastSalesperson.name
'    Else
'         hasUnfinishedSalespersonWork = False
'    End If
    
End Function

#End If
