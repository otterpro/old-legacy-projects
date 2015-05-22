VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "InventoryControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


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


'Public currentOperation As String '"in","out",or "view" , "saleIn" , "saleOut"mode
Public Enum FIELD_NAME_ENUM
    INDEX_NUMBER    ' "index"
    PRODUCT_QUANTITY        ' "quantity
    PRODUCT_NAME    'formerly PRODUCT_ID "id"
    PRODUCT_DESC    '"note" or "desc".
    LAST_MODIFIED   '"lastModified"
    PRODUCT_NUMBER
    LAST_OPERATION
    LAST_SALESPERSON_ID
    WORK_TABLE_SIZE
    CONFIG_TABLE_KEY
    CONFIG_TABLE_VALUE
End Enum

Public Enum OPERATION_MODE
    INVENTORY_IN
    INVENTORY_OUT
    VIEW_INVENTORY
    EDIT_INVENTORY
    SALE_IN
    SALE_OUT
End Enum
Public currentOperation As OPERATION_MODE
Public WithEvents productTable As TheAdoTable2
Attribute productTable.VB_VarHelpID = -1
Public WithEvents salespersonTable As TheAdoTable2
Attribute salespersonTable.VB_VarHelpID = -1
Public salespersonInventoryTable As New TheAdoTable2
Public balanceQuery As New TheAdoTable2
Public inOutReportTable As New TheAdoTable2
Public workTable As New TheAdoTable2
Public ProductEditTable As New TheAdoTable2
Public ProductTypeTable As New TheAdoTable2

Private Type SalespersonType
    name As String ' * SALESPERSON_NAME_MAX_SIZE    '
    id As Long          '
    tableName As String 'inventory table name
    out As Long         ' # of items out.
End Type

'Public multiplier As Long
    ' multiplies the quantity by this value when doing inventory
    
' DB and Tables
Private myDb As TheAdoDb
Private Const SALESPERSON_NAME_MAX_SIZE = 64
    'workTable is used as a temp table for keeping track of user's data-entry.
    
'external db/xls/ole filename
'Private Const DB_FILENAME = "sams.mdb"

'Public Const IN_OUT_REPORT_FILENAME_TEST = "inOutReport.ole"

'names of tables in sams.mdb
Private Const SALESPERSON_INVENTORY_TABLE_PREFIX = "salesInventory"
Private Const PRODUCT_TABLE = "product"
Private Const WORK_TABLE = "inventoryWork"
    'temp table to hold the In/Out/etc item
'Public Const PRODUCT_TABLE_TEMPLATE = "productTemplate"
    'just use WORK_TABLE.
Private Const IN_OUT_REPORT_FILENAME_PREFIX = "inout"
    'ex: inout1.xls, inout2.xls, inout3.xls, ...
Private Const IN_OUT_REPORT_TABLE = "inOutReport"
Private Const IN_OUT_REPORT_FOLDER = "inOutReport"
    ' all inout*.xls is stored in <appPath>/inOutReport/ folder.

' PRODUCT TABLE
Private Const INDEX_FIELD = "index"        'autonumber field
Private Const QUANTITY_FIELD = "quantity"
Private Const PRODUCT_ID_FIELD = "id"      ' "RG-1000/A"
Private Const DESCRIPTION_FIELD = "note"   ' "This is a turquoise ring"
Private Const LAST_MODIFIED_FIELD = "lastModified"
Private Const PRODUCT_NUMBER_FIELD = "productNumber"   ' 1000
Private Const BALANCE_FIELD = "balance"



' Currently Not used
Private Const PRODUCT_TYPE = "type"  '

Private Const REPORT_ID = "id"
Private Const REPORT_DATE = "date"
Private Const REPORT_NOTE = "note"
Private Const REPORT_IN = "in"
Private Const REPORT_OUT = "out"
Private Const REPORT_SALESPERSON = "salesperson"
Private Const REPORT_SHEET_FILENAME = "sheetFilename"
'Private Const DB_PRODUCT_BALANCE = "balance"

' WORK TABLE have same structure as PRODUCT TABLE
Private Const WORK_INDEX_FIELD = "index"
Private Const WORK_PRODUCT_ID = "id"
Private Const WORK_QUANTITY = "quantity"
Private Const WORK_NOTE = "note"

'SALESPERSON TABLE
Private Const SALESPERSON_TABLE = "salesperson"
Private Const SALESPERSON_OUT = "out"
Private Const SALESPERSON_ID = "id"
Private Const SALESPERSON_NAME = "name"
Private Const SALESPERSON_INVENTORY_TABLE_NAME = "tableName"

Private Const CONFIG_TABLE = "config"
Private Const CONFIG_WORK_TABLE_OPERATION = "workTableOperation"
    'in,out,saleIn,saleOut (see OperationMode and currentOperation)
    ' is in numbercode (enumerated)?
Private Const CONFIG_WORK_TABLE_SALESPERSON_ID = "workTableSalesperson"
    'id # of salesperson if the workTable operation is saleIn or saleOut
Private Const CONFIG_WORK_TABLE_SIZE = "workTableSize"
Private Const CONFIG_KEY_FIELD = "key"
Private Const CONFIG_VALUE_FIELD = "value"

'Public Salesperson As String
    ' currently selected salesperson, chosen in chooseSalespersonDialog. Used for In/Sales, Out/sales
'Public SalespersonID As Long
    ' currently selected salesperson's id. Convenient to store it in a var.
'Public SalespersonInventoryTableName As String

'
Private Const PRODUCT_TYPE_TABLE = "productType"
Private Const PRODUCT_TYPE_PREFIX = "prefix"

Private Const ERROR_CODE = 1002 'used to raise error

Private Salesperson As SalespersonType   'current salesperson
                    'salesperson can be selected from In-Sales, Out-Sales

' CONFIG TABLE
Public ConfigTable As New TheAdoTable2

    ' the size of work table the last time it was being used.
    ' if 0, there is no need to worry about checking which operation we're doing.

Public importExcelTable As New TheTable
    ' if the user imports excel sheet for "in", it stores the data here temporarily
    ' until it is retrieved by the dataWindow form.

'Dialog related
'Public ChooseSalesperson As Boolean
    'set by ChooseSalespersonDialog. If the user successfully chooses
    'a salesperson, it is true. If not, it is false.

Public Function getFieldName(ByVal fieldName As FIELD_NAME_ENUM) As String
    Dim result As String
    Select Case fieldName
    Case INDEX_NUMBER
        result = INDEX_FIELD
    Case PRODUCT_QUANTITY
        result = QUANTITY_FIELD
    Case PRODUCT_NAME
        result = PRODUCT_ID_FIELD
    Case PRODUCT_DESC
        result = DESCRIPTION_FIELD
    Case LAST_MODIFIED
        result = LAST_MODIFIED_FIELD
    Case PRODUCT_NUMBER
        result = PRODUCT_NUMBER_FIELD
    Case LAST_OPERATION
        result = CONFIG_WORK_TABLE_OPERATION
    Case LAST_SALESPERSON_ID
        result = CONFIG_WORK_TABLE_SALESPERSON_ID
    Case WORK_TABLE_SIZE
        result = CONFIG_WORK_TABLE_SIZE
    Case CONFIG_TABLE_KEY
        result = CONFIG_KEY_FIELD
    Case CONFIG_TABLE_VALUE
        result = CONFIG_VALUE_FIELD

    End Select
    
    getFieldName = result
End Function


Public Property Get SalespersonOut() As Long
    SalespersonOut = Salesperson.out
End Property

Public Property Get salespersonName() As String
    salespersonName = Salesperson.name
End Property

Public Property Get SalespersonId() As Long
    SalespersonId = Salesperson.id
End Property

Public Sub openFile(ByVal dbPath As String)
    Dim tempName As String
    tempName = PRODUCT_TABLE
    tempName = SALESPERSON_TABLE
    
    Set myDb = New TheAdoDb
    myDb.openFile dbPath
    Set productTable = New TheAdoTable2
    Dim result As Boolean
    result = productTable.openFile("SELECT * FROM " & PRODUCT_TABLE, myDb)
    Debug.Assert (result)
    
    Set salespersonTable = New TheAdoTable2
    result = salespersonTable.openFile("SELECT * FROM " & SALESPERSON_TABLE, myDb)
    
    'result = salespersonTable.openFile(joinField("select", sqlField(SALESPERSON_NAME,
    '    SALESPERSON_OUT, SALESPERSON_INVENTORY_TABLE_NAME, SALESPERSON_ID), "from", SALESPERSON_TABLE), myDb)
    Debug.Assert (result)
    result = balanceQuery.openFile(joinField("SELECT sum(", QUANTITY_FIELD, ") as", BALANCE_FIELD, _
                                "FROM ", PRODUCT_TABLE), myDb)
    Debug.Assert (result)
    result = inOutReportTable.openFile("SELECT * FROM " & IN_OUT_REPORT_TABLE, myDb)
    Debug.Assert (result)
    'inOutReportTable.gotoLast
    
    result = ConfigTable.openFile("SELECT * FROM " & CONFIG_TABLE, myDb)
    Debug.Assert (result)
    result = workTable.openFile("SELECT * FROM " & WORK_TABLE, myDb)
    Debug.Assert (result)
    
    result = ProductEditTable.openFile("SELECT * FROM " & PRODUCT_TABLE, myDb)
    Debug.Assert (result)
    
    result = ProductTypeTable.openFile("SELECT * FROM " & PRODUCT_TYPE_TABLE, myDb)
    Debug.Assert (result)
End Sub

Public Function findProduct(ByRef adoTable As TheAdoTable2, ByVal productId As String) As Boolean
    findProduct = adoTable.findRecord(PRODUCT_ID_FIELD, "=", productId)
End Function

Public Function findDb(ByVal productId)
    findDb = productTable.findRecord(PRODUCT_ID_FIELD, "=", productId)
End Function

Public Function findInSalespersonTable(ByVal id) As Boolean
    findInSalespersonTable = salespersonInventoryTable.findRecord(PRODUCT_ID_FIELD, "=", id)
End Function

' save as excel sheet and update data
Public Sub saveExcelSheet()
    Dim sheet As New TheExcelTable
    Dim table As New TheTable
    sheet.openFile visible:=False
    table.addRow PRODUCT_ID_FIELD, QUANTITY_FIELD, REPORT_NOTE
    'table.add dataWindow.productGrid
    'sheet.copy table
    TheTableBase.copyTo workTable, table
    sheet.add table
    sheet.setVisible
End Sub


Public Sub clearWorkTable(Optional askUser As Boolean = True)
    If workTable.getSize() > 0 And askUser Then
        If vbOK <> MsgBox("Click OK to erase everything. Click Cancel to go back.", vbOKCancel + vbQuestion, _
                        "Are you sure you want to delete the current sheet?") Then
            Exit Sub    'NO. User doesn't want to erase.
        End If
    End If
    workTable.clear 'erase the content.
    ConfigTable.setConfig CONFIG_WORK_TABLE_SIZE, 0 ' save 0 to config file.
End Sub

' update Products table and Report table.
' updateInOutReport
' Returns # of items that were either placed in the "In" or the "Out" box.
Public Sub updateDb()  'Optional sourceTable, Optional updateInOutReport As Boolean = True)
    Dim sqlStmt As String
    Dim salespersonInventoryTable As New TheAdoTable2
    
    'IN-Out Report Table datas
    Dim inCount As Long, outCount As Long, note As String
    
    If workTable.getSize() < 1 And currentOperation <> SALE_IN Then
        Exit Sub    'work table is empty. Don't do anything.
    End If
    
    'saveExcelSheet  ' TODO:Show Excel Spreadsheet
    
    If currentOperation = INVENTORY_IN Then
        'productT += workT
        note = "In"
        workTable.commit
        inCount = workTable.getSum(QUANTITY_FIELD)
        productTable.merge workTable, PRODUCT_ID_FIELD, PRODUCT_ID_FIELD, QUANTITY_FIELD, _
                                    QUANTITY_FIELD, MERGE_ADD
        
    ElseIf currentOperation = INVENTORY_OUT Then
        'productT -= workT
        note = "Out"
        workTable.commit
        outCount = workTable.getSum(QUANTITY_FIELD)
        productTable.merge workTable, PRODUCT_ID_FIELD, PRODUCT_ID_FIELD, QUANTITY_FIELD, _
                                    QUANTITY_FIELD, MERGE_SUBTRACT
    ElseIf currentOperation = SALE_IN Or currentOperation = SALE_OUT Then
        'open salesperson's inventory Table
        sqlStmt = "SELECT * FROM " & salespersonTable.getValue( _
                                        SALESPERSON_INVENTORY_TABLE_NAME)
        salespersonInventoryTable.openFile sqlStmt, myDb
        If currentOperation = SALE_IN Then  'SOLD
            ' formula:
            '1. salespersonT -= workTable  # of items sold
            '2. productT += salespersonT   # of returned items to inventory
            note = "Sale-In: " & Salesperson.name
            salespersonTable.setValue SALESPERSON_OUT, 0
            salespersonTable.commit
            workTable.commit
            outCount = workTable.getSum(QUANTITY_FIELD)  'SOLD
            salespersonInventoryTable.merge workTable, PRODUCT_ID_FIELD, PRODUCT_ID_FIELD, _
                                            QUANTITY_FIELD, QUANTITY_FIELD, MERGE_SUBTRACT
            inCount = salespersonInventoryTable.getSum(QUANTITY_FIELD)
            productTable.merge salespersonInventoryTable, PRODUCT_ID_FIELD, PRODUCT_ID_FIELD, _
                                            QUANTITY_FIELD, QUANTITY_FIELD, MERGE_ADD
            'Warning: Merge will change the position! So do this at the latest moment
            
            salespersonInventoryTable.clear
            salespersonInventoryTable.gotoFirst
        Else    ' SALE_OUT, append to salesperson's out table
            ' salespersonT += workTable
            ' productT -=workTable, OUT
            note = "Sale-Out: " & Salesperson.name
            workTable.gotoFirst
            outCount = workTable.getSum(QUANTITY_FIELD)
            salespersonInventoryTable.merge workTable, PRODUCT_ID_FIELD, PRODUCT_ID_FIELD, _
                                                QUANTITY_FIELD, QUANTITY_FIELD, MERGE_ADD
            Dim outTotal As Long  'updated quantity after adding current out.
            'newOut = salespersonTable.getValue(SALESPERSON_OUT, 0)
            outTotal = salespersonInventoryTable.getSum(QUANTITY_FIELD)
            'newOut = outCount
            'newOut = newOut + workTable.getSum(QUANTITY_FIELD, cache:=False)
            salespersonTable.setValue SALESPERSON_OUT, outTotal
            salespersonTable.commit
            productTable.merge salespersonInventoryTable, PRODUCT_ID_FIELD, PRODUCT_ID_FIELD, _
                                            QUANTITY_FIELD, QUANTITY_FIELD, MERGE_SUBTRACT
            'Warning: Merge will change the position! So do this at the latest moment
                       
        End If
        changeSalesperson
    Else
        eprint "UpdateDB() Undefined operation attempted."
    End If
    updateInOutReportTable inCount, outCount, note  'TODO: impl table::getSum(field)
    clearWorkTable askUser:=False
    ProductEditTable.refresh
    
End Sub

Public Sub updateInOutReportTable(Optional inCount As Long = 0, Optional outCount As Long = 0, _
                                Optional note As String = "", Optional salespersonName = 0)
    Dim keys, values
    keys = Array(REPORT_DATE, REPORT_IN, REPORT_OUT, REPORT_NOTE, REPORT_SALESPERSON)
    values = Array(Now(), inCount, outCount, note, salespersonName)
    
    inOutReportTable.addRow keys, values
    inOutReportTable.commit
    
End Sub

' get current balance.
Public Function getInventoryCount()
    balanceQuery.refresh
    'getInventoryCount = balanceQuery.getCell(0, 0)  '(DB_PRODUCT_BALANCE)
    getInventoryCount = balanceQuery.getValue(BALANCE_FIELD)
    
    If IsNull(getInventoryCount) Then
        getInventoryCount = 0
    End If
End Function

Public Function getOperationString(ByVal op As OPERATION_MODE) As String
    Dim returnString As String
    Select Case op
    Case INVENTORY_IN
        returnString = "In"
    Case INVENTORY_OUT
        returnString = "Out"
    Case VIEW_INVENTORY
        returnString = "View Inventory"
    Case EDIT_INVENTORY
        returnString = "Edit Inventory"
    Case SALE_IN
        returnString = "Sale In"
    Case SALE_OUT
        returnString = "Sale Out"
    Case Else
        eprint "getOperationEnumString(): this operationMode not implemented"
    End Select
    getOperationString = returnString
End Function


Public Function getCurrentOperationString() As String
    getCurrentOperationString = getOperationString(currentOperation)
End Function


Public Sub addNewSalesperson(ByVal name As String)
        'creating a new salesperson
    salespersonTable.addRow SALESPERSON_NAME, name
    Dim id, tableName
    
    id = salespersonTable.getValue(SALESPERSON_ID)
    tableName = SALESPERSON_INVENTORY_TABLE_PREFIX & CStr(id)
                        ' tableName can be out23,out45,etc.
    salespersonTable.setValue SALESPERSON_INVENTORY_TABLE_NAME, _
                                tableName
    salespersonTable.commit

    
    
    ' Create Inventory table for this salesperson
    Dim tempTable As New TheAdoTable2
    tempTable.createDuplicate tableName, _
                                PRODUCT_TABLE, myDb
    tempTable.closeFile
    
End Sub
' should be called only the _Change Evt() in ChooseSalesperson Form
Public Sub changeSalesperson(Optional id As Long = -1)
    Dim found As Boolean
    found = True
    
    salespersonInventoryTable.closeFile
    
    If id >= 0 Then 'FIND THE RECORD
        found = salespersonTable.findRecord(SALESPERSON_ID, "=", id)
    End If
    If salespersonTable.isEof() Then
        salespersonTable.gotoFirst
    End If
    If salespersonTable.getSize() < 1 Or found = False Then
        Salesperson.name = ""
        Salesperson.id = -1
        Salesperson.out = 0
    Else    'GET CURRENT SALESPERSON
        'If salespersonTable.getSize() = 1 Then
        '    salespersonTable.gotoFirst
        'End If
        Salesperson.name = salespersonTable.getValue(SALESPERSON_NAME)
        Salesperson.id = salespersonTable.getValue(SALESPERSON_ID)
        Salesperson.out = salespersonTable.getValue(SALESPERSON_OUT, 0)
        Dim tempTableName
        tempTableName = salespersonTable.getValue( _
                                    SALESPERSON_INVENTORY_TABLE_NAME)
        If Not IsNull(tempTableName) Then
            Salesperson.tableName = tempTableName
            'salespersonInventoryTable.closeFile 'close old file
            salespersonInventoryTable.openFile "SELECT * FROM " & _
                        Salesperson.tableName, myDb
        End If

    End If
End Sub

'True if delete operation is possible. It isn't an indication that delete operation
'   was successful.
' False if delete operation can't be done because user has inventory items.
Public Function deleteSalesperson() As Boolean
    If salespersonTable.getSize() <= 0 Then
        deleteSalesperson = True
            'it is T because deleting an empty rec is always true.
            ' ie 0 - 0 = 0. It won't let the user of this func flip-out
            ' by putting out error message either.
        Exit Function   ' no record to delete
    End If
    
    If Salesperson.out > 0 Then
        deleteSalesperson = False
        ' CAN'T delete a salesperson when he/she has some items in the inventory.
        Exit Function
    End If
    Dim oldTable As String
    oldTable = Salesperson.tableName
    salespersonTable.remove
    salespersonTable.gotoFirst
    changeSalesperson
    If oldTable <> "" Then
        myDb.dropTable oldTable
    End If
    deleteSalesperson = True
End Function


Public Function hasUnfinishedWork() As Boolean
    'table is already empty.
    If ConfigTable.getConfig(CONFIG_WORK_TABLE_SIZE) <= 0 Then
        hasUnfinishedWork = False
        Exit Function
    End If
    
    'table is active. What was the last activity?
    Dim previousOperation As OPERATION_MODE
    previousOperation = ConfigTable.getConfig(CONFIG_WORK_TABLE_OPERATION)
    If previousOperation <> currentOperation Then
        hasUnfinishedWork = True
        MsgBox "You still have unfinished " & getOperationString(previousOperation)
    Else
        hasUnfinishedWork = False   'same operation. Resume.
    End If

End Function

Public Function getPastSalespersonId() As Long
    getPastSalespersonId = ConfigTable.getConfig(CONFIG_WORK_TABLE_SALESPERSON_ID)
End Function

' validates the productId before it is added to the worktable.
' Returns the productID after it is formatted correctly (uppercase, etc)
' If error occurs, it will return "" (empty string) and raise an error.
' Err.Description contains error msg. If it is empty "", then it is a
' non-critical error and the calling procedure should exit without displaying any msg.
Public Function validateProductId(ByVal productId As String) As String
    
    Dim itExistsOnTheProductTable As Boolean
    itExistsOnTheProductTable = False
    Dim functionName As String
    'On Error GoTo 0
    functionName = "InventoryControl.validateProductId()"
    ' DO Validation. Check spelling, etc.
    productId = stripIllegalChar(productId, "#;" & QUOTE_CHARS) ' remove [#:"']
    productId = the.strip(productId) 'productId = lstrip(productId)
    If productId = "" Then
        Err.Raise ERROR_CODE, functionName, "" 'user just entered an empty line "Return key"
        Exit Function
    End If
    productId = UCase(getFirstWord(productId)) 'ignore trailing word. Just get the 1st part.
    If Len(productId) = 1 Then
        Err.Raise ERROR_CODE, functionName, _
                            "Very short product ID was entered.Please try longer ID."
        Exit Function
    End If
    validateProductId = productId ' put it here instead of at the bottom because
        'this func can exit at any moment if the error is raised. The calling func
        'is expecting a returned value and may accidentally overwrite its previous var.
    ' make sure that the product id is appropriate for current operation.
    Dim valueList As TheList
    Dim currentQuantity As Long, workItemQuantity As Long
    
    itExistsOnTheProductTable = findDb(productId)
    If (workTable.findRecord(PRODUCT_ID_FIELD, "=", productId)) Then
        workItemQuantity = workTable.getValue(QUANTITY_FIELD, 0)
        workItemQuantity = workItemQuantity + 1     '++ not in VB, only VB.net
            ' add 1 to include this productId as well as the ones already on the work Table
    End If
    If itExistsOnTheProductTable Then
        currentQuantity = productTable.getValue(QUANTITY_FIELD, 0)
    End If
    Select Case currentOperation
    Case INVENTORY_OUT  'usually items sold on Show
        If Not itExistsOnTheProductTable Then    'make sure item exists in the inventory
            Err.Raise ERROR_CODE, functionName, "The item " & productId & _
                                " is not in the inventory. You cannot take this item out. " & _
                                "Please do an 'In' for this item first."
            Exit Function
        ElseIf workItemQuantity > currentQuantity Then
            Err.Raise ERROR_CODE, functionName, "There are only item " & CStr(currentQuantity) & _
                " of " & productId & ". You cannot take out more than is available in the inventory."
            Exit Function
        
        End If
    Case INVENTORY_IN   'new inventory item and returned items...
        'Nothing to do.
    Case SALE_IN    'person is entering what was sold, not what was placed In.
        Dim salespersonQuantity As Long
        Debug.Assert (Not IsNull(salespersonInventoryTable))
        
        If Not findInSalespersonTable(productId) Then   'make sure salesperson has this item.
            Err.Raise ERROR_CODE, functionName, "The salesperson does not have any " & _
                                    productId & ". You cannot take this item out."
            Exit Function
        End If
        
        salespersonQuantity = salespersonInventoryTable.getValue(QUANTITY_FIELD, 0)
        If workItemQuantity > salespersonQuantity Then
            Err.Raise ERROR_CODE, functionName, "You are trying sell items that salesperson" & _
                " doesn't have enough of. " & _
                "Currently the salesperson only has taken out " & _
                CStr(salespersonQuantity) & "."
            Exit Function
        End If
    Case SALE_OUT   'add to current salesperson's inventory.
        If Not itExistsOnTheProductTable Then   'make sure item exists in the inventory
            Err.Raise ERROR_CODE, functionName, "The item " & productId & _
                                    " is not in the inventory. Cannot take this item out."
            Exit Function
        ElseIf workItemQuantity > currentQuantity Then
             Err.Raise ERROR_CODE, functionName, "There are only " & CStr(currentQuantity) & _
                " of " & productId & ". You cannot take out more than what is available in the inventory."
            Exit Function
        End If
    End Select
    'validateProductId = productId 'Moved Up (See above somewhere)
End Function

Public Function productTypeIsValid(ByVal productId As String) As Boolean
    Dim prefix As String
    prefix = Left$(productId, 3)    'get 1st 3 char ie "RG-", "ER-", "NL-", etc.
    ' Require that Product has "-". Like "RG-1000", "ER-1000/a", etc.
    ' Checking findRecord() will find it since "-" is already included as the index name.
    'If the.getCharAt(productId, 2) <> "-" Then
    '    productTypeIsValid = False
    '    Exit Function
    'End If
    If (ProductTypeTable.findRecord(PRODUCT_TYPE_PREFIX, "=", prefix) = False) Then
        productTypeIsValid = False
    Else
        productTypeIsValid = True
    End If
End Function


Public Sub addNewWorkItem(ByVal productId As String)
    Dim productNumber As Long
    
    If findProduct(workTable, productId) Then 'ADD ITEM THAT ALREADY EXISTS ON THE TABLE -------
        incCurrentWorkItem
    Else
        productNumber = TheRegExpTool.getNumber(productId)
        inventory.workTable.addRow Array(WORK_PRODUCT_ID, WORK_QUANTITY, LAST_MODIFIED, _
                                PRODUCT_NUMBER_FIELD), Array(productId, 1, Now(), productNumber)
    End If        '

End Sub
Public Sub incCurrentWorkItem(Optional incBy As Long = 1)
    Dim newQuantity As Long
    newQuantity = inventory.workTable.getValue(WORK_QUANTITY) + incBy 'we're adding
    workTable.setValue WORK_QUANTITY, newQuantity
    workTable.setValue LAST_MODIFIED_FIELD, Now()

End Sub
Private Sub productTable_cursorMoved()
    the.dprint "cursor moved"
    
End Sub

Private Sub salespersonTable_cursorMoved()
    'salespersonInventoryTable.
    changeSalesperson
    the.dprint "salesperson mmoved"
End Sub



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
'    dbReportCursor.Open "select [date],[in],[out],[balance],[note] from report", myDb
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
'    myDb.openDb App.path() & "\" & DB_FILENAME
'    myDb.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" _
'            & "Data Source=" & dbPath & ";Persist Security Info=False;" _
'            & "Mode=Read|Write"
'    myDb.Open
'    ' MsgBox myDb.State
'
'    dbCursor.cursorLocation = adUseClient
'    dbCursor.cursorType = adOpenKeyset
'    dbCursor.lockType = adLockPessimistic
'    dbCursor.Open "select * from products", myDb
'
'
'    dbSalespersonCursor.cursorLocation = adUseClient
'    dbSalespersonCursor.cursorType = adOpenKeyset
'    dbSalespersonCursor.lockType = adLockPessimistic
'    dbSalespersonCursor.Open "select [name],[out],[tableName],[id] from salesperson", myDb
'
'    'Debug.Print dbCursor(0)
'    'Debug.Print dbCursor(1)
'    'TODO: dbCursor.MoveFirst
'
'    dbBalance.Open "select sum(" & QUANTITY_IN_STOCK & ") as " & DB_PRODUCT_BALANCE & " from products", myDb
'
'    'salespersonTable.openTable "select [name],[out],[tableName],[id] from salesperson", myDb
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
    
'End Function

#End If


