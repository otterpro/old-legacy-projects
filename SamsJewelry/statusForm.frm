VERSION 5.00
Begin VB.Form statusForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Updating..."
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton closeStatusForm 
      Caption         =   "Close Window"
      Enabled         =   0   'False
      Height          =   975
      Left            =   240
      TabIndex        =   1
      Top             =   2160
      Width           =   4215
   End
   Begin VB.ListBox statusList 
      Height          =   1815
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4215
   End
End
Attribute VB_Name = "statusForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub closeStatusForm_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    Dim i, keys, found As Boolean, count, _
        outputText As String, _
        excelSheet As New TheExcelApp, _
        excelCurrentRow As Long

    excelCurrentRow = 2 ' start from row2. Row1=heading
    
    excelSheet.setCell 1, 1, "Product ID"
    excelSheet.setCell 2, 1, "Quantity"
    excelSheet.setCell 3, 1, "Note"
    Debug.Print "--- statusForm LOADED---"
    'inventoryCount.dprint
    'Debug.Print "testNumber=" & Str(testNumber)
    keys = inventoryCount.dict.keys()

    For Each i In keys
        outputText = ""
        'CHANGEME, uncomment below line fast
        found = findDb(i)
        count = inventoryCount.getCount(i)
        If currentTabKey = "+" Then 'Add'
            outputText = "+"
            If found Then
                dbCursor(QUANTITY_IN_STOCK).value = _
                    dbCursor(QUANTITY_IN_STOCK).value + _
                    count
                    
            Else
                ' Create a new item.
                dbCursor.AddNew
                dbCursor(PRODUCT_ID).value = i
                dbCursor(QUANTITY_IN_STOCK).value = _
                    count
                outputText = outputText & "Adding New Product "
                excelSheet.setCell 3, excelCurrentRow, "Added new product"
            End If
            dbCursor.Update
        Else    '- sub'
            outputText = "-"
            If found Then
                Dim tempNumber
                tempNumber = dbCursor(QUANTITY_IN_STOCK).value - _
                    count
                If tempNumber >= 0 Then
                    dbCursor(QUANTITY_IN_STOCK).value = tempNumber
                Else
                    dbCursor(QUANTITY_IN_STOCK).value = 0
                    outputText = outputText & "Error: Negative Number"
                    ' ERROR: can't go below 0. TODO: Warn the user?
                End If
                dbCursor.Update
            Else
                ' TODO: item not found. Should I warn the user?
            End If
        End If
        outputText = outputText & i & Str(count)
        excelSheet.setCell 1, excelCurrentRow, i
        excelSheet.setCell 2, excelCurrentRow, count
        excelCurrentRow = excelCurrentRow + 1
        statusList.addItem outputText
        statusList.TopIndex = statusList.NewIndex
    Next
    
    closeStatusForm.Enabled = True
    
End Sub

