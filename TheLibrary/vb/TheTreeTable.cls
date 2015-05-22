VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TheTreeTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=============================================================================
'
'=============================================================================
Option Explicit

'GLOBAL/PUBLIC Member Var/Properties

'TODO: Replace TheAdoDB with TheDB class when TheDB is finished.
Public Function openFile(ByVal openStatement As String, Optional db) As Boolean
End Function

Public Function closeFile() As Boolean
End Function

Public Function find(ByVal whereClause As String) As Boolean
End Function

Public Function commit() As Boolean
End Function

Public Sub refresh()
End Sub

Public Function gotoFirst() As Boolean
End Function

Public Function gotoNext()
End Function

Public Function gotoLast()
End Function

Public Function getSize() As Long
End Function

Public Function getColumnSize() As Long
End Function

Public Function getValue(Optional fieldName, Optional defaultValue) As Variant
End Function

Public Function isEof() As Boolean
End Function

Public Sub setValue(ByVal fieldName As String, ByVal value)
End Sub

Public Function addRow(Optional keys, Optional values) As Boolean
End Function

' goes to the 1st record and returns the result in Collection.
' In Perl's DBI, it is same as doing fetch() right after executing.
' In JDBC, it is same as executing, and then doing next() and then fetch().
Public Function getFirst(Optional fieldName)
End Function

Public Function getNext(Optional fieldName)
End Function

Public Sub clear()
End Sub

Public Function remove() As Boolean
End Function

' Either adds or subtracts "quantity" field from the two table. Used for updating the quantity.
' Should return the # of rows(records) that were affected.
Public Function merge(ByRef foreignTable As TheTreeTable, _
                        ByVal primaryKey As String, ByVal foreignKey As String, _
                        ByVal myField As String, ByVal foreignField As String, _
                        ByVal operationMode As MERGE_OPERATION)
End Function

'    clone (srcTable)    same as .clear() + .copy(srcTable)
'    clone()             clone itself and returns new TheAdoTable
'        makes another view of the same table. Making changes to
'        the cloned table makes changes to the original table that it refers to.
'        NOT_TESTED.Not sure if that is what Recordset.Clone() does.
Public Function clone(Optional srcTable As TheTreeTable) As TheTreeTable
End Function

Public Function copy(srcTable As TheTreeTable)
End Function

Public Function append(srcTable As TheTreeTable)
End Function



Public Function copyCurrentRow(ByRef sourceTable As TheTreeTable, _
                                Optional ignoreEmptyText = True)
End Function
                                
                                
Public Function isAutoIncrement(ByVal fieldName) As Boolean
End Function

Public Function isPrimaryKey(ByVal fieldName) As Boolean
End Function

Public Function getPrimaryKey() As String
End Function

Public Function getAutoincrementField() As String
End Function

'see fieldExists()
Public Function keyExists(ByVal key As String) As Boolean
End Function

'alias to keyExists()
Public Function fieldExists(ByVal key As String) As Boolean
End Function

'=============================================================================
'
'=============================================================================
#If OLD_CODE Then
#End If





