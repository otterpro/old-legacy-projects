VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TheStack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim myStack As Collection   'collection = 1-based, not 0-based

Private Sub Class_Initialize()
    Set myStack = New Collection
End Sub

Public Sub push(item)
    myStack.Add item
End Sub

Public Function pop()
    If myStack.count = 0 Then
        ' Nothing to pop.
        pop = Null
    Else
        pop = myStack.item(myStack.count())
        myStack.remove (myStack.count())
    End If
End Function

Public Function getSize() As Long
    getSize = myStack.count()
End Function
