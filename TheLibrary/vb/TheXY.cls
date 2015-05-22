VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TheXY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private myX ' could be long or float.  Prefer Long. As Long
Private myY ' As Long
'Private myZ As Long

Public Property Get x()
    x = myX
End Property

Public Property Get y()
    y = myY
End Property
Public Property Let x(x)
    myX = x
End Property
Public Property Let y(y)
    myY = y
End Property
Public Sub setValue(param1, Optional param2)
    If TypeOf param1 Is TheXY Then
        myX = param1.x
        myY = param1.y
    Else
        myX = param1
        myY = param2
    End If
End Sub


