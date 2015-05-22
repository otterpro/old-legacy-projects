VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TheLargeInteger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Hasn't been tested. Test it.

Option Explicit

Private myLow As Long, myHigh As Long

Public Sub setNumber(low As Long, high As Long)
    myLow = low
    myHigh = high
End Sub
' LARGE_INTEGER or BIG_INTEGER
' Use the Currency format as it can hold really large integer.
' This class only helps to convert from long to currency and vice versa.
Public Function getCurrency() As Currency
    'convert LARGE_INTEGER to Currency
    Dim lowBig As Currency, highBig As Currency
    If myLow < 0 Then
        lowBig = 2 ^ 32 + myLow
    Else
        lowBig = myLow
    End If
    If myHigh < 0 Then
        highBig = 2 ^ 32 + myHigh
    Else
        highBig = myHigh
    End If
    getCurrency = lowBig + highBig * 2 ^ 32
End Function

