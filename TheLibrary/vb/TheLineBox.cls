VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TheLineBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' DO NOT USE: NOT TESTED and NOT FINISHED.

' Box made of 4 lines instead of 1 picture object. The reason for doing this is to be able to
' control each lines independently. This could be made into polygon later.
' WISH: Convert it into Polygon class.
Option Explicit



Private myLines(0 To 3) As Line

' myLines(0) = left, (1) = right (2) = top (3) = bottom
Private Enum LINE_NAME
    LEFT_LINE
    RIGHT_LINE
    TOP_LINE
    BOTTOM_LINE
End Enum

'TODO: isRectangle(), isSquare() returns true if rect or square.

Public Sub setValue(left As Line, right As Line, top As Line, bottom As Line)
    Set myLines(LEFT_LINE) = left
    Set myLines(RIGHT_LINE) = right
    Set myLines(TOP_LINE) = top
    Set myLines(BOTTOM_LINE) = bottom
End Sub

Public Sub setVisible(Optional showBox = True)
    Dim i As Line
    For Each i In myLines
        i.visible = showBox
    Next
End Sub

Public Sub moveTo(ByVal x, ByVal y)
    Dim i As Line
    Dim origin As New TheXY
    Dim delta As New TheXY
    
    For Each i In myLines
        origin.setValue i.x1, i.y1
        delta.setValue i.x2 - i.x1, i.y2 - y.y1
        the.moveLine i, x, y, x + delta.x, y + delta.y
    Next
End Sub





'=============================================================================
' FUTURE
'=============================================================================
#If FUTURE_CODE Then
Private myWidth As Long
Private myHeight As Long

Property Get width()
    width = myWidth
End Property

' Works only with Rectangle. Doesn't work with other 4 sided polygon.
Property Set width(value)
    myWidth = value
    Dim i As Line
    Set i = myLines(RIGHT_LINE)
    i.x1 = value
    i.x2 = value
End Property

#End If
