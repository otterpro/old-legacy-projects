VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TheRandom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mySeed As Long  ' seed value for random. If the seed value is same,
        ' it will repeat the same sequence. One can save the seed value to registry/ini
        ' to repeat the same order or change mySeed to different value for different random # sequence.

Public Sub init(Optional repeat As Boolean = False)
    ' randomize the seed.
    Rnd -1  ' -1 value indicates that randomize will use same sequence if needed.
    If (repeat = True) Then
        Dim config As New TheSimpleRegistry
        mySeed = CLng(config.getValue("random", "seed", Timer))
    Else
        mySeed = Timer
    End If
    Randomize mySeed
    
    
End Sub

' get next random value (or next value from the sequence)
Public Function getNext(ByVal lowerbound As Long, ByVal upperbound As Long) As Long
    Debug.Assert (mySeed)   'make sure seed was created via init().
    getNext = (upperbound - lowerbound + 1) * Rnd + lowerbound
End Function

Private Sub Class_Terminate()
    'save seed value
        Dim config As New TheSimpleRegistry
        config.setValue "random", "seed", mySeed
End Sub
