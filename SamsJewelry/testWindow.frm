VERSION 5.00
Begin VB.Form testWindow 
   Caption         =   "testWindow"
   ClientHeight    =   7260
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8805
   LinkTopic       =   "Form1"
   ScaleHeight     =   7260
   ScaleWidth      =   8805
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command12 
      Caption         =   "testRemoveQuoteFunc"
      Height          =   1335
      Left            =   6840
      TabIndex        =   11
      Top             =   5400
      Width           =   1695
   End
   Begin VB.CommandButton Command11 
      Caption         =   "TheUser , Interface Class"
      Height          =   1695
      Left            =   6960
      TabIndex        =   10
      Top             =   3240
      Width           =   1335
   End
   Begin VB.CommandButton Command10 
      Caption         =   "getStringFromArrayOrCollection"
      Height          =   1215
      Left            =   5040
      TabIndex        =   9
      Top             =   4560
      Width           =   1455
   End
   Begin VB.CommandButton Command9 
      Caption         =   "regexp"
      Height          =   1695
      Left            =   6240
      TabIndex        =   8
      Top             =   1440
      Width           =   2055
   End
   Begin VB.CommandButton Command8 
      Caption         =   "class_instance"
      Height          =   1215
      Left            =   3120
      TabIndex        =   7
      Top             =   5040
      Width           =   1695
   End
   Begin VB.CommandButton Command7 
      Caption         =   "scripting.dictionary"
      Height          =   1575
      Left            =   480
      TabIndex        =   6
      Top             =   4680
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "TheDictionary"
      Height          =   855
      Left            =   5760
      TabIndex        =   5
      Top             =   720
      Width           =   2175
   End
   Begin VB.CommandButton Command6 
      Caption         =   "TheListTest"
      Height          =   1215
      Left            =   3000
      TabIndex        =   4
      Top             =   3120
      Width           =   2535
   End
   Begin VB.CommandButton Command5 
      Caption         =   "max(),min()"
      Height          =   735
      Left            =   3480
      TabIndex        =   3
      Top             =   1920
      Width           =   1935
   End
   Begin VB.CommandButton Command4 
      Caption         =   "testArrayVsParamArray"
      Height          =   855
      Left            =   3240
      TabIndex        =   2
      Top             =   480
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "cstr() test"
      Height          =   975
      Left            =   240
      TabIndex        =   1
      Top             =   2880
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "getString"
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2175
   End
End
Attribute VB_Name = "testWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const a1 = 34.5
Const a2 = 44
Const a3 = True
Const a4 = "hello there"
Const a5 = False

Private Sub Command1_Click()
    the.test "getString"
End Sub



Private Sub Command10_Click()
    Dim tempArr(0 To 2)
    tempArr(0) = "abc"
    tempArr(1) = "yoyo"
    tempArr(2) = "susan"
    dprint "-----BEG"
    'dprint getStringFromArrayOrCollection(tempArr)

    dprint getStringFromArrayOrCollection(tempArr, "+")
    dprint getStringFromArrayOrCollection(tempArr, vbTab)
    dprint "----- END"
End Sub

Private Sub Command11_Click()
    Dim user As New TheUser
    user.p1
End Sub

Private Sub Command12_Click()
    Dim a As String, b As String
    a = Chr(34) + "abc" + Chr(34)
    the.dprint "abc without removal", a
    b = removeLeadingAndTrailingQuote(a)
    the.dprint "abc", b
    
End Sub

Private Sub Command2_Click()
    Dim d As New TheDictionary
    d.test
End Sub

Private Sub Command3_Click()
    dprint "true", CStr(True)
    dprint "false", CStr(False)
    dprint "31.4", CStr(31.4)
    dprint "string", CStr("hello")
    dprint "now()", CStr(Now())
    'dprint "str", Str("hello")
End Sub

Function arraySub(ByRef Data())
Dim i
    dprint "--------------------"
    For Each i In Data
        dprint "arraySub", "<" & i & ">"
    Next
    arraySub = 1
End Function

Sub paramarraySub(ParamArray Data())
    Dim i
    dprint "--------------------"
    For Each i In Data
        If IsArray(i) Then
            Dim j
            For Each j In i
            dprint "param isArray", j
            Next
        Else
        dprint "paramarraySub", "<" & i & ">"
        End If
    Next

End Sub

Private Sub Command4_Click()
    Dim arr(0 To 3) As String
    arr(0) = "hi"
    arr(1) = "there"
    arr(2) = "good"
    arr(3) = "bye"
    Dim i
    'i = arraySub(arr)
    paramarraySub "abc", "hello"
    paramarraySub arr
End Sub


Private Sub Command5_Click()
    dprint "max", CStr(Max(1, 5))
    dprint "Min", CStr(Min(1, 5))
End Sub

Private Sub Command6_Click()
    Dim l As New TheList
    l.test
    
End Sub

Private Sub Command8_Click()
    Dim a As TheList
    Set a = New TheList
    
    a.Add "abc"
    dprint "Hello"
End Sub

