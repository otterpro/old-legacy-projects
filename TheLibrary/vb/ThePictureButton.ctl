VERSION 5.00
Begin VB.UserControl ThePictureButton 
   ClientHeight    =   2295
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3735
   EditAtDesignTime=   -1  'True
   PaletteMode     =   0  'Halftone
   ScaleHeight     =   153
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   249
   Begin VB.PictureBox myPicture 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   360
      ScaleHeight     =   95
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   207
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label ButtonText 
      BackStyle       =   0  'Transparent
      Caption         =   "ButtonText"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   1800
      Width           =   975
      WordWrap        =   -1  'True
   End
   Begin VB.Line bottomBorder 
      BorderColor     =   &H80000015&
      BorderWidth     =   2
      X1              =   0
      X2              =   248
      Y1              =   144
      Y2              =   144
   End
   Begin VB.Line rightBorder 
      BorderColor     =   &H80000015&
      BorderWidth     =   2
      X1              =   248
      X2              =   248
      Y1              =   144
      Y2              =   0
   End
   Begin VB.Line topBorder 
      BorderColor     =   &H80000014&
      BorderWidth     =   2
      X1              =   248
      X2              =   -8
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line leftBorder 
      BorderColor     =   &H80000014&
      BorderWidth     =   2
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   144
   End
End
Attribute VB_Name = "ThePictureButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=============================================================================
'   pictureButton
'
'   Written by Dan Kim (A-Team and CoconutStudio)
'   Copyright (c) 2003
'
'=============================================================================
' WISH
'   property Cancel=, Default=, BorderWidth=
'
Option Explicit

Private Const DEFAULT_PICTURE_SPACING = 2   'in Pixels (see Control's prop)
Private Const DEFAULT_BORDER_SPACING = 0  ' in Pixels
Private Const DEFAULT_CAPTION = ""
Private Const DEFAULT_PICTURE_APPEARANCE = 0    '0=flat, 1=3-d
Private Const DEFAULT_PICTURE_BORDER_STYLE = 1  '0=no border. 1=single-line
'Private Const DEFAULT_BORDER_THICKNESS = 1      '

Private myPictureSpacing    'usually 1
Private myBorderSpacing    'usually 0
Private myBorderThickness   'usually 1 pixel thick.

'Public Property Get Color() As OLE_COLOR
  'Color = Shape1.FillColor
'End Property

'Public Property Let Color(ByVal c As OLE_COLOR)
  'Shape1.FillColor = c
'End Property

Event Click()

'MappingInfo=UserControl,UserControl,-1,Click

Event DblClick()

Public Property Get Caption() As String
Attribute Caption.VB_Description = "Enter the button text."
Attribute Caption.VB_ProcData.VB_Invoke_Property = ";Text"
    Caption = ButtonText.Caption
End Property

Public Property Let Caption(ByVal value As String)
    ButtonText.Caption = value
End Property

Public Property Get PictureSpacing() As Long
Attribute PictureSpacing.VB_Description = "Spacing between the picture and the edge of the frame (in pixel)"
Attribute PictureSpacing.VB_ProcData.VB_Invoke_Property = ";Appearance"
    PictureSpacing = myPictureSpacing
End Property

Public Property Let PictureSpacing(ByVal value As Long)
    myPictureSpacing = value
    UserControl_Resize
End Property

Public Property Get BorderSpacing() As Long
Attribute BorderSpacing.VB_Description = "Space between the raised surface and the edge of the control. "
Attribute BorderSpacing.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BorderSpacing = myBorderSpacing
End Property

Public Property Let BorderSpacing(ByVal value As Long)
    myBorderSpacing = value
    UserControl_Resize
End Property

Public Property Get AutoSize() As Boolean
    AutoSize = myPicture.AutoSize
End Property
Public Property Let AutoSize(ByVal value As Boolean)
    myPicture.AutoSize = value
End Property

Public Property Get PictureAppearance() As AppearanceConstants
    PictureAppearance = myPicture.Appearance
End Property
Public Property Let PictureAppearance(ByVal value As AppearanceConstants)
    myPicture.Appearance = value
End Property

Public Property Get PictureBorderStyle() As BorderStyleConstants
    PictureBorderStyle = myPicture.BorderStyle
End Property
Public Property Let PictureBorderStyle(ByVal value As BorderStyleConstants)
    myPicture.BorderStyle = value
End Property

Public Property Get DataSource()
    Set DataSource = myPicture.DataSource
End Property
Public Property Set DataSource(value)
    Set myPicture.DataSource = value
End Property
Public Property Get DataField()
    DataField = myPicture.DataField
End Property
Public Property Let DataField(value)
    myPicture.DataField = value
End Property
Public Property Get Picture()
    Picture = myPicture.Picture
End Property
Public Property Let Picture(value)
    'Loads external file relative to the current path!
    If (InStr(value, ":") > 0) Then
        MsgBox "ThePictureButton::Let Picture() - " & _
                "DO not use fullpath! Use Relative Path!"
    Else
        openPicture (App.path & "\" & value)
    End If
End Property

Public Property Get Font() As Font
    Set Font = ButtonText.Font
End Property
Public Property Set Font(ByVal newFont As Font)
     Set ButtonText.Font = newFont
    PropertyChanged "Font"
    'comdlg
End Property


'Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
'Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
'Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp


'MappingInfo=UserControl,UserControl,-1,DblClick



Private Sub ButtonText_Click()
    UserControl_Click
End Sub

Private Sub ButtonText_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    UserControl_MouseDown Button, Shift, x, y
End Sub

Private Sub ButtonText_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    UserControl_MouseUp Button, Shift, x, y
End Sub

Private Sub myPicture_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    UserControl_MouseDown Button, Shift, x, y
End Sub

Private Sub myPicture_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    UserControl_MouseUp Button, Shift, x, y
End Sub

Public Function openPicture(ByVal file)
    On Error GoTo PICTURE_NOT_FOUND
    Debug.Assert InStr(file, ":") > 0 'make sure it is full-path
    myPicture.Picture = LoadPicture(file)
    openPicture = True
PICTURE_NOT_FOUND:
    openPicture = False
End Function

Public Sub clearPicture()
    myPicture.Picture = LoadPicture()   'empty file means clear.
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub


Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_EnterFocus()
  'Shape1.BorderColor = &HFFFFFF
End Sub

Private Sub UserControl_ExitFocus()
  'Shape1.BorderColor = &H0
End Sub


Private Sub UserControl_Initialize()
    myPictureSpacing = DEFAULT_PICTURE_SPACING
    myBorderSpacing = DEFAULT_BORDER_SPACING  ' in Pixels
    Caption = DEFAULT_CAPTION
    myPicture.Appearance = DEFAULT_PICTURE_APPEARANCE
    myPicture.BorderStyle = DEFAULT_PICTURE_BORDER_STYLE
    
    the.moveLine topBorder, myBorderSpacing, myBorderSpacing, myBorderSpacing, myBorderSpacing
    the.moveLine bottomBorder, myBorderSpacing, myBorderSpacing, myBorderSpacing, myBorderSpacing
    the.moveLine rightBorder, myBorderSpacing, myBorderSpacing, myBorderSpacing, myBorderSpacing
    the.moveLine leftBorder, myBorderSpacing, myBorderSpacing, myBorderSpacing, myBorderSpacing

    showReleasedState
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    showPressedState
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    showReleasedState
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    ButtonText.Caption = PropBag.ReadProperty("Caption", DEFAULT_CAPTION)
    myPictureSpacing = PropBag.ReadProperty("PictureSpacing", _
                                            DEFAULT_PICTURE_SPACING)
    myBorderSpacing = PropBag.ReadProperty("BorderSpacing", _
                                            DEFAULT_BORDER_SPACING)
    myPicture.AutoSize = PropBag.ReadProperty("AutoSize", False)
    myPicture.Appearance = PropBag.ReadProperty("PictureAppearance", _
                                            DEFAULT_PICTURE_APPEARANCE)
    myPicture.BorderStyle = PropBag.ReadProperty("PictureBorderStyle", _
                                                DEFAULT_PICTURE_BORDER_STYLE)
    Set myPicture.DataSource = PropBag.ReadProperty("DataSource", Nothing)
    
    myPicture.DataField = PropBag.ReadProperty("DataField", "")
    Set myPicture.Picture = PropBag.ReadProperty("Picture", Nothing)
    Set ButtonText.Font = PropBag.ReadProperty("Font", Ambient.Font)

    UserControl_Resize
End Sub
'
'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("BorderSpacing", myBorderSpacing, DEFAULT_BORDER_SPACING)
    Call PropBag.WriteProperty("PictureSpacing", myPictureSpacing, DEFAULT_PICTURE_SPACING)
    
    Call PropBag.WriteProperty("AutoSize", myPicture.AutoSize, False)
    Call PropBag.WriteProperty("PictureAppearance", myPicture.Appearance _
                                , DEFAULT_PICTURE_APPEARANCE)
    Call PropBag.WriteProperty("PictureBorderStyle", myPicture.BorderStyle, _
                                DEFAULT_PICTURE_BORDER_STYLE)
    Call PropBag.WriteProperty("DataSource", myPicture.DataSource)
    Call PropBag.WriteProperty("DataField", myPicture.DataField)
    Call PropBag.WriteProperty("Picture", myPicture.Picture)
    Call PropBag.WriteProperty("Caption", ButtonText.Caption, DEFAULT_CAPTION)
    Call PropBag.WriteProperty("Font", ButtonText.Font, Ambient.Font)
End Sub

Private Sub UserControl_Resize()
    'resize Pic
    myPicture.Left = myPictureSpacing
    myPicture.Top = myPictureSpacing
    myPicture.Width = UserControl.ScaleWidth - (myPictureSpacing * 2)
    myPicture.Height = UserControl.ScaleHeight - (myPictureSpacing * 2) _
                        - ButtonText.Height
    'resize Button border
    'the.moveLineBox 0, 0, UserControl.width, UserControl.Height
    
    'Width
    Dim newWidth
    newWidth = UserControl.ScaleWidth - (myBorderSpacing * 2) - 1
        ' sub 1 more pixel because it isn't shown in the 3d-border
    topBorder.x2 = newWidth
    bottomBorder.x2 = newWidth
    rightBorder.x1 = newWidth
    rightBorder.x2 = newWidth
    
    'Height
    Dim newHeight
    newHeight = UserControl.ScaleHeight - (myBorderSpacing * 2) - 1
    leftBorder.y2 = newHeight
    rightBorder.y2 = newHeight
    bottomBorder.y1 = newHeight
    bottomBorder.y2 = newHeight
    
    'move button text
    ButtonText.Left = myPicture.Left
    ButtonText.Top = myPicture.Top + myPicture.Height
    ButtonText.Width = myPicture.Width
    
    
End Sub


Private Sub showPressedState()
    topBorder.BorderColor = vb3DDKShadow
    bottomBorder.BorderColor = vb3DHighlight
    leftBorder.BorderColor = vb3DDKShadow
    rightBorder.BorderColor = vb3DHighlight

End Sub

Private Sub showReleasedState()
    topBorder.BorderColor = vb3DHighlight
    bottomBorder.BorderColor = vb3DDKShadow
    leftBorder.BorderColor = vb3DHighlight
    rightBorder.BorderColor = vb3DDKShadow
End Sub

