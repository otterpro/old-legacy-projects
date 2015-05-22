Attribute VB_Name = "TheMSWin"
Option Explicit

Public Enum MOUSEMESSAGES
  MOUSEOVER = &H200
  
  Rem Left Mouse Button Members...
  LEFTBUTTONDOUBLECLICK = &H203
  LEFTBUTTONUP = &H202
  LEFTBUTTONDOWN = &H201
  
  Rem Middle Mouse Button Members...
  MIDDLEBUTTONDOUBLECLICK = &H209
  MIDDLEBUTTONDOWN = &H207
  MIDDLEBUTTONUP = &H208
  
  Rem Right Mouse Button Members...
  RIGHTBUTTONDOUBLECLICK = &H206
  RIGHTBUTTONDOWN = &H204
  RIGHTBUTTONUP = &H205
End Enum

Public Enum OPERATING_SYSTEM
  UNKNOWN = &H8                           ' Non Microsoft Operating System
  WINDOWS95 = &H0                         ' Microsoft Windows® 95
  WINDOWS98 = &H1                         ' Microsoft Windows® 98
  WINDOWSME = &H2                         ' Microsoft Windows® Millenium Edition
  WINDOWSNT351 = &H3                      ' Microsoft Windows® NT v3.5.1.
  WINDOWSNT400 = &H4                      ' Microsoft Windows® NT 4.0.0.
  WINDOWS2000 = &H5                       ' Microsoft Windows® 2000 Professional, Server, Advanced Server
  WINDOWSXP = &H6                         ' Microsoft Windows® XP Personal, Home Edition
  WINDOWSNETSERVER = &H7                  ' Microsoft Windows® .nET® Server
End Enum


Public Enum MENU_THEME      ' /* Menu theme to be used for owner drawn menus. */
  MENU_THEME_AUTO = -(&H1)  ' /* Automatically detect appropriately menu style. */
  MENU_THEME_WIN2K = &H0&   ' /* Use regular style for menus. */
  MENU_THEME_WINXP = &H1&   ' /* Use WinXP style for menus. */
End Enum



'From: SystemInteroperability.DLL
'This structure contains basic information about a physical font.
'Public Type TEXTMETRIC
'  tmHeight                As Long ' Specifies the height (ascent descent) of characters.
'  tmAscent                As Long ' Specifies the ascent (units above the base line) of characters.
'  tmDescent               As Long ' Specifies the descent (units below the base line) of characters.
'  tmInternalLeading       As Long ' Specifies the amount of leading (space) inside the bounds set by
'                                  ' the tmHeight member. Accent marks and other diacritical characters
'                                  ' may occur in this area. The designer may set this member to zero.
'  tmExternalLeading       As Long ' Specifies the amount of extra leading (space) that the application adds between rows.
'  tmAveCharWidth          As Long ' Specifies the average width of characters in the font
'                                  ' (generally defined as the width of the letter x).
'  tmMaxCharWidth          As Long ' Specifies the width of the widest character in the font.
'  tmWeight                As Long ' Specifies the weight of the font.
'  tmOverhang              As Long ' Specifies the extra width per string that may be added to some
'                                  ' synthesized fonts. When synthesizing some attributes, such as bold
'                                  ' or italic, graphics device interface (GDI) or a device may have to
'                                  ' add width to a string on both a per-character and per-string basis.
'  tmDigitizedAspectX      As Long ' Specifies the horizontal aspect of the device for which the font was
'                                  ' designed.
'  tmDigitizedAspectY      As Long ' Specifies the vertical aspect of the device for which the font was
'                                  ' designed. The ratio of the tmDigitizedAspectX and tmDigitizedAspectY
'                                  ' members is the aspect ratio of the device for which the font was designed.
'  tmFirstChar             As Byte ' Specifies the value of the first character defined in the font.
'  tmLastChar              As Byte ' Specifies the value of the last character defined in the font.
'  tmDefaultChar           As Byte ' Specifies the value of the character to be substituted for characters
'                                  ' not in the font.
'  tmBreakChar             As Byte ' Specifies the value of the character that will be used to define
'                                  ' word breaks for text justification.
'  tmItalic                As Byte ' Specifies an italic font if it is nonzero.
'  tmUnderlined            As Byte ' Specifies an underlined font if it is nonzero.
'  tmStruckOut             As Byte ' Specifies a strikeout font if it is nonzero.
'  tmPitchAndFamily        As TEXTMETRIC_PITCH ' Specifies information about the pitch, the technology,
'                                  ' and the family of a physical font.
'                                  ' Four low-order bits of this member specify information about the pitch and
'                                  ' the technology of the font
'                                  ' The four high-order bits of tmPitchAndFamily designate the font's font
'                                  ' family. An application can use the value 0xF0 and the bitwise AND operator
'                                  ' to mask out the four low-order bits of tmPitchAndFamily, thus obtaining a
'                                  ' value that can be directly compared with font family names to find an
'                                  ' identical match.
'  tmCharSet               As LOGICAL_FONT_CHARSET ' Specifies the character set of the font.
'End Type

'Option Private Module: Option Explicit: Option Compare Text: Option Base 0
Rem-------------------------------------------------------------------------
Rem @Name                 :                 APIEnumerations
Rem @Type                 :                 Standard
Rem @Scope Qualifier      :                 Private
Rem @Purpose              :                 Provides and coordinates various
Rem                                         useful Win32 Constants used by
Rem                                         the solution.
Rem @Creation Date        :                 Saturday, 28 January 2002.
Rem @Creation Author      :                 Shantibhushan
Rem-------------------------------------------------------------------------

Public Const WM_USER                        As Long = &H400
Public Const TOOLTIP_CLASS                  As String = "tooltips_class32"

Public Const NONESSENTIAL                   As Long = &H0 ' Non-essential value...
Public Const NOHANDLE                       As Long = &H0 ' No handle present...
Public Const DEFAULT                        As Long = &H80000000

Public Const OFFSET                         As Long = &H16  ' Offset for positioning the tooltip...

#If (MAC = 0) Then
 Public Const MAX_COMPUTERNAME_LENGTH       As Long = &H1F    ' 31
#Else
 Public Const MAX_COMPUTERNAME_LENGTH       As Long = &HF     ' 15
#End If

Public Const UNLEN                          As Long = (&HFF + 1) ' Maximum user name length
Public Const MAX_PATH                       As Long = (&HFF + 5)    '260

Rem VB always has it's parent window named as the following...
Public Const VB_PARENT_WINDOW_NAME          As String = "ThunderMain"



'Option Private Module: Option Explicit: Option Compare Text: Option Base 0
Rem-------------------------------------------------------------------------
Rem @Name                 :                 APIEnumerations
Rem @Type                 :                 Standard
Rem @Scope Qualifier      :                 Private
Rem @Purpose              :                 Provides and coordinates various
Rem                                         useful Win32 Enumerations used by
Rem                                         the solution.
Rem @Creation Date        :                 Saturday, 28 January 2002.
Rem @Creation Author      :                 Shantibhushan
Rem-------------------------------------------------------------------------

Public Enum COMMON_CONTROL_TYPES
  ICC_LISTVIEW_CLASSES = &H1       ' listview, header
  ICC_TREEVIEW_CLASSES = &H2       ' treeview, tooltips
  ICC_BAR_CLASSES = &H4            ' toolbar, statusbar, trackbar, tooltips
  ICC_TAB_CLASSES = &H8            ' tab, tooltips
  ICC_UPDOWN_CLASS = &H10          ' updown
  ICC_PROGRESS_CLASS = &H20        ' progress
  ICC_HOTKEY_CLASS = &H40          ' hotkey
  ICC_ANIMATE_CLASS = &H80         ' animate
  ICC_WIN95_CLASSES = &HFF
  ICC_DATE_CLASSES = &H100         ' month picker, date picker, time picker, updown
  ICC_USEREX_CLASSES = &H200       ' comboex
  ICC_COOL_CLASSES = &H400         ' rebar (coolbar) control
#If (WIN32_IE >= &H400) Then
  ICC_INTERNET_CLASSES = &H800
  ICC_PAGESCROLLER_CLASS = &H1000  ' page scroller
  ICC_NATIVEFNTCTL_CLASS = &H2000  ' native font control
#End If
End Enum

Public Enum WINDOW_STYLE
  WS_OVERLAPPED = &H0&                ' Creates an overlapped window. An overlapped window usually has a caption and a border.
  WS_POPUP = &H80000000               ' Creates a pop-up window. Cannot be used with the WS_CHILD style.
  WS_CHILD = &H40000000               ' Creates a child window. Cannot be used with the WS_POPUP style.
  WS_MINIMIZE = &H20000000            ' Creates a window that is initially minimized. For use with the WS_OVERLAPPED style only.
  WS_VISIBLE = &H10000000
  WS_DISABLED = &H8000000             ' Creates a window that is initially disabled.
  WS_CLIPSIBLINGS = &H4000000         ' Clips child windows relative to each other; that is, when a particular child window receives
                                      ' a paint message, the WS_CLIPSIBLINGS style clips all other overlapped child windows out of the
                                      ' region of the child window to be updated. (If WS_CLIPSIBLINGS is not given and child windows
                                      ' overlap, when you draw within the client area of a child window, it is possible to draw within
                                      ' the client area of a neighboring child window.) For use with the WS_CHILD style only.
  WS_CLIPCHILDREN = &H2000000         ' Excludes the area occupied by child windows when you draw within the parent window. Used when you
                                      ' create the parent window.
  WS_MAXIMIZE = &H1000000             ' Creates a window of maximum size.
  WS_CAPTION = &HC00000               ' Creates a window that has a title bar (implies the WS_BORDER style). Cannot be used with the WS_DLGFRAME style.
  WS_BORDER = &H800000                ' Creates a window that has a border.
  WS_DLGFRAME = &H400000              ' Creates a window with a double border but no title.
  WS_VSCROLL = &H200000               ' Creates a window that has a vertical scroll bar.
  WS_HSCROLL = &H100000               ' Creates a window that has a horizontal scroll bar.
  WS_SYSMENU = &H80000                ' Creates a window that has a Control-menu box in its title bar. Used only for windows with title bars.
  WS_THICKFRAME = &H40000             ' Creates a window with a thick frame that can be used to size the window.
  WS_GROUP = &H20000                  ' Specifies the first control of a group of controls in which the user can move from one control to
                                      ' the next with the arrow keys. All controls defined with the WS_GROUP style FALSE after the first control
                                      ' belong to the same group. The next control with the WS_GROUP style starts the next group (that is, one
                                      ' group ends where the next begins).
  WS_TABSTOP = &H10000                ' Specifies one of any number of controls through which the user can move by using the TAB key. The TAB key moves
                                      ' the user to the next control specified by the WS_TABSTOP style.
  WS_MINIMIZEBOX = &H20000            ' Creates a window that is initially minimized. For use with the WS_OVERLAPPED style only.
  WS_MAXIMIZEBOX = &H10000            ' Creates a window that has a Maximize button.

  WS_TILED = WS_OVERLAPPED            ' Creates an overlapped window. An overlapped window usually has a caption and a border.
  WS_ICONIC = WS_MINIMIZE
  WS_SIZEBOX = WS_THICKFRAME

  WS_CHILDWINDOW = WS_CHILD
Rem Common Window Styles
  WS_OVERLAPPEDWINDOW = WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX
  WS_POPUPWINDOW = WS_POPUP Or WS_BORDER Or WS_SYSMENU  ' Creates a pop-up window with the WS_BORDER, WS_POPUP, and WS_SYSMENU styles. The WS_CAPTION style
                                                        ' must be combined with the WS_POPUPWINDOW style to make the Control menu visible.
  WS_TILEDWINDOW = WS_OVERLAPPEDWINDOW
End Enum

Public Enum WINDOW_STYLE_EXTENDED
  WS_EX_DLGMODALFRAME = &H1&       ' Creates a window that has a double border; the window can, optionally, be created with a title bar by specifying the WS_CAPTION style in the dwStyle parameter.
  WS_EX_NOPARENTNOTIFY = &H4&      ' Specifies that a child window created with this style does not send the WM_PARENTNOTIFY message to its parent window when it is created or destroyed.
  WS_EX_TOPMOST = &H8&             ' Specifies that a window created with this style should be placed above all non-topmost windows and should stay above them, even when the window is deactivated.
                                   ' To add or remove this style, use the SetWindowPos function.
  WS_EX_ACCEPTFILES = &H10&        ' Specifies that a window created with this style accepts drag-drop files.
  WS_EX_TRANSPARENT = &H20&        ' Specifies that a window created with this style should not be painted until siblings beneath the window (that were created by the same thread) have been painted.
                                   ' The window appears transparent because the bits of underlying sibling windows have already been painted. To achieve transparency without these restrictions, use
                                   ' the SetWindowRgn function.

  WS_EX_LAYERED = &H80000          ' Creates a layered window. Layered windows supported translucency effects.
#If (WIN32_IE >= &H400) Then
  WS_EX_MDICHILD = &H40&           ' Creates an MDI child window.
  WS_EX_TOOLWINDOW = &H80&         ' Creates a tool window; that is, a window intended to be used as a floating toolbar. A tool window has a title bar that is shorter than a normal title bar, and the
                                   ' window title is drawn using a smaller font. A tool window does not appear in the taskbar or in the dialog that appears when the user presses ALT+TAB. If a tool window
                                   ' has a system menu, its icon is not displayed on the title bar. However, you can display the system menu by right-clicking or by typing ALT+SPACE.
  WS_EX_WINDOWEDGE = &H100&        ' Specifies that a window has a border with a raised edge.
  WS_EX_CLIENTEDGE = &H200&        ' Specifies that a window has a border with a sunken edge.
  WS_EX_CONTEXTHELP = &H400&       ' Includes a question mark in the title bar of the window. When the user clicks the question mark, the cursor changes to a question mark with a pointer. If the user then
                                   ' clicks a child window, the child receives a WM_HELP message.

  WS_EX_RIGHT = &H1000&            ' The window has generic gright-alignedh properties. This depends on the window class. This style has an effect only if the shell language is Hebrew, Arabic,
                                   ' or another language that supports reading-order alignment; otherwise, the style is ignored.
  WS_EX_LEFT = &H0&                ' Creates a window that has generic left-aligned properties. This is the default.
  WS_EX_RTLREADING = &H2000&       ' If the shell language is Hebrew, Arabic, or another language that supports reading-order alignment, the window text is displayed using right-to-left reading-order
                                   ' properties. For other languages, the style is ignored.
  WS_EX_LTRREADING = &H0&          ' The window text is displayed using left-to-right reading-order properties. This is the default.
  WS_EX_LEFTSCROLLBAR = &H4000&    ' If the shell language is Hebrew, Arabic, or another language that supports reading order alignment, the vertical scroll bar (if present) is to the left of the
                                   ' client area. For other languages, the style is ignored.
  WS_EX_RIGHTSCROLLBAR = &H0&      ' Vertical scroll bar (if present) is to the right of the client area. This is the default.

  WS_EX_CONTROLPARENT = &H10000    ' Allows the user to navigate among the child windows of the window by using the TAB key.
  WS_EX_STATICEDGE = &H20000       ' Creates a window with a three-dimensional border style intended to be used for items that do not accept user input.
  WS_EX_APPWINDOW = &H40000        ' Forces a top-level window onto the taskbar when the window is visible.

  WS_EX_OVERLAPPEDWINDOW = WS_EX_WINDOWEDGE Or WS_EX_CLIENTEDGE ' Combines the WS_EX_CLIENTEDGE and WS_EX_WINDOWEDGE styles.
  WS_EX_PALETTEWINDOW = WS_EX_WINDOWEDGE Or WS_EX_TOOLWINDOW Or WS_EX_TOPMOST ' Combines the WS_EX_WINDOWEDGE, WS_EX_TOOLWINDOW, and WS_EX_TOPMOST styles.
  
  
#End If
End Enum

Public Enum WINDOW_OFFSETS
  GWL_EXSTYLE = (-20)             ' Sets a new extended window style. For more information, see CreateWindowEx.
  GWL_STYLE = (-16)               ' Sets a new window style.
  GWL_WNDPROC = (-4)              ' Sets a new address for the window procedure.
  GWL_HINSTANCE = (-6)            ' Sets a new application instance handle.
  GWL_ID = (-12)                  ' Sets a new identifier of the window.
  GWL_USERDATA = (-21)            ' Sets the user data associated with the window. This data is intended for use by the application that created the window.
                                  ' Its value is initially zero.
  Rem The following values are also available when the hWnd parameter identifies a dialog box.
  DWL_DLGPROC = 4                 ' Sets the new pointer to the dialog box procedure.
  DWL_MSGRESULT = 0               ' Sets the return value of a message processed in the dialog box procedure.
  DWL_USER = 8                    ' Sets new extra information that is private to the application, such as handles or pointers.
End Enum

Public Enum WINDOW_PAINT_MESSAGE
  WM_PAINT = &HF
  WM_PRINT = &H317
  WM_PRINTCLIENT = &H318
End Enum

Public Enum KEYBOARD_MESSAGE
  WM_SYSCOMMAND = &H112           ' /* A window receives this message when the user chooses a
                                  '  * command from the window menu (formerly known as the system
                                  '  * or control menu) or when the user chooses the maximize
                                  '  * button, minimize button, restore button, or close button.
                                  '  */
  WM_COMMAND = &H111              ' /* The WM_COMMAND message is sent when the user selects a
                                  '  * command item from a menu, when a control sends a
                                  '  * notification message to its parent window, or when an
                                  '  * accelerator keystroke is translated.
                                  '  */
  WM_KEYFIRST = &H100             ' /* This message filters for keyboard messages. */
  WM_KEYDOWN = &H100              ' /* The WM_KEYDOWN message is posted to the window with the
                                  '  * keyboard focus when a nonsystem key is pressed. A nonsystem
                                  '  * key is a key that is pressed when the ALT key is not
                                  '  * pressed.
                                  '  */
  WM_KEYUP = &H101                ' /* The WM_KEYUP message is posted to the window with the
                                  '  * keyboard focus when a nonsystem key is released. A
                                  '  * nonsystem key is a key that is pressed when the ALT key is
                                  '  * not pressed, or a keyboard key that is pressed when a
                                  '  * window has the keyboard focus.
                                  '  */
  WM_CHAR = &H102                 ' /* The WM_CHAR message is posted to the window with the
                                  '  * keyboard focus when a WM_KEYDOWN message is translated by
                                  '  * the TranslateMessage function. The WM_CHAR message contains
                                  '  * the character code of the key that was pressed.
                                  '  */
  WM_DEADCHAR = &H103             ' /* The WM_DEADCHAR message is posted to the window with the
                                  '  * keyboard focus when a WM_KEYUP message is translated by the
                                  '  * TranslateMessage function. WM_DEADCHAR specifies a character
                                  '  * code generated by a dead key. A dead key is a key that
                                  '  * generates a character, such as the umlaut (double-dot), that
                                  '  * is combined with another character to form a composite character.
                                  '  * For example, the umlaut-O character () is generated by typing
                                  '  * the dead key for the umlaut character, and then typing the
                                  '  * O key.
                                  '  */
  WM_SYSKEYDOWN = &H104           ' /* The WM_SYSKEYDOWN message is posted to the window with the
                                  '  * keyboard focus when the user presses the F10 key (which
                                  '  * activates the menu bar) or holds down the ALT key and then
                                  '  * presses another key. It also occurs when no window currently
                                  '  * has the keyboard focus; in this case, the WM_SYSKEYDOWN message
                                  '  * is sent to the active window. The window that receives the
                                  '  * message can distinguish between these two contexts by checking
                                  '  * the context code in the lParam parameter.
                                  '  */
  WM_SYSKEYUP = &H105             ' /* The WM_SYSKEYUP message is posted to the window with the
                                  '  * keyboard focus when the user releases a key that was pressed
                                  '  * while the ALT key was held down. It also occurs when no window
                                  '  * currently has the keyboard focus; in this case, the WM_SYSKEYUP
                                  '  * message is sent to the active window. The window that receives
                                  '  * the message can distinguish between these two contexts by checking
                                  '  * the context code in the lParam parameter.
                                  '  */
  WM_SYSCHAR = &H106              ' /* The WM_SYSCHAR message is posted to the window with the
                                  '  * keyboard focus when a WM_SYSKEYDOWN message is translated
                                  '  * by the TranslateMessage function. It specifies the character
                                  '  * code of a system character key that is, a character key that is
                                  '  * pressed while the ALT key is down.
                                  '  */
  WM_SYSDEADCHAR = &H107          ' /* The WM_SYSDEADCHAR message is sent to the window with the
                                  '  * keyboard focus when a WM_SYSKEYDOWN message is translated by
                                  '  * the TranslateMessage function. WM_SYSDEADCHAR specifies the
                                  '  * character code of a system dead key that is, a dead key that is
                                  '  * pressed while holding down the ALT key.
                                  '  */
  WM_KEYLAST = &H108              ' /* Keyboard message filter value. */
  WM_SETFOCUS = &H7               ' /* The WM_SETFOCUS message is sent to a window after it has gained the keyboard focus. */
  WM_KILLFOCUS = &H8              ' /* The WM_KILLFOCUS message is sent to a window immediately before it loses the keyboard focus. */
End Enum

Public Enum WINDOW_MESSAGES
  WM_NULL = &H0
  WM_CREATE = &H1
  WM_SPLWND = &H1E2
  WM_DESTROY = &H2
  WM_NCDESTROY = &H82
  WM_MOVE = &H3
  WM_SIZE = &H5
  WM_ACTIVATE = &H6
  WM_PAINTICON = &H26
  WM_ICONERASEBKGND = &H27
  WM_NEXTDLGCTL = &H28
  WM_SPOOLERSTATUS = &H2A
  WM_DRAWITEM = &H2B
  WM_MEASUREITEM = &H2C
  WM_DELETEITEM = &H2D
  WM_INITMENU = &H116
  WM_INITMENUPOPUP = &H117
  WM_NCPAINT = &H85
  WM_NCCALCSIZE = &H83
  WM_VKEYTOITEM = &H2E
  WM_CHARTOITEM = &H2F
  WM_SETFONT = &H30
  WM_GETFONT = &H31
  WM_SETHOTKEY = &H32
  WM_GETHOTKEY = &H33
  WM_QUERYDRAGICON = &H37
  WM_COMPAREITEM = &H39
#If (WIN32_IE >= &H500) Then
  WM_GETOBJECT = &H3D
#End If '/* WINVER >= &H0500 */
  WM_COMPACTING = &H41
  WM_COMMNOTIFY = &H44                      ' /* no longer suported */
  WM_WINDOWPOSCHANGING = &H46
  WM_WINDOWPOSCHANGED = &H47
  WM_POWER = &H48
  WM_SHOWWINDOW = &H18
  WM_ERASEBKGND = &H14
  WM_EXITMENULOOP = &H211
#If (WIN32_IE >= &H400) Then
  WM_NEXTMENU = &H213
#End If
  WM_ENTERIDLE = &H121
  WM_TIMER = &H113
  WM_NCHITTEST = &H84
End Enum

Public Enum SHOW_WINDOW_MSGS
 SW_HIDE = 0
 SW_SHOWNORMAL = 1
 SW_NORMAL = 1
 SW_SHOWMINIMIZED = 2
 SW_SHOWMAXIMIZED = 3
 SW_MAXIMIZE = 3
 SW_SHOWNOACTIVATE = 4
 SW_SHOW = 5
 SW_MINIMIZE = 6
 SW_SHOWMINNOACTIVE = 7
 SW_SHOWNA = 8
 SW_RESTORE = 9
 SW_SHOWDEFAULT = 10
 SW_FORCEMINIMIZE = 11
 SW_MAX = 11
End Enum

Rem --This section is used exclusively for tooltips and associated members--
Public Enum TOOLTIP_STYLES         ' Styles applicable for a tooltip...
  TTS_ALWAYSTIP = &H1              ' Always shown as tooltip...
  TTS_NOPREFIX = &H2               ' Unknown...
  TTS_BALLOON = &H40               ' Shown as balloon tooltip...
  TTS_NOANIMATE = &H10             ' Plain, with no animation...
  TTS_NOFADE = &H20                ' No fade effects for tooltip...
End Enum

Public Enum TOOLTIP_FLAGS          ' Flags applicable for tooltip...
  TTF_IDISHWND = &H1               ' The ID is the HWND for the tooltip...
  TTF_CENTERTIP = &H2              ' Show tooltip stem in center...
  TTF_RTLREADING = &H4             ' Show tooltip text in right-to-left reading...
  TTF_SUBCLASS = &H10              ' Use the tooltip window as a subclass of parent...
#If (WIN32_IE >= &H300) Then
  TTF_TRACK = &H20                 ' Track the mouse position...
  TTF_ABSOLUTE = &H80              ' Position the tooltip to absolute pixels...
  TTF_TRANSPARENT = &H100          ' Show transparent tooltips...
  TTF_DI_SETITEM = &H8000          ' Unknown...
#End If
End Enum

Public Enum TOOLTIP_DELAYTIME      ' Sets the initial, pop-up, and reshow durations for a ToolTip control...
  TTDT_AUTOMATIC = &H0             ' Set all three delay times to default proportions. The autopop time will be ten times the initial time and the reshow
                                   ' time will be one fifth the initial time. If this flag is set, use a positive value of iTime to specify the initial time,
                                   ' in milliseconds. Set iTime to a negative value to return all three delay times to their default values.
  TTDT_RESHOW = &H1                ' Set the length of time it takes for subsequent ToolTip windows to appear as the pointer moves from one tool to another.
                                   ' To return the reshow delay time to its default value, set iTime to -1.
  TTDT_AUTOPOP = &H2               ' Set the length of time a ToolTip window remains visible if the pointer is stationary within a tool's bounding rectangle.
                                   ' To return the autopop delay time to its default value, set iTime to -1....
  TTDT_INITIAL = &H3               ' Set the length of time a pointer must remain stationary within a tool's bounding rectangle before the ToolTip window appears.
                                   ' To return the initial delay time to its default value, set iTime to -1...
End Enum

Public Enum TOOLTIP_MESSAGES
  TTM_ACTIVATE = (WM_USER + 1)
  TTM_SETDELAYTIME = (WM_USER + 3)
  TTM_ADDTOOLA = (WM_USER + 4)
  TTM_ADDTOOLW = (WM_USER + 50)
  TTM_DELTOOLA = (WM_USER + 5)
  TTM_DELTOOLW = (WM_USER + 51)
  TTM_NEWTOOLRECTA = (WM_USER + 6)
  TTM_NEWTOOLRECTW = (WM_USER + 52)
  TTM_RELAYEVENT = (WM_USER + 7)

  TTM_GETTOOLINFOA = (WM_USER + 8)
  TTM_GETTOOLINFOW = (WM_USER + 53)

  TTM_SETTOOLINFOA = (WM_USER + 9)
  TTM_SETTOOLINFOW = (WM_USER + 54)

  TTM_HITTESTA = (WM_USER + 10)
  TTM_HITTESTW = (WM_USER + 55)
  TTM_GETTEXTA = (WM_USER + 11)
  TTM_GETTEXTW = (WM_USER + 56)
  TTM_UPDATETIPTEXTA = (WM_USER + 12)
  TTM_UPDATETIPTEXTW = (WM_USER + 57)
  TTM_GETTOOLCOUNT = (WM_USER + 13)
  TTM_ENUMTOOLSA = (WM_USER + 14)
  TTM_ENUMTOOLSW = (WM_USER + 58)
  TTM_GETCURRENTTOOLA = (WM_USER + 15)
  TTM_GETCURRENTTOOLW = (WM_USER + 59)
  TTM_WINDOWFROMPOINT = (WM_USER + 16)
#If (WIN32_IE >= &H300) Then
  TTM_TRACKACTIVATE = (WM_USER + 17)
  TTM_TRACKPOSITION = (WM_USER + 18)
  TTM_SETTIPBKCOLOR = (WM_USER + 19)
  TTM_SETTIPTEXTCOLOR = (WM_USER + 20)
  TTM_GETDELAYTIME = (WM_USER + 21)
  TTM_GETTIPBKCOLOR = (WM_USER + 22)
  TTM_GETTIPTEXTCOLOR = (WM_USER + 23)
  TTM_SETMAXTIPWIDTH = (WM_USER + 24)
  TTM_GETMAXTIPWIDTH = (WM_USER + 25)
  TTM_SETMARGIN = (WM_USER + 26)
  TTM_GETMARGIN = (WM_USER + 27)
  TTM_POP = (WM_USER + 28)
#End If

#If (WIN32_IE >= &H400) Then
  TTM_UPDATE = (WM_USER + 29)
#End If

  TTM_SETTITLE = &H420
#If (UNICODE) Then
  TTM_ADDTOOL = TTM_ADDTOOLW
  TTM_DELTOOL = TTM_DELTOOLW
  TTM_NEWTOOLRECT = TTM_NEWTOOLRECTW
  TTM_GETTOOLINFO = TTM_GETTOOLINFOW
  TTM_SETTOOLINFO = TTM_SETTOOLINFOW
  TTM_HITTEST = TTM_HITTESTW
  TTM_GETTEXT = TTM_GETTEXTW
  TTM_UPDATETIPTEXT = TTM_UPDATETIPTEXTW
  TTM_ENUMTOOLS = TTM_ENUMTOOLSW
  TTM_GETCURRENTTOOL = TTM_GETCURRENTTOOLW
#Else
  TTM_ADDTOOL = TTM_ADDTOOLA
  TTM_DELTOOL = TTM_DELTOOLA
  TTM_NEWTOOLRECT = TTM_NEWTOOLRECTA
  TTM_GETTOOLINFO = TTM_GETTOOLINFOA
  TTM_SETTOOLINFO = TTM_SETTOOLINFOA
  TTM_HITTEST = TTM_HITTESTA
  TTM_GETTEXT = TTM_GETTEXTA
  TTM_UPDATETIPTEXT = TTM_UPDATETIPTEXTA
  TTM_ENUMTOOLS = TTM_ENUMTOOLSA
  TTM_GETCURRENTTOOL = TTM_GETCURRENTTOOLA
#End If
End Enum

Public Enum TOOLTIP_NOTIFICATIONS
  TTN_FIRST = -520&               '(0U-520U)
  TTN_LAST = -549&                '(0U-549U)

  TTN_GETDISPINFOA = (TTN_FIRST - 0)
  TTN_GETDISPINFOW = (TTN_FIRST - 10)
  TTN_SHOW = (TTN_FIRST - 1)
  TTN_POP = (TTN_FIRST - 2)

#If (UNICODE) Then
  TTN_GETDISPINFO = TTN_GETDISPINFOW
#Else
  TTN_GETDISPINFO = TTN_GETDISPINFOA
#End If

  TTN_NEEDTEXT = TTN_GETDISPINFO
  TTN_NEEDTEXTA = TTN_GETDISPINFOA
  TTN_NEEDTEXTW = TTN_GETDISPINFOW
End Enum

Public Enum TOOLTIP_ICON
  TTI_NO_ICON = &H0
  TTI_INFO_ICON = &H1
  TTI_WARN_ICON = &H2
  TTI_ERROR_ICON = &H3
End Enum

Public Enum DWORD
  HIWORD_MASK = &H10000
  LOWORD_MASK = &HFFFF&
End Enum

Public Enum TOOLTIP_ACTIVATION_DIRECTION
  TAD_LEFT = &H0                              ' Activate the tooltip to the left of the control...
  TAD_RIGHT = &H1                             ' Activate the tooltip to the right of the control...
  TAD_CENTER = &H2                            ' Activate the tooltip to the center of the control...
  TAD_CUSTOM = &H3                            ' Activate the tooltip to the custom point of the control...
End Enum
Rem --This ends the section for tooltips--

Rem --The next section contains members that have been used for members to system tray--
Public Enum TRAY_ICONFLAGS
  NIF_ICON = &H2                              ' Show the icon in the status bar...
  NIF_MESSAGE = &H1                           ' Check the Callback member...
  NIF_TIP = &H4                               ' Consider the tip member for tooltip...
  NIF_INFO = &H10                             ' Use a balloon tooltip instead of a simple one...
  NIF_STATE = &H8                             ' WinXP uses it to show hidden icons...
' NIF_GUID                                    ' The value could not be found for this member.
                                              ' It's reserved.
End Enum

Public Enum TRAY_MESSAGES
  NIM_ADD = &H0                               ' Add icon to system tray...
  NIM_DELETE = &H2                            ' Remove icon from system tray...
  NIM_MODIFY = &H1                            ' Modify icon present in system tray...
  NIM_SETFOCUS = &H4                          ' Position focus on icon in system tray...
  NIM_SETVERSION = &H8                        ' Use the version flag specified in the structure...
End Enum

Public Enum TRAY_MESSAGEICONTYPE
'  NIIF_ERROR = &H10                           ' Shows up an exclamation icon in the system tray
'                                              ' balloon tooltip...
'  NIIF_INFO = &H40                            ' Information...
'  NIIF_NONE = &H0                             ' No icon. This is default.
'  NIIF_WARNING = &H30                         ' Shows up a critical icon in the balloon tooltip...

  NIIF_ERROR = &H3                             ' Shows up an exclamation icon in the system tray
                                               ' balloon tooltip...
  NIIF_INFO = &H1                              ' Information...
  NIIF_NONE = &H0                              ' No icon. This is default.
  NIIF_WARNING = &H2                           ' Shows up a critical icon in the balloon tooltip...
End Enum

Public Enum TRAY_ICONSTATES
  NIS_HIDDEN = &H1                            ' WinXP style hidden icon...
  NIS_SHAREDICON = &H2                        ' WinXP style shared icon...
End Enum

Rem --This ends the section for system tray members--

Rem --The next section contains some specific enumerations for detecting specific OS Major Version Numbers --
Public Enum MAJORVERSION_NUMBERS
  MAJORVERSION_WIN2000 = &H5
  MAJORVERSION_WIN95 = &H4
  MAJORVERSION_WIN98 = &H4
  MAJORVERSION_WINME = &H4
  MAJORVERSION_WINNETSERVER = &H5
  MAJORVERSION_WINNT351 = &H3
  MAJORVERSION_WINNT400 = &H4
  MAJORVERSION_WINXP = &H5
End Enum

Public Enum MINORVERSION_NUMBERS
  MINORVERSION_WIN2000 = &H0            ' 0
  MINORVERSION_WIN95 = &H0              ' 0
  MINORVERSION_WIN98 = &HA              ' 10
  MINORVERSION_WINME = &H5A             ' 90
  MINORVERSION_WINNETSERVER = &H1       ' 1
  MINORVERSION_WINNT351 = &H33          ' 51
  MINORVERSION_WINNT400 = &H0           ' 0
  MINORVERSION_WINXP = &H1              ' 1
End Enum

Public Enum PLATFORM_ID
  VER_PLATFORM_WIN32_NT = &H2&          ' Identifies the operating system platform. This member can be VER_PLATFORM_WIN32_NT.
End Enum

Public Enum SUITE_MASKS
  VER_SUITE_BACKOFFICE = &H4           ' Microsoft BackOffice components are installed. 0 x00000004 (VER_SUITE_BACKOFFICE)
  VER_SUITE_DATACENTER = &H80          ' Windows 2000 DataCenter Server is installed. 0 x00000080 (VER_SUITE_DATACENTER)
  VER_SUITE_PERSONAL = &H200           ' Windows XP: Windows XP Home Edition is installed. 0 x00000200 (VER_SUITE_PERSONAL)
  VER_SUITE_ENTERPRISE = &H2           ' Windows 2000 Advanced Server or Windows .NET Enterprise Server is installed. 0 x00000002 (VER_SUITE_ENTERPRISE)
  VER_SUITE_SMALLBUSINESS = &H1        ' Microsoft Small Business Server is installed. 0 x00000001 (VER_SUITE_SMALLBUSINESS)
  VER_SUITE_SMALLBUSINESS_RESTRICTED = &H20 ' Microsoft Small Business Server is installed with the restrictive client license in force.
                                       ' 0 x00000020 (VER_SUITE_SMALLBUSINESS_RESTRICTED)
  VER_SUITE_TERMINAL = &H10            ' Terminal Services is installed. 0 x00000010 (VER_SUITE_TERMINAL)
  VER_SUITE_COMMNICATIONS = &H8        ' 0 x00000008 (VER_SUITE_COMMUNICATIONS)
  VER_SUITE_EMBEDDEDNT = &H40          ' 0 x00000040 (VER_SUITE_EMBEDDEDNT)
  VER_SUITE_SINGLEUSERTS = &H100       ' 0 x00000100 (VER_SUITE_SINGLEUSERTS)
  VER_SUITE_SERVERAPPLIANCE = &H400    ' 0 x00000400 (VER_SUITE_SERVERAPPLIANCE)
End Enum

Public Enum PRODUCT_TYPE
  VER_NT_WORKSTATION = &H1              ' The system is running Windows NT 4.0 Workstation, Windows 2000 Professional, Windows XP Home
                                        ' Edition, or Windows XP Professional. 0 x0000001 (VER_NT_WORKSTATION)
  VER_NT_DOMAIN_CONTROLLER = &H2        ' The system is a domain controller. 0 x0000002 (VER_NT_DOMAIN_CONTROLLER)
  VER_NT_SERVER = &H3                   ' The system is a server. 0 x0000003 (VER_NT_SERVER)
End Enum
Rem --This ends the section for OS members--

Rem --This section deals with enumerations required for retrieving folder paths--
Rem The following enumerations are used to retrieve special folder paths for a given user...
Public Enum CSIDL
  CSIDL_DESKTOP = &H0                       ' Windows Desktop-virtual folder that is the root of the namespace.
  CSIDL_INTERNET = &H1                      ' Virtual folder representing the Internet.
  CSIDL_PROGRAMS = &H2                      ' File system directory that contains the user's program groups (which are also file system directories).
                                            ' A typical path is C:\Documents and Settings\username\Start Menu\Programs.
  CSIDL_CONTROLS = &H3                      ' Virtual folder containing icons for the Control Panel applications.
  CSIDL_PRINTERS = &H4                      ' Virtual folder containing installed printers.
  CSIDL_PERSONAL = &H5                      ' File system directory that serves as a common repository for documents. A typical path is C:\Documents
                                            ' and Settings\username\My Documents. This should be distinguished from the virtual My Documents folder in
                                            ' the namespace. To access that virtual folder, use the technique described in Managing the File System.
  CSIDL_FAVORITES = &H6                     ' File system directory that serves as a common repository for the user's favorite items. A typical path is
                                            ' C:\Documents and Settings\username\Favorites.
  CSIDL_STARTUP = &H7                       ' File system directory that corresponds to the user's Startup program group. The system starts these programs
                                            ' whenever any user logs onto Windows NTR or starts WindowsR 95. A typical path is
                                            ' C:\Documents and Settings\username\Start Menu\Programs\Startup.
  CSIDL_RECENT = &H8                        ' File system directory that contains the user's most recently used documents. A typical path is
                                            ' C:\Documents and Settings\username\Recent. To create a shortcut in this folder, use SHAddToRecentDocs. In
                                            ' addition to creating the shortcut, this function updates the Shell's list of recent documents and adds the
                                            ' shortcut to the Documents submenu of the Start menu.
  CSIDL_SENDTO = &H9                        ' File system directory that contains Send To menu items. A typical path is C:\Documents and Settings\username\SendTo.
  CSIDL_BITBUCKET = &HA                     ' Virtual folder containing the objects in the user's Recycle Bin.
  CSIDL_STARTMENU = &HB                     ' File system directory containing Start menu items. A typical path is C:\Documents and Settings\username\Start Menu.
  CSIDL_DESKTOPDIRECTORY = &H10             ' File system directory used to physically store file objects on the desktop (not to be confused with the desktop folder
                                            ' itself). A typical path is C:\Documents and Settings\username\Desktop
  CSIDL_DRIVES = &H11                       ' My Computer?virtual folder containing everything on the local computer: storage devices, printers, and Control Panel.
                                            ' The folder may also contain mapped network drives.
  CSIDL_NETWORK = &H12                      ' Network Neighborhood?virtual folder representing the root of the network namespace hierarchy.
  CSIDL_NETHOOD = &H13                      ' A file system folder containing the link objects that may exist in the My Network Places virtual folder. It
                                            ' is not the same as CSIDL_NETWORK, which represents the network namespace root.
                                            ' A typical path is C:\Documents and Settings\username\NetHood
  CSIDL_FONTS = &H14                        ' Virtual folder containing fonts. A typical path is C:\WINNT\Fonts.
  CSIDL_TEMPLATES = &H15                    ' File system directory that serves as a common repository for document templates.
  CSIDL_COMMON_STARTMENU = &H16             ' File system directory that contains the programs and folders that appear on the Start menu for all users. A
                                            ' typical path is C:\Documents and Settings\All Users\Start Menu. Valid only for Windows NT(R) systems.
  CSIDL_COMMON_PROGRAMS = &H17              ' File system directory that contains the directories for the common program groups that appear on the Start menu
                                            ' for all users. A typical path is C:\Documents and Settings\All Users\Start Menu\Programs. Valid only for Windows NT(R) systems.
  CSIDL_COMMON_STARTUP = &H18               ' File system directory that contains the programs that appear in the Startup folder for all users. A typical path
                                            ' is C:\Documents and Settings\All Users\Start Menu\Programs\Startup. Valid only for Windows NT(R) systems.
  CSIDL_COMMON_DESKTOPDIRECTORY = &H19      ' File system directory that contains files and folders that appear on the desktop for all users.
                                            ' A typical path is C:\Documents and Settings\All Users\Desktop. Valid only for Windows NT(R) systems.
  CSIDL_APPDATA = &H1A                      ' Version 4.71. File system directory that serves as a common repository for application-specific data.
                                            ' A typical path is C:\Documents and Settings\username\Application Data.
                                            ' This CSIDL is supported by the redistributable ShFolder.dll for systems that do not have the Internet Explorer 4.0
                                            ' integrated Shell installed.
  CSIDL_LOCAL_APPDATA = &H1C                '{user}\Local Settings\Application Data (non roaming)
  CSIDL_PRINTHOOD = &H1B                    ' File system directory that contains the link objects that may exist in the Printers virtual folder.
                                            ' A typical path is C:\Documents and Settings\username\PrintHood.
  CSIDL_ALTSTARTUP = &H1D ' DBCS            ' File system directory that corresponds to the user's nonlocalized Startup program group.
  CSIDL_COMMON_ALTSTARTUP = &H1E ' DBCS     ' File system directory that corresponds to the nonlocalized Startup program group for all users. Valid only for Windows NT(R) systems.
  CSIDL_COMMON_FAVORITES = &H1F             ' File system directory that serves as a common repository for all users' favorite items. Valid only for Windows NT(R) systems.
  CSIDL_INTERNET_CACHE = &H20               ' Version 4.72. File system directory that serves as a common repository for temporary Internet files.
                                            ' A typical path is C:\Documents and Settings\username\Temporary Internet Files.
  CSIDL_COOKIES = &H21                      ' File system directory that serves as a common repository for Internet cookies. A typical path is C:\Documents and Settings\username\Cookies.
  CSIDL_HISTORY = &H22                      ' File system directory that serves as a common repository for Internet history items.
  CSIDL_COMMON_APPDATA = &H23               ' All Users\Application Data
  CSIDL_WINDOWS = &H24                      ' GetWindowsDirectory()
  CSIDL_SYSTEM = &H25                       ' GetSystemDirectory()
  CSIDL_PROGRAM_FILES = &H26                ' C:\Program Files
  CSIDL_MYPICTURES = &H27                   ' C:\Program Files\My Pictures
  CSIDL_PROFILE = &H28                      ' USERPROFILE
  CSIDL_SYSTEMX86 = &H29                    ' x86 system directory on RISC
  CSIDL_PROGRAM_FILESX86 = &H2A             ' x86 C:\Program Files on RISC
  CSIDL_PROGRAM_FILES_COMMON = &H2B         ' C:\Program Files\Common
  CSIDL_PROGRAM_FILES_COMMONX86 = &H2C      ' x86 Program Files\Common on RISC
  CSIDL_COMMON_TEMPLATES = &H2D             ' All Users\Templates
  CSIDL_COMMON_DOCUMENTS = &H2E             ' All Users\Documents
  CSIDL_COMMON_ADMINTOOLS = &H2F            ' All Users\Start Menu\Programs\Administrative Tools
  CSIDL_ADMINTOOLS = &H30                   ' {user}\Start Menu\Programs\Administrative Tools
End Enum

Public Enum CSIDL_FOLDERPATH
  SHGFP_TYPE_CURRENT = &H0                  ' Return the folder's current path.
  SHGFP_TYPE_DEFAULT = &H1                  ' Return the folder's default path.
End Enum

Public Enum CSIDL_MASK
  CSIDL_FLAG_CREATE = &H8000&               ' combine with CSIDL_ value to force create on SHGetSpecialFolderLocation()
  CSIDL_FLAG_DONT_VERIFY = &H4000           ' combine with CSIDL_ value to force create on SHGetSpecialFolderLocation()
  CSIDL_FLAG_MASK = &HFF00                  ' mask for all possible flag values
End Enum
Rem --End of retrieving folder path's section--

Rem --The next section deals with retrieving and setting menu item bitmaps--
Rem +-------------------------------------+--------------------------------------------------------+
Rem | Value                               | Meaning                                                |
Rem +-------------------------------------+--------------------------------------------------------+
Public Enum MENU_ITEM_INFO_FLAGS '        |                                                        |
#If (WIN32_IE >= &H400) Then '            |                                                        |
      MIIM_STATE = &H1 '                  | Retrieves or sets the fState member.                   |
Rem | [0x00000001]                        |                                                        |
Rem +-------------------------------------+--------------------------------------------------------+
      MIIM_ID = &H2 '                     | Retrieves or sets the wID member.                      |
Rem | [0x00000002]                        |                                                        |
Rem +-------------------------------------+--------------------------------------------------------+
      MIIM_SUBMENU = &H4 '                | Retrieves or sets the hSubMenu member.                 |
Rem | [0x00000004]                        |                                                        |
Rem +-------------------------------------+--------------------------------------------------------+
      MIIM_CHECKMARKS = &H8 '             | Retrieves or sets the hbmpChecked and hbmpUnchecked    |
Rem | [0x00000008]                        | members.                                               |
Rem +-------------------------------------+--------------------------------------------------------+
      MIIM_TYPE = &H10 '                  | Retrieves or sets the fType and dwTypeData members.    |
Rem | [0x00000010]                        | Windows 98/Me, Windows 2000/XP: MIIM_TYPE is replaced  |
Rem |                                     | by MIIM_BITMAP, MIIM_FTYPE and MIIM_STRING.            |
Rem +-------------------------------------+--------------------------------------------------------+
      MIIM_DATA = &H20 '                  | Retrieves or sets the dwItemData member.               |
Rem | [0x00000020]                        |                                                        |
Rem +-------------------------------------+--------------------------------------------------------+
#End If '                                 |                                                        |
#If (WIN32_IE >= &H500) Then '            |                                                        |
      MIIM_STRING = &H40 '                | Windows 98/Me, Windows 2000/XP: Retrieves or sets the  |
Rem | [0x00000040]                        | dwTypeData member.                                     |
Rem +-------------------------------------+--------------------------------------------------------+
      MIIM_BITMAP = &H80 '                | Windows 98/Me, Windows 2000/XP: Retrieves or sets the  |
Rem | [0x00000080]                        | hbmpItem member.                                       |
Rem +-------------------------------------+--------------------------------------------------------+
      MIIM_FTYPE = &H100 '                | Windows 98/Me, Windows 2000/XP: Retrieves or sets the  |
Rem | [0x00000100]                        | fType member.                                          |
Rem +-------------------------------------+--------------------------------------------------------+
#End If '                                 |                                                        |
End Enum '                                |                                                        |
Rem +-------------------------------------+--------------------------------------------------------+

Rem +-------------------------------------+--------------------------------------------------------+
Rem | Value                               | Meaning                                                |
Rem +-------------------------------------+--------------------------------------------------------+
Public Enum MENU_ITEM_TYPE_FLAGS '        |                                                        |
#If (WIN32_IE >= &H400) Then '            |                                                        |
      MFT_BITMAP = &H4& '                 | Displays the menu item using a bitmap. The low-order   |
Rem | [0x00000004L]                       | word of the dwTypeData member is the bitmap handle, and|
Rem |                                     | the cch member is ignored.                             |
Rem |                                     | Windows 98/Me, Windows 2000/XP: MFT_BITMAP is replaced |
Rem |                                     | by MIIM_BITMAP and hbmpItem                            |
Rem +-------------------------------------+--------------------------------------------------------+
      MFT_STRING = &H0& '                 | Displays the menu item using a text string. The        |
Rem | [0x00000000L]                       | dwTypeData member is the pointer to a null-terminated  |
Rem |                                     | string, and the cch member is the length of the string.|
Rem |                                     | Windows 98/Me, Windows 2000/XP: MFT_STRING is replaced |
Rem |                                     | by MIIM_STRING                                         |
Rem +-------------------------------------+--------------------------------------------------------+
      MFT_MENUBARBREAK = &H20& '          | Places the menu item on a new line (for a menu bar) or |
Rem | [0x00000020L]                       | in a new column (for a drop-down menu, submenu, or     |
Rem |                                     | shortcut menu). For a drop-down menu, submenu, or      |
Rem |                                     | shortcut menu, a vertical line separates the new column|
Rem |                                     | from the old.                                          |
Rem +-------------------------------------+--------------------------------------------------------+
      MFT_MENUBREAK = &H40& '             | Places the menu item on a new line (for a menu bar) or |
Rem | [0x00000040L]                       | in a new column (for a drop-down menu, submenu, or     |
Rem |                                     | shortcut menu). For a drop-down menu, submenu, or      |
Rem |                                     | shortcut menu, the columns are not separated by a      |
Rem |                                     | vertical line.                                         |
Rem +-------------------------------------+--------------------------------------------------------+
      MFT_OWNERDRAW = &H100& '            | Assigns responsibility for drawing the menu item to the|
Rem | [0x00000100L]                       | window that owns the menu. The window receives a       |
Rem |                                     | WM_MEASUREITEM message before the menu is displayed for|
Rem |                                     | the first time, and a WM_DRAWITEM message whenever the |
Rem |                                     | appearance of the menu item must be updated. If this   |
Rem |                                     | value is specified, the dwTypeData member contains an  |
Rem |                                     | application-defined value.                             |
Rem +-------------------------------------+--------------------------------------------------------+
      MFT_RADIOCHECK = &H200& '           | Displays selected menu items using a radio-button mark |
Rem | [0x00000200L]                       | instead of a check mark if the hbmpChecked member is   |
Rem |                                     | NULL.                                                  |
Rem +-------------------------------------+--------------------------------------------------------+
      MFT_SEPARATOR = &H800& '            | Specifies that the menu item is a separator. A menu    |
Rem | [0x00000800L]                       | item separator appears as a horizontal dividing line.  |
Rem |                                     | The dwTypeData and cch members are ignored. This value |
Rem |                                     | is valid only in a drop-down menu, submenu, or shortcut|
Rem |                                     | menu.                                                  |
Rem +-------------------------------------+--------------------------------------------------------+
      MFT_RIGHTORDER = &H2000& '          | Windows 95/98/Me, Windows 2000/XP: Specifies that menus|
Rem | [0x00002000L]                       | cascade right-to-left (the default is left-to-right).  |
Rem |                                     | This is used to support right-to-left languages, such  |
Rem |                                     | as Arabic and Hebrew.                                  |
Rem +-------------------------------------+--------------------------------------------------------+
      MFT_RIGHTJUSTIFY = &H4000& '        | Right-justifies the menu item and any subsequent items.
Rem | [0x00004000L]                       | This value is valid only if the menu item is in a menu
Rem |                                     | bar.
Rem +-------------------------------------+--------------------------------------------------------+
      MFS_GRAYED = &H3& '                 | Disables a menu item.                                  |
Rem | [0x00000003L]                       |                                                        |
Rem +-------------------------------------+--------------------------------------------------------+
      MFS_DISABLED = MFS_GRAYED '         | Disables a menu item.                                  |
Rem +-------------------------------------+--------------------------------------------------------+
      MFS_CHECKED = &H8& '                | Checks the menu item. Handle to the bitmap to display  |
Rem | [0x00000008L]                       | next to the item if it is checked. If this member is   |
Rem |                                     | NULL, a default bitmap is used. If the MFT_RADIOCHECK  |
Rem |                                     | type value is specified, the default bitmap is a bullet.
Rem |                                     | Otherwise, it is a check mark.                         |
Rem +-------------------------------------+--------------------------------------------------------+
      MFS_HILITE = &H80& '                | Highlights the menu item.                              |
Rem | [0x00000080L]                       |                                                        |
Rem +-------------------------------------+--------------------------------------------------------+
      MFS_ENABLED = &H0& '                | Enables the menu item so that it can be selected. This |
Rem | [0x00000000L]                       | is the default state.                                  |
Rem +-------------------------------------+--------------------------------------------------------+
      MFS_UNCHECKED = &H0& '              | Unchecks the menu item. Handle to the bitmap to display|
Rem | [0x00000000L]                       | next to the item if it is not checked. If this member is
Rem |                                     | NULL, no bitmap is used.                               |
Rem +-------------------------------------+--------------------------------------------------------+
      MFS_UNHILITE = &H0& '               | Removes the highlight from the menu item. This is the  |
Rem | [0x00000000L]                       | default state.                                         |
Rem +-------------------------------------+--------------------------------------------------------+
      MFS_DEFAULT = &H1000& '             | Default menu item state.                               |
Rem | [0x00001000L]                       |                                                        |
Rem +-------------------------------------+--------------------------------------------------------+
#If (WIN32_IE >= &H500) Then '            |                                                        |
      MFS_MASK = &H108B& '                |                                                        |
Rem | [0x0000108BL]                       |                                                        |
Rem +-------------------------------------+--------------------------------------------------------+
      MFS_HOTTRACKDRAWN = &H10000000 '    |                                                        |
Rem | [0x10000000L]                       |                                                        |
Rem +-------------------------------------+--------------------------------------------------------+
      MFS_CACHEDBMP = &H20000000 '        |                                                        |
Rem | [0x20000000L]                       |                                                        |
Rem +-------------------------------------+--------------------------------------------------------+
      MFS_BOTTOMGAPDROP = &H40000000 '    |                                                        |
Rem | [0x40000000L]                       |                                                        |
Rem +-------------------------------------+--------------------------------------------------------+
      MFS_TOPGAPDROP = &H80000000 '       |                                                        |
Rem | [0x80000000L]                       |                                                        |
Rem +-------------------------------------+--------------------------------------------------------+
      MFS_GAPDROP = &HC0000000 '          |                                                        |
Rem | [0xC0000000L]                       |                                                        |
Rem +-------------------------------------+--------------------------------------------------------+
#End If '/* WINVER >= 0x0500 */           |                                                        |
#End If '/* WINVER >= 0x0400 */           |                                                        |
End Enum '                                |                                                        |
Rem +-------------------------------------+--------------------------------------------------------+

Rem +-------------------------------------+--------------------------------------------------------+
Rem | Value                               | Meaning                                                |
Rem +-------------------------------------+--------------------------------------------------------+
Public Enum GMDI_FLAGS '                 |                                                        |
      GMDI_USEDEFAULT = &H0&   '          |  Specifies to use default search for menu items.       |
Rem | [0x0000L]                           |                                                        |
      GMDI_GOINTOPOPUPS = &H1& '          |  Specifies that if the default item is one that opens a|
Rem | [0x0001L]                           |  submenu, the function is to search recursively in the |
Rem |                                     |  corresponding submenu. If the submenu has no default  |
Rem |                                     |  item, the return value identifies the item that opens |
Rem |                                     |  the submenu.                                          |
Rem |                                     |  By default, the function returns the first default    |
Rem |                                     |  item on the specified menu, regardless of whether it  |
Rem |                                     |  is an item that opens a submenu.                      |
Rem +-------------------------------------+--------------------------------------------------------+
      GMDI_USEDISABLED = &H2&  '          |   Specifies that the function is to return a default   |
Rem | [0x0002L]                           |   item, even if it is disabled. By default, the function
Rem |                                     |   skips disabled or grayed items.                      |
End Enum '                                |                                                        |
Rem +-------------------------------------+--------------------------------------------------------+

Rem +-------------------------------------+--------------------------------------------------------+
Rem | Value                               | Meaning                                                |
Rem +-------------------------------------+--------------------------------------------------------+
Public Enum SETMENUITEMBITMAPFLAGS '      |                                                        |
      MF_BYCOMMAND = &H0&      '          |  Specifies to use default search for menu items.       |
Rem +-------------------------------------+--------------------------------------------------------+
      MF_BYPOSITION = &H400&   '          |  Specifies that if the default item is one that opens a|
Rem |                                     |   skips disabled or grayed items.                      |
End Enum '                                |                                                        |
Rem +-------------------------------------+--------------------------------------------------------+

Rem --This ends the section for menu item bitmaps--

Rem --This section holds enumerations used for animating windows--
Public Enum ANIMATE_WINDOW_STYLE
  AW_SLIDE = &H40000            ' Uses slide animation. By default, roll animation is used. This flag is
                                ' ignored when used with AW_CENTER.
  AW_ACTIVATE = &H20000         ' Activates the window. Do not use this value with AW_HIDE.
  AW_BLEND = &H80000            ' Uses a fade effect. This flag can be used only if hwnd is a top-level window.
  AW_HIDE = &H10000             ' Hides the window. By default, the window is shown.
  AW_CENTER = &H10              ' Makes the window appear to collapse inward if AW_HIDE is used or expand
                                ' outward if the AW_HIDE is not used.
  AW_HOR_POSITIVE = &H1         ' Animates the window from left to right. This flag can be used with roll or
                                ' slide animation. It is ignored when used with AW_CENTER or AW_BLEND.
  AW_HOR_NEGATIVE = &H2         ' Animates the window from right to left. This flag can be used with roll or
                                ' slide animation. It is ignored when used with AW_CENTER or AW_BLEND.
  AW_VER_POSITIVE = &H4         ' Animates the window from top to bottom. This flag can be used with roll or
                                ' slide animation. It is ignored when used with AW_CENTER or AW_BLEND.
  AW_VER_NEGATIVE = &H8         ' Animates the window from bottom to top. This flag can be used with roll or
                                ' slide animation. It is ignored when used with AW_CENTER or AW_BLEND.
End Enum
Rem --End of animating windows section--

Rem --The next section deals with system menus--
Public Enum ENABLE_MENU_ITEM_FLAGS
  MF_DISABLED = &H2&            ' Indicates that the menu item is disabled, but not grayed, so it cannot be selected.
  MF_ENABLED = &H0&             ' Indicates that the menu item is enabled and restored from a grayed state so that it can be selected.
  MF_GRAYED = &H1&              ' Indicates that the menu item is disabled and grayed so that it cannot be selected.
  MF_BITMAP = &H4&              ' Uses a bitmap as the menu item.
  MF_POPUP = &H10&              ' Specifies that the menu item opens a drop-down menu or submenu.
End Enum
Rem --This ends the section for system menus--

Rem --The next section deals with MCI devices--
Public Enum MCI_COMMANDS
  MCI_DEVICE_DOOR_OPEN = &H0
  MCI_DEVICE_DOOR_CLOS = &H1
End Enum
Rem --This ends the section for MCI devices--

Rem --The next section deals with owner drawn menus--
Public Enum OWNER_DRAWN_CONTROL_TYPE
  ODT_MENU = 1                  ' Owner-drawn menu item.
  ODT_LISTBOX = 2               ' Owner-drawn list box.
  ODT_COMBOBOX = 3              ' Owner-drawn combo box.
  ODT_BUTTON = 4                ' Owner-drawn button.
#If (WIN32_IE >= &H400) Then
  ODT_STATIC = 5                ' Owner-drawn static control.
  ODT_LISTVIEW = 6              ' List view control.
  ODT_TAB = 7                   ' Tab control.
#End If ' /* WINVER >= 0x0400 */
End Enum

Public Enum OWNER_DRAWN_CONTROL_ACTION
  ODA_DRAWENTIRE = &H1          ' The entire control needs to be drawn.
  ODA_SELECT = &H2              ' The control has lost or gained the keyboard focus. The itemState
                                ' member should be checked to determine whether the control has the
                                ' focus.
  ODA_FOCUS = &H4               ' The selection status has changed. The itemState member should be
                                ' checked to determine the new selection state
End Enum

Public Enum OWNER_DRAWN_CONTROL_STATE
  ODS_SELECTED = &H1            ' The menu item's status is selected.
  ODS_GRAYED = &H2              ' The item is to be grayed. This bit is used only in a menu.
  ODS_DISABLED = &H4            ' The item is to be drawn as disabled.
  ODS_CHECKED = &H8             ' The menu item is to be checked. This bit is used only in a menu.
  ODS_FOCUS = &H10              ' The item has the keyboard focus.
#If (WIN32_IE >= &H400) Then
  ODS_DEFAULT = &H20            ' The item is the default item.
  ODS_COMBOBOXEDIT = &H1000     ' The drawing takes place in the selection field (edit control) of an owner-drawn combo box.
#End If ' /* WINVER >= 0x0400 */
#If (WIN32_IE >= &H500) Then
  ODS_HOTLIGHT = &H40           ' Windows 98/Me, Windows 2000/XP: The item is being hot-tracked, that is
                                ' , the item will be highlighted when the mouse is on the item.
  ODS_INACTIVE = &H80           ' Windows 98/Me, Windows 2000/XP: The item is inactive and the window
                                ' associated with the menu is inactive.
  ODS_NOACCEL = &H100                ' Windows 2000/XP: The control is drawn without the keyboard accelerator
                                ' cues.
  ODS_NOFOCUSRECT = &H200       ' Windows 2000/XP: The control is drawn without focus indicator cues.
#End If ' /* WINVER >= 0x0500 */
End Enum

Public Enum COLOR_INDEX
  COLOR_SCROLLBAR = 0          ' Scroll bar gray area.
  COLOR_BACKGROUND = 1         ' Desktop.
  COLOR_ACTIVECAPTION = 2      ' Active window title bar.
  COLOR_INACTIVECAPTION = 3    ' Inactive window caption.
  COLOR_MENU = 4               ' Menu background.
  COLOR_MENUBAR = 30           ' Menubar color.
  COLOR_WINDOW = 5             ' Window background.
  COLOR_WINDOWFRAME = 6        ' Window frame.
  COLOR_MENUTEXT = 7           ' Text in menus.
  COLOR_WINDOWTEXT = 8         ' Text in windows.
  COLOR_CAPTIONTEXT = 9        ' Text in caption, size box, and scroll bar arrow box.
  COLOR_ACTIVEBORDER = 10      ' Active window border.
  COLOR_INACTIVEBORDER = 11    ' Inactive window border.
  COLOR_APPWORKSPACE = 12      ' Background color of multiple document interface (MDI) applications.
  COLOR_HIGHLIGHT = 13         ' Item(s) selected in a control.
  COLOR_HIGHLIGHTTEXT = 14     ' Text of item(s) selected in a control.
  COLOR_BTNFACE = 15           ' Face color for three-dimensional display elements.
  COLOR_BTNSHADOW = 16         ' Shadow color for three-dimensional display elements (for edges facing
                               ' away from the light source).
  COLOR_GRAYTEXT = 17          ' Grayed (disabled) text. This color is set to 0 if the current display
                               ' driver does not support a solid gray color.
  COLOR_BTNTEXT = 18           ' Text on push buttons.
  COLOR_INACTIVECAPTIONTEXT = 19 ' Color of text in an inactive caption.
  COLOR_BTNHIGHLIGHT = 20      ' Highlight color for three-dimensional display elements (for edges
                               ' facing the light source.)
#If (WIN32_IE >= &H400) Then
  COLOR_3DDKSHADOW = 21        ' Dark shadow for three-dimensional display elements.
  COLOR_3DLIGHT = 22           ' Light color for three-dimensional display elements (for edges facing
                               ' the light source.)
  COLOR_INFOTEXT = 23          ' Text color for tooltip controls.
  COLOR_INFOBK = 24            ' Background color for tooltip controls.
#End If ' /* WINVER >= 0x0400 */

#If (WIN32_IE >= &H500) Then
  COLOR_HOTLIGHT = 26          ' Windows 98/Me, Windows 2000/XP: Color for a hot-tracked item. Single
                               ' clicking a hot-tracked item executes the item.
  COLOR_GRADIENTACTIVECAPTION = 27 ' Windows 98/Me, Windows 2000/XP: Right side color in the color
                               ' gradient of an active window's title bar. COLOR_ACTIVECAPTION specifies
                               ' the left side color. Use SPI_GETGRADIENTCAPTIONS with the SystemParametersInfo
                               ' function to determine whether the gradient effect is enabled.
  COLOR_GRADIENTINACTIVECAPTION = 28 ' Windows 98/Me, Windows 2000/XP: Right side color in the color
                               ' gradient of an inactive window's title bar. COLOR_INACTIVECAPTION
                               ' specifies the left side color.
#End If ' /* WINVER >= 0x0500 */

#If (WIN32_IE >= &H400) Then
  COLOR_DESKTOP = COLOR_BACKGROUND
  COLOR_3DFACE = COLOR_BTNFACE
  COLOR_3DSHADOW = COLOR_BTNSHADOW
  COLOR_3DHIGHLIGHT = COLOR_BTNHIGHLIGHT
  COLOR_3DHILIGHT = COLOR_BTNHIGHLIGHT
  COLOR_BTNHILIGHT = COLOR_BTNHIGHLIGHT
#End If ' /* WINVER >= 0x0400 */
End Enum

Public Enum SYSTEM_METRICS_ITEM           ' All SM_CX* values are widths. All SM_CY* values are heights.
  SM_CXSCREEN = 0
  SM_CYSCREEN = 1
  SM_CXVSCROLL = 2
  SM_CYHSCROLL = 3
  SM_CYCAPTION = 4
  SM_CXBORDER = 5
  SM_CYBORDER = 6
  SM_CXDLGFRAME = 7
  SM_CYDLGFRAME = 8
  SM_CYVTHUMB = 9
  SM_CXHTHUMB = 10
  SM_CXICON = 11
  SM_CYICON = 12
  SM_CXCURSOR = 13
  SM_CYCURSOR = 14
  SM_CYMENU = 15
  SM_CXFULLSCREEN = 16
  SM_CYFULLSCREEN = 17
  SM_CYKANJIWINDOW = 18
  SM_MOUSEPRESENT = 19
  SM_CYVSCROLL = 20
  SM_CXHSCROLL = 21
  SM_DEBUG = 22
  SM_SWAPBUTTON = 23
  SM_RESERVED1 = 24
  SM_RESERVED2 = 25
  SM_RESERVED3 = 26
  SM_RESERVED4 = 27
  SM_CXMIN = 28
  SM_CYMIN = 29
  SM_CXSIZE = 30
  SM_CYSIZE = 31
  SM_CXFRAME = 32
  SM_CYFRAME = 33
  SM_CXMINTRACK = 34
  SM_CYMINTRACK = 35
  SM_CXDOUBLECLK = 36
  SM_CYDOUBLECLK = 37
  SM_CXICONSPACING = 38
  SM_CYICONSPACING = 39
  SM_MENUDROPALIGNMENT = 40
  SM_PENWINDOWS = 41
  SM_DBCSENABLED = 42
  SM_CMOUSEBUTTONS = 43

#If (WIN32_IE >= &H400) Then
  SM_CXFIXEDFRAME = SM_CXDLGFRAME ' /*=;win40=name=change=*/
  SM_CYFIXEDFRAME = SM_CYDLGFRAME '/*=;win40=name=change=*/
  SM_CXSIZEFRAME = SM_CXFRAME     '/*=;win40=name=change=*/
  SM_CYSIZEFRAME = SM_CYFRAME     '/*=;win40=name=change=*/

  SM_SECURE = 44
  SM_CXEDGE = 45
  SM_CYEDGE = 46
  SM_CXMINSPACING = 47
  SM_CYMINSPACING = 48
  SM_CXSMICON = 49
  SM_CYSMICON = 50
  SM_CYSMCAPTION = 51
  SM_CXSMSIZE = 52
  SM_CYSMSIZE = 53
  SM_CXMENUSIZE = 54
  SM_CYMENUSIZE = 55
  SM_ARRANGE = 56
  SM_CXMINIMIZED = 57
  SM_CYMINIMIZED = 58
  SM_CXMAXTRACK = 59
  SM_CYMAXTRACK = 60
  SM_CXMAXIMIZED = 61
  SM_CYMAXIMIZED = 62
  SM_NETWORK = 63
  SM_CLEANBOOT = 67
  SM_CXDRAG = 68
  SM_CYDRAG = 69
#End If '/*=WINVER=>==0x0400=*/

  SM_SHOWSOUNDS = 70
#If (WIN32_IE >= &H400) Then
  SM_CXMENUCHECK = 71           '/*=Use=instead=of=GetMenuCheckMarkDimensions()!=*/
  SM_CYMENUCHECK = 72
  SM_SLOWMACHINE = 73
  SM_MIDEASTENABLED = 74
#End If  '/*=WINVER=>==0x0400=*/

#If (WIN32_IE >= &H500) Then
  SM_MOUSEWHEELPRESENT = 75
#End If

#If (WIN32_IE >= &H500) Then
  SM_XVIRTUALSCREEN = 76
  SM_YVIRTUALSCREEN = 77
  SM_CXVIRTUALSCREEN = 78
  SM_CYVIRTUALSCREEN = 79
  SM_CMONITORS = 80
  SM_SAMEDISPLAYFORMAT = 81
#End If '/*=WINVER=>==0x0500=*/

#If (WIN32_IE <= &H500) Then '(!defined(_WIN32_WINNT)=||=(_WIN32_WINNT=<=0x0400))
  SM_CMETRICS = 76
#Else
  SM_CMETRICS = 83
#End If
End Enum

Public Enum SELECT_OBJECT_RETURN_VALUE
  NULLREGION = 1                  ' Region is empty.
  SIMPLEREGION = 2                ' Region consists of a single rectangle.
  COMPLEXREGION = 3               ' Region consists of more than one rectangle.
End Enum

Public Enum LOGICAL_FONT
  LF_FACESIZE = &H20
End Enum

Public Enum LOGICAL_FONT_WEIGHT
  FW_DONTCARE = 0
  FW_THIN = 100
  FW_EXTRALIGHT = 200
  FW_LIGHT = 300
  FW_NORMAL = 400
  FW_MEDIUM = 500
  FW_SEMIBOLD = 600
  FW_BOLD = 700
  FW_EXTRABOLD = 800
  FW_HEAVY = 900

  FW_ULTRALIGHT = FW_EXTRALIGHT
  FW_REGULAR = FW_NORMAL
  FW_DEMIBOLD = FW_SEMIBOLD
  FW_ULTRABOLD = FW_EXTRABOLD
  FW_BLACK = FW_HEAVY
End Enum

Public Enum LOGICAL_FONT_CHARSET
  ANSI_CHARSET = 0
  DEFAULT_CHARSET = 1             ' You can use the DEFAULT_CHARSET value to allow the name and size of
                                  ' a font to fully describe the logical font. If the specified font
                                  ' name does not exist, a font from any character set can be substituted
                                  ' for the specified font, so you should use DEFAULT_CHARSET sparingly
                                  ' to avoid unexpected results.
  SYMBOL_CHARSET = 2
  SHIFTJIS_CHARSET = 128
  HANGEUL_CHARSET = 129
  HANGUL_CHARSET = 129
  GB2312_CHARSET = 134
  CHINESEBIG5_CHARSET = 136
  OEM_CHARSET = 255
#If (WIN32_IE >= &H400) Then
  JOHAB_CHARSET = 130
  HEBREW_CHARSET = 177
  ARABIC_CHARSET = 178
  GREEK_CHARSET = 161
  TURKISH_CHARSET = 162
  VIETNAMESE_CHARSET = 163
  THAI_CHARSET = 222
  EASTEUROPE_CHARSET = 238
  RUSSIAN_CHARSET = 204

  MAC_CHARSET = 77
  BALTIC_CHARSET = 186

  FS_LATIN1 = &H1&
  FS_LATIN2 = &H2&
  FS_CYRILLIC = &H4&
  FS_GREEK = &H8&
  FS_TURKISH = &H10&
  FS_HEBREW = &H20&
  FS_ARABIC = &H40&
  FS_BALTIC = &H80&
  FS_VIETNAMESE = &H100&
  FS_THAI = &H10000
  FS_JISJAPAN = &H20000
  FS_CHINESESIMP = &H40000
  FS_WANSUNG = &H80000
  FS_CHINESETRAD = &H100000
  FS_JOHAB = &H200000
  FS_SYMBOL = &H80000000
#End If '=/*=WINVER=>==&H0400=*/
End Enum

Public Enum LOGICAL_FONT_OUTPUT_PRECISION
  OUT_DEFAULT_PRECIS = 0            ' Specifies the default font mapper behavior.
  OUT_STRING_PRECIS = 1             ' This value is not used by the font mapper, but it is returned when
                                    ' raster fonts are enumerated.
  OUT_CHARACTER_PRECIS = 2
  OUT_STROKE_PRECIS = 3
  OUT_TT_PRECIS = 4
  OUT_DEVICE_PRECIS = 5
  OUT_RASTER_PRECIS = 6             ' Instructs the font mapper to choose a raster font when the system
                                    ' contains multiple fonts with the same name.
  OUT_TT_ONLY_PRECIS = 7
  OUT_OUTLINE_PRECIS = 8
  OUT_SCREEN_OUTLINE_PRECIS = 9
End Enum

Public Enum LOGICAL_FONT_CLIP_PRECISION
  CLIP_DEFAULT_PRECIS = 0           ' Specifies default clipping behavior.
  CLIP_CHARACTER_PRECIS = 1         ' Not used.
  CLIP_STROKE_PRECIS = 2            ' Not used by the font mapper, but is returned when raster, vector,
                                    ' or TrueType fonts are enumerated.
  CLIP_MASK = &HF
  CLIP_LH_ANGLES = &H10&
  CLIP_TT_ALWAYS = &H20&
  CLIP_EMBEDDED = &H40&
End Enum

Public Enum LOGICAL_FONT_OUTPUT_QUALITY
  DEFAULT_QUALITY = 0               ' Appearance of the font does not matter.
  DRAFT_QUALITY = 1                 ' Appearance of the font is less important than when PROOF_QUALITY
                                    ' is used. For GDI raster fonts, scaling is enabled, which means
                                    ' that more font sizes are available, but the quality may be lower.
                                    ' Bold, italic, underline, and strikeout fonts are synthesized if
                                    ' necessary.
  PROOF_QUALITY = 2                 ' Character quality of the font is more important than exact
                                    ' matching of the logical-font attributes. For GDI raster fonts,
                                    ' scaling is disabled and the font closest in size is chosen.
                                    ' Although the chosen font size may not be mapped exactly when
                                    ' PROOF_QUALITY is used, the quality of the font is high and there
                                    ' is no distortion of appearance. Bold, italic, underline, and
                                    ' strikeout fonts are synthesized if necessary.
#If (WIN32_IE >= &H400) Then
  NONANTIALIASED_QUALITY = 3
  ANTIALIASED_QUALITY = 4
#End If '/* WINVER >= 0x0400 */
End Enum

Public Enum LOGICAL_FONT_PITCH
  DEFAULT_PITCH = 0
  FIXED_PITCH = 1
  VARIABLE_PITCH = 2
#If (WIN32_IE >= &H400) Then
  MONO_FONT = 8
#End If '/* WINVER >= 0x0400 */
End Enum

Public Enum LOGICAL_FONT_FAMILY
  FF_DONTCARE = (0 * &HF&)    ' /* Don't care or don't know. */
  FF_ROMAN = (1 * &H10&)      ' /* Variable stroke width, serifed. */
                              ' /* Times Roman, Century Schoolbook, etc. */
  FF_SWISS = (2 * &H10&)      ' /* Variable stroke width, sans-serifed. */
                              ' /* Helvetica, Swiss, etc. */
  FF_MODERN = (3 * &H10&)     ' /* Constant stroke width, serifed or sans-serifed. */
                              ' /* Pica, Elite, Courier, etc. */
  FF_SCRIPT = (4 * &H10&)     ' /* Cursive, etc. */
  FF_DECORATIVE = (5 * &H10&) '/* Old English, etc. */
End Enum

Public Enum TEXTMETRIC_PITCH
  TMPF_FIXED_PITCH = &H1
  TMPF_VECTOR = &H2
  TMPF_DEVICE = &H8
  TMPF_TRUETYPE = &H4
End Enum

Public Enum TEXTOUT_ALIGN_OPTION
  ETO_OPAQUE = &H2
  ETO_CLIPPED = &H4
#If (WIN32_IE >= &H400) Then
  ETO_GLYPH_INDEX = &H10
  ETO_RTLREADING = &H80
  ETO_NUMERICSLOCAL = &H400
  ETO_NUMERICSLATIN = &H800
  ETO_IGNORELANGUAGE = &H1000
#End If '/* WINVER >= 0x0400 */
#If (WIN32_IE >= &H500) Then
  ETO_PDY = &H2000
#End If ' (_WIN32_WINNT >= 0x0500)
End Enum

Public Enum DEVICE_CAPABILITY
 DRIVERVERSION = 0    ' /* Device driver version                    */
 TECHNOLOGY = 2       ' /* Device classification                    */
 HORZSIZE = 4         ' /* Horizontal size in millimeters           */
 VERTSIZE = 6         ' /* Vertical size in millimeters             */
 HORZRES = 8          ' /* Horizontal width in pixels               */
 VERTRES = 10         ' /* Vertical height in pixels                */
 BITSPIXEL = 12       ' /* Number of bits per pixel                 */
 PLANES = 14          ' /* Number of planes                         */
 NUMBRUSHES = 16      ' /* Number of brushes the device has         */
 NUMPENS = 18         ' /* Number of pens the device has            */
 NUMMARKERS = 20      ' /* Number of markers the device has         */
 NUMFONTS = 22        ' /* Number of fonts the device has           */
 NUMCOLORS = 24       ' /* Number of colors the device supports     */
 PDEVICESIZE = 26     ' /* Size required for device descriptor      */
 CURVECAPS = 28       ' /* Curve capabilities                       */
 LINECAPS = 30        ' /* Line capabilities                        */
 POLYGONALCAPS = 32   ' /* Polygonal capabilities                   */
 TEXTCAPS = 34        ' /* Text capabilities                        */
 CLIPCAPS = 36        ' /* Clipping capabilities                    */
 RASTERCAPS = 38      ' /* Bitblt capabilities                      */
 ASPECTX = 40         ' /* Length of the X leg                      */
 ASPECTY = 42         ' /* Length of the Y leg                      */
 ASPECTXY = 44        ' /* Length of the hypotenuse                 */

#If (WIN32_IE >= &H500) Then
 SHADEBLENDCAPS = 45  ' /* Shading and blending caps                */
#End If ' /* WINVER >= &H0500 */

 LOGPIXELSX = 88      ' /* Logical pixels/inch in X                 */
 LOGPIXELSY = 90      ' /* Logical pixels/inch in Y                 */

 SIZEPALETTE = 104    ' /* Number of entries in physical palette    */
 NUMRESERVED = 106    ' /* Number of reserved entries in palette    */
 COLORRES = 108       ' /* Actual color resolution                  */


Rem /* Printing related DeviceCaps. These replace the appropriate Escapes */
 PHYSICALWIDTH = 110  ' /* Physical Width in device units           */
 PHYSICALHEIGHT = 111 ' /* Physical Height in device units          */
 PHYSICALOFFSETX = 112 ' /* Physical Printable Area x margin        */
 PHYSICALOFFSETY = 113 ' /* Physical Printable Area y margin        */
 SCALINGFACTORX = 114 ' /* Scaling factor x                         */
 SCALINGFACTORY = 115 ' /* Scaling factor y                         */

Rem /* Display driver specific */
 VREFRESH = 116        ' /* Current vertical refresh rate of the    */
                       ' /* display device (for displays only) in Hz*/
 DESKTOPVERTRES = 117  ' /* Horizontal width of entire desktop in   */
                             ' /* pixels                            */
 DESKTOPHORZRES = 118  ' /* Vertical height of entire desktop in    */
                       ' /* pixels                                  */
 BLTALIGNMENT = 119    ' /* Preferred blt alignment                 */

#If (NOGDICAPMASKS = 1) Then
' /* Device Capability Masks: */
' /* Device Technologies */
 DT_PLOTTER = 0           ' /* Vector plotter                   */
 DT_RASDISPLAY = 1        ' /* Raster display                   */
 DT_RASPRINTER = 2        ' /* Raster printer                   */
 DT_RASCAMERA = 3         ' /* Raster camera                    */
 DT_CHARSTREAM = 4        ' /* Character-stream, PLP            */
 DT_METAFILE = 5          ' /* Metafile, VDM                    */
 DT_DISPFILE = 6          ' /* Display-file                     */

' /* Curve Capabilities */
 CC_NONE = 0              ' /* Curves not supported             */
 CC_CIRCLES = 1           ' /* Can do circles                   */
 CC_PIE = 2               ' /* Can do pie wedges                */
 CC_CHORD = 4             ' /* Can do chord arcs                */
 CC_ELLIPSES = 8          ' /* Can do ellipese                  */
 CC_WIDE = 16             ' /* Can do wide lines                */
 CC_STYLED = 32           ' /* Can do styled lines              */
 CC_WIDESTYLED = 64       ' /* Can do wide styled lines         */
 CC_INTERIORS = 128       ' /* Can do interiors                 */
 CC_ROUNDRECT = 256       ' /*                                  */

' /* Line Capabilities */
 LC_NONE = 0              ' /* Lines not supported              */
 LC_POLYLINE = 2          ' /* Can do polylines                 */
 LC_MARKER = 4            ' /* Can do markers                   */
 LC_POLYMARKER = 8        ' /* Can do polymarkers               */
 LC_WIDE = 16             ' /* Can do wide lines                */
 LC_STYLED = 32           ' /* Can do styled lines              */
 LC_WIDESTYLED = 64       ' /* Can do wide styled lines         */
 LC_INTERIORS = 128       ' /* Can do interiors                 */

' /* Polygonal Capabilities */
 PC_NONE = 0              ' /* Polygonals not supported         */
 PC_POLYGON = 1           ' /* Can do polygons                  */
 PC_RECTANGLE = 2         ' /* Can do rectangles                */
 PC_WINDPOLYGON = 4       ' /* Can do winding polygons          */
 PC_TRAPEZOID = 4         ' /* Can do trapezoids                */
 PC_SCANLINE = 8          ' /* Can do scanlines                 */
 PC_WIDE = 16             ' /* Can do wide borders              */
 PC_STYLED = 32           ' /* Can do styled borders            */
 PC_WIDESTYLED = 64       ' /* Can do wide styled borders       */
 PC_INTERIORS = 128       ' /* Can do interiors                 */
 PC_POLYPOLYGON = 256     ' /* Can do polypolygons              */
 PC_PATHS = 512           ' /* Can do paths                     */

' /* Clipping Capabilities */
 CP_NONE = 0              ' /* No clipping of output            */
 CP_RECTANGLE = 1         ' /* Output clipped to rects          */
 CP_REGION = 2            ' /* obsolete                         */

' /* Text Capabilities */
 TC_OP_CHARACTER = &H1            ' /* Can do OutputPrecision   CHARACTER      */
 TC_OP_STROKE = &H2               ' /* Can do OutputPrecision   STROKE         */
 TC_CP_STROKE = &H4               ' /* Can do ClipPrecision     STROKE         */
 TC_CR_90 = &H8                   ' /* Can do CharRotAbility    90             */
 TC_CR_ANY = &H10                 ' /* Can do CharRotAbility    ANY            */
 TC_SF_X_YINDEP = &H20            ' /* Can do ScaleFreedom      X_YINDEPENDENT */
 TC_SA_DOUBLE = &H40              ' /* Can do ScaleAbility      DOUBLE         */
 TC_SA_INTEGER = &H80             ' /* Can do ScaleAbility      INTEGER        */
 TC_SA_CONTIN = &H100             ' /* Can do ScaleAbility      CONTINUOUS     */
 TC_EA_DOUBLE = &H200             ' /* Can do EmboldenAbility   DOUBLE         */
 TC_IA_ABLE = &H400               ' /* Can do ItalisizeAbility  ABLE           */
 TC_UA_ABLE = &H800               ' /* Can do UnderlineAbility  ABLE           */
 TC_SO_ABLE = &H1000              ' /* Can do StrikeOutAbility  ABLE           */
 TC_RA_ABLE = &H2000              ' /* Can do RasterFontAble    ABLE           */
 TC_VA_ABLE = &H4000              ' /* Can do VectorFontAble    ABLE           */
 TC_RESERVED = &H8000
 TC_SCROLLBLT = &H10000           ' /* Don't do text scroll with blt           */

#End If ' /* NOGDICAPMASKS */

' /* Raster Capabilities */
 RC_NONE = 0
 RC_BITBLT = 1                ' /* Can do standard BLT.             */
 RC_BANDING = 2               ' /* Device requires banding support  */
 RC_SCALING = 4               ' /* Device requires scaling support  */
 RC_BITMAP64 = 8              ' /* Device can support >64K bitmap   */
 RC_GDI20_OUTPUT = &H10           ' /* has 2.0 output calls         */
 RC_GDI20_STATE = &H20
 RC_SAVEBITMAP = &H40
 RC_DI_BITMAP = &H80              ' /* supports DIB to memory       */
 RC_PALETTE = &H100               ' /* supports a palette           */
 RC_DIBTODEV = &H200              ' /* supports DIBitsToDevice      */
 RC_BIGFONT = &H400               ' /* supports >64K fonts          */
 RC_STRETCHBLT = &H800            ' /* supports StretchBlt          */
 RC_FLOODFILL = &H1000            ' /* supports FloodFill           */
 RC_STRETCHDIB = &H2000           ' /* supports StretchDIBits       */
 RC_OP_DX_OUTPUT = &H4000
 RC_DEVBITS = &H8000

#If (WIN32_IE >= &H500) Then
' /* Shading and blending caps                */
 SB_NONE = &H0
 SB_CONST_ALPHA = &H1
 SB_PIXEL_ALPHA = &H2
 SB_PREMULT_ALPHA = &H4

 SB_GRAD_RECT = &H10
 SB_GRAD_TRI = &H20
#End If ' /* WINVER >= &H0500 */
End Enum

Public Enum DRAWTEXT_OPTION
  DT_TOP = &H0                    ' Justifies the text to the top of the rectangle.
  DT_LEFT = &H0                   ' Aligns text to the left.
  DT_CENTER = &H1                 ' Centers text horizontally in the rectangle.
  DT_RIGHT = &H2                  ' Aligns text to the right.
  DT_VCENTER = &H4                ' Centers text vertically. This value is used only with the
                                  ' DT_SINGLELINE value.
  DT_BOTTOM = &H8                 ' Justifies the text to the bottom of the rectangle.
                                  ' This value is used only with the DT_SINGLELINE value.
  DT_WORDBREAK = &H10             ' Breaks words. Lines are automatically broken between words if a
                                  ' word would extend past the edge of the rectangle specified by the
                                  ' lpRect parameter. A carriage return-line feed sequence also breaks
                                  ' the line.
  DT_SINGLELINE = &H20            ' Displays text on a single line only. Carriage returns and
                                  ' line feeds do not break the line.
  DT_EXPANDTABS = &H40            ' Expands tab characters. The default number of characters
                                  ' per tab is eight.
  DT_TABSTOP = &H80               ' Sets tab stops. Bits 15?8 (high-order byte of the low-order word)
                                  ' of the uFormat parameter specify the number of characters for each
                                  ' tab. The default number of characters per tab is eight.
  DT_NOCLIP = &H100               ' Draws without clipping. DrawText is somewhat faster when
                                  ' DT_NOCLIP is used.
  DT_EXTERNALLEADING = &H200      ' Includes the font external leading in line height. Normally,
                                  ' external leading is not included in the height of a line of text.
  DT_CALCRECT = &H400             ' Determines the width and height of the rectangle.
  DT_NOPREFIX = &H800             ' Turns off processing of prefix characters.
  DT_INTERNAL = &H1000            ' Uses the system font to calculate text metrics.

#If (WIN32_IE >= &H400) Then
  DT_NOFULLWIDTHCHARBREAK = &H80000 ' Windows 98/Me, Windows 2000/XP: Prevents a line break at a
                                  ' DBCS (double-wide character string), so that the line breaking
                                  ' rule is equivalent to SBCS strings. For example, this can be
                                  ' used in Korean windows, for more readability of icon labels.
                                  ' This value has no effect unless DT_WORDBREAK is specified.
  DT_HIDEPREFIX = &H100000        ' Windows 2000/XP: Ignores the ampersand (&) prefix character in
                                  ' the text. The letter that follows will not be underlined, but
                                  ' other mnemonic-prefix characters are still processed.
  DT_EDITCONTROL = &H2000         ' Duplicates the text-displaying characteristics of a multiline
                                  ' edit control.
  DT_PATH_ELLIPSIS = &H4000       ' For displayed text, replaces characters in the middle of the string
                                  ' with ellipses so that the result fits in the specified rectangle.
  DT_END_ELLIPSIS = &H8000        ' For displayed text, if the end of a string does not fit in the
                                  ' rectangle, it is truncated and ellipses are added.
  DT_MODIFYSTRING = &H10000       ' Modifies the specified string to match the displayed text.
  DT_RTLREADING = &H20000         ' Layout in right-to-left reading order for bi-directional text when
                                  ' the font selected into the hdc is a Hebrew or Arabic font. The
                                  ' default reading order for all text is left-to-right.
  DT_WORD_ELLIPSIS = &H40000      ' Truncates any word that does not fit in the rectangle and adds ellipses.
#End If
End Enum

Public Enum DRAWICON_FLAG
  DI_MASK = &H1                   ' Draws the icon or cursor using the mask.
  DI_IMAGE = &H2                  ' Draws the icon or cursor using the image.
  DI_NORMAL = &H3                 ' Combination of DI_IMAGE and DI_MASK.
  DI_COMPAT = &H4                 ' Draws the icon or cursor using the system default image rather than
                                  ' the user-specified image.
  DI_DEFAULTSIZE = &H8            ' Draws the icon or cursor using the width and height specified by the
                                  ' system metric values for cursors or icons, if the cxWidth and
                                  ' cyWidth parameters are set to zero
End Enum

Public Enum INNER_BORDER_FLAGS
  BDR_INNER = &HC
  BDR_SUNKEN = &HA
  BDR_RAISEDINNER = &H4           ' Raised inner edge.
  BDR_SUNKENINNER = &H8           ' Sunken inner edge.
End Enum

Public Enum OUTER_BORDER_FLAGS
  BDR_OUTER = &H3
  BDR_RAISED = &H5
  BDR_RAISEDOUTER = &H1 ' Raised outer edge.
  BDR_SUNKENOUTER = &H2           ' Sunken outer edge.
End Enum

Public Enum EDGE_FLAGS
  EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)
  EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
  EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
  EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)
End Enum

Public Enum BORDER_TYPE_FLAGS
  BF_ADJUST = &H2000                ' If this flag is passed, shrink the rectangle pointed to by the
                                    ' qrc parameter to exclude the edges that were drawn. If this flag
                                    ' is not passed, then do not change the rectangle pointed to by the
                                    ' qrc parameter.
  BF_BOTTOM = &H8                   ' Bottom of border rectangle.
  BF_FLAT = &H4000                  ' Flat border.
  BF_LEFT = &H1                     ' Left side of border rectangle.
  BF_MIDDLE = &H800                 ' Interior of rectangle to be filled.
  BF_MONO = &H8000                  ' One-dimensional border.
  BF_RIGHT = &H4                    ' Right side of border rectangle.
  BF_SOFT = &H1000                  ' Soft buttons instead of tiles.
  BF_TOP = &H2                      ' Top of border rectangle.
  BF_BOTTOMLEFT = (BF_BOTTOM Or BF_LEFT) ' Bottom and left side of border rectangle.
  BF_BOTTOMRIGHT = (BF_BOTTOM Or BF_RIGHT)   ' Bottom and right side of border rectangle.
  BF_DIAGONAL = &H10                ' Diagonal border.
  BF_DIAGONAL_ENDBOTTOMLEFT = (BF_DIAGONAL Or BF_BOTTOM Or BF_LEFT)   ' Diagonal border. The end point is the bottom-left corner
                                                                      ' of the rectangle; the origin is top-right corner.
  BF_DIAGONAL_ENDBOTTOMRIGHT = (BF_DIAGONAL Or BF_BOTTOM Or BF_RIGHT) ' Diagonal border. The end point is the bottom-right corner
                                                                      ' of the rectangle; the origin is top-left corner.
  BF_DIAGONAL_ENDTOPLEFT = (BF_DIAGONAL Or BF_TOP Or BF_LEFT)         ' Diagonal border. The end point is the top-left corner of
                                                                      ' the rectangle; the origin is bottom-right corner.
  BF_DIAGONAL_ENDTOPRIGHT = (BF_DIAGONAL Or BF_TOP Or BF_RIGHT)       ' Diagonal border. The end point is the top-right corner of
                                                                      ' the rectangle; the origin is bottom-left corner.
  BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)  ' Entire border rectangle.
  BF_TOPLEFT = (BF_TOP Or BF_LEFT)  ' Top and left side of border rectangle.
  BF_TOPRIGHT = (BF_TOP Or BF_RIGHT) ' Top and right side of border rectangle.
End Enum

Public Enum RASTER_OPERATION
  ' /* Binary raster ops */
  R2_BLACK = 1              ' /*  0       */
  R2_NOTMERGEPEN = 2       ' /* DPon     */
  R2_MASKNOTPEN = 3        ' /* DPna     */
  R2_NOTCOPYPEN = 4        ' /* PN       */
  R2_MASKPENNOT = 5        ' /* PDna     */
  R2_NOT = 6               ' /* Dn       */
  R2_XORPEN = 7            ' /* DPx      */
  R2_NOTMASKPEN = 8        ' /* DPan     */
  R2_MASKPEN = 9           ' /* DPa      */
  R2_NOTXORPEN = 10        ' /* DPxn     */
  R2_NOP = 11              ' /* D        */
  R2_MERGENOTPEN = 12      ' /* DPno     */
  R2_COPYPEN = 13          ' /* P        */
  R2_MERGEPENNOT = 14      ' /* PDno     */
  R2_MERGEPEN = 15         ' /* DPo      */
  R2_WHITE = 16            ' /*  1       */
  R2_LAST = 16

  ' /* Ternary raster operations */
  SRCCOPY = &HCC0020              ' /* dest = source                   */
  SRCPAINT = &HEE0086             ' /* dest = source OR dest           */
  SRCAND = &H8800C6               ' /* dest = source AND dest          */
  SRCINVERT = &H660046            ' /* dest = source XOR dest          */
  SRCERASE = &H440328             ' /* dest = source AND (NOT dest )   */
  NOTSRCCOPY = &H330008           ' /* dest = (NOT source)             */
  NOTSRCERASE = &H1100A6          ' /* dest = (NOT src) AND (NOT dest) */
  MERGECOPY = &HC000CA            ' /* dest = (source AND pattern)     */
  MERGEPAINT = &HBB0226           ' /* dest = (NOT source) OR dest     */
  PATCOPY = &HF00021              ' /* dest = pattern                  */
  PATPAINT = &HFB0A09             ' /* dest = DPSnoo                   */
  PATINVERT = &H5A0049            ' /* dest = pattern XOR dest         */
  DSTINVERT = &H550009            ' /* dest = (NOT dest)               */
  BLACKNESS = &H42                ' /* dest = BLACK                    */
  WHITENESS = &HFF0062            ' /* dest = WHITE                    */
  TRANSPARENCY = &HB8074A         ' /* dest = TRANSPARENT              */
  ' /* Quaternary raster codes */
  ' MAKEROP4(fore,back) (DWORD)((((back) << 8) & &HFF000000) | (fore))
End Enum

Public Enum DRAWSTATE_OPTION
  ' /* Image type */
  DST_COMPLEX = &H0               ' The image is application defined. To render the image, DrawState
                                  ' calls the callback function specified by the lpOutputFunc parameter.
  DST_TEXT = &H1                  ' The image is text. The lData parameter is a pointer to the string,
                                  ' and the wData parameter specifies the length. If wData is zero,
                                  ' the string must be null-terminated.
  DST_PREFIXTEXT = &H2            ' The image is text that may contain an accelerator mnemonic.
                                  ' DrawState interprets the ampersand (&) prefix character as a
                                  ' directive to underscore the character that follows.
                                  ' The lData parameter is a pointer to the string, and the wData
                                  ' parameter specifies the length. If wData is zero, the string must
                                  ' be null-terminated.
  DST_ICON = &H3                  ' The image is an icon. The lData parameter is the icon handle.
  DST_BITMAP = &H4                ' The image is a bitmap. The lData parameter is the bitmap handle.
                                  ' Note that the bitmap cannot already be selected into an existing
                                  ' device context.
  ' /* For all states except DSS_NORMAL, the image is converted to monochrome before the visual effect
  '  * is applied.
  '  */
  ' /* State type */
  DSS_NORMAL = &H0                ' Draws the image without any modification.
  DSS_UNION = &H10                ' /* Gray string appearance */ Dithers the image.
  DSS_DISABLED = &H20             ' Embosses the image.
  DSS_MONO = &H80                 ' Draws the image using the brush specified by the hbr parameter.
  DSS_RIGHT = &H8000              ' Aligns the text to the right.
End Enum

Public Enum BACKMODE
  ' /* Background Modes */
  TRANSPARENT = 1                 ' Background remains untouched.
  OPAQUE = 2                      ' Background is filled with the current background color before the
                                  ' text, hatched brush, or pen is drawn.
  BKMODE_LAST = 2
End Enum
Rem --This ends the section for owner drawn menus--

Public Enum LOADIMGPARAM
  LR_DEFAULTCOLOR = &H0           ' The default flag; it does nothing. All it means is "not LR_MONOCHROME".
  LR_MONOCHROME = &H1             ' Loads the image in black and white.
  LR_COLOR = &H2                  ' None.
  LR_COPYRETURNORG = &H4          ' None.
  LR_COPYDELETEORG = &H8          ' None.
  LR_LOADFROMFILE = &H10          ' Loads the image from the file specified by the lpszName parameter.
                                  ' If this flag is not specified, lpszName is the name of the resource.
  LR_LOADTRANSPARENT = &H20       ' Retrieves the color value of the first pixel in the image and replaces
                                  ' the corresponding entry in the color table with the default window
                                  ' color (COLOR_WINDOW). All pixels in the image that use that entry
                                  ' become the default window color. This value applies only to images
                                  ' that have corresponding color tables.
                                  ' Do not use this option if you are loading a bitmap with a color depth
                                  ' greater than 8bpp.
                                  ' If fuLoad includes both the LR_LOADTRANSPARENT and LR_LOADMAP3DCOLORS
                                  ' values, LRLOADTRANSPARENT takes precedence. However, the color table
                                  ' entry is replaced with COLOR_3DFACE rather than COLOR_WINDOW.
  LR_DEFAULTSIZE = &H40           ' Uses the width or height specified by the system metric values for
                                  ' cursors or icons, if the cxDesired or cyDesired values are set to
                                  ' zero. If this flag is not specified and cxDesired and cyDesired are
                                  ' set to zero, the function uses the actual resource size. If the
                                  ' resource contains multiple images, the function uses the size of the
                                  ' first image.
  LR_VGACOLOR = &H80              ' Uses true VGA colors.
  LR_LOADMAP3DCOLORS = &H1000     ' Searches the color table for the image and replaces the shades of
                                  ' gray with the corresponding 3-D color.
  LR_CREATEDIBSECTION = &H2000    ' When the uType parameter specifies IMAGE_BITMAP, causes the function
                                  ' to return a DIB section bitmap rather than a compatible bitmap. This
                                  ' flag is useful for loading a bitmap without mapping it to the colors
                                  ' of the display device.
  LR_COPYFROMRESOURCE = &H4000    ' None.
  LR_SHARED = &H8000              ' Shares the image handle if the image is loaded multiple times. If
                                  ' LR_SHARED is not set, a second call to LoadImage for the same resource
                                  ' will load the image again and return a different handle.
                                  ' When you use this flag, the system will destroy the resource when it
                                  ' is no longer needed.
                                  ' Do not use LR_SHARED for images that have non-standard sizes, that may
                                  ' change after loading, or that are loaded from a file.
                                  ' Windows 95/98 : The function finds the first image with the
                                  ' requested resource name in the cache, regardless of the size requested.
End Enum

Public Enum LOADIMGTYP
  IMAGE_BITMAP = 0                ' Loads a bitmap.
  IMAGE_ICON = 1                  ' Loads an icon.

  IMAGE_CURSOR = 2                ' Loads a cursor.
#If (WIN32_IE >= &H400) Then
  IMAGE_ENHMETAFILE = 3           ' Load an enhanced metafile.
#End If
End Enum

Public Enum WIN32DEFINEDBMP
  OBM_CLOSE = 32754
  OBM_UPARROW = 32753
  OBM_DNARROW = 32752
  OBM_RGARROW = 32751
  OBM_LFARROW = 32750
  OBM_REDUCE = 32749
  OBM_ZOOM = 32748
  OBM_RESTORE = 32747
  OBM_REDUCED = 32746
  OBM_ZOOMD = 32745
  OBM_RESTORED = 32744
  OBM_UPARROWD = 32743
  OBM_DNARROWD = 32742
  OBM_RGARROWD = 32741
  OBM_LFARROWD = 32740
  OBM_MNARROW = 32739
  OBM_COMBO = 32738
  OBM_UPARROWI = 32737
  OBM_DNARROWI = 32736
  OBM_RGARROWI = 32735
  OBM_LFARROWI = 32734

  OBM_OLD_CLOSE = 32767
  OBM_SIZE = 32766
  OBM_OLD_UPARROW = 32765
  OBM_OLD_DNARROW = 32764
  OBM_OLD_RGARROW = 32763
  OBM_OLD_LFARROW = 32762
  OBM_BTSIZE = 32761
  OBM_CHECK = 32760
  OBM_CHECKBOXES = 32759
  OBM_BTNCORNERS = 32758
  OBM_OLD_REDUCE = 32757
  OBM_OLD_ZOOM = 32756
  OBM_OLD_RESTORE = 32755

  OCR_NORMAL = 32512
  OCR_IBEAM = 32513
  OCR_WAIT = 32514
  OCR_CROSS = 32515
  OCR_UP = 32516
  OCR_SIZE = 32640             ' /* OBSOLETE: use OCR_SIZEALL */
  OCR_ICON = 32641             ' /* OBSOLETE: use OCR_NORMAL */
  OCR_SIZENWSE = 32642
  OCR_SIZENESW = 32643
  OCR_SIZEWE = 32644
  OCR_SIZENS = 32645
  OCR_SIZEALL = 32646
  OCR_ICOCUR = 32647           ' /* OBSOLETE: use OIC_WINLOGO */
  OCR_NO = 32648
#If (WIN32_IE >= &H500) Then
  OCR_HAND = 32649
#End If ' /* WINVER >= &H0500 */
#If (WIN32_IE >= &H400) Then
  OCR_APPSTARTING = 32650
#End If ' /* WINVER >= &H0400 */

  OIC_SAMPLE = 32512
  OIC_HAND = 32513
  OIC_QUES = 32514
  OIC_BANG = 32515
  OIC_NOTE = 32516
#If (WIN32_IE >= &H400) Then
  OIC_WINLOGO = 32517
  OIC_WARNING = OIC_BANG
  OIC_ERROR = OIC_HAND
  OIC_INFORMATION = OIC_NOTE
#End If ' /* WINVER >= 0x0400 */
End Enum

Public Enum MAPPING_MODES
  Rem /* Mapping Modes */
  MM_TEXT = 1                 ' /* Each logical unit is mapped to one device pixel. Positive x is to
                              '  * the right; positive y is down.
                              '  */
  MM_LOMETRIC = 2             ' /* Each logical unit is mapped to 0.1 millimeter. Positive x is to the
                              '  * right; positive y is up.
                              '  */
  MM_HIMETRIC = 3             ' /* Each logical unit is mapped to 0.01 millimeter. Positive x is to the
                              '  * right; positive y is up.
                              '  */
  MM_LOENGLISH = 4            ' /* Each logical unit is mapped to 0.001 inch. Positive x is to the right;
                              ' * positive y is up.
                              ' */
  MM_HIENGLISH = 5            ' /* Each logical unit is mapped to 0.001 inch. Positive x is to the right;
                              '  * positive y is up.
                              '  */
  MM_TWIPS = 6                ' /* Each logical unit is mapped to one twentieth of a printer's point
                              '  * (1/1440 inch, also called a twip). Positive x is to the right;
                              '  * positive y is up.
                              '  */
  MM_ISOTROPIC = 7            ' /* Logical units are mapped to arbitrary units with equally scaled axes;
                              '  * that is, one unit along the x-axis is equal to one unit along the
                              '  * y-axis. Use the SetWindowExtEx and SetViewportExtEx functions to
                              '  * specify the units and the orientation of the axes. Graphics device
                              '  * interface (GDI) makes adjustments as necessary to ensure the x and y
                              '  * units remain the same size (When the window extent is set, the
                              '  * viewport will be adjusted to keep the units isotropic).
                              '  */
  MM_ANISOTROPIC = 8          ' /* Logical units are mapped to arbitrary units with arbitrarily scaled
                              '  * axes. Use the SetWindowExtEx and SetViewportExtEx functions to
                              '  * specify the units, orientation, and scaling.
                              '  */
  Rem /* Min and Max Mapping Mode values */
  MM_MIN = MM_TEXT
  MM_MAX = MM_ANISOTROPIC
  MM_MAX_FIXEDSCALE = MM_TWIPS
End Enum

Public Enum DRAWCAPTIONOPTIONS
  DC_ACTIVE = &H1             ' /* The function uses the colors that denote an active caption. */
  DC_SMALLCAP = &H2           ' /* The function draws a small caption, using the current small caption font. */
  DC_ICON = &H4               ' /* The function draws the icon when drawing the caption text. */
  DC_TEXT = &H8               ' /* The function draws the caption text when drawing the caption. */
  DC_INBUTTON = &H10          ' /* The function draws the caption as a button. */
#If (WIN32_IE >= &H500) Then
  DC_GRADIENT = &H20          ' /* Windows 98, Windows 2000: When this flag is set, the function uses
                              '  * COLOR_GRADIENTACTIVECAPTION (if the DC_ACTIVE flag was set) or
                              '  * COLOR_GRADIENTINACTIVECAPTION for the title-bar color.
                              '  * If this flag is not set, the function uses COLOR_ACTIVECAPTION or
                              '  * COLOR_INACTIVECAPTION for both colors.
                              '  */
#End If ' /* WINVER >= 0x0500 */
End Enum

Public Enum GRADFILLMODE
  GRADIENT_FILL_RECT_H = &H0  ' /* In this mode, two endpoints describe a rectangle. The rectangle is
                              '  * defined to have a constant color (specified by the TRIVERTEX
                              '  * structure) for the left and right edges. GDI interpolates the
                              '  * color from the left to right edge and fills the interior.
                              '  */
  GRADIENT_FILL_RECT_V = &H1  ' /* In this mode, two endpoints describe a rectangle. The rectangle is
                              '  * defined to have a constant color (specified by the TRIVERTEX
                              '  * structure) for the top and bottom edges. GDI interpolates the color
                              '  * from the top to bottom edge and fills the interior.
                              '  */
  GRADIENT_FILL_TRIANGLE = &H2 ' /* In this mode, an array of TRIVERTEX structures is passed to GDI
                              '  * along with a list of array indexes that describe separate triangles.
                              '  * GDI performs linear interpolation between triangle vertices and fills
                              '  * the interior. Drawing is done directly in 24- and 32-bpp modes.
                              '  * Dithering is performed in 16-, 8-, 4-, and 1-bpp mode.
                              '  */
  GRADIENT_FILL_OP_FLAG = &HFF
End Enum

Public Enum DEV_CNTXT_ACCUMUL_FLAG
  Rem /* Bounds Accumulation APIs */
  DCB_RESET = &H1             ' /* Clears the bounding rectangle after returning it. If this flag is
                              '  * not set, the bounding rectangle will not be cleared.
                              '  */
  DCB_ACCUMULATE = &H2
  DCB_DIRTY = DCB_ACCUMULATE
  DCB_SET = (DCB_RESET Or DCB_ACCUMULATE) ' /* The bounding rectangle is not empty. */
  DCB_ENABLE = &H4            ' /* Boundary accumulation is on. */
  DCB_DISABLE = &H8           ' /* Boundary accumulation is off. */
End Enum

Public Enum DRAWFRAMECONTROL_TYPE   ' /* Specifies the type of frame control to draw. */
  DFC_CAPTION = 1                   ' /* Title bar */
  DFC_MENU = 2                      ' /* Menu bar */
  DFC_SCROLL = 3                    ' /* Scroll bar */
  DFC_BUTTON = 4                    ' /* Standard button */
#If (WIN32_IE >= &H500) Then
  DFC_POPUPMENU = 5                 ' /* Windows 98/Me, Windows 2000/XP: Popup menu item. */
#End If ' /* WINVER >= 0x0500 */
End Enum

Public Enum DRAWFRAMECONTROL_STATE
  ' /* If uType is DFC_CAPTION, uState can be one of the following values.  */
  DFCS_CAPTIONCLOSE = &H0           ' /* Close button */
  DFCS_CAPTIONMIN = &H1             ' /* Minimize button */
  DFCS_CAPTIONMAX = &H2             ' /* Maximize button */
  DFCS_CAPTIONRESTORE = &H3         ' /* Restore button */
  DFCS_CAPTIONHELP = &H4            ' /* Help button */

  Rem /* If uType is DFC_MENU, uState can be one of the following values.  */
  DFCS_MENUARROW = &H0              ' /* Submenu arrow */
  DFCS_MENUCHECK = &H1              ' /* Check mark */
  DFCS_MENUBULLET = &H2             ' /* Bullet */
  DFCS_MENUARROWRIGHT = &H4         ' /* Submenu arrow pointing left. This is used for the right-to-left
                                    '  * cascading menus used with right-to-left languages such as
                                    '  * Arabic or Hebrew.
                                    '  */

  Rem /* If uType is DFC_SCROLL, uState can be one of the following values.  */
  DFCS_SCROLLUP = &H0               ' /* Up arrow of scroll bar */
  DFCS_SCROLLDOWN = &H1             ' /* Down arrow of scroll bar */
  DFCS_SCROLLLEFT = &H2             ' /* Left arrow of scroll bar */
  DFCS_SCROLLRIGHT = &H3            ' /* Right arrow of scroll bar */
  DFCS_SCROLLCOMBOBOX = &H5         ' /* Combo box scroll bar */
  DFCS_SCROLLSIZEGRIP = &H8         ' /* Size grip in bottom-right corner of window */
  DFCS_SCROLLSIZEGRIPRIGHT = &H10   ' /* Size grip in bottom-left corner of window. This is used with
                                    '  * right-to-left languages such as Arabic or Hebrew.
                                    '  */

  Rem /* If uType is DFC_BUTTON, uState can be one of the following values.  */
  DFCS_BUTTONCHECK = &H0            ' /* Check box */
  DFCS_BUTTONRADIOIMAGE = &H1       ' /* Image for radio button (nonsquare needs image) */
  DFCS_BUTTONRADIOMASK = &H2        ' /* Mask for radio button (nonsquare needs mask) */
  DFCS_BUTTONRADIO = &H4            ' /* Radio button */
  DFCS_BUTTON3STATE = &H8           ' /* Three-state button */
  DFCS_BUTTONPUSH = &H10            ' /* Push button */

  Rem /* One or more of the following values can be used to set the state of the control to be drawn. */
  DFCS_INACTIVE = &H100             ' /* Button is inactive (grayed). */
  DFCS_PUSHED = &H200               ' /* Button is pushed. */
  DFCS_CHECKED = &H400              ' /* Button is checked. */

#If (WIN32_IE >= &H500) Then
  DFCS_TRANSPARENT = &H800          ' /* Windows 98/Me, Windows 2000/XP: The background remains untouched. */
  DFCS_HOT = &H1000                 ' /* Windows 98/Me, Windows 2000/XP: Button is hot-tracked. */
#End If ' /* WINVER >= 0x0500 */

  DFCS_ADJUSTRECT = &H2000          ' /* Bounding rectangle is adjusted to exclude the surrounding edge
                                    '  * of the push button.
                                    '  */
  DFCS_FLAT = &H4000                ' /* Button has a flat border. */
  DFCS_MONO = &H8000                ' /* Button has a monochrome border. */
End Enum

Public Enum STOCK_OBJ_TYPE
Rem /* Stock Logical Objects */
  WHITE_BRUSH = 0                    ' /* White brush. */
  LTGRAY_BRUSH = 1                   ' /* Light gray brush. */
  GRAY_BRUSH = 2                     ' /* Gray brush. */
  DKGRAY_BRUSH = 3                   ' /* Dark gray brush. */
  BLACK_BRUSH = 4                    ' /* Black brush. */
  NULL_BRUSH = 5                     ' /* Null brush (equivalent to HOLLOW_BRUSH). */
  HOLLOW_BRUSH = NULL_BRUSH          ' /* Hollow brush (equivalent to NULL_BRUSH). */
  WHITE_PEN = 6                      ' /* White pen. */
  BLACK_PEN = 7                      ' /* Black pen. */
  NULL_PEN = 8
  OEM_FIXED_FONT = 10                ' /* Original equipment manufacturer (OEM) dependent fixed-pitch
                                     '  * (monospace) font.
                                     '  */
  ANSI_FIXED_FONT = 11               ' /* Windows fixed-pitch (monospace) system font. */
  ANSI_VAR_FONT = 12                 ' /* Windows variable-pitch (proportional space) system font. */
  SYSTEM_FONT = 13                   ' /* System font. By default, the system uses the system font to
                                     '  * draw menus, dialog box controls, and text.
                                     '  * Windows 95/98 and NT: The system font is MS Sans Serif.
                                     '  * Windows 2000: The system font is Tahoma
                                     '  */
  DEVICE_DEFAULT_FONT = 14           ' /* Windows NT/2000: Device-dependent font. */
  DEFAULT_PALETTE = 15               ' /* Default palette. This palette consists of the static colors
                                     '  * in the system palette.
                                     '  */
  SYSTEM_FIXED_FONT = 16             ' /* Fixed-pitch (monospace) system font. This stock object is
                                     '  * provided only for compatibility with 16-bit Windows versions
                                     '  * earlier than 3.0.
                                     '  */
#If (WIN32_IE >= &H400) Then
  DEFAULT_GUI_FONT = 17              ' /* Default font for user interface objects such as menus and
                                     '  * dialog boxes. This is MS Sans Serif. Compare this with
                                     '  * SYSTEM_FONT.
                                     '  */
#End If ' /* WINVER >= 0x0400 */
#If (WIN32_IE >= &H500) Then
  DC_BRUSH = 18                      ' /* Windows 98, Windows 2000: Solid color brush. The default
                                     '  * color is white. The color can be changed by using the
                                     '  * SetDCBrushColor function.
                                     '  */
  DC_PEN = 19                        ' /* Windows 98, Windows 2000: Solid pen color. The default color
                                     '  * is white. The color can be changed by using the SetDCPenColor
                                     '  * function.
                                     '  */
#End If
#If (WIN32_IE >= &H500) Then
  STOCK_LAST = 19
#ElseIf (WIN32_IE >= &H400) Then
  STOCK_LAST = 17
#Else
  STOCK_LAST = 16
#End If
End Enum

Public Enum EDITCONTROL_STYLE
  Rem /* Edit Control Styles */
  ES_LEFT = &H0                   ' /* Left-aligns text. */
  ES_CENTER = &H1                 ' /* Centers text in a multiline edit control. */
  ES_RIGHT = &H2                  ' /* Right-aligns text in a multiline edit control. */
  ES_MULTILINE = &H4              ' /* Designates a multiline edit control. The default is a single-line
                                  '  * edit control.
                                  '  * When the multiline edit control is in a dialog box, the default
                                  '  * response to pressing the ENTER key is to activate the default button.
                                  '  * To use the ENTER key as a carriage return, use the ES_WANTRETURN style.
                                  '  * When the multiline edit control is not in a dialog box and the
                                  '  * ES_AUTOVSCROLL style is specified, the edit control shows as many
                                  '  * lines as possible and scrolls vertically when the user presses the
                                  '  * ENTER key. If you do not specify ES_AUTOVSCROLL, the edit control shows
                                  '  * as many lines as possible and beeps if the user presses the ENTER key
                                  '  * when no more lines can be displayed.
                                  '  * If you specify the ES_AUTOHSCROLL style, the multiline edit control
                                  '  * automatically scrolls horizontally when the caret goes past the right
                                  '  * edge of the control. To start a new line, the user must press the ENTER
                                  '  * key. If you do not specify ES_AUTOHSCROLL, the control automatically
                                  '  * wraps words to the beginning of the next line when necessary. A new line
                                  '  * is also started if the user presses the ENTER key. The window size
                                  '  * determines the position of the word wrap. If the window size changes,
                                  '  * the word wrapping position changes and the text is redisplayed.
                                  '  * Multiline edit controls can have scroll bars. An edit control with
                                  '  * scroll bars processes its own scroll bar messages. Edit controls without
                                  '  * scroll bars scroll as described in the previous paragraphs and process
                                  '  * any scroll messages that are sent by the parent window.
                                  '  */
  ES_UPPERCASE = &H8              ' /* Converts all characters to uppercase as they are typed into the
                                  '  * edit control.
                                  '  */
  ES_LOWERCASE = &H10             ' /* Converts all characters to lowercase as they are typed into the
                                  '  * edit control.
                                  '  */
  ES_PASSWORD = &H20              ' /* Displays an asterisk (*) for each character that is typed into
                                  '  * the edit control. You can use the EM_SETPASSWORDCHAR message to
                                  '  * change the displayed character.
                                  '  */
  ES_AUTOVSCROLL = &H40           ' /* Scrolls text up one page when the user presses the ENTER key on
                                  '  * the last line.
                                  '  */
  ES_AUTOHSCROLL = &H80           ' /* Automatically scrolls text to the right by 10 characters when the
                                  '  * user types a character at the end of the line. When the user
                                  '  * presses the ENTER key, the control scrolls all text back to the
                                  '  * zero position.
                                  '  */
  ES_NOHIDESEL = &H100            ' /* Negates the default behavior for an edit control. The default
                                  '  * behavior hides the selection when the control loses the input
                                  '  * focus and inverts the selection when the control receives the
                                  '  * input focus. If you specify ES_NOHIDESEL, the selected text is
                                  '  * inverted, even if the control does not have the focus.
                                  '  */
  ES_OEMCONVERT = &H400           ' /* Converts text typed in the edit control from the Windows CE
                                  '  * character set to the OEM character set and then converts it
                                  '  * back to the Windows CE set. This style is most useful for edit
                                  '  * controls that contain file names.
                                  '  */
  ES_READONLY = &H800             ' /* Prevents the user from typing or editing text in the edit control. */
  ES_WANTRETURN = &H1000          ' /* Specifies that a carriage return be inserted when the user presses
                                  '  * the ENTER key while typing text into a multiline edit control in
                                  '  * a dialog box. If you do not specify this style, pressing the
                                  '  * ENTER key has the same effect as pressing the dialog box's
                                  '  * default push button. This style has no effect on a single-line
                                  '  * edit control.
                                  '  */
#If (WIN32_IE >= &H400) Then
  ES_NUMBER = &H2000              ' /* Accepts into the edit control only digits to be typed. */
#End If '/* WINVER >= 0x0400 */
End Enum

Public Enum BMPCOMPRESSION
  BI_RGB = &H0&                 ' /* An uncompressed format. */
  BI_RLE8 = &H1&                ' /* A run-length encoded (RLE) format for bitmaps with 8 bpp.
                                '  * The compression format is a 2-byte format consisting of a count
                                '  * byte followed by a byte containing a color index.
                                '  */
  BI_RLE4 = &H2&                ' /* An RLE format for bitmaps with 4 bpp. The compression format is a
                                '  * 2-byte format consisting of a count byte followed by two
                                '  * word-length color indexes.
                                '  */
  BI_BITFIELDS = &H3&           ' /* Specifies that the bitmap is not compressed and that the color
                                '  * table consists of three DWORD color masks that specify the red,
                                '  * green, and blue components, respectively, of each pixel. This is
                                '  * valid when used with 16- and 32-bpp bitmaps.
                                '  */
#If (WIN32_IE >= &H400) Then
  BI_JPEG = &H4&                ' /* Windows 98/Me, Windows 2000/XP: Indicates that the image is a
                                '  * JPEG image. [To be found]
                                '  */
  BI_PNG = &H5&                 ' /* Windows 98/Me, Windows 2000/XP: Indicates that the image is a
                                '  * PNG image. [To be found]
                                '  */
#End If
End Enum

Public Enum BMPBITCOUNT
  BC_IMPLIED = &H0&             ' /* Windows 98/Me, Windows 2000/XP: The number of bits-per-pixel is
                                '  * specified or is implied by the JPEG or PNG format.
                                '  */
  BC_MONOCHROME = &H1&          ' /* The bitmap is monochrome, and the bmiColors member of BITMAPINFO
                                '  * contains two entries. Each bit in the bitmap array represents a
                                '  * pixel. If the bit is clear, the pixel is displayed with the color
                                '  * of the first entry in the bmiColors table; if the bit is set, the
                                '  * pixel has the color of the second entry in the table.
                                '  */
  BC_16COLOR = &H4&             ' /* The bitmap has a maximum of 16 colors, and the bmiColors member of
                                '  * BITMAPINFO contains up to 16 entries. Each pixel in the bitmap is
                                '  * represented by a 4-bit index into the color table. For example, if
                                '  * the first byte in the bitmap is 0x1F, the byte represents two
                                '  * pixels. The first pixel contains the color in the second table
                                '  * entry, and the second pixel contains the color in the sixteenth
                                '  * table entry.
                                '  */
 BC_256COLOR = &H8&             ' /* The bitmap has a maximum of 256 colors, and the bmiColors member
                                '  * of BITMAPINFO contains up to 256 entries. In this case, each byte
                                '  * in the array represents a single pixel.
                                '  */
 BC_5BITRGB = &H10&             ' /* The bitmap has a maximum of 2^16 colors. If the biCompression
                                '  * member of the BITMAPINFOHEADER is BI_RGB, the bmiColors member of
                                '  * BITMAPINFO is NULL. Each WORD in the bitmap array represents a
                                '  * single pixel. The relative intensities of red, green, and blue are
                                '  * represented with five bits for each color component. The value for
                                '  * blue is in the least significant five bits, followed by five bits
                                '  * each for green and red. The most significant bit is not used. The
                                '  * bmiColors color table is used for optimizing colors used on
                                '  * palette-based devices, and must contain the number of entries
                                '  * specified by the biClrUsed member of the BITMAPINFOHEADER.
                                '  * If the biCompression member of the BITMAPINFOHEADER is BI_BITFIELDS,
                                '  * the bmiColors member contains three DWORD color masks that specify
                                '  * the red, green, and blue components, respectively, of each pixel.
                                '  * Each WORD in the bitmap array represents a single pixel.
                                '  * Windows NT/Windows 2000/XP: When the biCompression member is
                                '  * BI_BITFIELDS, bits set in each DWORD mask must be contiguous and
                                '  * should not overlap the bits of another mask. All the bits in the
                                '  * pixel do not have to be used.
                                '  * Windows 95/98/Me: When the biCompression member is BI_BITFIELDS,
                                '  * the system supports only the following 16bpp color masks: A
                                '  * 5-5-5 16-bit image, where the blue mask is 0x001F, the green mask
                                '  * is 0x03E0, and the red mask is 0x7C00; and a 5-6-5 16-bit image,
                                '  * where the blue mask is 0x001F, the green mask is 0x07E0, and the
                                '  * red mask is 0xF800.
                                '  */
End Enum

Public Enum LAYERED_WINDOW_ATTRIB
  LWA_COLORKEY = &H1&           ' /* Use crKey as the transparency color. */
  LWA_ALPHA = &H2&              ' /* Use bAlpha to determine the opacity of the layered window. */
End Enum

Public Enum MENU_PARAMETERS_INFO
  SPI_GETDROPSHADOW = &H1024&   ' /* Windows XP: Indicates whether the drop shadow effect is enabled.
                                '  * The pvParam parameter must point to a BOOL variable that
                                '  * returns TRUE if enabled or FALSE if disabled.
                                '  */
  SPI_GETFLATMENU = &H1022&     ' /* Windows XP: Indicates whether native User menus have flat menu
                                '  * appearance. The pvParam parameter must point to a BOOL variable
                                '  * that returns TRUE if the flat menu appearance is set, or FALSE
                                '  * otherwise.
                                '  */
  SPI_GETMENUUNDERLINES = &H100A ' /* Windows 98/Me, Windows 2000/XP: Indicates whether menu access
                                 '  * keys are always underlined. The pvParam parameter must point
                                 '  * to a BOOL variable that receives TRUE if menu access keys are
                                 '  * always underlined, and FALSE if they are underlined only when
                                 '  * the menu is activated by the keyboard.
                                 '  * Note that SPI_GETMENUUNDERLINES is the same as SPI_GETKEYBOARDCUES,
                                 '  * and SPI_SETMENUUNDERLINES is the same as SPI_SETKEYBOARDCUES.
                                 '  */
End Enum

Public Enum WINHOOKID
Rem /*
Rem  * SetWindowsHook() codes
Rem  */
 WH_MIN = (-1)
 WH_MSGFILTER = (-1)
 WH_JOURNALRECORD = 0
 WH_JOURNALPLAYBACK = 1
 WH_KEYBOARD = 2
 WH_GETMESSAGE = 3
 WH_CALLWNDPROC = 4
 WH_CBT = 5
 WH_SYSMSGFILTER = 6
 WH_MOUSE = 7
 WH_HARDWARE = 8
 WH_DEBUG = 9
 WH_SHELL = 10
 WH_FOREGROUNDIDLE = 11
#If (WIN32_IE >= &H400) Then
 WH_CALLWNDPROCRET = 12
#End If ' /* WINVER >= 0x0400 */

#If (WIN32_IE >= &H400) Then
 WH_KEYBOARD_LL = 13
 WH_MOUSE_LL = 14
#End If ' /* (_WIN32_WINNT >= 0x0400) */

#If (WIN32_IE >= &H400) Then
  #If (WIN32_IE >= &H400) Then
   WH_MAX = 14
  #Else
   WH_MAX = 12
  #End If ' /* (_WIN32_WINNT >= 0x0400) */
#Else
 WH_MAX = 11
#End If

 WH_MINHOOK = WH_MIN
 WH_MAXHOOK = WH_MAX
End Enum

Public Enum HOOK_CODE
Rem /*
Rem  * Hook Codes
Rem  */
 HC_ACTION = 0
 HC_GETNEXT = 1
 HC_SKIP = 2
 HC_NOREMOVE = 3
 HC_NOREM = HC_NOREMOVE
 HC_SYSMODALON = 4
 HC_SYSMODALOFF = 5

Rem /*
Rem  * CBT Hook Codes
Rem  */
 HCBT_MOVESIZE = 0
 HCBT_MINMAX = 1
 HCBT_QS = 2
 HCBT_CREATEWND = 3
 HCBT_DESTROYWND = 4
 HCBT_ACTIVATE = 5
 HCBT_CLICKSKIPPED = 6
 HCBT_KEYSKIPPED = 7
 HCBT_SYSCOMMAND = 8
 HCBT_SETFOCUS = 9
End Enum

Public Enum DC_WINDOW_RGN
Rem /*
Rem  * GetDCEx() flags
Rem  */
 DCX_WINDOW = &H1&
 DCX_CACHE = &H2&
 DCX_NORESETATTRS = &H4&
 DCX_CLIPCHILDREN = &H8&
 DCX_CLIPSIBLINGS = &H10&
 DCX_PARENTCLIP = &H20&

 DCX_EXCLUDERGN = &H40&
 DCX_INTERSECTRGN = &H80&

 DCX_EXCLUDEUPDATE = &H100&
 DCX_INTERSECTUPDATE = &H200&

 DCX_LOCKWINDOWUPDATE = &H400&

 DCX_VALIDATE = &H200000
End Enum

Public Enum WM_PRINT_OPTIONS
#If (WIN32_IE >= &H400) Then
  Rem /* WM_PRINT flags */
  PRF_CHECKVISIBLE = &H1&       ' /* Draws the window only if it is visible. */
  PRF_NONCLIENT = &H2&          ' /* Draws the nonclient area of the window. */
  PRF_CLIENT = &H4              ' /* Draws the client area of the window. */
  PRF_ERASEBKGND = &H8&         ' /* Erases the background before drawing the window. */
  PRF_CHILDREN = &H10&          ' /* Draws all visible children windows. */
  PRF_OWNED = &H20&             ' /* Draws all owned windows. */
#End If
End Enum

Public Enum CLASSINFO_OPTION
  GCL_MENUNAME = (-8)           ' /* Replaces the pointer to the menu name string. The string
                                '  * identifies the menu resource associated with the class.
                                '  */
  GCL_HBRBACKGROUND = (-10)     ' /* Replaces a handle to the background brush associated with the
                                '  * class.
                                '  */
  GCL_HCURSOR = (-12)           ' /* Replaces a handle to the cursor associated with the class. */
  GCL_HICON = (-14)             ' /* Replaces a handle to the icon associated with the class. */
  GCL_HMODULE = (-16)           ' /* Replaces a handle to the module that registered the class. */
  GCL_CBWNDEXTRA = (-18)        ' /* Sets the size, in bytes, of the extra window memory associated
                                '  * with each window in the class. Setting this value does not
                                '  * change the number of extra bytes already allocated. For
                                '  * information on how to access this memory,
                                '  * see SetWindowLongPtr.
                                '  */
  GCL_CBCLSEXTRA = (-20)        ' /* Sets the size, in bytes, of the extra memory associated with
                                '  * the class. Setting this value does not change the number of
                                '  * extra bytes already allocated.
                                '  */
  GCL_WNDPROC = (-24)           ' /* Replaces the pointer to the window procedure associated with
                                '  * the class.
                                '  */
  GCL_STYLE = (-26)             ' /* Replaces the window-class style bits. */
  GCW_ATOM = (-32)

#If (WIN32_IE >= &H400) Then
  GCL_HICONSM = (-34)           ' /* Retrieves a handle to the small icon associated with the class. */
#End If ' /* WINVER >= 0x0400 */
End Enum

Public Enum ZORDER
  HWND_TOP = (0)         ' /* Places the window at the top of the Z order. */
  HWND_BOTTOM = (1)      ' /* Places the window at the bottom of the Z order. If the hWnd parameter
                         '  * identifies a topmost window, the window loses its topmost status and is
                         '  * placed at the bottom of all other windows.
                         '  */
  HWND_TOPMOST = (-1)    ' /* Places the window above all non-topmost windows. The window maintains
                         '  * its topmost position even when it is deactivated.
                         '  */
  HWND_NOTOPMOST = (-2)  ' /* Places the window above all non-topmost windows (that is, behind all
                         '  * topmost windows). This flag has no effect if the window is already a
                         '  * non-topmost window.
                         '  */
End Enum

Public Enum WINDOWPOSFLAGS
  SWP_NOSIZE = &H1             ' /* Retains current size (ignores the cx and cy members). */
  SWP_NOMOVE = &H2             ' /* Retains current position (ignores the x and y members). */
  SWP_NOZORDER = &H4           ' /* Retains current ordering (ignores the hwndInsertAfter member). */
  SWP_NOREDRAW = &H8           ' /* Does not redraw changes. */
  SWP_NOACTIVATE = &H10        ' /* Does not activate the window. */
  SWP_FRAMECHANGED = &H20      ' /* The frame changed: send WM_NCCALCSIZE */
  SWP_SHOWWINDOW = &H40        ' /*  Displays the window. */
  SWP_HIDEWINDOW = &H80        ' /* Hides the window. */
  SWP_NOCOPYBITS = &H100       ' /* Discards the entire contents of the client area. If this flag
                               '  * is not specified, the valid contents of the client area are
                               '  * saved and copied back into the client area after the window is
                               '  * sized or repositioned.
                               '  */
  SWP_NOOWNERZORDER = &H200    ' /* Don't do owner Z ordering */
  SWP_NOSENDCHANGING = &H400   ' /* Don't send WM_WINDOWPOSCHANGING */

  SWP_DRAWFRAME = SWP_FRAMECHANGED ' /* Draws a frame (defined in the class description for the
                               '  * window) around the window. The window receives a WM_NCCALCSIZE
                               '  * message.
                               '  */
  SWP_NOREPOSITION = SWP_NOOWNERZORDER

#If (WIN32_IE >= &H400) Then
  SWP_DEFERERASE = &H2000      ' /* Prevents generation of the WM_SYNCPAINT message. */
  SWP_ASYNCWINDOWPOS = &H4000  ' /* If the calling thread and the thread that owns the window are
                               '  * attached to different input queues, the system posts the
                               '  * request to the thread that owns the window. This prevents the
                               '  * calling thread from blocking its execution while other threads
                               '  * process the request.
                               '  */
#End If ' /* WINVER >= =0x0400 */
End Enum

Public Enum GETWINDOW_OPTION
  GW_HWNDFIRST = 0            ' /* The retrieved handle identifies the window of the same type that
                              '  * is highest in the Z order. If the specified window is a topmost
                              '  * window, the handle identifies the topmost window that is highest
                              '  * in the Z order. If the specified window is a top-level window,
                              '  * the handle identifies the top-level window that is highest in
                              '  * the Z order. If the specified window is a child window, the
                              '  * handle identifies the sibling window that is highest in the Z
                              '  * order.
                              '  */
  GW_HWNDLAST = 1             ' /* The retrieved handle identifies the window of the same type that
                              '  * is lowest in the Z order. If the specified window is a topmost
                              '  * window, the handle identifies the topmost window that is lowest
                              '  * in the Z order. If the specified window is a top-level window,
                              '  * the handle identifies the top-level window that is lowest in the
                              '  * Z order. If the specified window is a child window, the handle
                              '  * identifies the sibling window that is lowest in the Z order.
                              '  */
  GW_HWNDNEXT = 2             ' /* The retrieved handle identifies the window below the specified
                              '  * window in the Z order. If the specified window is a topmost
                              '  * window, the handle identifies the topmost window below the
                              '  * specified window. If the specified window is a top-level window,
                              '  * the handle identifies the top-level window below the specified
                              '  * window. If the specified window is a child window, the handle
                              '  * identifies the sibling window below the specified window.
                              '  */
  GW_HWNDPREV = 3             ' /* The retrieved handle identifies the window above the specified
                              '  * window in the Z order. If the specified window is a topmost
                              '  * window, the handle identifies the topmost window above the specified
                              '  * window. If the specified window is a top-level window, the handle
                              '  * identifies the top-level window above the specified window. If the
                              '  * specified window is a child window, the handle identifies the sibling
                              '  * window above the specified window.
                              '  */
  GW_OWNER = 4                ' /* The retrieved handle identifies the specified window's owner
                              '  * window, if any.
                              '  */
  GW_CHILD = 5                ' /* The retrieved handle identifies the child window at the top of
                              '  * the Z order, if the specified window is a parent window;
                              '  * otherwise, the retrieved handle is NULL. The function examines
                              '  * only child windows of the specified window. It does not examine
                              '  * descendant windows.
                              '  */
#If (WIN32_IE <= &H400&) Then
  GW_MAX = 5
#Else
  GW_ENABLEDPOPUP = 6         ' /* Windows 2000/XP: The retrieved handle identifies the enabled popup
                              '  * window owned by the specified window (the search uses the first
                              '  * such window found using GW_HWNDNEXT); otherwise, if there are no
                              '  * enabled popup windows, the retrieved handle is that of the
                              '  * specified window.
                              '  */
  GW_MAX = 6
#End If
End Enum

Public Enum MSGFILTER_CODE    ' /* Type of input event that generated the message. */
  MSGF_DIALOGBOX = 0          ' /* The input event occurred in a message box or dialog box. */
  MSGF_MESSAGEBOX = 1
  MSGF_MENU = 2               ' /* The input event occurred in a menu. */
  MSGF_SCROLLBAR = 5          ' /* The input event occurred in a scroll bar. */
  MSGF_NEXTWINDOW = 6         ' /* The next window action is about to take place. */
  MSGF_MAX = 8                ' /* unused */
  MSGF_USER = 4096            ' /* Define MSGF_HOOK to a value greater than or equal to MSGF_USER
                              '  * defined in WINDOWS.H to prevent collision with values used by
                              '  * Windows.
                              '  */
  MSGF_DDEMGR = &H8001        ' /* The input event occurred while the Dynamic Data Exchange
                              '  * Management Library (DDEML) was waiting for a synchronous
                              '  * transaction to finish.
                              '  */
End Enum

Public Enum SYSCOMMAND        ' Specifies the type of system command requested for WM_COMMAND.
 SC_SIZE = &HF000             ' /* Sizes the window. */
 SC_MOVE = &HF010             ' /* Moves the window. */
 SC_MINIMIZE = &HF020         ' /* Minimizes the window. */
 SC_MAXIMIZE = &HF030         ' /* Maximizes the window. */
 SC_NEXTWINDOW = &HF040       ' /* Moves to the next window. */
 SC_PREVWINDOW = &HF050       ' /* Moves to the previous window. */
 SC_CLOSE = &HF060            ' /* Closes the window. */
 SC_VSCROLL = &HF070          ' /* Scrolls vertically. */
 SC_HSCROLL = &HF080          ' /* Scrolls horizontally. */
 SC_MOUSEMENU = &HF090        ' /* Retrieves the window menu as a result of a mouse click. */
 SC_KEYMENU = &HF100          ' /* Retrieves the window menu as a result of a keystroke. */
 SC_ARRANGE = &HF110
 SC_RESTORE = &HF120          ' /* Restores the window to its normal position and size. */
 SC_TASKLIST = &HF130         ' /* Activates the Start menu. */
 SC_SCREENSAVE = &HF140       ' /* Executes the screen saver application specified in the [boot]
                              '  * section of the System.ini file.
                              '  */
 SC_HOTKEY = &HF150           ' /* Activates the window associated with the application-specified
                              '  * hot key. The lParam parameter identifies the window to activate.
                              '  */
#If (WIN32_IE >= &H400) Then
 SC_DEFAULT = &HF160          ' /* Selects the default item; the user double-clicked the window
                              '  * menu.
                              '  */
 SC_MONITORPOWER = &HF170     ' /* Sets the state of the display. This command supports devices
                              '  * that have power-saving features, such as a battery-powered
                              '  * personal computer.
                              '  * The lParam parameter can have the following values:
                              '  * 1 - the display is going to low power
                              '  * 2 - the display is being shut off
                              '  */
 SC_CONTEXTHELP = &HF180      ' /* Changes the cursor to a question mark with a pointer. If the
                              '  * user then clicks a control in the dialog box, the control
                              '  * receives a WM_HELP message.
                              '  */
 SC_SEPARATOR = &HF00F
#End If '/* WINVER >= 0x0400 */
'/*
' * Obsolete names
' */
 SC_ICON = SC_MINIMIZE
 SC_ZOOM = SC_MAXIMIZE
End Enum

Public Enum ANCESTOR_WINDOW   ' /* Specifies the ancestor to be retrieved. */
  GA_MIC = 1
  GA_PARENT = 1               ' /* Retrieves the parent window. This does not include the owner,
                              '  * as it does with the GetParent function.
                              '  */
  GA_ROOT = 2                 ' /* Retrieves the root window by walking the chain of parent
                              '  * windows.
                              '  */
  GA_ROOTOWNER = 3            ' /* Retrieves the owned root window by walking the chain of parent
                              '  * and owner windows returned by GetParent.
                              '  */
  GA_MAC = 4
End Enum

Public Enum PEN_STYLE
  Rem /* Pen Styles */
  PS_SOLID = 0              ' /* The pen is solid. */
  '/* -------  */
  PS_DASH = 1               ' /* The pen is dashed. This style is valid only when the pen width is
                            '  * one or less in device units.
                            '  */
  '/* .......  */
  PS_DOT = 2                ' /* The pen is dotted. This style is valid only when the pen width is
                            '  * one or less in device units.
                            '  */
  '/* _._._._  */
  PS_DASHDOT = 3            ' /* The pen has alternating dashes and dots. This style is valid only
                            '  * when the pen width is one or less in device units.
                            '  */
  '/* _.._.._  */
  PS_DASHDOTDOT = 4         ' /* The pen has alternating dashes and double dots. This style is valid
                            '  * only when the pen width is one or less in device units.
                            '  */
  PS_NULL = 5               ' /* The pen is invisible. */
  PS_INSIDEFRAME = 6        ' /* The pen is solid. When this pen is used in any GDI drawing function
                            '  * that takes a bounding rectangle, the dimensions of the figure are
                            '  * shrunk so that it fits entirely in the bounding rectangle, taking into
                            '  * account the width of the pen. This applies only to geometric pens.
                            '  */
  PS_USERSTYLE = 7          ' /* Windows NT/2000/XP: The pen uses a styling array supplied by the
                            '  * user.
                            '  */
  PS_ALTERNATE = 8          ' /* Windows NT/2000/XP: The pen sets every other pixel. (This style is
                            '  * applicable only for cosmetic pens.)
                            '  */
  PS_STYLE_MASK = &HF

  Rem /* The end cap is only specified for geometric pens. */
  PS_ENDCAP_ROUND = &H0     ' /* End caps are round. */
  PS_ENDCAP_SQUARE = &H100  ' /* End caps are square. */
  PS_ENDCAP_FLAT = &H200    ' /* End caps are flat. */
  PS_ENDCAP_MASK = &HF00

  Rem /* The join is only specified for geometric pens. */
  PS_JOIN_ROUND = &H0       ' /* Joins are round. */
  PS_JOIN_BEVEL = &H1000    ' /* Joins are beveled. */
  PS_JOIN_MITER = &H2000    ' /* Joins are mitered when they are within the current limit set by the
                            '  * SetMiterLimit function. If it exceeds this limit, the join is beveled.
                            '  */
  PS_JOIN_MASK = &HF000

  PS_COSMETIC = &H0         ' /* The pen is cosmetic. */
  PS_GEOMETRIC = &H10000    ' /* The pen is geometric. */
  PS_TYPE_MASK = &HF0000
End Enum

Public Enum EXIT_WINDOWS_FLAGS
  EWX_LOGOFF = 0            ' /* Shuts down all processes running in the security context of the process
                            '  * that called the ExitWindowsEx function. Then it logs the user off.
                            '  */
  EWX_SHUTDOWN = &H1        ' /* Shuts down the system to a point at which it is safe to turn off the
                            '  * power. All file buffers have been flushed to disk, and all running
                            '  * processes have stopped. If the system supports the power-off feature,
                            '  * the power is also turned off.
                            '  * Windows NT/2000/XP: The calling process must have the SE_SHUTDOWN_NAME
                            '  * privilege.
                            '  */
  EWX_REBOOT = &H2          ' /* Shuts down the system and then restarts the system.
                            '  * Windows NT/2000/XP: The calling process must have the SE_SHUTDOWN_NAME
                            '  * privilege.
                            '  */
  EWX_POWEROFF = &H8        ' /* Shuts down the system and turns off the power. The system must support
                            '  * the power-off feature.
                            '  * Windows NT/2000/XP: The calling process must have the SE_SHUTDOWN_NAME
                            '  * privilege.
                            '  */
    
  Rem /* This parameter can optionally include the following values. */
  EWX_FORCE = &H4           ' /* Forces processes to terminate. When this flag is set, the system does
                            '  * not send the WM_QUERYENDSESSION and WM_ENDSESSION messages. This can cause
                            '  * the applications to lose data. Therefore, you should only use this flag
                            '  * in an emergency.
                            '  */
#If (WIN32_IE >= &H500) Then
  EWX_FORCEIFHUNG = &H10    ' /* Windows 2000/XP: Forces processes to terminate if they do not respond
                            '  * to the WM_QUERYENDSESSION or WM_ENDSESSION message. This flag is
                            '  * ignored if EWX_FORCE is used.
                            '  */
#End If '/* _WIN32_WINNT >= 0x0500 */
End Enum

Public Enum SHUTDOWN_ACTION
  LOGOFF = EXIT_WINDOWS_FLAGS.EWX_LOGOFF
  REBOOT = EXIT_WINDOWS_FLAGS.EWX_REBOOT
  SHUTDWN = EXIT_WINDOWS_FLAGS.EWX_SHUTDOWN
  POWEROFF = EXIT_WINDOWS_FLAGS.EWX_POWEROFF
End Enum


Public Enum SHUTDOWN_REASON
  Rem /* Starting with Windows XP, the system allows the user to document the reason for shutting down
  Rem  * or restarting the system. This feature is called the Shutdown Event Tracker. It is enabled by
  Rem  * default in Windows .NET Server. The user is prompted to fill in information when selecting
  Rem  * Shut Down from the Start menu, or when using Shutdown.exe. The information is stored in the
  Rem  * event log. For more information, see the help for Shutdown Event Tracker included in the
  Rem  * operating system.
  Rem  * The ExitWindowsEx and InitiateSystemShutdownEx functions have been updated to support shutdown
  Rem  * reason codes in the dwReason parameter. Use the values defined in Reason.h to construct a
  Rem  * shutdown reason code. A shutdown reason code is constructed from a major flag, a minor flag,
  Rem  * and two additional flags.
  Rem  */
  
  Rem  /* The following are the major reason flags. They indicates the general issue type. */
  SHTDN_REASON_MAJOR_APPLICATION      ' /* Application issue. */
  SHTDN_REASON_MAJOR_HARDWARE         ' /* Hardware issue. */
  SHTDN_REASON_MAJOR_LEGACY_API       ' /* The InitiateSystemShutdown function was used instead of
                                      '  * InitiateSystemShutdownEx.
                                      '  */
  SHTDN_REASON_MAJOR_OPERATINGSYSTEM  ' /* Operating system issue. */
  SHTDN_REASON_MAJOR_OTHER            ' /* Other issue. */
  SHTDN_REASON_MAJOR_POWER            ' /* Power failure. */
  SHTDN_REASON_MAJOR_SAFE_MODE        ' /* Safe mode. */
  SHTDN_REASON_MAJOR_SOFTWARE         ' /* Software issue. */
  SHTDN_REASON_MAJOR_SYSTEM           ' /* System failure. */
  
  Rem /* The following are the minor reason flags. They modify the specified major reason flag. You
  Rem  * can use any minor reason in conjunction with any major reason, but some combinations do not
  Rem  * make sense.
  Rem  */
  SHTDN_REASON_MINOR_BLUESCREEN       ' /* Blue screen crash event. */
  SHTDN_REASON_MINOR_CORDUNPLUGGED    ' /* Unplugged. */
  SHTDN_REASON_MINOR_DISK             ' /* Disk. */
  SHTDN_REASON_MINOR_ENVIRONMENT      ' /* Environment. */
  SHTDN_REASON_MINOR_HARDWARE_DRIVER  ' /* Driver. */
  SHTDN_REASON_MINOR_HOTFIX           ' /* Hot fix. */
  SHTDN_REASON_MINOR_HUNG             ' /* Unresponsive. */
  SHTDN_REASON_MINOR_INSTALLATION     ' /* Installation. */
  SHTDN_REASON_MINOR_MAINTENANCE      ' /* Maintenance. */
  SHTDN_REASON_MINOR_NETWORKCARD      ' /* Network card. */
  SHTDN_REASON_MINOR_OTHER            ' /* Other issue. */
  SHTDN_REASON_MINOR_OTHERDRIVER      ' /* Other driver event. */
  SHTDN_REASON_MINOR_POWER_SUPPLY     ' /* Power supply. */
  SHTDN_REASON_MINOR_PROCESSOR        ' /* Processor. */
  SHTDN_REASON_MINOR_RECONFIG         ' /* Reconfigure. */
  SHTDN_REASON_MINOR_SECURITYFIX      ' /* Security fix. */
  SHTDN_REASON_MINOR_SERVICEPACK      ' /* Service pack. */
  SHTDN_REASON_MINOR_UNSTABLE         ' /* Unstable. */
  SHTDN_REASON_MINOR_UPGRADE          ' /* Upgrade. */
  
  Rem /* The following flags provide additional information about the event. */
  SHTDN_REASON_FLAG_USER_DEFINED      ' /* The reason code is defined by the user.
                                      '  *If this flag is not present, the reason code is defined by
                                      '  * the system.
                                      '  */
  SHTDN_REASON_FLAG_PLANNED           ' /* The shutdown was planned. On Windows .NET Server, the system
                                      '  * generates a state snapshot. For more information, see the
                                      '  * help for Shutdown Event Tracker.
                                      '  * If this flag is not present, the shutdown was unplanned.
                                      '  */

  
  Rem /* The following combinations are recognized by the system. The description text is stored with the
  Rem  * error code in the event log.
  Rem  */
  Rem /* ' /* A restart or shutdown to troubleshoot an unresponsive application. */
  SHTDN_REASON_TROUBLESHOOT = (SHTDN_REASON_MAJOR_APPLICATION Or SHTDN_REASON_MINOR_HUNG)
  Rem /* A restart or shutdown to perform application installation. */
  SHTDN_REASON_APPINSTL = (SHTDN_REASON_MAJOR_APPLICATION Or SHTDN_REASON_MINOR_INSTALLATION)
  Rem /* A restart or shutdown to service an application. */
  SHTDN_REASON_APPSVC = (SHTDN_REASON_MAJOR_APPLICATION Or SHTDN_REASON_MINOR_MAINTENANCE)
  Rem /* A restart or shutdown to perform planned maintenance on an application. */
  SHTDN_REASON_PLAN_MAINT = (SHTDN_REASON_MAJOR_APPLICATION Or SHTDN_REASON_MINOR_MAINTENANCE Or SHTDN_REASON_FLAG_PLANNED)
  Rem /* A restart or shutdown to troubleshoot an unstable application. */
  SHTDN_REASON_UNSTBL_APP_TRBLESHT = (SHTDN_REASON_MAJOR_APPLICATION Or SHTDN_REASON_MINOR_UNSTABLE)
  Rem /* A restart or shutdown to begin or complete hardware installation. */
  SHTDN_REASON_HWINSTL = (SHTDN_REASON_MAJOR_HARDWARE Or SHTDN_REASON_MINOR_INSTALLATION)
  Rem /* A restart or shutdown to service hardware on the system. */
  SHTDN_REASON_HWSVC = (SHTDN_REASON_MAJOR_HARDWARE Or SHTDN_REASON_MINOR_MAINTENANCE)
  Rem /* This shutdown was initiated by the legacy InitiateSystemShutdown function. */
  SHTDN_REASON_LEGACY = (SHTDN_REASON_MAJOR_LEGACY_API)
  Rem /* A restart or shutdown to install a hot fix. */
  SHTDN_REASON_HOTFIX = (SHTDN_REASON_MAJOR_OPERATINGSYSTEM Or SHTDN_REASON_MINOR_HOTFIX)
  Rem /* A restart or shutdown to change the operating system configuration. */
  SHTDN_REASON_OS_CNFG = (SHTDN_REASON_MAJOR_OPERATINGSYSTEM Or SHTDN_REASON_MINOR_RECONFIG)
  Rem /* A restart or shutdown to install a security fix. */
  SHTDN_REASON_SECURITY_FIX = (SHTDN_REASON_MAJOR_OPERATINGSYSTEM Or SHTDN_REASON_MINOR_SECURITYFIX)
  Rem /* A restart or shutdown to install a service pack. */
  SHTDN_REASON_SP_INSTL = (SHTDN_REASON_MAJOR_OPERATINGSYSTEM Or SHTDN_REASON_MINOR_SERVICEPACK)
  Rem /* A restart or shutdown to upgrade the operating system configuration. */
  SHTDN_REASON_OS_CNFG_UPGD = (SHTDN_REASON_MAJOR_OPERATINGSYSTEM Or SHTDN_REASON_MINOR_UPGRADE)
  Rem /* A shutdown or restart for an unknown reason. */
  SHTDN_REASON_UNKNOWN = (SHTDN_REASON_MAJOR_OTHER Or SHTDN_REASON_MINOR_OTHER)
  Rem /* The system became unresponsive. */
  SHTDN_REASON_SYS_HUNG = (SHTDN_REASON_MAJOR_OTHER Or SHTDN_REASON_MINOR_HUNG)
  Rem  /* The computer was unplugged. */
  SHTDN_REASON_PCUNPLUG = (SHTDN_REASON_MAJOR_POWER Or SHTDN_REASON_MINOR_CORDUNPLUGGED)
  Rem /* There was a power outage. */
  SHTDN_REASON_POWEROUT = (SHTDN_REASON_MAJOR_POWER Or SHTDN_REASON_MINOR_ENVIRONMENT)
  Rem /* This is a safe mode shutdown. */
  SHTDN_REASON_SAFEMODE = (SHTDN_REASON_MAJOR_SAFE_MODE)
  Rem /* The computer displayed a blue screen crash event. */
  SHTDN_REASON_BLUESCREEN = (SHTDN_REASON_MAJOR_SYSTEM Or SHTDN_REASON_MINOR_BLUESCREEN)
  Rem /* You can also define your own shutdown reasons and add them to the registry. */
  SHTDN_REASON_UNSPECIFIED = &HFFFFFFFF ' /* Reason used to specify for Win95/98/Me OSs'. */
End Enum

Public Enum PRIVILEGE_ATTR
  SE_PRIVILEGE_ENABLED_BY_DEFAULT = (&H1) ' /* The privilege is enabled by default. */
  SE_PRIVILEGE_ENABLED = (&H2)    ' /* The privilege is enabled. */
  SE_PRIVILEGE_USED_FOR_ACCESS = (&H80000000) ' /* The privilege was used to gain access to an object
                                  '  * or service. This flag is used to identify the relevant privileges
                                  '  * in a set passed by a client application that may contain unnecessary
                                  '  * privileges.
                                  '  */
End Enum

Public Enum PRIVILEGE_CONTROL_ATTR
  PRIVILEGE_SET_ALL_NECESSARY = (1) ' /* The PRIVILEGE_SET_ALL_NECESSARY control flag is currently defined.
                                  '  * It indicates that all of the specified privileges must be held
                                  '  * by the process requesting access.
                                  '  */
End Enum

Public Enum TOKEN_INFORMATION_CLASS ' /* The TOKEN_INFORMATION_CLASS enumeration type contains values
                                    '  * that specify the type of information being assigned to or
                                    '  * retrieved from an access token.
                                    '  * The GetTokenInformation function uses these enumerator values
                                    '  * to indicate the type of token information to retrieve
                                    '  */
  TokenUser = 1                     ' /* The buffer receives a TOKEN_USER structure containing the
                                    '  * token's user account.
                                    '  */
  TokenGroups = 2                   ' /* The buffer receives a TOKEN_GROUPS structure containing the
                                    '  * group accounts associated with the token.
                                    '  */
  TokenPrivileges = 3               ' /* The buffer receives a TOKEN_PRIVILEGES structure containing
                                    '  * the token's privileges.
                                    '  */
  TokenOwner = 4                    ' /* The buffer receives a TOKEN_OWNER structure containing the
                                    '  * default owner SID for newly created objects.
                                    '  */
  TokenPrimaryGroup = 5             ' /* The buffer receives a TOKEN_PRIMARY_GROUP structure containing
                                    '  * the default primary group SID for newly created objects.
                                    '  */
  TokenDefaultDacl = 6              ' /* The buffer receives a TOKEN_DEFAULT_DACL structure containing
                                    '  * the default DACL for newly created objects.
                                    '  */
  TokenSource = 7                   ' /* The buffer receives a TOKEN_SOURCE structure containing the
                                    '  * source of the token. TOKEN_QUERY_SOURCE access is needed to
                                    '  * retrieve this information.
                                    '  */
  TokenType = 8                     ' /* The buffer receives a TOKEN_TYPE value indicating whether the
                                    '  * token is a primary or impersonation token.
                                    '  */
  TokenImpersonationLevel = 9       ' /* The buffer receives a SECURITY_IMPERSONATION_LEVEL value
                                    '  * indicating the impersonation level of the token. If the
                                    '  * access token is not an impersonation token, the function
                                    '  * fails.
                                    '  */
  TokenStatistics = 10              ' /* The buffer receives a TOKEN_STATISTICS structure containing
                                    '  * various token statistics.
                                    '  */
  TokenRestrictedSids = 11          ' /* The buffer receives a TOKEN_GROUPS structure containing the
                                    '  * list of restricting SIDs in a restricted token.
                                    '  */
  TokenSessionId = 12               ' /* Terminal Services: The buffer receives a DWORD value that
                                    '  * indicates the Terminal Services session identifier associated
                                    '  * with the token. If the token is associated with the Terminal
                                    '  * Server console session, the session identifier is zero. A nonzero
                                    '  * session identifier indicates a Terminal Services client session.
                                    '  * In a non-Terminal Services environment, the session identifier is
                                    '  * zero. If TokenSessionId is set with SetTokenInformation, the
                                    '  * application must have the Act As Part Of the Operating System
                                    '  * privilege and the application must be enabled to set the session
                                    '  * ID in a token.
                                    '  */
  TokenGroupsAndPrivileges = 13     ' /* The buffer receives a TOKEN_GROUPS_AND_PRIVILEGES structure
                                    '  * containing the user SID, the group accounts, the restricted
                                    '  * SIDs, and the authentication ID associated with the token.
                                    '  */
  TokenSessionReference = 14        ' /* Reserved for internal use. */
  TokenSandBoxInert = 15            ' /* The buffer receives a DWORD value that is nonzero if the
                                    '  * token includes the SANDBOX_INERT flag.
                                    '  */
End Enum

Public Enum STD_ACCESS_RIGHTS ' /* Each type of securable object has a set of access rights that
                              '  * correspond to operations specific to that type of object. In
                              '  * addition to these object-specific access rights, there is a set
                              '  * of standard access rights that correspond to operations common to
                              '  * most types of securable objects.
                              '  * The Windows 2000/Windows NT access mask format includes a set of
                              '  * bits for the standard access rights.
                              '  */
  DELETE_CONTROL = (&H10000)  ' /* The right to delete the object. */
  READ_CONTROL = (&H20000)    ' /* The right to read the information in the object's security
                              '  * descriptor, not including the information in the SACL.
                              '  */
  WRITE_DAC = (&H40000)       ' /* The right to modify the DACL in the object's security descriptor. */
  WRITE_OWNER = (&H80000)     ' /* The right to change the owner in the object's security descriptor. */
  Synchronize = (&H100000)    ' /* The right to use the object for synchronization. This enables a
                              '  * thread to wait until the object is in the signaled state. Some
                              '  * object types do not support this access right.
                              '  */

  STANDARD_RIGHTS_REQUIRED = (&HF0000)  ' /* Combines DELETE, READ_CONTROL, WRITE_DAC, and
                                        '  * WRITE_OWNER access.
                                        '  */

  STANDARD_RIGHTS_READ = (READ_CONTROL)    ' /* Currently defined to equal READ_CONTROL. */
  STANDARD_RIGHTS_WRITE = (READ_CONTROL)   ' /* Currently defined to equal READ_CONTROL. */
  STANDARD_RIGHTS_EXECUTE = (READ_CONTROL) ' /* Currently defined to equal READ_CONTROL. */

  STANDARD_RIGHTS_ALL = (&H1F0000)  ' /* Combines DELETE, READ_CONTROL, WRITE_DAC, WRITE_OWNER, and
                                    '  * SYNCHRONIZE access.
                                    '  */
  SPECIFIC_RIGHTS_ALL = (&HFFFF)
End Enum

Public Enum TOKEN_RIGHTS
  Rem /* Token Specific Access Rights. */
  TOKEN_ASSIGN_PRIMARY = (&H1)
  TOKEN_DUPLICATE = (&H2)
  TOKEN_IMPERSONATE = (&H4)
  TOKEN_QUERY = (&H8)
  TOKEN_QUERY_SOURCE = (&H10)
  TOKEN_ADJUST_PRIVILEGES = (&H20)
  TOKEN_ADJUST_GROUPS = (&H40)
  TOKEN_ADJUST_DEFAULT = (&H80)
  TOKEN_ADJUST_SESSIONID = (&H100)
  
  TOKEN_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or TOKEN_ASSIGN_PRIMARY Or TOKEN_DUPLICATE Or TOKEN_IMPERSONATE Or TOKEN_QUERY Or TOKEN_QUERY_SOURCE Or TOKEN_ADJUST_PRIVILEGES Or TOKEN_ADJUST_GROUPS Or TOKEN_ADJUST_SESSIONID Or TOKEN_ADJUST_DEFAULT)
  
  TOKEN_READ = (STANDARD_RIGHTS_READ Or TOKEN_QUERY)
  
  TOKEN_WRITE = (STANDARD_RIGHTS_WRITE Or TOKEN_ADJUST_PRIVILEGES Or TOKEN_ADJUST_GROUPS Or TOKEN_ADJUST_DEFAULT)
  
  TOKEN_EXECUTE = (STANDARD_RIGHTS_EXECUTE)
End Enum

Rem /* Standard Windows NT defined priviliges, */
Public Const SE_CREATE_TOKEN_NAME               As String = "SeCreateTokenPrivilege"
Public Const SE_ASSIGNPRIMARYTOKEN_NAME         As String = "SeAssignPrimaryTokenPrivilege"
Public Const SE_LOCK_MEMORY_NAME                As String = "SeLockMemoryPrivilege"
Public Const SE_INCREASE_QUOTA_NAME             As String = "SeIncreaseQuotaPrivilege"
Public Const SE_UNSOLICITED_INPUT_NAME          As String = "SeUnsolicitedInputPrivilege"
Public Const SE_MACHINE_ACCOUNT_NAME            As String = "SeMachineAccountPrivilege"
Public Const SE_TCB_NAME                        As String = "SeTcbPrivilege"
Public Const SE_SECURITY_NAME                   As String = "SeSecurityPrivilege"
Public Const SE_TAKE_OWNERSHIP_NAME             As String = "SeTakeOwnershipPrivilege"
Public Const SE_LOAD_DRIVER_NAME                As String = "SeLoadDriverPrivilege"
Public Const SE_SYSTEM_PROFILE_NAME             As String = "SeSystemProfilePrivilege"
Public Const SE_SYSTEMTIME_NAME                 As String = "SeSystemtimePrivilege"
Public Const SE_PROF_SINGLE_PROCESS_NAME        As String = "SeProfileSingleProcessPrivilege"
Public Const SE_INC_BASE_PRIORITY_NAME          As String = "SeIncreaseBasePriorityPrivilege"
Public Const SE_CREATE_PAGEFILE_NAME            As String = "SeCreatePagefilePrivilege"
Public Const SE_CREATE_PERMANENT_NAME           As String = "SeCreatePermanentPrivilege"
Public Const SE_BACKUP_NAME                     As String = "SeBackupPrivilege"
Public Const SE_RESTORE_NAME                    As String = "SeRestorePrivilege"
Public Const SE_SHUTDOWN_NAME                   As String = "SeShutdownPrivilege"
Public Const SE_DEBUG_NAME                      As String = "SeDebugPrivilege"
Public Const SE_AUDIT_NAME                      As String = "SeAuditPrivilege"
Public Const SE_SYSTEM_ENVIRONMENT_NAME         As String = "SeSystemEnvironmentPrivilege"
Public Const SE_CHANGE_NOTIFY_NAME              As String = "SeChangeNotifyPrivilege"
Public Const SE_REMOTE_SHUTDOWN_NAME            As String = "SeRemoteShutdownPrivilege"

Public Enum CLASS_STYLE         ' /* The class styles define additional elements of the window class.
                                '  * Two or more styles can be combined by using the bitwise OR (|)
                                '  * operator. To assign a style to a window class, assign the style
                                '  * to the style member of the WNDCLASSEX structure.
                                '  */
  CS_BYTEALIGNCLIENT = &H1000   ' /* Aligns the window's client area on a byte boundary (in the x
                                '  * direction). This style affects the width of the window and its
                                '  * horizontal placement on the display.
                                '  */
  CS_BYTEALIGNWINDOW = &H2000   ' /* Aligns the window on a byte boundary (in the x direction). This
                                '  * style affects the width of the window and its horizontal placement
                                '  * on the display.
                                '  */
  CS_CLASSDC = &H40             ' /* Allocates one device context to be shared by all windows in the
                                '  * class. Because window classes are process specific, it is possible
                                '  * for multiple threads of an application to create a window of the
                                '  * same class. It is also possible for the threads to attempt to use
                                '  * the device context simultaneously. When this happens , the system
                                '  * allows only one thread to successfully finish its drawing operation.
                                '  */
  CS_DBLCLKS = &H8              ' /* Sends a double-click message to the window procedure when the user
                                '  * double-clicks the mouse while the cursor is within a window belonging
                                '  * to the class.
                                '  */
  CS_HREDRAW = &H2              ' /* Redraws the entire window if a movement or size adjustment changes
                                '  * the width of the client area.
                                '  */
  CS_INSERTCHAR = &H2000
  CS_KEYCVTWINDOW = &H4
  CS_NOCLOSE = &H200            ' /* Disables Close on the window menu. */
  CS_NOKEYCVT = &H100
  CS_NOMOVECARET = &H4000
  CS_OWNDC = &H20               ' /* Allocates a unique device context for each window in the class.  */
  CS_PARENTDC = &H80            ' /* Sets the clipping rectangle of the child window to that of the parent
                                '  * window so that the child can draw on the parent. A window with the
                                '  * CS_PARENTDC style bit receives a regular device context from the system's
                                '  * cache of device contexts. It does not give the child the parent's device
                                '  * context or device context settings. Specifying CS_PARENTDC enhances
                                '  * an application's performance.
                                '  */
  CS_PUBLICCLASS = &H4000       ' /* Same as CS_GLOBALCLASS. */
  CS_SAVEBITS = &H800           ' /* Saves, as a bitmap, the portion of the screen image obscured by a window
                                '  * of this class. When the window is removed, the system uses the saved bitmap
                                '  * to restore the screen image, including other windows that were obscured.
                                '  * Therefore, the system does not send WM_PAINT messages to windows that were
                                '  * obscured if the memory used by the bitmap has not been discarded and if other
                                '  * screen actions have not invalidated the stored image.
                                '  * This style is useful for small windows (for example, menus or dialog boxes)
                                '  * that are displayed briefly and then removed before other screen activity takes
                                '  * place. This style increases the time required to display the window, because
                                '  * the system must first allocate memory to store the bitmap.
                                '  */
  CS_VREDRAW = &H1              ' /* Redraws the entire window if a movement or size adjustment changes the height
                                '  * of the client area.
                                '  */
  CS_DROPSHADOW = &H20000       ' /* Windows XP: Enables the drop shadow effect on a window. The effect is
                                '  * turned on and off through SPI_SETDROPSHADOW. Typically, this is enabled
                                '  * for small, short-lived windows such as menus to emphasize their Z order
                                '  * relationship to other windows.
                                '  */
  CS_GLOBALCLASS = &H4000       ' /* Specifies that the window class is an application global class. For more
                                '  * information, see Application Global Classes.
                                '  */
End Enum




'get network username. see GetNetUserName()
Private Declare Function WNetGetUser Lib "mpr.dll" Alias "WNetGetUserA" ( _
    ByVal lpName As String, lpUserName As String, lpnLength As Long) As Long
    

Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
'   Finds Window Handle, given the window title.

Public Declare Function ShowCursor& Lib "user32" (ByVal bShow As Long)
    'The Parameter bShow is Set To True (non-zero) to display the cursor, False to hide it.

Public Declare Function GetTickCount Lib "kernel32" () As Long

' Same as Ambient.LocaleID (but only the language version of Windows)
Public Declare Function GetSystemDefaultLCID Lib "kernel32" () As Long

' same as Ambient.LocaleID (but based on user's setting. Use this instead
' of GetSystemDefaultLCID)
Public Declare Function GetUserDefaultLCID Lib "kernel32" () As Long


'Rem This function fills the specified buffer with the metrics for the currently selected font.
'Rem @hDC            :   [in] Handle to the device context (DC).
'Rem @lptm           :   [out] Long pointer to the TEXTMETRIC structure that is to receive the metrics.
'Rem Return Values   :   Nonzero indicates success. Zero indicates failure.
'Public Declare Function GetTextMetrics Lib "gdi32" _
'                  Alias "GetTextMetricsA" (ByVal hDC As stdole.OLE_HANDLE, _
'                                           ByRef lpMetrics As TEXTMETRIC) As Long



Rem-------------------------------------------------------------------------
Rem @Name                 :                 APICoordinator
Rem @Type                 :                 Standard
Rem @Scope Qualifier      :                 Private
Rem @Purpose              :                 Provides and coordinates various
Rem                                         useful Win32 APIs' used by the
Rem                                         solution.
Rem @Creation Date        :                 Saturday, 27 January 2002.
Rem @Creation Author      :                 Shantibhushan
Rem-------------------------------------------------------------------------

Rem This function registers specific common controls classes from the common control
Rem dynamic-link library (DLL).
Rem @lpInitCtrls    :   Long pointer to an LPINITCOMMONCONTROLSEX structure that contains
Rem                     information specifying which control classes are registered.
Rem @Return Values  :   TRUE indicates success. FALSE indicates failure.
Public Declare Function _
  InitCommonControlsEx Lib "COMCTL32.DLL" (ByRef lpInitCtrls As LPINITCOMMONCONTROLSEX) _
As Boolean                                                                                ' Required to initialize the common
                                                                                          ' controls...

Rem This function sends the specified message to a window or windows. SendMessage calls the
Rem window procedure for the specified window and does not return until the window procedure
Rem has processed the message.
Rem @hWnd             :   [in] Handle to the window whose window procedure will receive the message. If
Rem                           this parameter is HWND_BROADCAST, the message is sent to all top-level
Rem                           windows in the system, including disabled or invisible unowned windows,
Rem                           overlapped windows, and pop-up windows; but the message is not sent to
Rem                           child windows.
Rem @Msg              :   [in] Specifies the message to be sent.
Rem @wParam           :   [in] Specifies additional message-specific information.
Rem @lParam           :   [in] Specifies additional message-specific information.
Rem @Return Values    :   The return value specifies the result of the message processing and
Rem                       depends on the message sent.
Public Declare Function _
  SendMessage Lib "user32" Alias _
  "SendMessageA" (ByVal hWnd As stdole.OLE_HANDLE, _
                  ByVal wMsg As Long, _
                  ByVal wParam As Long, _
                  lParam As Any) _
As Long                                                       ' Message to be sent to members...

Rem This function creates an overlapped, pop-up, or child window with an extended style
Public Declare Function _
  CreateWindowEx Lib "user32" Alias _
  "CreateWindowExA" (ByVal dwExStyle As WINDOW_STYLE_EXTENDED, _
                     ByVal lpClassName As String, _
                     ByVal lpWindowName As String, _
                     ByVal dwStyle As TOOLTIP_STYLES, _
                     ByVal x As stdole.OLE_XPOS_PIXELS, _
                     ByVal y As stdole.OLE_YPOS_PIXELS, _
                     ByVal nWidth As stdole.OLE_XSIZE_PIXELS, _
                     ByVal nHeight As stdole.OLE_YSIZE_PIXELS, _
                     ByVal hWndParent As stdole.OLE_HANDLE, _
                     ByVal hMenu As stdole.OLE_HANDLE, _
                     ByVal hInstance As stdole.OLE_HANDLE, _
                     lpParam As Any) _
As Long                                                       ' Create a window...

Rem This function destroys the specified window. The function sends a WM_DESTROY message to the
Rem window to deactivate it and removes the keyboard focus from it. The function also destroys
Rem the window's menu, destroys timers, removes clipboard ownership, and breaks the clipboard
Rem viewer chain (if the window is at the top of the viewer chain). If the specified window is a
Rem parent or owner window, DestroyWindow automatically destroys the associated child or owned
Rem windows when it destroys the parent or owner window. The function first destroys child or owned
Rem windows, and then it destroys the parent or owner window. DestroyWindow also destroys modeless
Rem dialog boxes created by the CreateDialog function.
Rem @hWnd             :    Handle to the window to be destroyed.
Rem @Return Values    :   Nonzero indicates success. Zero indicates failure. To get extended error
Rem                       information, call GetLastError.
Public Declare Function _
  DestroyWindow Lib "user32" (ByVal hWnd As stdole.OLE_HANDLE) _
As Long                                                       ' Destroy a window...

Rem The CopyMemory function copies a block of memory from one location to another.
Rem @Destination      :   [in] Pointer to the starting address of the copied block's destination.
Rem @Source           :   [in] Pointer to the starting address of the block of memory to copy.
Rem @Length           :   [in] Specifies the size, in bytes, of the block of memory to copy.
Rem @Return Values    :   This function has no return value.
Public Declare Sub _
  CopyMemory Lib "kernel32" Alias _
  "RtlMoveMemory" (ByRef Destination As Any, _
                   ByRef Source As Any, _
                   ByVal Length As Long)                      ' Copy a given block of memory from
                                                              ' source to destination...

Rem This function retrieves the dimensions of the bounding rectangle of the specified window. The
Rem dimensions are given in screen coordinates that are relative to the upper-left corner of the
Rem screen.
Rem @hWnd             :   [in] Handle to the window.
Rem @lpRect           :   [in] Long pointer to a RECT structure that receives the screen coordinates
Rem                            of the upper-left and lower-right corners of the window.
Rem @Return Values    :   Nonzero indicates success. Zero indicates failure. To get extended error
Rem                       information, call GetLastError.
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As stdole.OLE_HANDLE, lpRect As RECT) As Long

Rem The GetClientRect function retrieves the coordinates of a window's client area. The client
Rem coordinates specify the upper-left and lower-right corners of the client area. Because client
Rem coordinates are relative to the upper-left corner of a window's client area, the coordinates of
Rem the upper-left corner are (0,0).
Rem @hWnd             :   [in] Handle to the window.
Rem @lpRect           :   [out] Pointer to a RECT structure that receives the client coordinates.
Rem                       The left and top members are zero. The right and bottom members contain the
Rem                       width and height of the window.
Rem @Return Values    :   Nonzero indicates success. Zero indicates failure. To get extended error
Rem                       information, call GetLastError.
Public Declare Function GetClientRect Lib "user32" (ByVal hWnd As stdole.OLE_HANDLE, lpRect As RECT) As Long

Rem Sends a message to the taskbar's status area.
Rem @dwMessage        :   [in] Variable of type DWORD that specifies the action to be taken.
Rem @lpData           :   Address of a NOTIFYICONDATA structure. The content of the structure depends
Rem                       on the value of dwMessage.
Rem @Return Values    :   Returns TRUE if successful or FALSE otherwise. If dwMessage is set to NIM_SETVERSION,
Rem                       the function returns TRUE if the version was successfully changed or FALSE if the
Rem                       requested version is not supported.
Public Declare Function _
  Shell_NotifyIcon Lib "shell32.dll" Alias _
  "Shell_NotifyIconA" (ByVal dwMessage As TRAY_MESSAGES, _
                       ByRef lpData As NOTIFYICONDATAEX) As Long

Rem The GetVersionEx function obtains extended information about the version of the operating system that is currently
Rem running.
Rem Windows 2000/XP: To compare the current system version to a required version, use the VerifyVersionInfo function instead
Rem of using GetVersionEx to perform the comparison yourself.
Rem @lpVersionInformation :   [in/out] Pointer to an OSVERSIONINFO data structure that the function fills with operating system
Rem                           version information. Before calling the GetVersionEx function, set the dwOSVersionInfoSize member
Rem                           of the OSVERSIONINFO data structure to sizeof(OSVERSIONINFO). Windows NT 4.0 SP6 and later: This
Rem                           member can be a pointer to an OSVERSIONINFOEX structure. Set the dwOSVersionInfoSize member to
Rem                           sizeof(OSVERSIONINFOEX) to identify the structure type
Rem @Return Values    :   If the function succeeds, the return value is a nonzero value.
Public Declare Function _
  GetVersionEx Lib "kernel32" Alias _
  "GetVersionExA" (ByRef lpVersionInformation As OSVERSIONINFOEX) As Long

Rem The VerifyVersionInfo function compares a set of operating system version requirements to the corresponding values for the
Rem currently running version of the system.
Rem @lpVersionInfo       :    [in] Pointer to an OSVERSIONINFOEX structure containing the operating system version requirements to
Rem                           compare. The dwTypeMask parameter indicates the members of this structure that contain information to
Rem                           compare. You must set the dwOSVersionInfoSize member of this structure to sizeof(OSVERSIONINFOEX). You
Rem                           must also specify valid data for the members indicated by dwTypeMask. The function ignores structure
Rem                           members for which the corresponding dwTypeMask bit is not set.
Rem @dwTypeMask          :    Specifies the members of the OSVERSIONINFOEX structure to test.
Rem @dwlConditionMask    :    [in] A 64-bit value that indicates the type of comparison to use for each lpVersionInfo member being compared.
Rem                           To build this value, call the VerSetConditionMask function or the VER_SET_CONDITION macro once for each
Rem                           OSVERSIONINFOEX member being compared.
Rem @Return Values       :    If the currently running operating system satisfies the specified requirements, the return value is a nonzero value.
Rem                           The error value is ERROR_OLD_WIN_VERSION.
Public Declare Function _
  VerifyVersionInfo Lib "kernel32" Alias _
  "VerifyVersionInfoA" (ByRef lpVersionInfo As OSVERSIONINFOEX, _
                        ByVal dwTypeMask As Long, _
                        ByVal dwlConditionMask As Long) As Long

Rem The GetComputerName function retrieves the NetBIOS name of the local computer. This name is established at system startup, when the system
Rem reads it from the registry. If the local computer is a node in a cluster, GetComputerName returns the name of the node.
Rem Windows 2000/XP: GetComputerName retrieves only the NetBIOS name of the local computer. To retrieve the DNS host name, DNS domain name, or
Rem the fully qualified DNS name, call the GetComputerNameEx function.
Rem Windows 2000/XP: Additional information is provided by the IADsADSystemInfo interface.
Rem @lpBuffer           :     [out] Pointer to a buffer that receives a null-terminated string containing the computer name. The buffer size should be
Rem                           large enough to contain MAX_COMPUTERNAME_LENGTH + 1 characters.
Rem @nSize              :     [in/out] On input, specifies the size, in TCHARs, of the buffer. On output, receives the number of TCHARs copied to the
Rem                           destination buffer, not including the terminating null character.
Rem                           Windows 95/98/Me: GetComputerName fails if the input size is less than MAX_COMPUTERNAME_LENGTH + 1.
Rem @Return Values      :     If the function succeeds, the return value is a nonzero value.
Public Declare Function _
  GetComputerName Lib "kernel32" Alias _
  "GetComputerNameA" (ByVal lpBuffer As String, _
                      ByRef nSize As Long) _
As Long

Rem The GetUserName function retrieves the user name of the current thread. This is the name of the user currently logged onto the system.
Rem Windows 2000/XP: Use the GetUserNameEx function to retrieve the user name in a specified format. Additional information is provided by the IADsADSystemInfo
Rem interface.
Rem @lpBuffer           :     [out] Pointer to the buffer to receive the null-terminated string containing the user's logon name. If this buffer is not large
Rem                           enough to contain the entire user name, the function fails. A buffer size of (UNLEN + 1) characters will hold the maximum length user
Rem                           name including the terminating null character.
Rem @nSize              :     [in/out] On input, specifies the maximum size, in TCHARs, of the buffer specified by the lpBuffer parameter. On output, receives the
Rem                           number of characters copied to the buffer, including the terminating null character.
Rem @Return Values      :     If the function succeeds, the return value is a nonzero value, and the variable pointed to by nSize contains the number of TCHARs copied
Rem                           to the buffer specified by lpBuffer, including the terminating null character.
Rem                           If the function fails, the return value is zero.
Public Declare Function _
  GetUserName Lib "advapi32" Alias _
  "GetUserNameA" (ByVal lpBuffer As String, _
                  ByRef nSize As Long) _
As Long

Rem The GetWindowsDirectory function retrieves the path of the Windows directory. The Windows directory contains such files as applications, initialization files, and help files.
Rem @lpBuffer           :     [out] Pointer to the buffer to receive the null-terminated string containing the path. This path does not end with a backslash unless the Windows
Rem                           directory is the root directory. For example, if the Windows directory is named Windows on drive C, the path of the Windows directory retrieved by
Rem                           this function is C:\Windows. If the system was installed in the root directory of drive C, the path retrieved is C:\.
Rem @nSize              :     [in] Specifies the maximum size, in TCHARs, of the buffer specified by the lpBuffer parameter. This value should be set to MAX_PATH+1 to allow sufficient
Rem                           space for the path and the null terminator.
Rem @Return Values      :     If the function succeeds, the return value is the length, in TCHARs, of the string copied to the buffer, not including the terminating null character.
Rem                           If the length is greater than the size of the buffer, the return value is the size of the buffer required to hold the path. If the function fails, the
Rem                           return value is zero.
Public Declare Function _
  GetWindowsDirectory Lib "kernel32" Alias _
  "GetWindowsDirectoryA" (ByVal lpBuffer As String, _
                          ByVal nSize As Long) _
As Long

Rem The GetCurrentDirectory function retrieves the current directory for the current process.
Rem @nBufferLength      :     [in] Specifies the length, in TCHARs, of the buffer for the current directory string. The buffer length must include room for a terminating null character.
Rem @lpBuffer           :     [out] Pointer to the buffer that receives the current directory string. This null-terminated string specifies the absolute path to the current directory.
Rem @Return Values      :     If the function succeeds, the return value specifies the number of characters written to the buffer, not including the terminating null character. If the
Rem                           function fails, the return value is zero.
Public Declare Function _
  GetCurrentDirectory Lib "kernel32" Alias _
  "GetCurrentDirectoryA" (ByVal nBufferLength As Long, _
                          ByVal lpBuffer As String) _
As Long

Rem The GetSystemDirectory function retrieves the path of the system directory. The system directory contains such files as dynamic-link libraries, drivers, and font files.
Rem @lpBuffer           :     [out] Pointer to the buffer to receive the null-terminated string containing the path. This path does not end with a backslash unless the Windows
Rem                           directory is the root directory. For example, if the Windows directory is named Windows on drive C, the path of the Windows directory retrieved by
Rem                           this function is C:\Windows. If the system was installed in the root directory of drive C, the path retrieved is C:\.
Rem @nSize              :     [in] Specifies the maximum size, in TCHARs, of the buffer specified by the lpBuffer parameter. This value should be set to MAX_PATH+1 to allow sufficient
Rem                           space for the path and the null terminator.
Rem @Return Values      :     If the function succeeds, the return value is the length, in TCHARs, of the string copied to the buffer, not including the terminating null character.
Rem                           If the length is greater than the size of the buffer, the return value is the size of the buffer required to hold the path. If the function fails, the
Rem                           return value is zero.
Public Declare Function _
  GetSystemDirectory Lib "kernel32" Alias _
  "GetSystemDirectoryA" (ByVal lpBuffer As String, _
                         ByVal nSize As Long) _
As Long

Rem Takes the CSIDL of a folder and returns the path.
Rem @hwndOwner          :     Handle to an owner window. This parameter is typically set to NULL. If it is not NULL, and a dial-up connection needs to be made to access the folder, a UI
Rem                           prompt will appear in this window.
Rem @nFolder            :     CSIDL value that identifies the folder whose path is to be retrieved. Only real folders are valid. If a virtual folder is specified, this function will fail.
Rem                           You can force creation of a folder with SHGetFolderPath by combining the folder's CSIDL with CSIDL_FLAG_CREATE.
Rem @hToken             :     An access token that can be used to represent a particular user. For systems earlier than MicrosoftR WindowsR 2000, it should be set to NULL. For later systems,
Rem                           hToken is usually set to NULL. However, you might need to assign a value to hToken for those folders that can have multiple users but are treated as belonging to
Rem                           a single user. The most commonly used folder of this type is My Documents.
Rem                           The caller is responsible for correct impersonation when hToken is non-NULL. It must have appropriate security privileges for the particular user, including TOKEN_QUERY
Rem                           and TOKEN_IMPERSONATE, and the user's registry hive must be currently mounted. For more information about access control issues, see Access Control.
Rem                           Assigning the hToken parameter a value of -1 indicates the default user. This allows clients of SHGetFolderPath to find folder locations (such as the Desktop folder)
Rem                           for the default user without querying the registry.
Rem @dwFlags            :     Flags to specify which path is to be returned. It is used for cases where the folder associated with a CSIDL might be moved or renamed by the user.
Rem @pszPath            :     Pointer to a null-terminated string of length MAX_PATH that will receive the path. If an error occurs or S_FALSE is returned, this string will be empty.
Rem @Return Values      :     Value               Description
Rem                           S_OK                Success.
Rem                           S_FALSE             The CSIDL in nFolder is valid, but the folder does not exist.
Rem                           E_INVALIDARG        The CSIDL in nFolder is not valid
Public Declare Function _
  SHGetFolderPath Lib "shfolder" Alias _
  "SHGetFolderPathA" (ByVal hwndOwner As stdole.OLE_HANDLE, _
                      ByVal nFolder As CSIDL, _
                      ByVal hToken As stdole.OLE_HANDLE, _
                      ByVal dwFlags As CSIDL_FOLDERPATH, _
                      ByVal pszPath As String) As Long

Rem The GetMenu function retrieves a handle to the menu assigned to the specified window.
Rem @hWnd           :   [in] Handle to the window whose menu handle is to be retrieved.
Rem Return Values   :   The return value is a handle to the menu. If the specified window
Rem                     has no menu, the return value is NULL. If the window is a child window,
Rem                     the return value is undefined.
Public Declare Function GetMenu Lib "user32" (ByVal hWnd As stdole.OLE_HANDLE) As stdole.OLE_HANDLE

Rem The GetMenuDefaultItem function determines the default menu item on the specified menu.
Rem hMenu           :   [in] Handle to the menu for which to retrieve the default menu item.
Rem fByPos          :   [in] Specifies whether to retrieve the menu item's identifier or its
Rem                     position. If this parameter is FALSE, the identifier is returned. Otherwise,
Rem                     the position is returned.
Rem gmdiFlags       :   [in] Specifies how the function searches for menu items. This parameter
Rem                     can be zero or more of the following values.
Public Declare Function GetMenuDefaultItem Lib "user32" (ByVal hMenu As stdole.OLE_HANDLE, _
                                                          ByVal fByPos As Long, _
                                                          ByVal gmdiFlags As GMDI_FLAGS) As Long

Rem The GetMenuItemCount function determines the number of items in the specified menu.
Rem @hMenu          :   [in] Handle to the menu to be examined.
Rem Return Values   :   If the function succeeds, the return value specifies the number of items in
Rem                     the menu.If the function fails, the return value is -1.
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As stdole.OLE_HANDLE) As Long

Rem The GetMenuItemID function retrieves the menu item identifier of a menu item located at
Rem the specified position in a menu.
Rem @hMenu          :   [in] Handle to the menu that contains the item whose identifier is to be
Rem                     retrieved.
Rem @nPos           :   [in] Specifies the zero-based relative position of the menu item whose
Rem                     identifier is to be retrieved.
Rem Return Values   :   The return value is the identifier of the specified menu item. If the menu
Rem                     item identifier is NULL or if the specified item opens a submenu, the return
Rem                     value is -1.
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As stdole.OLE_HANDLE, _
                                                     ByVal nPos As Long) As Long

Rem The GetMenuItemInfo function retrieves information about a menu item.
Rem @hMenu          :   [in] Handle to the menu that contains the menu item.
Rem @uItem          :   [in] Identifier or position of the menu item to get information about. The
Rem                     meaning of this parameter depends on the value of fByPosition.
Rem @fByPosition    :   [in] Specifies the meaning of uItem. If this parameter is FALSE, uItem is a
Rem                     menu item identifier. Otherwise, it is a menu item position.
Rem @lpMenuItemInfo :   [in/out] Pointer to a MENUITEMINFO structure that specifies the information
Rem                     to retrieve and receives information about the menu item.
Rem Return Values   :   If the function succeeds, the return value is nonzero. If the function fails,
Rem                     the return value is zero.
Public Declare Function GetMenuItemInfo Lib "user32" Alias _
                        "GetMenuItemInfoA" (ByVal hMenu As stdole.OLE_HANDLE, _
                                            ByVal uItem As Long, _
                                            ByVal fByPosition As Long, _
                                            ByRef lpMenuItemInfo As MENUITEMINFOEX) As Long

Rem The GetMenuItemRect function retrieves the bounding rectangle for the specified menu item.
Rem @hWnd           :   [in] Handle to the window containing the menu.
Rem @hMenu          :   [in] Handle to a menu.
Rem @uItem          :   [in] Zero-based position of the menu item.
Rem @lprcItem       :   [out] Pointer to a RECT structure that receives the bounding rectangle of the
Rem                     specified menu item expressed in screen coordinates.
Rem Return Values   :   If the function succeeds, the return value is nonzero. If the function fails,
Rem                     the return value is zero.
Public Declare Function GetMenuItemRect Lib "user32" (ByVal hWnd As stdole.OLE_HANDLE, _
                                                       ByVal hMenu As stdole.OLE_HANDLE, _
                                                       ByVal uItem As Long, _
                                                       ByRef lprcItem As RECT) As Long

Rem The GetSubMenu function retrieves a handle to the drop-down menu or submenu activated by
Rem the specified menu item.
Rem @hMenu          :   [in] Handle to the menu.
Rem @nPos           :   [in] Specifies the zero-based relative position in the specified menu
Rem                     of an item that activates a drop-down menu or submenu.
Rem Return Values   :   If the function succeeds, the return value is a handle to the drop-down
Rem                     menu or submenu activated by the menu item. If the menu item does not
Rem                     activate a drop-down menu or submenu, the return value is NULL.
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As stdole.OLE_HANDLE, _
                                                  ByVal nPos As Long) As stdole.OLE_HANDLE

Rem The GetSystemMenu function allows the application to access the window menu (also known as
Rem the system menu or the control menu) for copying and modifying.
Rem @hWnd           :   [in] Handle to the window that will own a copy of the window menu.
Rem @bRevert        :   [in] Specifies the action to be taken. If this parameter is FALSE, GetSystemMenu
Rem                     returns a handle to the copy of the window menu currently in use. The copy
Rem                     is initially identical to the window menu, but it can be modified. If this
Rem                     parameter is TRUE, GetSystemMenu resets the window menu back to the default
Rem                     state. The previous window menu, if any, is destroyed.
Rem Return Values   :   If the bRevert parameter is FALSE, the return value is a handle to a copy of
Rem                     the window menu. If the bRevert parameter is TRUE, the return value is NULL.
Public Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As stdole.OLE_HANDLE, _
                                                     ByVal bRevert As Long) As stdole.OLE_HANDLE

Rem The SetMenu function assigns a new menu to the specified window.
Rem @hWnd           :   [in] Handle to the window to which the menu is to be assigned.
Rem @hMenu          :   [in] Handle to the new menu. If this parameter is NULL, the window's current
Rem                     menu is removed.
Rem Return Values   :   If the function succeeds, the return value is nonzero. If the function fails,
Rem                     the return value is zero.
Public Declare Function SetMenu Lib "user32" (ByVal hWnd As stdole.OLE_HANDLE, _
                                               ByVal hMenu As stdole.OLE_HANDLE) As Long

Rem The SetMenuDefaultItem function sets the default menu item for the specified menu.
Rem @hMenu          :   [in] Handle to the menu to set the default item for.
Rem @uItem          :   [in] Identifier or position of the new default menu item or -1 for no default
Rem                     item. The meaning of this parameter depends on the value of fByPos.
Rem @fByPos         :   [in] Value specifying the meaning of uItem. If this parameter is FALSE, uItem
Rem                     is a menu item identifier. Otherwise, it is a menu item position.
Rem Return Values   :   If the function succeeds, the return value is nonzero. If the function fails,
Rem                     the return value is zero.
Public Declare Function SetMenuDefaultItem Lib "user32" (ByVal hMenu As stdole.OLE_HANDLE, _
                                                          ByVal uItem As Long, _
                                                          ByVal fByPos As Long) As Long

Rem The SetMenuItemBitmaps function associates the specified bitmap with a menu item. Whether the menu
Rem item is selected or clear, the system displays the appropriate bitmap next to the menu item.
Rem @hMenu          :   [in] Handle to the menu containing the item to receive new check-mark
Rem                     bitmaps.
Rem @nPosition      :   [in] Specifies the menu item to be changed, as determined by the uFlags
Rem                     parameter.
Rem @wFlags         :   [in] Specifies how the uPosition parameter is interpreted. The uFlags
Rem                     parameter must be one of the following values.
Rem @hBitMapUnchecked:  [in] Handle to the bitmap displayed when the menu item is not selected.
Rem @hBitMapChecked :   [in] Handle to the bitmap displayed when the menu item is selected.
Rem Return Values   :   If the function succeeds, the return value is nonzero. If the function fails,
Rem                     the return value is zero.
Public Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As stdole.OLE_HANDLE, _
                                                          ByVal nPosition As Long, _
                                                          ByVal wFlags As SETMENUITEMBITMAPFLAGS, _
                                                          ByVal hBitmapUnchecked As stdole.OLE_HANDLE, _
                                                          ByVal hBitmapChecked As stdole.OLE_HANDLE) As Long

Rem The SetMenuItemInfo function changes information about a menu item.
Rem @hMenu          :   [in] Handle to the menu that contains the menu item.
Rem @uItem          :   [in] Identifier or position of the menu item to change. The meaning of this
Rem                     parameter depends on the value of fByPosition.
Rem @fByPosition    :   [in] Value specifying the meaning of uItem. If this parameter is FALSE,
Rem                     uItem is a menu item identifier. Otherwise, it is a menu item position.
Rem @lpcMenuItemInfo:   [in] Pointer to a MENUITEMINFO structure that contains information about
Rem                     the menu item and specifies which menu item attributes to change.
Rem Return Values   :   If the function succeeds, the return value is nonzero. If the function
Rem                     fails, the return value is zero.
Public Declare Function SetMenuItemInfo Lib "user32" Alias _
                        "SetMenuItemInfoA" (ByVal hMenu As stdole.OLE_HANDLE, _
                                            ByVal uItem As Long, _
                                            ByVal fByPosition As Boolean, _
                                            ByRef lpcMenuItemInfo As MENUITEMINFOEX) As Long

Rem The AnimateWindow function enables you to produce special effects when showing or hiding windows.
Rem There are three types of animation: roll, slide, and alpha-blended fade.
Rem @hWnd           :   [in] Handle to the window to animate. The calling thread must own this window.
Rem @dwTime         :   [in] Specifies how long it takes to play the animation, in milliseconds. Typically,
Rem                     an animation takes 200 milliseconds to play.
Rem @dwFlags        :   [in] Specifies the type of animation.
Rem Return Values   :   If the function succeeds, the return value is nonzero. If the function fails, the
Rem                     return value is zero.
Public Declare Function AnimateWindow Lib "user32" (ByVal hWnd As stdole.OLE_HANDLE, _
                                                    ByVal dwTime As Long, _
                                                    ByVal dwFlags As ANIMATE_WINDOW_STYLE) As Long

Rem The GetWindowLongPtr function retrieves information about the specified window. The function also retrieves
Rem the value at a specified offset into the extra window memory.
Rem If you are retrieving a pointer or a handle, this function supersedes the GetWindowLong function. (Pointers
Rem and handles are 32 bits on 32-bit Windows and 64 bits on 64-bit Windows.) To write code that is compatible with
Rem both 32-bit and 64-bit versions of Windows, use GetWindowLongPtr.
Rem @hWnd           :   [in] Handle to the window and, indirectly, the class to which the window belongs.
Rem @nIndex         :   [in] Specifies the zero-based offset to the value to be retrieved. Valid values are in the
Rem                     range zero through the number of bytes of extra window memory, minus the size of an integer.
Rem Return Values   :   If the function succeeds, the return value is the requested value. If the function fails,
Rem                     the return value is zero.
Public Declare Function _
  GetWindowLongPtr Lib "user32" Alias _
  "GetWindowLongA" (ByVal hWnd As stdole.OLE_HANDLE, _
                   ByVal nIndex As WINDOW_OFFSETS) As Long

Rem The SetWindowLongPtr function changes an attribute of the specified window. The function also sets a value at the
Rem specified offset in the extra window memory.
Rem This function supersedes the SetWindowLong function. To write code that is compatible with both 32-bit and 64-bit
Rem versions of Windows, use SetWindowLongPtr.
Rem @hWnd           :   [in] Handle to the window and, indirectly, the class to which the window belongs. The SetWindowLongPtr
Rem                     function fails if the window specified by the hWnd parameter does not belong to the same process
Rem                     as the calling thread.
Rem @nIndex         :   [in] Specifies the zero-based offset to the value to be retrieved. Valid values are in the
Rem                     range zero through the number of bytes of extra window memory, minus the size of an integer.
Rem @dwNewLong      :   [in] Specifies the replacement value.
Rem Return Values   :   If the function succeeds, the return value is the requested value. If the function fails,
Rem                     the return value is zero.
Public Declare Function _
  SetWindowLongPtr Lib "user32" Alias _
  "SetWindowLongA" (ByVal hWnd As stdole.OLE_HANDLE, _
                   ByVal nIndex As WINDOW_OFFSETS, _
                   ByVal dwNewLong As Long) As Long

Rem The GetProp function retrieves a data handle from the property list of the specified window. The character string identifies
Rem the handle to be retrieved. The string and handle must have been added to the property list by a previous call to the SetProp
Rem function.
Rem @hWnd           :   [in] Handle to the window whose property list is to be searched.
Rem @lpString       :   [in] Pointer to a null-terminated character string or contains an atom that identifies a string. If this
Rem                     parameter is an atom, it must have been created by using the GlobalAddAtom function. The atom, a 16-bit
Rem                     value, must be placed in the low-order word of the lpString parameter; the high-order word must be zero.
Rem Return Values   :   If the property list contains the string, the return value is the associated data handle. Otherwise, the
Rem                     return value is NULL.
Public Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As stdole.OLE_HANDLE, _
                                                               ByVal lpString As String) As Long

Rem The SetProp function adds a new entry or changes an existing entry in the property list of the specified window. The function
Rem adds a new entry to the list if the specified character string does not exist already in the list. The new entry contains the
Rem string and the handle. Otherwise, the function replaces the string's current handle with the specified handle.
Rem @hWnd           :   [in] Handle to the window whose property list is to be searched.
Rem @lpString       :   [in] Pointer to a null-terminated character string or contains an atom that identifies a string. If this
Rem                     parameter is an atom, it must have been created by using the GlobalAddAtom function. The atom, a 16-bit
Rem                     value, must be placed in the low-order word of the lpString parameter; the high-order word must be zero.
Rem @hData          :   [in] Handle to the data to be copied to the property list. The data handle can identify any value useful
Rem                     to the application.
Rem Return Values   :   If the data handle and string are added to the property list, the return value is nonzero. If the function
Rem                     fails, the return value is zero.
Public Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hWnd As stdole.OLE_HANDLE, _
                                                               ByVal lpString As String, _
                                                               ByVal hData As Long) As Long

Rem The RemoveProp function removes an entry from the property list of the specified window. The specified character string identifies
Rem the entry to be removed.
Rem @hWnd           :   [in] Handle to the window whose property list is to be searched.
Rem @lpString       :   [in] Pointer to a null-terminated character string or contains an atom that identifies a string. If this
Rem                     parameter is an atom, it must have been created by using the GlobalAddAtom function. The atom, a 16-bit
Rem                     value, must be placed in the low-order word of the lpString parameter; the high-order word must be zero.
Public Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hWnd As stdole.OLE_HANDLE, _
                                                                     ByVal lpString As String) As Long

Rem The CallWindowProc function passes message information to the specified window procedure.
Rem @lpPrevWndFunc  :   [in] Pointer to the previous window procedure.
Rem                     If this value is obtained by calling the GetWindowLong function with the nIndex parameter set to GWL_WNDPROC
Rem                     or DWL_DLGPROC, it is actually either the address of a window or dialog box procedure, or a special internal
Rem                     value meaningful only to CallWindowProc.
Rem @hWnd           :   [in] Handle to the window procedure to receive the message.
Rem @Msg            :   [in] Specifies the message.
Rem @wParam         :   [in] Specifies additional message-specific information. The contents of this parameter depend on the value of
Rem                     the Msg parameter.
Rem @lParam         :   [in] Specifies additional message-specific information. The contents of this parameter depend on the value of
Rem                     the Msg parameter.
Rem Return Values   :   The return value specifies the result of the message processing and depends on the message sent.
Public Declare Function _
  CallWindowProc Lib "user32" Alias _
  "CallWindowProcA" (ByVal lpPrevWndFunc As stdole.OLE_HANDLE, _
                     ByVal hWnd As stdole.OLE_HANDLE, _
                     ByVal MSG As Long, _
                     ByVal wParam As Long, _
                     ByVal lParam As Long) As Long

Rem Converts an OLE_COLOR type to a COLORREF.
Rem @clr            :   [in] The OLE color to be converted into a COLORREF.
Rem @hpal           :   [in] Palette used as a basis for the conversion.
Rem @pcolorref      :   [out] Pointer to the caller's variable that receives the converted COLORREF result. This can be NULL, indicating
Rem                     that the caller wants only to verify that a converted color exists.
Rem Return Values   :   This function supports the standard return values E_INVALIDARG and E_UNEXPECTED, as well as the following:
Rem                     S_OK : The color was translated successfully.
Public Declare Function _
  OleTranslateColor Lib "oleaut32" (ByVal clr As stdole.OLE_COLOR, _
                                    ByVal hpal As stdole.OLE_HANDLE, _
                                    ByRef pcolorref As stdole.OLE_COLOR) As Long

Rem The CreateSolidBrush function creates a logical brush that has the specified solid color.
Rem @crColor        :   [in] Specifies the color of the brush. To create a COLORREF color value, use the RGB macro.
Rem Return Values   :   If the function succeeds, the return value identifies a logical brush.
Rem                     If the function fails, the return value is NULL.
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As stdole.OLE_COLOR) As stdole.OLE_HANDLE

Rem The FillRect function fills a rectangle by using the specified brush. This function includes the left and top borders, but excludes the
Rem right and bottom borders of the rectangle.
Rem @hDC            :   [in] Handle to the device context.
Rem @lprc           :   [in] Pointer to a RECT structure that contains the logical coordinates of the rectangle to be filled.
Rem @hbr            :   [in] Handle to the brush used to fill the rectangle.
Rem Return Value    :   If the function succeeds, the return value is nonzero. If the function fails, the return value is zero.
Public Declare Function FillRect Lib "user32" (ByVal hDC As stdole.OLE_HANDLE, _
                                               ByRef lprc As RECT, _
                                               ByVal hbr As stdole.OLE_HANDLE) As Long

Rem The DeleteObject function deletes a logical pen, brush, font, bitmap, region, or palette, freeing all system resources associated with the
Rem object. After the object is deleted, the specified handle is no longer valid.
Rem @hObject        :   [in] Handle to a logical pen, brush, font, bitmap, region, or palette.
Rem Return Values   :   If the function succeeds, the return value is nonzero. If the specified handle is not valid or is currently selected into
Rem                     a DC, the return value is zero.
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As stdole.OLE_HANDLE) As Long

Rem The FindWindowEx function retrieves a handle to a window whose class name and window name match the specified strings. The function searches
Rem child windows, beginning with the one following the specified child window. This function does not perform a case-sensitive search.
Rem @hWndParent     :   [in] Handle to the parent window whose child windows are to be searched.
Rem                     If hwndParent is NULL, the function uses the desktop window as the parent window. The function searches among windows that
Rem                     are child windows of the desktop.
Rem                     Windows 2000/XP: If hwndParent is HWND_MESSAGE, the function searches all message-only windows.
Rem @hWndChildAfter :   [in] Handle to a child window. The search begins with the next child window in the Z order. The child window must be a direct
Rem                     child window of hwndParent, not just a descendant window.
Rem                     If hwndChildAfter is NULL, the search begins with the first child window of hwndParent.
Rem                     Note that if both hwndParent and hwndChildAfter are NULL, the function searches all top-level and message-only windows.
Rem @lpszClass      :   [in] Pointer to a null-terminated string that specifies the class name or a class atom created by a previous call to the
Rem                     RegisterClass or RegisterClassEx function. The atom must be placed in the low-order word of lpszClass; the high-order word
Rem                     must be zero.
Rem                     If lpszClass is a string, it specifies the window class name. The class name can be any name registered with RegisterClass
Rem                     or RegisterClassEx, or any of the predefined control-class names.
Rem @lpszWindow     :   [in] Pointer to a null-terminated string that specifies the window name (the window's title). If this parameter is NULL, all
Rem                     window names match.
Rem Return Value    :   If the function succeeds, the return value is a handle to the window that has the specified class and window names
Public Declare _
  Function FindWindow Lib "user32" _
    Alias "FindWindowExA" (ByVal hWndParent As stdole.OLE_HANDLE, _
                           ByVal hWndChildAfter As stdole.OLE_HANDLE, _
                           ByVal lpszClass As String, _
                           ByVal lpszWindow As Long) As Long

Rem The RemoveMenu function deletes a menu item or detaches a submenu from the specified menu. If the menu item opens a drop-down menu or submenu,
Rem RemoveMenu does not destroy the menu or its handle, allowing the menu to be reused. Before this function is called, the GetSubMenu function
Rem should retrieve a handle to the drop-down menu or submenu
Rem @hMenu          :   [in] Handle to the menu to be changed.
Rem @uPosition      :   [in] Specifies the menu item to be deleted, as determined by the uFlags parameter.
Rem @uFlags         :   [in] Specifies how the uPosition parameter is interpreted.
Rem @Return Values  :   If the function succeeds, the return value is nonzero. If the function fails, the return value is zero.
Rem @Note           :   The application must call the DrawMenuBar function whenever a menu changes, whether or not the menu is in a displayed window.
Public Declare Function RemoveMenu Lib "user32" (ByVal hMenu As stdole.OLE_HANDLE, _
                                                 ByVal uPosition As Long, _
                                                 ByVal uFlags As SETMENUITEMBITMAPFLAGS) As Long

Rem The DrawMenuBar function redraws the menu bar of the specified window. If the menu bar changes after the system has created the window, this function
Rem must be called to draw the changed menu bar
Rem @hMenu          :   [in] Handle to the window whose menu bar needs redrawing.
Rem @Return Values  :   If the function succeeds, the return value is nonzero. If the function fails, the return value is zero.
Public Declare Function DrawMenuBar Lib "user32" (ByVal hMenu As stdole.OLE_HANDLE) As Long

Rem The EnableMenuItem function enables, disables, or grays the specified menu item.
Rem @hMenu          :   [in] Handle to the menu.
Rem @uIDEnableItem  :   [in] Specifies the menu item to be enabled, disabled, or grayed, as determined by the uEnable parameter. This parameter specifies
Rem                     an item in a menu bar, menu, or submenu.
Rem @uEnable        :   [in] Controls the interpretation of the uIDEnableItem parameter and indicate whether the menu item is enabled, disabled, or grayed.
Rem                     This parameter must be a combination of
Rem                     either MF_BYCOMMAND or MF_BYPOSITION
Rem                     and MF_ENABLED, MF_DISABLED, or MF_GRAYED.
Rem @Return Values  :   The return value specifies the previous state of the menu item (it is either MF_DISABLED, MF_ENABLED, or MF_GRAYED).
Rem                     If the menu item does not exist, the return value is -1
Public Declare Function EnableMenuItem Lib "user32" (ByVal hMenu As stdole.OLE_HANDLE, _
                                                      ByVal uIDEnableItem As Long, _
                                                      ByVal uEnable As SETMENUITEMBITMAPFLAGS) As Long

Rem The ShowWindow function sets the specified window's show state.
Rem @hWnd           :   [in] Handle to the window.
Rem @nCmdShow       :   [in] Specifies how the window is to be shown.
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As stdole.OLE_HANDLE, _
                                                 ByVal nCmdShow As SHOW_WINDOW_MSGS) As Long

Rem The AppendMenu function appends a new item to the end of the specified menu bar, drop-down menu, submenu, or shortcut menu. You can use this function to
Rem specify the content, appearance, and behavior of the menu item.
Rem @hMenu          :   [in] Handle to the menu bar, drop-down menu, submenu, or shortcut menu to be changed.
Rem @uFlags         :   [in] Specifies flags to control the appearance and behavior of the new menu item. This parameter can be a combination of the values.
Rem @uIDNewItem     :   [in] Specifies either the identifier of the new menu item or, if the uFlags parameter is set to MF_POPUP, a handle to the drop-down menu or submenu.
Rem @lpNewItem      :   [in] Specifies the content of the new menu item. The interpretation of lpNewItem depends on whether the uFlags parameter includes the MF_BITMAP,
Rem                     MF_OWNERDRAW, or MF_STRING flag
Rem @Return Values  :   If the function succeeds, the return value is nonzero. If the function fails, the return value is zero.
Public Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As stdole.OLE_HANDLE, _
                                                                     ByVal uFlags As Long, _
                                                                     ByVal uIDNewItem As Long, _
                                                                     ByVal lpNewItem As Long) As Long

Rem The mciSendString function sends a command string to an MCI device. The device that the command is sent to is specified in the command string.
Rem @lpszCommand    :   Pointer to a null-terminated string that specifies an MCI command string.
Rem @lpszReturnString:  Pointer to a buffer that receives return information. If no return information is needed, this parameter can be NULL.
Rem @cchReturn      :   Size, in characters, of the return buffer specified by the lpszReturnString parameter.
Rem @hwndCallback   :   Handle to a callback window if the "notify" flag was specified in the command string.
Rem @Return Values  :   Returns zero if successful or an error otherwise. The low-order word of the returned DWORD value contains the error return value. If the error is
Rem                     device-specific, the high-order word of the return value is the driver identifier; otherwise, the high-order word is zero.
Public Declare Function mciSendString Lib "Winmm.dll" (ByVal lpszCommand As String, _
                                                       ByVal lpszReturnString As String, _
                                                       ByVal cchReturn As Long, _
                                                       ByVal hwndCallback As stdole.OLE_HANDLE) As Long

Rem The mciSendString function sends a command string to an MCI device. The device that the command is sent to is specified in the command string.
Rem @lpszCommand    :   Pointer to a null-terminated string that specifies an MCI command string.
Rem @lpszReturnString:  Pointer to a buffer that receives return information. If no return information is needed, this parameter can be NULL.
Rem @cchReturn      :   Size, in characters, of the return buffer specified by the lpszReturnString parameter.
Rem @hwndCallback   :   Handle to a callback window if the "notify" flag was specified in the command string.
Rem @Return Values  :   Returns zero if successful or an error otherwise. The low-order word of the returned DWORD value contains the error return value. If the error is
Rem                     device-specific, the high-order word of the return value is the driver identifier; otherwise, the high-order word is zero.
Rem @Note           :   Implemented as Unicode and ANSI versions on Windows NT/2000/XP.
Public Declare Function mciSendStringNT Lib "Winmm.dll" _
         Alias "mciSendStringA" (ByVal lpszCommand As String, _
                                 ByVal lpszReturnString As String, _
                                 ByVal cchReturn As Long, _
                                 ByVal hwndCallback As stdole.OLE_HANDLE) As Long

Rem This function retrieves the current color of the specified display element. Display elements are the
Rem parts of a window and the Windows display that appear on the system display screen.
Rem @nIndex         :   Specifies the display element whose color is to be retrieved.
Rem Return Values   :   The red, green, blue (RGB) color value that determines the color of the
Rem                     specified element indicates success. Zero indicates failure.
Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As COLOR_INDEX) As stdole.OLE_COLOR

Rem This function retrieves the dimensions?widths and heights?of Windows display elements and system
Rem configuration settings. All dimensions retrieved by GetSystemMetrics are in pixels.
Rem @nIndex         :   Specifies the system metric or configuration setting to retrieve.
Rem                     All SM_CX* values are widths. All SM_CY* values are heights.
Rem Return Values   :
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As SYSTEM_METRICS_ITEM) As Long

Rem This function fills the specified buffer with the metrics for the currently selected font.
Rem @hDC            :   [in] Handle to the device context (DC).
Rem @lptm           :   [out] Long pointer to the TEXTMETRIC structure that is to receive the metrics.
Rem Return Values   :   Nonzero indicates success. Zero indicates failure.
Public Declare Function GetTextMetrics Lib "gdi32" _
                  Alias "GetTextMetricsA" (ByVal hDC As stdole.OLE_HANDLE, _
                                           ByRef lpMetrics As TEXTMETRIC) As Long

Rem This function retrieves the typeface name of the font that is selected into the specified device
Rem context (DC).
Rem @hDC            :   [in] Handle to the device context (DC).
Rem @nCount         :   [in] Specifies the size, in characters, of the buffer.
Rem @lpFacename     :   [out] Long pointer to the buffer that is to receive the typeface name. If this
Rem                     parameter is NULL, the function returns the number of characters in the name,
Rem                     including the terminating null character.
Rem Return Values   :   The number of characters copied to the buffer indicates success. Zero indicates failure.
Public Declare Function GetTextFace Lib "gdi32" _
                  Alias "GetTextFaceA" (ByVal hDC As stdole.OLE_HANDLE, _
                                        ByVal nCount As Long, _
                                        ByVal lpFacename As String) As Long

Rem This function selects an object into a specified device context. The new object replaces the
Rem previous object of the same type
Rem @hDC            :   [in] Handle to the device context.
Rem @hObject        :   [in] Handle to the object to be selected.
Rem Return Values   :   [Look Enum Return Type]
Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As stdole.OLE_HANDLE, _
                                                  ByVal hObject As stdole.OLE_HANDLE) As SELECT_OBJECT_RETURN_VALUE

Rem This function retrieves the current text color for the specified device context.
Rem @hDC            :   [in] Handle to the device context.
Rem Return Values   :   The current text color as a COLORREF value indicates success.
Rem                     CLR_INVALID indicates failure.
Public Declare Function GetTextColor Lib "gdi32" (ByVal hDC As stdole.OLE_HANDLE) As stdole.OLE_COLOR

Rem This function draws a character string by using the currently selected font. An optional rectangle
Rem may be provided, to be used for clipping, opaquing, or both.
Rem @hDC            :   [in] Handle to the device context (DC).
Rem @X              :   [in] Specifies the logical x-coordinate of the reference point used to position the string.
Rem @Y              :   [in] Specifies the logical y-coordinate of the reference point used to position the string.
Rem @fuOptions      :   [in] Specifies how to use the application-defined rectangle.
Rem @lprc           :   [in] Long pointer to an optional RECT structure that specifies the dimensions of
Rem                     a rectangle that is used for clipping, opaquing, or both.
Rem @lpString       :   [in] Long pointer to the character string to be drawn. The string does not need
Rem                     to be zero-terminated, since cbCount specifies the length of the string.
Rem @cbCount        :   [in] Specifies the number of characters in the string.
Rem @lpDx           :   [in] Long pointer to an optional array of values that indicate the distance
Rem                     between origins of adjacent character cells. For example, lpDx[i] logical units
Rem                     separate the origins of character cell i and character cell i + 1.
Rem Return Values   :   Nonzero indicates that the string is drawn. Zero indicates failure.
Public Declare Function ExtTextOut Lib "gdi32" (ByVal hDC As stdole.OLE_HANDLE, _
                                                ByVal x As stdole.OLE_XPOS_PIXELS, _
                                                ByVal y As stdole.OLE_YPOS_PIXELS, _
                                                ByVal fuOptions As TEXTOUT_ALIGN_OPTION, _
                                                ByRef lprc As RECT, _
                                                ByRef lpString As String, _
                                                ByVal cbCount As Long, _
                                                ByRef lpDx As Integer) As Long

Rem This function retrieves information about the capabilities of a specified device.
Rem @hDC            :   [in] Handle to the device context.
Rem @nIndex         :   [in] Specifies the item to return.
Rem Return Values   :   Returns the value of the desired item.
Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As stdole.OLE_HANDLE, ByVal nIndex As DEVICE_CAPABILITY) As Long

Rem The MulDiv function multiplies two 32-bit values and then divides the 64-bit result by a third
Rem 32-bit value. The return value is rounded up or down to the nearest integer.
Rem @nNumber        :   [in] Specifies the multiplicand.
Rem @nNumerator     :   [in] Specifies the multiplier.
Rem @nDenominator   :   [in] Specifies the number by which the result of the multiplication
Rem                     (nNumber * nNumerator) is to be divided.
Rem Return Values   :   If the function succeeds, the return value is the result of the multiplication
Rem                     and division. If either an overflow occurred or nDenominator was 0, the return
Rem                     value is -1.
Public Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Integer, _
                                               ByVal nNumerator As Integer, _
                                               ByVal nDenominator As Integer) As Integer

Rem The GetDC function retrieves a handle to a display device context (DC) for the client area of a
Rem specified window or for the entire screen. You can use the returned handle in subsequent GDI
Rem functions to draw in the DC.
Rem @hWnd           :   [in] Handle to the window whose DC is to be retrieved. If this value is NULL,
Rem                     GetDC retrieves the DC for the entire screen.
Rem Return Values   :   If the function succeeds, the return value is a handle to the DC for the
Rem                     specified window's client area. If the function fails, the return value is NULL.
Public Declare Function GetDC Lib "user32" (ByVal hWnd As stdole.OLE_HANDLE) As stdole.OLE_HANDLE

Rem The GetWindowDC function retrieves the device context (DC) for the entire window, including title
Rem bar, menus, and scroll bars. A window device context permits painting anywhere in a window, because
Rem the origin of the device context is the upper-left corner of the window instead of the client area.
Rem @hWnd           :   [in] Handle to the window whose DC is to be retrieved. If this value is NULL,
Rem                     GetWindowDC retrieves the DC for the entire screen.
Rem Return Values   :   If the function succeeds, the return value is a handle to a device context for
Rem                     the specified window. If the function fails, the return value is NULL,
Rem                     indicating an error or an invalid hWnd parameter.
Public Declare Function GetWindowDC Lib "user32" (ByVal hWnd As stdole.OLE_HANDLE) As stdole.OLE_HANDLE

Rem The ReleaseDC function releases a device context (DC), freeing it for use by other applications.
Rem The effect of the ReleaseDC function depends on the type of DC. It frees only common and window DCs.
Rem It has no effect on class or private DCs.
Rem @hWnd           :   [in] Handle to the window whose DC is to be released.
Rem @hDC            :   [in] Handle to the DC to be released.
Rem Return Values   :   The return value indicates whether the DC was released. If the DC was released,
Rem                     the return value is 1. If the DC was not released, the return value is zero.
Rem Remarks         :   The application must call the ReleaseDC function for each call to the GetWindowDC
Rem                     function and for each call to the GetDC function that retrieves a common DC.
Public Declare Function ReleaseDC Lib "user32" (ByVal hWnd As stdole.OLE_HANDLE, ByVal hDC As stdole.OLE_HANDLE) As Long

Rem The CreateFontIndirect function creates a logical font that has the specified characteristics.
Rem The font can subsequently be selected as the current font for any device context.
Rem @lpLogFont      :   [in] Pointer to a LOGFONT structure that defines the characteristics of the logical font.
Rem Return Values   :   If the function succeeds, the return value is a handle to a logical font.
Rem                     If the function fails, the return value is NULL.
Public Declare Function CreateFontIndirect Lib "gdi32" _
                  Alias "CreateFontIndirectA" (ByRef lpLogFont As LOGFONT) As stdole.OLE_HANDLE

Rem The DrawTextEx function draws formatted text in the specified rectangle.
Rem @hDC            :   [in] Handle to the device context in which to draw.
Rem @lpStr          :   [in/out] Pointer to the string that contains the text to draw. The string must
Rem                     be null-terminated if the cchText parameter is -1. If dwDTFormat includes
Rem                     DT_MODIFYSTRING, the function could add up to four additional characters to
Rem                     this string. The buffer containing the string should be large enough to
Rem                     accommodate these extra characters.
Rem @nCount         :   [in] Specifies the length of the string specified by the lpchText parameter.
Rem                     For the ANSI function it is a BYTE count and for the Unicode function it is a
Rem                     WORD count.
Rem @lpRect         :   [in/out] Pointer to a RECT structure that contains the rectangle, in logical
Rem                     coordinates, in which the text is to be formatted.
Rem @wFormat        :   [in] Specifies formatting options.
Rem Return Values   :   If the function succeeds, the return value is the text height in logical units.
Rem                     If DT_VCENTER or DT_BOTTOM is specified, the return value is the offset from
Rem                     lprc->top to the bottom of the drawn text. If the function fails, the return
Rem                     value is zero.
Public Declare Function DrawText Lib "user32" _
                  Alias "DrawTextA" (ByVal hDC As stdole.OLE_HANDLE, _
                                     ByVal lpStr As String, _
                                     ByVal nCount As Long, _
                                     ByRef lpRect As RECT, _
                                     ByVal wFormat As DRAWTEXT_OPTION) As Long

Rem The DrawIconEx function draws an icon or cursor into the specified device context, performing the
Rem specified raster operations, and stretching or compressing the icon or cursor as specified.
Rem @hDC            :   [in] Handle to the device context into which the icon or cursor will be drawn.
Rem @xLeft          :   [in] Specifies the logical x-coordinate of the upper-left corner of the icon
Rem                     or cursor.
Rem @yTop           :   [in] Specifies the logical y-coordinate of the upper-left corner of the icon
Rem                     or cursor.
Rem @hIcon          :   [in] Handle to the icon or cursor to be drawn. This parameter can identify
Rem                     an animated cursor.
Rem @cxWidth        :   [in] Specifies the logical width of the icon or cursor.
Rem @cyWidth        :   [in] Specifies the logical height of the icon or cursor.
Rem @istepIfAniCur  :   [in] Specifies the index of the frame to draw, if hIcon identifies an animated
Rem                     cursor.
Rem @hbrFlickerFreeDraw: [in] Handle to a brush that the system uses for flicker-free drawing.
Rem @diFlags        :   [in] Specifies the drawing flags.
Rem @Return Values  :   If the function succeeds, the return value is nonzero. If the function fails,
Rem                     the return value is zero
Public Declare Function DrawIconEx Lib "user32" (ByVal hDC As stdole.OLE_HANDLE, _
                                                 ByVal xLeft As stdole.OLE_XPOS_PIXELS, _
                                                 ByVal yTop As stdole.OLE_YPOS_PIXELS, _
                                                 ByVal hIcon As stdole.OLE_HANDLE, _
                                                 ByVal cxWidth As stdole.OLE_XSIZE_PIXELS, _
                                                 ByVal cyWidth As stdole.OLE_YSIZE_PIXELS, _
                                                 ByVal istepIfAniCur As Long, _
                                                 ByVal hbrFlickerFreeDraw As stdole.OLE_HANDLE, _
                                                 ByVal diFlags As DRAWICON_FLAG) As Long

Rem This function sets the pen capture to the specified window belonging to the current thread.
Rem After a window has captured the pen, all pen input is directed to that window. Only one
Rem window at a time can capture the pen.
Rem @hWnd           :   [in] Handle to the window in the current thread that is to capture the mouse.
Rem Return Values   :   The handle of the window that had previously captured the mouse indicates
Rem                     success. NULL indicates that there is no such window.
Public Declare Function SetCapture Lib "user32" (ByVal hWnd As stdole.OLE_HANDLE) As stdole.OLE_HANDLE

Rem The ReleaseCapture function releases the mouse capture from a window in the current thread and
Rem restores normal mouse input processing. A window that has captured the mouse receives all mouse
Rem input, regardless of the position of the cursor, except when a mouse button is clicked while the
Rem cursor hot spot is in the window of another thread.
Rem Return Values   : If the function succeeds, the return value is nonzero. If the function fails,
Rem                   the return value is zero.
Public Declare Function ReleaseCapture Lib "user32" () As Long

Rem The TextOut function writes a character string at the specified location, using the currently
Rem selected font, background color, and text color.
Rem @hdc        :   [in] Handle to the device context.
Rem @x          :   [in] Specifies the x-coordinate, in logical coordinates, of the reference point
Rem                 that the system uses to align the string.
Rem @y          :   [in] Specifies the y-coordinate, in logical coordinates, of the reference point
Rem                 that the system uses to align the string.
Rem @lpString   :   [in] Pointer to the string to be drawn. The string does not need to be zero-terminated,
Rem                 since cbString specifies the length of the string.
Rem @nCount     :   [in] Specifies the length of the string. For the ANSI function it is a BYTE count
Rem                 and for the Unicode function it is a WORD count. Note that for the ANSI function,
Rem                 characters in SBCS code pages take one byte each while most characters in DBCS code
Rem                 pages take two bytes; for the Unicode function, most currently defined Unicode characters
Rem                 (those in the Basic Multilingual Plane (BMP)) are one WORD while Unicode surrogates are
Rem                 two WORDs.
Rem                 Windows 95/98/Me: This value may not exceed 8192.
Rem Return Values:  If the function succeeds, the return value is nonzero. If the function fails, the return
Rem                 value is zero.
Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDC As stdole.OLE_HANDLE, _
                                                               ByVal x As stdole.OLE_XPOS_PIXELS, _
                                                               ByVal y As stdole.OLE_YPOS_PIXELS, _
                                                               ByVal lpString As String, _
                                                               ByVal nCount As Long) As Long

Rem The DrawEdge function draws one or more edges of rectangle.
Rem @hdc            :   [in] Handle to the device context.
Rem @qrc            :   [in/out] Pointer to a RECT structure that contains the logical coordinates of the
Rem                     rectangle.
Rem @edge           :   [in] Specifies the type of inner and outer edges to draw. This parameter must be a
Rem                     combination of one inner-border flag and one outer-border flag. The inner-border
Rem                     flags are as follows.
Rem                     The outer-border flags are as follows.
Rem                      Alternatively, the edge parameter can specify one of the following flags.
Rem @grfFlags       :   [in] Specifies the type of border. This parameter can be a combination of the following
Rem                     values.
Rem Return Values   :   If the function succeeds, the return value is nonzero. If the function fails,
Rem                     the return value is zero.
Public Declare Function DrawEdge Lib "user32" (ByVal hDC As stdole.OLE_HANDLE, _
                                                ByRef qrc As RECT, _
                                                ByVal edge As EDGE_FLAGS, _
                                                ByVal grfFlags As BORDER_TYPE_FLAGS) As Long

Rem The GetTextExtentPoint32 function computes the width and height of the specified string of text.
Rem @hdc            :   [in] Handle to the device context.
Rem @lpsz           :   [in] Pointer to a buffer that specifies the text string. The string does not
Rem                     need to be zero-terminated, because the cbString parameter specifies the length
Rem                     of the string.
Rem @cbString       :   [in] Specifies the length of the lpString buffer. For the ANSI function it is
Rem                     a BYTE count and for the Unicode function it is a WORD count. Note that for the
Rem                     ANSI function, characters in SBCS code pages take one byte each, while most characters
Rem                     in DBCS code pages take two bytes; for the Unicode function, most currently defined
Rem                     Unicode characters (those in the Basic Multilingual Plane (BMP)) are one WORD while
Rem                     Unicode surrogates are two WORDs.
Rem                     Windows 95/98/Me: This value may not exceed 8192.
Rem @lpSize         :   [out] Pointer to a SIZE structure that receives the dimensions of the string, in logical units.
Rem Return Values   :   If the function succeeds, the return value is nonzero. If the function fails, the return
Rem                     value is zero.
Public Declare Function GetTextExtentPoint32 Lib "gdi32" _
                  Alias "GetTextExtentPoint32A" (ByVal hDC As stdole.OLE_HANDLE, _
                                                 ByVal lpsz As String, _
                                                 ByVal cbString As Long, _
                                                 ByRef lpSize As Size) As Long

Rem This function creates a memory device context (DC) compatible with the specified device.
Rem @hDC            :   [in] Handle to an existing device context. If this handle is NULL, the function
Rem                     creates a memory device context compatible with the application's current screen.
Rem Return Values   :   The handle to a memory device context indicates success. NULL indicates failure.
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As stdole.OLE_HANDLE) As stdole.OLE_HANDLE

Rem This function creates a bitmap compatible with the device associated with the specified device.
Rem context.
Rem @hDC            :   [in] Handle to a device context.
Rem @nWidth         :   [in] Specifies the bitmap width, in pixels.
Rem @nHeight        :   [in] Specifies the bitmap height, in pixels.
Rem Return Values   :   A handle to the bitmap indicates success. NULL indicates failure.
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As stdole.OLE_HANDLE, _
                                                            ByVal nWidth As stdole.OLE_XSIZE_PIXELS, _
                                                            ByVal nHeight As stdole.OLE_YSIZE_PIXELS) As stdole.OLE_HANDLE

Rem This function paints the given rectangle using the brush that is currently selected into the
Rem specified device context. The brush pixels and the surface pixels are combined according to the
Rem specified raster operation.
Rem @hDC            :   [in] Handle to the device context.
Rem @x              :   [in] Specifies the x-coordinate, in logical units, of the upper-left corner
Rem                     of the rectangle to be filled.
Rem @y              :   [in] Specifies the y-coordinate, in logical units, of the upper-left corner
Rem                     of the rectangle to be filled.
Rem @nWidth         :   [in] Specifies the width, in logical units, of the rectangle.
Rem @nHeight        :   Specifies the height, in logical units, of the rectangle.
Rem @dwRop          :   [in] Specifies the raster operation code.
Rem Return Values   :
Public Declare Function PatBlt Lib "gdi32" (ByVal hDC As stdole.OLE_HANDLE, _
                                            ByVal x As stdole.OLE_XPOS_PIXELS, _
                                            ByVal y As stdole.OLE_YPOS_PIXELS, _
                                            ByVal nWidth As stdole.OLE_XSIZE_PIXELS, _
                                            ByVal nHeight As stdole.OLE_YSIZE_PIXELS, _
                                            ByVal dwRop As RASTER_OPERATION) As Long

Rem This function transfers pixels from a specified source rectangle to a specified destination
Rem rectangle, altering the pixels according to the selected raster operation (ROP) code.
Rem @hDestDC        :   [in] Handle to the destination device context.
Rem @x              :   [in] Specifies the logical x-coordinate of the upper-left corner of the
Rem                     destination rectangle.
Rem @y              :   [in] Specifies the logical y-coordinate of the upper-left corner of the
Rem                     destination rectangle.
Rem @nWidth         :   [in] Specifies the logical width of the source and destination rectangles.
Rem @nHeight        :   [in] Specifies the logical height of the source and the destination rectangles.
Rem @hSrcDC         :   [in] Handle to the source device context.
Rem @xSrc           :   [in] Specifies the logical x-coordinate of the upper-left corner of the
Rem                     source rectangle.
Rem @ySrc           :   [in] Specifies the logical y-coordinate of the upper-left corner of the
Rem                     source rectangle.
Rem @dwRop          :   [in] Specifies a raster-operation code. These codes define how the color data
Rem                     for the source rectangle is to be combined with the color data for the
Rem                     destination rectangle to achieve the final color.
Rem Return Values   :   Nonzero indicates success. Zero indicates failure.
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As stdole.OLE_HANDLE, _
                                            ByVal x As stdole.OLE_XPOS_PIXELS, _
                                            ByVal y As stdole.OLE_YPOS_PIXELS, _
                                            ByVal nWidth As stdole.OLE_XSIZE_PIXELS, _
                                            ByVal nHeight As stdole.OLE_YSIZE_PIXELS, _
                                            ByVal hSrcDC As stdole.OLE_HANDLE, _
                                            ByVal xSrc As stdole.OLE_XPOS_PIXELS, _
                                            ByVal ySrc As stdole.OLE_YPOS_PIXELS, _
                                            ByVal dwRop As RASTER_OPERATION) As Long

Rem The SetBkColor function sets the current background color to the specified color value, or to the
Rem nearest physical color if the device cannot represent the specified color value.
Rem @hDC            :   [in] Handle to the device context.
Rem @crColor        :   [in] Specifies the new background color. To make a COLORREF value, use the
Rem                     RGB macro.
Rem Return Values   :   If the function succeeds, the return value specifies the previous background
Rem                     color as a COLORREF value. If the function fails, the return value is
Rem                     CLR_INVALID.
Public Declare Function SetBkColor Lib "gdi32" (ByVal hDC As stdole.OLE_HANDLE, _
                                                ByVal crColor As stdole.OLE_COLOR) As stdole.OLE_COLOR

Rem The DrawState function displays an image and applies a visual effect to indicate a state, such as a
Rem disabled or default state.
Rem @hDC            :   [in] Handle to the device context to draw in.
Rem @hBrush         :   [in] Handle to the brush used to draw the image, if the state specified by the
Rem                     fuFlags parameter is DSS_MONO. This parameter is ignored for other states.
Rem @lpDrawStateProc:   [in] Pointer to an application-defined callback function used to render the
Rem                     image. This parameter is required if the image type in fuFlags is DST_COMPLEX.
Rem                     It is optional and can be NULL if the image type is DST_TEXT. For all other
Rem                     image types, this parameter is ignored.
Rem @lData          :   [in] Specifies information about the image. The meaning of this parameter depends
Rem                     on the image type.
Rem @wData          :   [in] Specifies information about the image. The meaning of this parameter depends
Rem                     on the image type.
Rem @x              :   [in] Specifies the horizontal location at which to draw the image.
Rem @y              :   [in] Specifies the vertical location at which to draw the image.
Rem @cx             :   [in] Specifies the width of the image, in device units. This parameter is
Rem                     required if the image type is DST_COMPLEX. Otherwise, it can be zero to calculate
Rem                     the width of the image.
Rem @cy             :   [in] Specifies the height of the image, in device units. This parameter is
Rem                     required if the image type is DST_COMPLEX. Otherwise, it can be zero to calculate
Rem                     the height of the image.
Rem @fuFlags        :   [in] Specifies the image type and state.
Public Declare Function DrawState Lib "user32" _
                  Alias "DrawStateA" (ByVal hDC As stdole.OLE_HANDLE, _
                                      ByVal hBrush As stdole.OLE_HANDLE, _
                                      ByVal lpDrawStateProc As stdole.OLE_HANDLE, _
                                      ByVal lData As Long, _
                                      ByVal wData As Long, _
                                      ByVal x As stdole.OLE_XPOS_PIXELS, _
                                      ByVal y As stdole.OLE_YPOS_PIXELS, _
                                      ByVal cx As stdole.OLE_XSIZE_PIXELS, _
                                      ByVal cy As stdole.OLE_YSIZE_PIXELS, _
                                      ByVal fuFlags As DRAWSTATE_OPTION) As Long

Rem The SetBkMode function sets the background mix mode of the specified device context. The background
Rem mix mode is used with text, hatched brushes, and pen styles that are not solid lines.
Rem @hDC            :   [in] Handle to the device context.
Rem @nBkMode        :   [in] Specifies the background mode.
Rem Return Values   :   If the function succeeds, the return value specifies the previous background
Rem                     mode. If the function fails, the return value is zero.
Public Declare Function SetBkMode Lib "gdi32" (ByVal hDC As stdole.OLE_HANDLE, ByVal nBkMode As BACKMODE) As BACKMODE

Rem This function sets the text color of the specified device context to the specified color.
Rem @hDC            :   [in] Handle to the device context.
Rem @crColor        :   [in] Specifies the color of the text.
Rem Return Values   :   A color reference for the previous text color indicates success.
Rem                     CLR_INVALID indicates failure.
Public Declare Function SetTextColor Lib "gdi32" (ByVal hDC As stdole.OLE_HANDLE, ByVal crColor As stdole.OLE_COLOR) As stdole.OLE_COLOR

Rem The GrayString function draws gray text at the specified location. The function draws the text by
Rem copying it into a memory bitmap, graying the bitmap, and then copying the bitmap to the screen.
Rem The function grays the text regardless of the selected brush and background. GrayString uses the
Rem font currently selected for the specified device context.
Rem @hDC            :   [in] Handle to the device context.
Rem @hBrush         :   [in] Handle to the brush to be used for graying. If this parameter is NULL,
Rem                     the text is grayed with the same brush that was used to draw window text.
Rem @lpOutputFunc   :   [in] Pointer to the application-defined function that will draw the string,
Rem                     or, if TextOut is to be used to draw the string, it is a NULL pointer.
Rem                     For details, see the OutputProc callback function.
Rem @lpData         :   [in] Specifies a pointer to data to be passed to the output function.
Rem                     If the lpOutputFunc parameter is NULL, lpData must be a pointer to the string
Rem                     to be output.
Rem @nCount         :   [in] Specifies the number of characters to be output. If the nCount parameter
Rem                     is zero, GrayString calculates the length of the string (assuming lpData is a
Rem                     pointer to the string). If nCount is -1 and the function pointed to by
Rem                     lpOutputFunc returns FALSE, the image is shown but not grayed.
Rem @X              :   [in] Specifies the device x-coordinate of the starting position of the
Rem                     rectangle that encloses the string.
Rem @Y              :   [in] Specifies the device y-coordinate of the starting position of the rectangle
Rem                     that encloses the string.
Rem @nWidth         :   [in] Specifies the width, in device units, of the rectangle that encloses the
Rem                     string. If this parameter is zero, GrayString calculates the width of the area,
Rem                     assuming lpData is a pointer to the string.
Rem @nHeight        :   [in] Specifies the height, in device units, of the rectangle that encloses the
Rem                     string. If this parameter is zero, GrayString calculates the height of the area,
Rem                     assuming lpData is a pointer to the string.
Rem Return Values   :   If the string is drawn, the return value is nonzero. If either the TextOut
Rem                     function or the application-defined output function returned zero, or there
Rem                     was insufficient memory to create a memory bitmap for graying, the return value
Rem                     is zero.
Public Declare Function GrayString Lib "user32" _
                  Alias "GrayStringA" (ByVal hDC As stdole.OLE_HANDLE, _
                                       ByVal hBrush As stdole.OLE_HANDLE, _
                                       ByVal lpOutputFunc As stdole.OLE_HANDLE, _
                                       ByVal lpData As Long, _
                                       ByVal nCount As Long, _
                                       ByVal x As stdole.OLE_XPOS_PIXELS, _
                                       ByVal y As stdole.OLE_YPOS_PIXELS, _
                                       ByVal nWidth As stdole.OLE_XSIZE_PIXELS, _
                                       ByVal nHeight As stdole.OLE_YSIZE_PIXELS) As Long

Rem This function deletes the specified device context (DC).
Rem @hdc            :   [in] Handle to the device context.
Rem Return Values   :   Nonzero indicates success. Zero indicates failure.
Public Declare Function DeleteDC Lib "gdi32" (ByVal hDC As stdole.OLE_HANDLE) As Long

Rem The LoadImage function loads an icon, cursor, animated cursor, or bitmap.
Rem @hInst          :   [in] Handle to an instance of the module that contains the image to be loaded.
Rem                     To load an OEM image, set this parameter to zero.
Rem @lpszName       :   [in] Specifies the image to load.
Rem @uType          :   [in] Specifies the type of image to be loaded.
Rem @cxDesired      :   [in] Specifies the width, in pixels, of the icon or cursor.If this parameter
Rem                     is zero and LR_DEFAULTSIZE is not used, the function uses the actual resource
Rem                     width.
Rem @cyDesired      :   [in] Specifies the height, in pixels, of the icon or cursor. If this parameter
Rem                     is zero and LR_DEFAULTSIZE is not used, the function uses the actual resource
Rem                     height.
Rem @fuLoad         :   [in] This parameter can be one or more of the following values. See Enum.
Rem Return Values   :   Returns a handle to the loaded image, 0 otherwise.
Public Declare Function LoadImage Lib "user32" _
                  Alias "LoadImageA" (ByVal hinst As stdole.OLE_HANDLE, _
                                      ByVal lpszName As WIN32DEFINEDBMP, _
                                      ByVal uType As LOADIMGTYP, _
                                      ByVal cxDesired As stdole.OLE_XSIZE_PIXELS, _
                                      ByVal cyDesired As stdole.OLE_YSIZE_PIXELS, _
                                      ByVal fuLoad As LOADIMGPARAM) As Long

Rem This function creates a bitmap with the specified width, height, and bit depth.
Rem @nWidth         :   [in] Specifies the bitmap width, in pixels.
Rem @nHeight        :   [in] Specifies the bitmap height, in pixels.
Rem @nPlanes        :   [in] Specifies the number of color planes used by the device. The value of this
Rem                     parameter must be 1.
Rem @nBitCount      :   [in] Specifies the number of bits required to identify the color of a single pixel.
Rem @lpBits         :   [in] Long void pointer to an array of color data used to set the colors in a
Rem                     rectangle of pixels. Each scan line in the rectangle must be word aligned
Rem                     (scan lines that are not word aligned must be padded with zeros). If this
Rem                     parameter is NULL, the new bitmap is undefined.
Rem Return Values   :   A handle to a bitmap indicates success. NULL indicates failure.
Public Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As stdole.OLE_XSIZE_PIXELS, _
                                                  ByVal nHeight As stdole.OLE_YSIZE_PIXELS, _
                                                  ByVal nPlanes As Long, _
                                                  ByVal nBitCount As Long, _
                                                  ByRef lpBits As Any) As stdole.OLE_HANDLE

Rem The CreatePatternBrush function creates a logical brush with the specified bitmap pattern. The
Rem bitmap can be a DIB section bitmap, which is created by the CreateDIBSection function.
Rem @hBitmap        :   Windows 95/98: Creating brushes from bitmaps or DIBs larger than 8 by 8 pixels
Rem                     is not supported. If a larger bitmap is specified, only a portion of the
Rem                     bitmap is used.
Rem                     Windows NT/ 2000: Brushes can be created from bitmaps or DIBs larger than 8 by
Rem                     8 pixels.
Rem Return Values   :   If the function succeeds, the return value identifies a logical brush.
Rem                     If the function fails, the return value is NULL.
Public Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As stdole.OLE_HANDLE) As stdole.OLE_HANDLE

Rem The ExcludeClipRect function creates a new clipping region that consists of the existing clipping
Rem region minus the specified rectangle.
Rem @hdc            :   [in] Handle to the device context.
Rem @nLeftRect      :   [in] Specifies the x-coordinate, in logical units, of the upper-left corner
Rem                     of the rectangle.
Rem @nTopRect       :   [in] Specifies the y-coordinate, in logical units, of the upper-left corner
Rem                     of the rectangle.
Rem @nRightRect     :   [in] Specifies the x-coordinate, in logical units, of the lower-right corner
Rem                     of the rectangle.
Rem @nBottomRect    :   [in] Specifies the y-coordinate, in logical units, of the lower-right corner
Rem                     of the rectangle.
Rem @Return Values  :   See @Enum
Public Declare Function ExcludeClipRect Lib "gdi32" (ByVal hDC As stdole.OLE_HANDLE, _
                                                     ByVal nLeftRect As stdole.OLE_XPOS_PIXELS, _
                                                     ByVal nTopRect As stdole.OLE_YPOS_PIXELS, _
                                                     ByVal nRightRect As stdole.OLE_XSIZE_PIXELS, _
                                                     ByVal nBottomRect As stdole.OLE_YSIZE_PIXELS) As SELECT_OBJECT_RETURN_VALUE

Rem The DPtoLP function converts device coordinates into logical coordinates. The conversion depends
Rem on the mapping mode of the device context, the settings of the origins and extents for the window
Rem and viewport, and the world transformation.
Rem @hDC            :   [in] Handle to the device context.
Rem @lpPoints       :   [in/out] Pointer to an array of POINT structures. The x- and y-coordinates
Rem                     contained in each POINT structure will be transformed.
Rem @nCount         :   [in] Specifies the number of points in the array.
Rem @Return Values  :   If the function succeeds, the return value is nonzero. If the function fails,
Rem                     the return value is zero.
Rem @Remarks        :   The DPtoLP function fails if the device coordinates exceed 27 bits, or if the
Rem                     converted logical coordinates exceed 32 bits. In the case of such an overflow,
Rem                     the results for all the points are undefined.
Public Declare Function DPtoLP Lib "gdi32" (ByVal hDC As stdole.OLE_HANDLE, _
                                            ByRef lpPoints As POINTAPI, _
                                            ByVal nCount As Integer) As Long

Rem The SetMapMode function sets the mapping mode of the specified device context. The mapping mode
Rem defines the unit of measure used to transform page-space units into device-space units, and also
Rem defines the orientation of the device's x and y axes.
Rem @hDC            :   [in] Handle to the device context.
Rem @fnMapMode      :   See Enum
Rem @Return Values  :   If the function succeeds, the return value identifies the previous mapping mode.
Rem                     If the function fails, the return value is zero.
Rem @Remarks        :   The MM_TEXT mode allows applications to work in device pixels, whose size varies
Rem                     from device to device.
Rem                     The MM_HIENGLISH, MM_HIMETRIC, MM_LOENGLISH, MM_LOMETRIC, and MM_TWIPS modes
Rem                     are useful for applications drawing in physically meaningful units (such as
Rem                     inches or millimeters).
Rem                     The MM_ISOTROPIC mode ensures a 1:1 aspect ratio.
Rem                     The MM_ANISOTROPIC mode allows the x-coordinates and y-coordinates to be
Rem                     adjusted independently.
Public Declare Function SetMapMode Lib "gdi32" (ByVal hDC As stdole.OLE_HANDLE, _
                                                ByVal fnMapMode As MAPPING_MODES) As MAPPING_MODES

Rem The GetMapMode function retrieves the current mapping mode.
Rem @hDC            :   [in] Handle to the device context.
Rem @Return Values  :   See Enum.
Rem                     If the function fails, the return value is zero.
Public Declare Function GetMapMode Lib "gdi32" (ByVal hDC As stdole.OLE_HANDLE) As MAPPING_MODES

Rem The DrawCaption function draws a window caption.
Rem  @hWnd          :   [in] Handle to a window that supplies text and an icon for the window caption.
Rem @hDC            :   [in] Handle to a device context. The function draws the window caption into
Rem                     this device context.
Rem @lprc           :   [in] Pointer to a RECT structure that specifies the bounding rectangle for the
Rem                     window caption.
Rem @uFlags         :   [in] Specifies drawing options.
Rem @Return Values  :   If the function succeeds, the return value is nonzero.
Rem                     If the function fails, the return value is zero.
Public Declare Function DrawCaption Lib "user32" (ByVal hWnd As stdole.OLE_HANDLE, _
                                                  ByVal hDC As stdole.OLE_HANDLE, _
                                                  ByRef lprc As RECT, _
                                                  ByVal uFlags As DRAWCAPTIONOPTIONS) As Long

Rem The GradientFill function fills rectangle and triangle structures.
Rem @hDC            :   [in] Handle to the destination device context.
Rem @pVertex        :   [in] Pointer to an array of TRIVERTEX structures that each define a triangle
Rem                     vertex.
Rem @dwNumVertex    :   [in] The number of vertices in pVertex.
Rem @pMesh          :   [in] Array of GRADIENT_TRIANGLE structures in triangle mode, or an array of
Rem                     GRADIENT_RECT structures in rectangle mode.
Rem @dwNumMesh      :   [in] The number of elements (triangles or rectangles) in pMesh.
Rem @dwMode         :   [in] Specifies gradient fill mode.
Rem @Return Values  :   If the function succeeds, the return value is TRUE.
Rem                     If the function fails, the return value is FALSE.
Rem @Remarks        :   To add smooth shading to a triangle, call the GradientFill function with the
Rem                     three triangle endpoints. GDI will linearly interpolate and fill the triangle.
Rem                     To add smooth shading to a rectangle, call GradientFill with the upper-left and
Rem                     lower-right coordinates of the rectangle. There are two shading modes used when
Rem                     drawing a rectangle. In horizontal mode, the rectangle is shaded from
Rem                     left-to-right. In vertical mode, the rectangle is shaded from top-to-bottom.
Rem                     The GradientFill function uses a mesh method to specify the endpoints of the
Rem                     object to draw. All vertices are passed to GradientFill in the pVertex array.
Rem                     The pMesh parameter specifies how these vertices are connected to form an
Rem                     object. When filling a rectangle, pMesh points to an array of GRADIENT_RECT
Rem                     structures. Each GRADIENT_RECT structure specifies the index of two vertices
Rem                     in the pVertex array. These two vertices form the upper-left and lower-right
Rem                     boundary of one rectangle.
Rem                     In the case of filling a triangle, pMesh points to an array of
Rem                     GRADIENT_TRIANGLE structures. Each GRADIENT_TRIANGLE structure specifies the
Rem                     index of three vertices in the pVertex array. These three vertices form one
Rem                     triangle.
Rem                     To simplify hardware acceleration, this routine is not required to be
Rem                     pixel-perfect in the triangle interior.
Public Declare Function GradientFill Lib "Msimg32.dll" (ByVal hDC As stdole.OLE_HANDLE, _
                                                        ByRef pVertex As TRIVERTEX, _
                                                        ByVal dwNumVertex As Long, _
                                                        ByRef pMesh As GRADIENT_RECT, _
                                                        ByVal dwNumMesh As Long, _
                                                        ByVal dwMode As GRADFILLMODE) As Long

Rem The GetPixel function retrieves the red, green, blue (RGB) color value of the pixel at the
Rem specified coordinates.
Rem @hDC            :   [in] Handle to the device context.
Rem @nXPos          :   [in] Specifies the x-coordinate, in logical units, of the pixel to be examined.
Rem @nYPos          :   [in] Specifies the y-coordinate, in logical units, of the pixel to be examined.
Rem @Return Values  :   The return value is the RGB value of the pixel. If the pixel is outside of the
Rem                     current clipping region, the return value is CLR_INVALID.
Public Declare Function GetPixel Lib "gdi32" (ByVal hDC As stdole.OLE_HANDLE, _
                                              ByVal nXPos As stdole.OLE_XPOS_PIXELS, _
                                              ByVal nYPos As stdole.OLE_YPOS_PIXELS) As stdole.OLE_COLOR

Rem The SetPixel function sets the pixel at the specified coordinates to the specified color.
Rem @hDC            :  [in] Handle to the device context
Rem @X              :  [in] Specifies the x-coordinate, in logical units, of the point to be set.
Rem @Y              :  [in] Specifies the y-coordinate, in logical units, of the point to be set.
Rem @crColor        :   [in] Specifies the color to be used to paint the point. To create a COLORREF
Rem                     color value, use the RGB macro.
Rem @Return Values  :   If the function succeeds, the return value is the RGB value that the function
Rem                     sets the pixel to. This value may differ from the color specified by crColor;
Rem                     that occurs when an exact match for the specified color cannot be found.
Rem                     If the function fails, the return value is -1.
Public Declare Function SetPixel Lib "gdi32" (ByVal hDC As stdole.OLE_HANDLE, _
                                              ByVal x As stdole.OLE_XPOS_PIXELS, _
                                              ByVal y As stdole.OLE_YPOS_PIXELS, _
                                              ByVal crColor As stdole.OLE_COLOR) As stdole.OLE_COLOR

Rem The SaveDC function saves the current state of the specified device context (DC) by copying data
Rem describing selected objects and graphic modes (such as the bitmap, brush, palette, font, pen,
Rem region, drawing mode, and mapping mode) to a context stack.
Rem @hDC            :   [in] Handle to the DC whose state is to be saved.
Rem @Return Values  :   If the function succeeds, the return value identifies the saved state.
Rem                     If the function fails, the return value is zero.
Public Declare Function SaveDC Lib "gdi32" (ByVal hDC As stdole.OLE_HANDLE) As Long

Rem The RestoreDC function restores a device context (DC) to the specified state. The DC is restored
Rem by popping state information off a stack created by earlier calls to the SaveDC function.
Rem @hDC            :   [in] Handle to the DC.
Rem @nSavedDC       :   [in] Specifies the saved state to be restored. If this parameter is positive,
Rem                     nSavedDC represents a specific instance of the state to be restored. If this
Rem                     parameter is negative, nSavedDC represents an instance relative to the current
Rem                     state. For example, -1 restores the most recently saved state.
Rem @Return Values  :   If the function succeeds, the return value is nonzero.
Rem                     If the function fails, the return value is zero.
Public Declare Function RestoreDC Lib "gdi32" (ByVal hDC As stdole.OLE_HANDLE, _
                                               ByVal nSavedDC As Long) As Long

Rem The GetBoundsRect function obtains the current accumulated bounding rectangle for a specified
Rem device context.
Rem The system maintains an accumulated bounding rectangle for each application. An application can
Rem retrieve and set this rectangle.
Rem @hdc            :   [in] Handle to the device context whose bounding rectangle the function will
Rem                     return.
Rem @lprcBounds     :   [out] Pointer to the RECT structure that will receive the current bounding
Rem                     rectangle. The application's rectangle is returned in logical coordinates,
Rem                     and the bounding rectangle is returned in screen coordinates.
Rem @flags          :   [in] Specifies how the GetBoundsRect function will behave.
Rem @Return Values  :   The return value specifies the state of the accumulated bounding rectangle;
Rem @Remarks        :   The DCB_SET value is a combination of the bit values DCB_ACCUMULATE and
Rem                     DCB_RESET. Applications that check the DCB_RESET bit to determine whether the
Rem                     bounding rectangle is empty must also check the DCB_ACCUMULATE bit. The
Rem                     bounding rectangle is empty only if the DCB_RESET bit is 1 and the
Rem                     DCB_ACCUMULATE bit is 0.
Public Declare Function GetBoundsRect Lib "gdi32" (ByVal hDC As stdole.OLE_HANDLE, _
                                                   ByRef lprcBounds As RECT, _
                                                   ByVal flags As DEV_CNTXT_ACCUMUL_FLAG) As DEV_CNTXT_ACCUMUL_FLAG

Rem The DrawFrameControl function draws a frame control of the specified type and style.
Rem @hdc            :   [in] Handle to the device context of the window in which to draw the control.
Rem @lprc           :   [in] Pointer to a RECT structure that contains the logical coordinates of the
Rem                     bounding rectangle for frame control
Rem @uType          :   [in] Specifies the type of frame control to draw.
Rem @uState         :   [in] Specifies the initial state of the frame control.
Rem @Return Values  :   If the function succeeds, the return value is nonzero.
Rem                     If the function fails, the return value is zero.
Rem @Remarks        :   If uType is either DFC_MENU or DFC_BUTTON and uState is not DFCS_BUTTONPUSH,
Rem                     the frame control is a black-on-white mask (that is, a black frame control on
Rem                     a white background). In such cases, the application must pass a handle to a
Rem                     bitmap memory device control. The application can then use the associated
Rem                     bitmap as the hbmMask parameter to the MaskBlt function, or it can use the
Rem                     device context as a parameter to the BitBlt function using ROPs such as
Rem                     SRCAND and SRCINVERT.
Public Declare Function DrawFrameControl Lib "user32" (ByVal hDC As stdole.OLE_HANDLE, _
                                                      ByRef lprc As RECT, _
                                                      ByVal uType As DRAWFRAMECONTROL_TYPE, _
                                                      ByVal uState As DRAWFRAMECONTROL_STATE) As Long

Rem The GetStockObject function retrieves a handle to one of the stock pens, brushes, fonts, or palettes.
Rem @fnObject       :   [in] Specifies the type of stock object.
Rem @Return Values  :   If the function succeeds, the return value is a handle to the requested logical
Rem                     object.
Rem                     If the function fails, the return value is NULL.
Public Declare Function GetStockObject Lib "gdi32" (ByVal fnObject As STOCK_OBJ_TYPE) As stdole.OLE_HANDLE

Rem The SetParent function changes the parent window of the specified child window.
Rem @hWndChild      :   [in] Handle to the child window.
Rem @hWndNewParent  :   [in] Handle to the new parent window. If this parameter is NULL, the desktop
Rem                     window becomes the new parent window.
Rem                     Windows 2000: If this parameter is HWND_MESSAGE, the child window becomes a
Rem                     message-only window.
Rem @Return Values  :   If the function succeeds, the return value is a handle to the previous parent
Rem                     window.
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As stdole.OLE_HANDLE, _
                                                ByVal hWndNewParent As stdole.OLE_HANDLE) As stdole.OLE_HANDLE

Rem The SetLayeredWindowAttributes function sets the opacity and transparency color key of a layered
Rem window.
Rem @hWnd           :   [in] Handle to the layered window. A layered window is created by specifying
Rem                     WS_EX_LAYERED when creating the window with the CreateWindowEx function or by
Rem                     setting WS_EX_LAYERED via SetWindowLong after the window has been created.
Rem @crKey          :   [in] Pointer to a COLORREF value that specifies the transparency color key to
Rem                     be used when composing the layered window. All pixels painted by the window in
Rem                     this color will be transparent. To generate a COLORREF, use the RGB macro.
Rem @bAlpha         :   [in] Alpha value used to describe the opacity of the layered window. Similar
Rem                     to the SourceConstantAlpha member of the BLENDFUNCTION structure. When bAlpha
Rem                     is 0, the window is completely transparent. When bAlpha is 255, the window is
Rem                     opaque.
Rem @dwFlags        :   [in] Specifies an action to take.
Rem @Return Values  :   If the function succeeds, the return value is nonzero.
Rem                     If the function fails, the return value is zero.
Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As stdole.OLE_HANDLE, _
                                                                 ByVal crKey As Byte, _
                                                                 ByVal bAlpha As Byte, _
                                                                 ByVal dwFlags As LAYERED_WINDOW_ATTRIB) As Long

Rem The GetNearestColor function retrieves a color value identifying a color from the system palette
Rem that will be displayed when the specified color value is used.
Rem @hdc            :   [in] Handle to the device context.
Rem @crColor        :   [in] Specifies a color value that identifies a requested color.
Rem @Return Values  :   If the function succeeds, the return value identifies a color from the system
Rem                     palette that corresponds to the given color value.
Rem                     If the function fails, the return value is CLR_INVALID.
Public Declare Function GetNearestColor Lib "gdi32" (ByVal hDC As stdole.OLE_HANDLE, _
                                                     ByVal crColor As stdole.OLE_COLOR) As stdole.OLE_COLOR

Rem The SystemParametersInfo function retrieves or sets the value of one of the system-wide
Rem parameters. This function can also update the user profile while setting a parameter.
Rem @uiAction       :   [in] Specifies the system-wide parameter to retrieve or set.
Rem @uiParam        :   [in] Depends on the system parameter being queried or set.
Rem                     For more information about system-wide parameters, see the uiAction
Rem                     parameter. If not otherwise indicated, you must specify zero for this
Rem                     parameter.
Rem @pvParam        :   [in/out] Depends on the system parameter being queried or set.
Rem                     For more information about system-wide parameters, see the uiAction
Rem                     parameter. If not otherwise indicated, you must specify NULL for this
Rem                     parameter.
Rem @fWinIni        :   [in] If a system parameter is being set, specifies whether the user profile
Rem                     is to be updated, and if so, whether the WM_SETTINGCHANGE message is to be
Rem                     broadcast to all top-level windows to notify them of the change.
Rem                     This parameter can be zero if you don't want to update the user profile or
Rem                     broadcast the WM_SETTINGCHANGE message
Rem @Return Values  :   If the function succeeds, the return value is a nonzero value.
Rem                     If the function fails, the return value is zero
Public Declare Function SystemParametersInfo Lib "user32" _
                 Alias "SystemParametersInfoA" (ByVal uiAction As Long, _
                                                ByVal uiParam As Long, _
                                                ByRef pvParam As Any, _
                                                ByVal fWinIni As Long) As Long

Rem The FrameRect function draws a border around the specified rectangle by using the specified brush.
Rem The width and height of the border are always one logical unit.
Rem @hDC            :   [in] Handle to the device context in which the border is drawn
Rem @lprc           :   [in] Pointer to a RECT structure that contains the logical coordinates of the
Rem                     upper-left and lower-right corners of the rectangle.
Rem @hbr            :   [in] Handle to the brush used to draw the border.
Rem @Return Values  :   If the function succeeds, the return value is nonzero.
Rem                     If the function fails, the return value is zero.
Public Declare Function FrameRect Lib "user32" (ByVal hDC As stdole.OLE_HANDLE, _
                                                ByRef lprc As RECT, _
                                                ByVal hbr As stdole.OLE_HANDLE) As Integer

Rem The GetClassName function retrieves the name of the class to which the specified window belongs.
Rem @hWnd           :   [in] Handle to the window and, indirectly, the class to which the window
Rem                     belongs.
Rem @lpClassName    :   [out] Pointer to the buffer that is to receive the class name string.
Rem @nMaxCount      :   [in] Specifies the length, in TCHARs, of the buffer pointed to by the
Rem                     lpClassName parameter. The class name string is truncated if it is longer
Rem                     than the buffer.
Rem @Return Values  :   If the function succeeds, the return value is the number of TCHARs copied to
Rem                     the specified buffer.
Rem                     If the function fails, the return value is zero.
Public Declare Function GetClassName Lib "user32" _
                      Alias "GetClassNameA" (ByVal hWnd As stdole.OLE_HANDLE, _
                                            ByVal lpClassName As String, _
                                            ByVal nMaxCount As Long) As Integer

Rem /* The SetWindowsHookEx function installs an application-defined hook procedure into a hook chain.
Rem  * You would install a hook procedure to monitor the system for certain types of events. These
Rem  * events are associated either with a specific thread or with all threads in the same desktop as
Rem  * the calling thread.
Rem  * @idHook      :   [in] Specifies the type of hook procedure to be installed.
Rem  * @lpfn        :   [in] Pointer to the hook procedure. If the dwThreadId parameter is zero or
Rem  *                  specifies the identifier of a thread created by a different process, the
Rem  *                  lpfn parameter must point to a hook procedure in a dynamic-link library (DLL).
Rem  *                  Otherwise, lpfn can point to a hook procedure in the code associated with the
Rem  *                  current process.
Rem  * @hMod        :   [in] Handle to the DLL containing the hook procedure pointed to by the lpfn
Rem  *                  parameter. The hMod parameter must be set to NULL if the dwThreadId parameter
Rem  *                  specifies a thread created by the current process and if the hook procedure is
Rem  *                  within the code associated with the current process.
Rem  * @dwThreadId  :   [in] Specifies the identifier of the thread with which the hook procedure is
Rem  *                  to be associated. If this parameter is zero, the hook procedure is associated
Rem  *                  with all existing threads running in the same desktop as the calling thread.
Rem  * @Return Values:  If the function succeeds, the return value is the handle to the hook procedure.
Rem  *                  If the function fails, the return value is NULL.
Rem  * /
Public Declare Function SetWindowsHookEx Lib "user32" _
                  Alias "SetWindowsHookExA" (ByVal idHook As WINHOOKID, _
                                             ByVal lpfn As stdole.OLE_HANDLE, _
                                             ByVal hMod As stdole.OLE_HANDLE, _
                                             ByVal dwThreadId As Long) As Long

Rem /* The CallNextHookEx function passes the hook information to the next hook procedure in the current
Rem  * hook chain. A hook procedure can call this function either before or after processing the hook
Rem  * information.
Rem  * @hhk         :   [in] Handle to the current hook. An application receives this handle as a
Rem  *                  result of a previous call to the SetWindowsHookEx function.
Rem  * @nCode       :   [in] Specifies the hook code passed to the current hook procedure. The next
Rem  *                  hook procedure uses this code to determine how to process the hook
Rem  *                  information.
Rem  * @wParam      :   [in] Specifies the wParam value passed to the current hook procedure. The
Rem  *                  meaning of this parameter depends on the type of hook associated with the
Rem  *                  current hook chain.
Rem  * @lParam      :   [in] Specifies the lParam value passed to the current hook procedure. The
Rem  *                  meaning of this parameter depends on the type of hook associated with the
Rem  *                  current hook chain.
Rem  * @Return Values:  The return value is the value returned by the next hook procedure in the
Rem  *                  chain. The current hook procedure must also return this value. The meaning of
Rem  *                  the return value depends on the hook type.
Rem  * /
Public Declare Function CallNextHookEx Lib "user32" (ByVal hhk As stdole.OLE_HANDLE, _
                                                     ByVal nCode As WINHOOKID, _
                                                     ByVal wParam As Long, _
                                                     ByVal lParam As Long) As Long

Rem /* The UnhookWindowsHookEx function removes a hook procedure installed in a hook chain by the
Rem  * SetWindowsHookEx function.
Rem  * @hhk         :   [in] Handle to the hook to be removed. This parameter is a hook handle obtained
Rem  *                  by a previous call to SetWindowsHookEx.
Rem  * @Return Values:  If the function succeeds, the return value is nonzero.
Rem  *                  If the function fails, the return value is zero.
Rem  * /
Public Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hhk As stdole.OLE_HANDLE) As Long

Rem /* The GetDCEx function retrieves a handle to a display device context (DC) for the client area of
Rem  * a specified window or for the entire screen. You can use the returned handle in subsequent GDI
Rem  * functions to draw in the DC.
Rem  * This function is an extension to the GetDC function, which gives an application more control
Rem  * over how and whether clipping occurs in the client area.
Rem  * @hWnd        :   [in] Handle to the window whose DC is to be retrieved. If this value is NULL,
Rem  *                  GetDCEx retrieves the DC for the entire screen.
Rem  *                  Windows 98/Me, Windows 2000/XP: To get the DC for a specific display monitor,
Rem  *                  use the EnumDisplayMonitors and CreateDC functions.
Rem  * @hrgnClip    :   [in] Specifies a clipping region that may be combined with the visible region
Rem  *                  of the DC. If the value of flags is DCX_INTERSECTRGN or DCX_EXCLUDERGN, then
Rem  *                  the operating system assumes ownership of the region and will automatically
Rem  *                  delete it when it is no longer needed. In this case, applications should not
Rem  *                  use the region-not even delete it-after a successful call to GetDCEx.
Rem  * @flags       :   [in] Specifies how the DC is created.
Rem  * @Return Values:  If the function succeeds, the return value is the handle to the DC for the
Rem  *                  specified window.
Rem  *                  If the function fails, the return value is NULL. An invalid value for the hWnd
Rem  *                  parameter will cause the function to fail.
Rem  */
Public Declare Function GetDCEx Lib "user32" (ByVal hWnd As stdole.OLE_HANDLE, _
                                              ByVal hrgnClip As stdole.OLE_HANDLE, _
                                              ByVal flags As DC_WINDOW_RGN) As stdole.OLE_HANDLE

Rem /* The GetDesktopWindow function returns a handle to the desktop window. The desktop window covers
Rem  * the entire screen. The desktop window is the area on top of which all icons and other
Rem  * windows are painted.
Rem  * @Return Values:  The return value is a handle to the desktop window.
Rem  */
Public Declare Function GetDesktopWindow Lib "user32" () As stdole.OLE_HANDLE

Rem /* The WindowFromDC function returns a handle to the window associated with the specified
Rem  * display device context (DC). Output functions that use the specified device context draw
Rem  * into this window.
Rem  * @hDC       :   [in] Handle to the device context from which a handle for the associated
Rem  *                 window is to be retrieved.
Rem  * @Return Values:  The return value is a handle to the window associated with the specified
Rem  *                  DC.
Rem  */
Public Declare Function WindowFromDC Lib "user32" (ByVal hDC As stdole.OLE_HANDLE) As stdole.OLE_HANDLE

Rem /* The IntersectClipRect function creates a new clipping region from the intersection of the
Rem  * current clipping region and the specified rectangle.
Rem  * @hdc       :   [in] Handle to the device context
Rem  * @nLeftRect :   [in] Specifies the x-coordinate, in logical units, of the upper-left corner of
Rem  *                the rectangle.
Rem  * @nTopRect  :   [in] Specifies the y-coordinate, in logical units, of the upper-left corner of
Rem  *                the rectangle.
Rem  * @nRightRect:   [in] Specifies the x-coordinate, in logical units, of the lower-right corner
Rem  *                of the rectangle.
Rem  * @nBottomRect:  [in] Specifies the y-coordinate, in logical units, of the lower-right corner
Rem  *                of the rectangle.
Rem  * @Return Values: The return value specifies the new clipping region's type.
Rem  */
Public Declare Function IntersectClipRect Lib "gdi32" (ByVal hDC As stdole.OLE_HANDLE, _
                                                       ByVal nLeftRect As stdole.OLE_XPOS_PIXELS, _
                                                       ByVal nTopRect As stdole.OLE_YPOS_PIXELS, _
                                                       ByVal nRightRect As stdole.OLE_XSIZE_PIXELS, _
                                                       ByVal nBottomRect As stdole.OLE_YSIZE_PIXELS) As SELECT_OBJECT_RETURN_VALUE

Rem /* The ClientToScreen function converts the client-area coordinates of a specified point to
Rem  * screen coordinates.
Rem  * @hWnd      :   [in] Handle to the window whose client area is used for the conversion.
Rem  * @lpPoint   :   [in/out] Pointer to a POINT structure that contains the client coordinates
Rem  *                to be converted. The new screen coordinates are copied into this structure if
Rem  *                the function succeeds.
Rem  * @Return Values:  If the function succeeds, the return value is nonzero. If the function
Rem  *                fails, the return value is zero.
Rem  */
Public Declare Function ClientToScreen Lib "user32" (ByVal hWnd As stdole.OLE_HANDLE, _
                                                     ByRef lpPoint As POINTAPI) As Boolean

Rem /* The GetClassInfoEx function retrieves information about a window class, including a handle
Rem  * to the small icon associated with the window class. The GetClassInfo function does not
Rem  * retrieve a handle to the small icon.
Rem  * @hinst     :   [in] Handle to the instance of the application that created the class.
Rem  *                To retrieve information about classes defined by the system (such as buttons
Rem  *                or list boxes), set this parameter to NULL.
Rem  * @lpszClass :   [in] Pointer to a null-terminated string containing the class name. The name
Rem  *                must be that of a preregistered class or a class registered by a previous call
Rem  *                to the RegisterClass or RegisterClassEx function.
Rem  *                Alternatively, this parameter can be a class atom created by a previous call
Rem  *                to RegisterClass or RegisterClassEx. The atom must be in the low-order word
Rem  *                of lpszClass; the high-order word must be zero.
Rem  * @lpwcx     :   [out] Pointer to a WNDCLASSEX structure that receives the information about
Rem  *                the class.
Rem  * @Return Values: If the function finds a matching class and successfully copies the data, the
Rem  *                return value is nonzero.
Rem  *                If the function does not find a matching class and successfully copy the
Rem  *                data, the return value is zero.
Rem  */
Public Declare Function GetClassInfoEx Lib "user32" _
                 Alias "GetClassInfoExA" (ByVal hinst As stdole.OLE_HANDLE, _
                                          ByRef lpszClass As String, _
                                          ByRef lpwcx As WNDCLASSEX) As Long

Rem /* The SetClassLong function replaces the specified value at the specified offset in the extra
Rem  * class memory or the WNDCLASSEX structure for the class to which the specified window belongs.
Rem  * This function supersedes the SetClassLong function. To write code that is compatible with
Rem  * both 32-bit and 64-bit versions of Windows, use SetClassLongPtr.
Rem  * @hWnd      :   [in] Handle to the window and, indirectly, the class to which the window
Rem  *                belongs.
Rem  * @nIndex    :   [in] Specifies the value to replace.
Rem  * @dwNewLong :   [in] Specifies the replacement value.
Rem  * @Return Values:  If the function succeeds, the return value is the previous value of the
Rem  *                specified offset. If this was not previously set, the return value is zero.
Rem  *                If the function fails, the return value is zero.
Rem  */
Public Declare Function SetClassLongPtr Lib "user32" _
                 Alias "SetClassLongA" (ByVal hWnd As stdole.OLE_HANDLE, _
                                        ByVal nIndex As CLASSINFO_OPTION, _
                                        ByVal dwNewLong As Long) As Long

Rem /* The GetClassLongPtr function retrieves the specified value from the WNDCLASSEX structure
Rem  * associated with the specified window.
Rem  * If you are retrieving a pointer or a handle, this function supersedes the GetClassLong
Rem  * function. (Pointers and handles are 32 bits on 32-bit Windows and 64 bits on 64-bit Windows.)
Rem  * To write code that is compatible with both 32-bit and 64-bit versions of Windows, use
Rem  * GetClassLongPtr
Rem  * @hWnd      :   [in] Handle to the window and, indirectly, the class to which the window
Rem  *                belongs.
Rem  * @nIndex    :   [in] Specifies the value to retrieve. To retrieve a value from the extra
Rem  *                class memory, specify the positive, zero-based byte offset of the value to be
Rem  *                retrieved. Valid values are in the range zero through the number of bytes of
Rem  *                extra class memory, minus eight; for example, if you specified 24 or more
Rem  *                bytes of extra class memory, a value of 16 would be an index to the third
Rem  *                integer.
Rem  *Return Values:  If the function succeeds, the return value is the requested value.
Rem  *                If the function fails, the return value is zero.
Rem  */
Public Declare Function GetClassLongPtr Lib "user32" _
                Alias "GetClassLongA" (ByVal hWnd As stdole.OLE_HANDLE, _
                                       ByVal nIndex As CLASSINFO_OPTION) As Long

Rem /* The GetCurrentThreadId function retrieves the thread identifier of the calling thread.
Rem  * Return Values: The return value is the thread identifier of the calling thread.
Rem  */
Public Declare Function GetCurrentThreadId Lib "kernel32" () As Long

Rem /* The FrameRgn function draws a border around the specified region by using the specified
Rem  * brush.
Rem  * @hdc       :   [in] Handle to the device context.
Rem  * @hrgn      :   [in] Handle to the region to be enclosed in a border. The region's coordinates
Rem  *                are presumed to be in logical units.
Rem  * @hbr       :   [in] Handle to the brush to be used to draw the border.
Rem  * @nWidth    :   [in] Specifies the width, in logical units, of vertical brush strokes.
Rem  * @nHeight   :   [in] Specifies the height, in logical units, of horizontal brush strokes.
Rem  * Return Values: If the function succeeds, the return value is nonzero. If the function fails,
Rem  *                the return value is zero.
Rem  */
Public Declare Function FrameRgn Lib "gdi32" (ByVal hDC As stdole.OLE_HANDLE, _
                                              ByVal hrgn As stdole.OLE_HANDLE, _
                                              ByVal hbr As stdole.OLE_HANDLE, _
                                              ByVal nWidth As stdole.OLE_XSIZE_PIXELS, _
                                              ByVal nHeight As stdole.OLE_YSIZE_PIXELS) As Boolean

Rem /* The GetWindowThreadProcessId function retrieves the identifier of the thread that created
Rem  * the specified window and, optionally, the identifier of the process that created the
Rem  * window.
Rem  * @hWnd      :   [in] Handle to the window.
Rem  * @lpdwProcessId:[out] Pointer to a variable that receives the process identifier. If this
Rem  *                parameter is not NULL, GetWindowThreadProcessId copies the identifier of the
Rem  *                process to the variable; otherwise, it does not.
Rem  * Return Values: The return value is the identifier of the thread that created the window.
Rem  */
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As stdole.OLE_HANDLE, _
                                                               ByRef lpdwProcessId As Long) As Long

Rem /* The SetTimer function creates a timer with the specified time-out value.
Rem  * @hWnd      :   [in] Handle to the window to be associated with the timer. This window must
Rem  *                be owned by the calling thread. If this parameter is NULL, no window is
Rem  *                associated with the timer and the nIDEvent parameter is ignored.
Rem  * @nIDEvent  :   [in] Specifies a nonzero timer identifier. If the hWnd parameter is NULL,
Rem  *                this parameter is ignored.
Rem  *                If the hWnd parameter is not NULL and the window specified by hWnd already
Rem  *                has a timer with the value nIDEvent, then the existing timer is replaced by
Rem  *                the new timer. When SetTimer replaces a timer, the timer is reset. Therefore,
Rem  *                a message will be sent after the current time-out value elapses, but the
Rem  *                previously set time-out value is ignored.
Rem  * @uElapse   :   [in] Specifies the time-out value, in milliseconds.
Rem  * @lpTimerFunc:  [in] Pointer to the function to be notified when the time-out value elapses.
Rem  *                For more information about the function, see TimerProc.
Rem  *                If lpTimerFunc is NULL, the system posts a WM_TIMER message to the application
Rem  *                queue. The hwnd member of the message's MSG structure contains the value of
Rem  *                the hWnd parameter.
Rem  * Return Values: If the function succeeds and the hWnd parameter is NULL, the return value is
Rem  *                an integer identifying the new timer. An application can pass this value to
Rem  *                the KillTimer function to destroy the timer.
Rem  *                If the function succeeds and the hWnd parameter is not NULL, then the return
Rem  *                value is a nonzero integer. An application can pass the value of the nIDTimer
Rem  *                parameter to the KillTimer function to destroy the timer.
Rem  *                If the function fails to create a timer, the return value is zero.
Rem  */
Public Declare Function SetTimer Lib "user32" (ByVal hWnd As stdole.OLE_HANDLE, _
                                              ByVal nIDEvent As Long, _
                                              ByVal uElapse As Long, _
                                              ByVal lpTimerFunc As stdole.OLE_HANDLE) As Long

Rem /* The KillTimer function destroys the specified timer.
Rem  * @hWnd      :   [in] Handle to the window associated with the specified timer. This value must
Rem  *                be the same as the hWnd value passed to the SetTimer function that created
Rem  *                the timer.
Rem  * @uIDEvent  :   [in] Specifies the timer to be destroyed. If the window handle passed to
Rem  *                SetTimer is valid, this parameter must be the same as the uIDEvent value
Rem  *                passed to SetTimer. If the application calls SetTimer with hWnd set to NULL,
Rem  *                this parameter must be the timer identifier returned by SetTimer.
Rem  * Return Values: If the function succeeds, the return value is nonzero.
Rem  *                If the function fails, the return value is zero.
Rem  */
Public Declare Function KillTimer Lib "user32" (ByVal hWnd As stdole.OLE_HANDLE, _
                                                ByVal uIDEvent As Long) As Long

Rem /* Changes the luminance of a RGB value. Hue and saturation are not affected.
Rem  * @clrRGB    :   Initial RGB value.
Rem  * @n         :   Luminance in units of 0.1 percent of the total range.
Rem  *                For example, a value of n = 50 corresponds to 5 percent of the maximum
Rem  *                luminance.
Rem  * @fScale    :   If fScale is set to TRUE, n specifies how much to increment or decrement
Rem  *                the current luminance. If fScale is set to FALSE, n specifies the absolute
Rem  *                luminance.
Rem  * @Return Value: Returns the modified RGB value.
Rem  */
Public Declare Function ColorAdjustLuma Lib "shlwapi.dll" (ByVal clrRGB As stdole.OLE_COLOR, _
                                                           ByVal n As Integer, _
                                                           ByVal fScale As Boolean) As stdole.OLE_COLOR

Rem /* The GetWindow function retrieves a handle to a window that has the specified relationship
Rem  * (Z order or owner) to the specified window.
Rem  * @hWnd      :   [in] Handle to a window. The window handle retrieved is relative to this
Rem  *                window, based on the value of the uCmd parameter.
Rem  * @uCmd      :   [in] Specifies the relationship between the specified window and the window
Rem  *                whose handle is to be retrieved.
Rem  * Return Values: If the function succeeds, the return value is a window handle. If no window
Rem  *                exists with the specified relationship to the specified window, the return
Rem  *                value is NULL.
Rem  */
Public Declare Function GetWindow Lib "user32" (ByVal hWnd As stdole.OLE_HANDLE, _
                                                ByVal wCmd As GETWINDOW_OPTION) As stdole.OLE_HANDLE

Rem /* The UpdateWindow function updates the client area of the specified window by sending a
Rem  * WM_PAINT message to the window if the window's update region is not empty. The function
Rem  * sends a WM_PAINT message directly to the window procedure of the specified window, bypassing
Rem  * the application queue. If the update region is empty, no message is sent.
Rem  * @hWnd      :   [in] Handle to the window to be updated.
Rem  * Return Values: If the function succeeds, the return value is nonzero. If the function fails,
Rem  *                the return value is zero.
Rem  */
Public Declare Function UpdateWindow Lib "user32" (ByVal hWnd As stdole.OLE_HANDLE) As Long

Rem /* The MoveWindow function changes the position and dimensions of the specified window. For a
Rem  * top-level window, the position and dimensions are relative to the upper-left corner of the
Rem  * screen. For a child window, they are relative to the upper-left corner of the parent window's
Rem  * client area.
Rem  * @hWnd      :   [in] Handle to the window.
Rem  * @x         :   [in] Specifies the new position of the left side of the window.
Rem  * @y         :   [in] Specifies the new position of the top of the window.
Rem  * @nWidth    :   [in] Specifies the new width of the window.
Rem  * @nHeight   :   [in] Specifies the new height of the window.
Rem  * @bRepaint  :   [in] Specifies whether the window is to be repainted. If this parameter is
Rem  *                TRUE, the window receives a WM_PAINT message. If the parameter is FALSE, no
Rem  *                repainting of any kind occurs. This applies to the client area, the nonclient
Rem  *                area (including the title bar and scroll bars), and any part of the parent
Rem  *                window uncovered as a result of moving a child window.
Rem  * Return Values: If the function succeeds, the return value is nonzero.
Rem  *                If the function fails, the return value is zero.
Rem  */
Public Declare Function MoveWindow Lib "user32" (ByVal hWnd As stdole.OLE_HANDLE, _
                                                 ByVal x As stdole.OLE_XPOS_PIXELS, _
                                                 ByVal y As stdole.OLE_YPOS_PIXELS, _
                                                 ByVal nWidth As stdole.OLE_XSIZE_PIXELS, _
                                                 ByVal nHeight As stdole.OLE_YSIZE_PIXELS, _
                                                 ByVal bRepaint As Long) As Long

Rem /* Creates a GUID, a unique 128-bit integer used for CLSIDs and interface identifiers.
Rem  * @pguid     :   [out] Pointer to the requested GUID on return.
Rem  * Return Value:  S_OK; The GUID was successfully created.
Rem  */
Public Declare Function CoCreateGuid Lib "OLE32.DLL" (ByRef pGuid As UUID) As Long

Rem /* The GetAncestor function retrieves the handle to the ancestor of the specified window.
Rem  * @hwnd      :   [in] Handle to the window whose ancestor is to be retrieved. If this
Rem  *                parameter is the desktop window, the function returns NULL.
Rem  * @gaFlags   :   [in] Specifies the ancestor to be retrieved.
Rem  * Return Values: The return value is the handle to the ancestor window.
Rem  */
Public Declare Function GetAncestor Lib "user32" (ByVal hWnd As stdole.OLE_HANDLE, _
                                                  ByVal gaFlags As ANCESTOR_WINDOW) As stdole.OLE_HANDLE

Rem /* The SetWindowPos function changes the size, position, and Z order of a child, pop-up, or
Rem  * top-level window. Child, pop-up, and top-level windows are ordered according to their
Rem  * appearance on the screen. The topmost window receives the highest rank and is the first
Rem  * window in the Z order.
Rem  * @hWnd      :   [in] Handle to the window.
Rem  * @hWndInsertAfter: [in] Handle to the window to precede the positioned window in the Z order.
Rem  */
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As stdole.OLE_HANDLE, _
                                                   ByVal hwndInsertAfter As ZORDER, _
                                                   ByVal x As stdole.OLE_XPOS_PIXELS, _
                                                   ByVal y As stdole.OLE_YPOS_PIXELS, _
                                                   ByVal cx As stdole.OLE_XSIZE_PIXELS, _
                                                   ByVal cy As stdole.OLE_YSIZE_PIXELS, _
                                                   ByVal uFlags As WINDOWPOSFLAGS) As Long

Rem /* The CreatePen function creates a logical pen that has the specified style, width, and color.
Rem  * The pen can subsequently be selected into a device context and used to draw lines and curves.
Rem  * @fnPenStyle  : [in] Specifies the pen style.
Rem  * @nWidth      : [in] Specifies the width of the pen, in logical units. If nWidth is zero, the
Rem  *                pen is a single pixel wide, regardless of the current transformation.
Rem  *                CreatePen returns a pen with the specified width bit with the PS_SOLID style if
Rem  *                you specify a width greater than one for the following styles: PS_DASH, PS_DOT,
Rem  *                PS_DASHDOT, PS_DASHDOTDOT.
Rem  * @crColor     : [in] Specifies a color reference for the pen color.
Rem  * Return Values: If the function succeeds, the return value is a handle that identifies a logical
Rem  *                pen. If the function fails, the return value is NULL
Rem  */
Public Declare Function CreatePen Lib "gdi32" (ByVal fnPenStyle As PEN_STYLE, _
                                      ByVal nWidth As stdole.OLE_XSIZE_PIXELS, _
                                      ByVal crColor As stdole.OLE_COLOR) As stdole.OLE_HANDLE

Rem /* The Rectangle function draws a rectangle. The rectangle is outlined by using the current pen
Rem  * and filled by using the current brush.
Rem  * @hdc         : [in] Handle to the device context.
Rem  * @nLeftRect   : [in] Specifies the x-coordinate, in logical coordinates, of the upper-left
Rem  *                corner of the rectangle.
Rem  * @nTopRect    : [in] Specifies the y-coordinate, in logical coordinates, of the upper-left
Rem  *                corner of the rectangle.
Rem  * @nRightRect  : [in] Specifies the x-coordinate, in logical coordinates, of the lower-right
Rem  *                corner of the rectangle.
Rem  * @nBottomRect : [in] Specifies the y-coordinate, in logical coordinates, of the lower-right
Rem  *                corner of the rectangle.
Rem  * Return Values: If the function succeeds, the return value is nonzero.
Rem  *                If the function fails, the return value is zero.
Rem  */
Public Declare Function Rectangle Lib "gdi32" (ByVal hDC As stdole.OLE_HANDLE, _
                                               ByVal nLeftRect As stdole.OLE_XPOS_PIXELS, _
                                               ByVal nTopRect As stdole.OLE_YPOS_PIXELS, _
                                               ByVal nRightRect As stdole.OLE_XSIZE_PIXELS, _
                                               ByVal nBottomRect As stdole.OLE_YSIZE_PIXELS) As Boolean
                                               
Rem /* The ExitWindowsEx function either logs off the current user, shuts down the system, or shuts
Rem  * down and restarts the system. It sends the WM_QUERYENDSESSION message to all applications to
Rem  * determine if they can be terminated.
Rem  * @uFlags      : [in] Specifies the type of shutdown.
Rem  * @dwReason    : Windows XP: [in] Specifies the reason for initiating the shutdown.
Rem  *                Windows 2000 and earlier, Windows 95/98/Me: This parameter is ignored.
Rem  * Return Values: If the function succeeds, the return value is nonzero.
Rem  *                If the function fails, the return value is zero.
Rem  * Remarks      :
Rem  * The ExitWindowsEx function returns as soon as it has initiated the shutdown. The shutdown or logoff
Rem  * then proceeds asynchronously.
Rem  * To set a shutdown priority for an application relative to other applications in the system,
Rem  * use the SetProcessShutdownParameters function.
Rem  * During a shutdown or log-off operation, applications that are shut down are allowed a specific
Rem  * amount of time to respond to the shutdown request. If the time expires, the system displays a
Rem  * dialog box that allows the user to forcibly shut down the application, to retry the shutdown,
Rem  * or to cancel the shutdown request. If the EWX_FORCE value is specified, the system always forces
Rem  * applications to close and does not display the dialog box.
Rem  * Windows 2000/XP: If the EWX_FORCEIFHUNG value is specified, the system forces hung applications
Rem  * to close and does not display the dialog box.
Rem  * Windows 95/98/Me: Because of the design of the shell, calling ExitWindowsEx with EWX_FORCE fails
Rem  * to completely log off the user (the system terminates the applications and displays the Enter
Rem  * Windows Password dialog box, however, the user's desktop remains.) To log off the user forcibly
Rem  * , terminate the Explorer process before calling ExitWindowsEx with EWX_LOGOFF and EWX_FORCE.
Rem  * Console processes receive a separate notification message, CTRL_SHUTDOWN_EVENT or CTRL_LOGOFF_EVENT
Rem  * , as the situation warrants. A console process routes these messages to its HandlerRoutine function.
Rem  * ExitWindowsEx sends these notification messages asynchronously; thus, an application cannot
Rem  * assume that the console notification messages have been handled when a call to ExitWindowsEx
Rem  * returns.
Rem  * Windows NT/2000/XP: To shut down or restart the system, the calling process must use the
Rem  * AdjustTokenPrivileges function to enable the SE_SHUTDOWN_NAME privilege.
Rem  * Windows 95/98/Me: ExitWindowEx does not work from a console application.
Rem  */
Public Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As EXIT_WINDOWS_FLAGS, _
                                                    ByVal dwReason As SHUTDOWN_REASON) As Long

Rem /* The AdjustTokenPrivileges function enables or disables privileges in the specified access
Rem  * token. Enabling or disabling privileges in an access token requires TOKEN_ADJUST_PRIVILEGES
Rem  * access.
Rem  * @TokenHandle   : [in] Handle to the access token that contains the privileges to be modified.
Rem  *                  The handle must have TOKEN_ADJUST_PRIVILEGES access to the token. If the
Rem  *                  PreviousState parameter is not NULL, the handle must also have TOKEN_QUERY
Rem  *                  access.
Rem  * @DisableAllPrivileges: [in] Specifies whether the function disables all of the token's
Rem  *                  privileges. If this value is TRUE, the function disables all privileges and
Rem  *                  ignores the NewState parameter. If it is FALSE, the function modifies privileges
Rem  *                  based on the information pointed to by the NewState parameter.
Rem  * @NewState      : [in] Pointer to a TOKEN_PRIVILEGES structure that specifies an array of
Rem  *                  privileges and their attributes. If the DisableAllPrivileges parameter is
Rem  *                  FALSE, AdjustTokenPrivileges enables or disables these privileges for the
Rem  *                  token. If you set the SE_PRIVILEGE_ENABLED attribute for a privilege, the
Rem  *                  function enables that privilege; otherwise, it disables the privilege.
Rem  *                  If DisableAllPrivileges is TRUE, the function ignores this parameter.
Rem  * @BufferLength  : [in] Specifies the size, in bytes, of the buffer pointed to by the
Rem  *                  PreviousState parameter. This parameter can be zero if the PreviousState
Rem  *                  parameter is NULL.
Rem  * @PreviousState : [out] Pointer to a buffer that the function fills with a TOKEN_PRIVILEGES
Rem  *                  structure that contains the previous state of any privileges that the
Rem  *                  function modifies. This parameter can be NULL.
Rem  *                  If you specify a buffer that is too small to receive the complete list of
Rem  *                  modified privileges, the function fails and does not adjust any privileges.
Rem  *                  In this case, the function sets the variable pointed to by the ReturnLength
Rem  *                  parameter to the number of bytes required to hold the complete list of
Rem  *                  modified privileges.
Rem  * @ReturnLength  : [out] Pointer to a variable that receives the required size, in bytes, of the
Rem  *                  buffer pointed to by the PreviousState parameter. This parameter can be NULL
Rem  *                  if PreviousState is NULL.
Rem  * Return Values  : If the function succeeds, the return value is nonzero.
Rem  */
Public Declare Function AdjustTokenPrivileges Lib "advapi32" (ByVal TokenHandle As stdole.OLE_HANDLE, _
                                                              ByVal DisableAllPrivileges As Boolean, _
                                                              ByRef NewState As TOKEN_PRIVILEGES, _
                                                              ByVal BufferLength As Long, _
                                                              ByRef PreviousState As TOKEN_PRIVILEGES, _
                                                              ByRef ReturnLength As Long) As Boolean

Rem /* The OpenProcessToken function opens the access token associated with a process.
Rem  * @ProcessHandle:  [in] Handle to the process whose access token is opened.
Rem  * @DesiredAccess:  [in] Specifies an access mask that specifies the requested types of access to
Rem  *                  the access token. These requested access types are compared with the token's
Rem  *                  discretionary access-control list (DACL) to determine which accesses are
Rem  *                  granted or denied. For a list of access rights for access tokens, see
Rem  *                  Access Rights for Access-Token Objects.
Rem  * @TokenHandle  :  [out] Pointer to a handle identifying the newly opened access token when the
Rem  *                  function returns.
Rem  */
Public Declare Function OpenProcessToken Lib "advapi32" (ByVal ProcessHandle As stdole.OLE_HANDLE, _
                                                         ByVal DesiredAccess As TOKEN_RIGHTS, _
                                                         ByRef TokenHandle As stdole.OLE_HANDLE) As Boolean

Rem /* The CloseHandle function closes an open object handle.
Rem  * @hObject     : [in/out] Handle to an open object.
Rem  * Return Values: If the function succeeds, the return value is nonzero.
Rem  *                If the function fails, the return value is zero. To get extended error
Rem  *                information, call GetLastError.
Rem  *                Windows NT/2000/XP: Closing an invalid handle raises an exception when the
Rem  *                application is running under a debugger. This includes closing a handle twice,
Rem  *                and using CloseHandle on a handle returned by the FindFirstFile function.
Rem  */
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As stdole.OLE_HANDLE) As Long

Rem /* The GetTokenInformation function retrieves a specified type of information about an access
Rem  * token. The calling process must have appropriate access rights to obtain the information.
Rem  * To determine if a user is a member of a specific group, use the CheckTokenMembership function.
Rem  * @TokenHandle : [in] Handle to an access token from which information is retrieved.
Rem  *                If TokenInformationClass specifies TokenSource, the handle must have TOKEN_QUERY_SOURCE
Rem  *                access. For all other TokenInformationClass values, the handle must have
Rem  *                TOKEN_QUERY access.
Rem  * @TokenInformationClass:  [in] Specifies a value from the TOKEN_INFORMATION_CLASS enumerated
Rem  *                type to identify the type of information the function retrieves.
Rem  * @TokenInformation: [out] Pointer to a buffer the function fills with the requested information.
Rem  *                The structure put into this buffer depends upon the type of information
Rem  *                specified by the TokenInformationClass parameter.
Rem  * @TokenInformationLength: [in] Specifies the size, in bytes, of the buffer pointed to by the
Rem  *                TokenInformation parameter. If TokenInformation is NULL, this parameter must be
Rem  *                zero.
Rem  * @ReturnLength  : [out] Pointer to a variable that receives the number of bytes needed for the
Rem  *                buffer pointed to by the TokenInformation parameter. If this value is larger
Rem  *                than the value specified in the TokenInformationLength parameter, the function
Rem  *                fails and stores no data in the buffer.
Rem  *                If the value of the TokenInformationClass parameter is TokenDefaultDacl and
Rem  *                the token has no default DACL, the function sets the variable pointed to by
Rem  *                ReturnLength to sizeof(TOKEN_DEFAULT_DACL) and sets the DefaultDacl member of
Rem  *                the TOKEN_DEFAULT_DACL structure to NULL.
Rem  * Return Values: If the function succeeds, the return value is nonzero.
Rem  *                If the function fails, the return value is zero.
Rem  */
Public Declare Function GetTokenInformation Lib "advapi32" (ByVal TokenHandle As stdole.OLE_HANDLE, _
                                                                ByVal TokenInformationClass As TOKEN_INFORMATION_CLASS, _
                                                                ByRef TokenInformation As Any, _
                                                                ByVal TokenInformationLength As Long, _
                                                                ByVal ReturnLength As Long) As Long

Rem /* The LookupPrivilegeValue function retrieves the locally unique identifier (LUID) used on a
Rem  * specified system to locally represent the specified privilege name.
Rem  * @lpSystemName : [in] Pointer to a null-terminated string specifying the name of the system on
Rem  *                 which the privilege name is looked up. If a null string is specified, the
Rem  *                 function attempts to find the privilege name on the local system.
Rem  * @lpName       : [in] Pointer to a null-terminated string that specifies the name of the
Rem  *                 privilege, as defined in the Winnt.h header file. For example, this parameter
Rem  *                 could specify the constant, SE_SECURITY_NAME, or its corresponding string,
Rem  *                 "SeSecurityPrivilege".
Rem  * @lpLuid      :  [out] Pointer to a variable that receives the locally unique identifier by
Rem  *                 which the privilege is known on the system, specified by the lpSystemName
Rem  *                 parameter.
Rem  * Return Values:  If the function succeeds, the return value is nonzero.
Rem  *                 If the function fails, the return value is zero.
Rem  */
Public Declare Function LookupPrivilegeValue Lib "advapi32" _
                   Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, _
                                                  ByVal lpName As String, _
                                                  ByRef lpLuid As Luid) As Boolean

Rem /* The GetCurrentProcess function retrieves a pseudo handle for the current process.
Rem  * Return Values: The return value is a pseudo handle to the current process.
Rem  */
Public Declare Function GetCurrentProcess Lib "kernel32" () As stdole.OLE_HANDLE


Rem-------------------------------------------------------------------------
Rem @Name                 :                 APIEnumerations
Rem @Type                 :                 Standard
Rem @Scope Qualifier      :                 Private
Rem @Purpose              :                 Provides and coordinates various
Rem                                         useful Win32 Types used by the
Rem                                         solution.
Rem @Creation Date        :                 Saturday, 28 January 2002.
Rem @Creation Author      :                 Shantibhushan
Rem-------------------------------------------------------------------------

Public Type LPINITCOMMONCONTROLSEX ' This structure carries information used to load common
                                   ' control classes from the dynamic-link library (DLL).
  dwSize      As Long              ' size of this structure
  dwICC       As COMMON_CONTROL_TYPES ' flags indicating which classes to be initialized
End Type

Public Type NMHDR                  ' This structure contains information about a notification message...
  hWndFrom    As stdole.OLE_HANDLE ' Window handle to the control sending a message...
  idfrom      As Long              ' Identifier of the control sending a message...
  code        As Long              ' Notification code. This member can be a control-specific notification
                                   ' code or it can be one of the common notification codes...
End Type

Public Type POINTAPI               ' Mouse Location...
  x           As stdole.OLE_XPOS_PIXELS ' X-position...
  y           As stdole.OLE_YPOS_PIXELS ' Y-position...
End Type

Public Type Size                   ' The SIZE structure specifies the width and height of a rectangle.
  cx          As stdole.OLE_XSIZE_PIXELS ' Specifies the rectangle's width. The units depend on which function uses this.
  cy          As stdole.OLE_YSIZE_PIXELS ' Specifies the rectangle's height. The units depend on which function uses this.
End Type

Public Type RECT                   ' This structure defines the coordinates of the upper-left and lower-right
                                   ' corners of a rectangle...
  Left        As stdole.OLE_XPOS_PIXELS ' Specifies the x-coordinate of the upper-left corner of the rectangle...
  Top         As stdole.OLE_YPOS_PIXELS ' Specifies the y-coordinate of the upper-left corner of the rectangle...
  Right       As stdole.OLE_XSIZE_PIXELS ' Specifies the x-coordinate of the lower-right corner of the rectangle...
  Bottom      As stdole.OLE_XSIZE_PIXELS ' Specifies the y-coordinate of the lower-right corner of the rectangle...
End Type

Public Type TOOLINFO               ' Structure that carries information for tooltip...
  cbSize      As Long              ' Size of this structure...
  uFlags      As TOOLTIP_FLAGS     ' ToolTip flags...
  hWnd        As stdole.OLE_HANDLE ' Handle of the member...
  uID         As Long              ' Message ID for tooltip...
  rec         As RECT              ' Rectangle for tooltip...
  hinst       As stdole.OLE_HANDLE ' Null...
  lpszText    As String            ' Message to be displayed in tooltip...
#If (WIN32_IE >= &H300) Then
  lParam      As Long              ' Additional information to be passed...
#End If
End Type

Public Type TOOLTIP_TITLE
  dwSize As Long                   ' DWORD that contains the size of the tooltip icon.
  uTitleBitmap As Long             ' UINT that specifies the tooltip icon.
  cch As Long                      ' UINT that specifies the number of characters in the title.
  pszTitle As String               ' Pointer to a wide character string that contains the title.
End Type

Public Type NMTTDISPINFO
  hdr         As NMHDR
  lpszText    As String
#If (UNICODE) Then
  szText      As String * 160
#Else
  szText      As String * 80
#End If
  hinst       As stdole.OLE_HANDLE
  uFlags      As Long
#If (WIN32_IE >= &H300) Then
  lParam      As Long
#End If
End Type

Public Type UUID
  Data1                   As Long
  Data2                   As Integer
  Data3                   As Integer
  Data4(&H0& To &H7&)    As Byte
End Type

Public Type NOTIFYICONDATAEX
  cbSize            As Long         ' Size of this structure, in bytes.
  hWnd              As stdole.OLE_HANDLE ' Handle to the window that receives notification messages associated with      an icon in the taskbar status area. The Shell uses hWnd and uID to identify which icon to operate on when Shell_NotifyIcon is invoked.

  uID               As Long         ' Application-defined identifier of the taskbar icon. The Shell uses hWnd and uID to identify which icon to operate on when Shell_NotifyIcon is invoked. You can have multiple icons associated with a single hWnd by assigning each a different uID.

  uFlags            As Long         ' Flags that indicate which of the other members contain valid data.
  uCallbackMessage  As Long         ' Application-defined message identifier.
  hIcon             As stdole.OLE_HANDLE ' Handle to the icon to be added, modified, or deleted.

  #If (WIN32_IE < &H500) Then         ' [0x0500]
                                    ' Pointer to a null-terminated string with the text for a standard ToolTip.
      szTip         As String * 64  ' It can have a maximum of 64 characters including the terminating NULL.
  #Else
      szTip         As String * 128 ' For Version 5.0 and later, szTip can have a maximum of 128 characters, including the terminating NULL.
  #End If

  #If (WIN32_IE >= &H500) Then
      dwState       As Long         ' State of the icon.
      dwStateMask   As Long         ' A value that specifies which bits of the state member are retrieved or modified. For example, setting this member to NIS_HIDDEN causes only the item's hidden state to be retrieved.
      szInfo        As String * 256 ' Pointer to a null-terminated string with the text for a balloon ToolTip. It can have a maximum of 255 characters. To remove the ToolTip, set the NIF_INFO flag in uFlags and set szInfo to an empty string.
      uTimeoutOrVersion As Long     ' The timeout value, in milliseconds, for a balloon ToolTip, alongwith the version number.

      szInfoTitle   As String * 64  ' Pointer to a null-terminated string containing a title for a balloon ToolTip. This title appears in boldface above the text. It can have a maximum of 63 characters.
      dwInfoFlags   As Long         ' Flags that can be set to add an icon to a balloon ToolTip. It is placed to the left of the title. If the szTitleInfo member is zero-length, the icon is not shown.
  #End If

  #If (WIN32_IE >= &H600) Then
      guidItem      As Guid         ' Reserved.
  #End If
End Type

Public Type OSVERSIONINFOEX                         ' The OSVERSIONINFOEX structure contains operating system version information. The information includes major and minor version numbers,
                                                    ' a build number, a platform identifier, and information about product suites and the latest Service Pack installed on the system. This
                                                    ' structure is used with the GetVersionEx and VerifyVersionInfo functions.
  dwOSVersionInfoSize             As Long           ' Specifies the size, in bytes, of this data structure.
  dwMajorVersion                  As MAJORVERSION_NUMBERS ' Identifies the major version number of the operating system
  dwMinorVersion                  As MINORVERSION_NUMBERS ' Identifies the minor version number of the operating system
  dwBuildNumber                   As Long           ' Identifies the build number of the operating system.
  dwPlatformId                    As PLATFORM_ID    ' Identifies the operating system platform. This member can be VER_PLATFORM_WIN32_NT.
  szCSDVersion                    As String * 128   ' Contains a null-terminated string, such as "Service Pack 3", that indicates the latest Service Pack installed on the system. If no Service Pack has been installed, the string is empty.
  wServicePackMajor               As Integer        ' Identifies the major version number of the latest Service Pack installed on the system. For example, for Service Pack 3, the major version number is 3. If no Service Pack has been installed, the value is zero.
  wServicePackMinor               As Integer        ' Identifies the minor version number of the latest Service Pack installed on the system. For example, for Service Pack 3, the minor version number is 0.
  wSuiteMask                      As Integer        ' A set of bit flags that identify the product suites available on the system.
  wProductType                    As Byte           ' Indicates additional information about the system.
  wReserved                       As Byte           ' Reserved for future use.
End Type

Rem The MENUITEMINFOEX structure contains information about a menu item.
Rem @cbSize           :   Size of structure, in bytes.
Rem @fMask            :   Members to retrieve or set.
Rem @fType            :   Menu item type.
Rem @fMask            :   Members to retrieve or set. This member can be one or more of these values.
Rem @fState           :   Menu item state. It can be one or more of these values:
Rem                       MFS_CHECKED
Rem                       MFS_ENABLED
Rem                       MFS_HILITE
Rem                       MFS_UNCHECKED
Rem                       MFS_UNHILITE
Rem  @wID             :   Application-defined 16-bit value that identifies the menu item.
Rem  @hSubMenu        :   Handle to the drop-down menu or submenu associated with the menu item. If
Rem                       the menu item is not an item that opens a drop-down menu or submenu, this
Rem                       member is NULL.
Rem  @hbmpChecked     :   Handle to the bitmap to display next to the item if it is checked. If this
Rem                       member is NULL, a default bitmap is used. If the MFT_RADIOCHECK type value
Rem                       is specified, the default bitmap is a bullet. Otherwise, it is a check mark.
Rem  @hbmpUnchecked   :   Handle to the bitmap to display next to the item if it is not checked. If this
Rem                       member is NULL, no bitmap is used.
Rem  @dwItemData      :   Specifies the application-defined value associated with the menu item.
Rem  @dwTypeData      :   Specifies the content of the menu item. This member is used only if the
Rem                       MIIM_TYPE flag is set in the fMask member. Before calling GetMenuItemInfo,
Rem                       the application must set this member to point to a buffer whose length is specified
Rem                       by the cch member. If the retrieved menu item is of the type MFT_STRING, then
Rem                       GetMenuItemInfo copies the menu item text to the buffer. If the retrieved menu item
Rem                       is of some other type, then GetMenuItemInfo sets the dwTypeData member to a value whose
Rem                       type is specified by the fType member.
Rem                       When using with the SetMenuItemInfo function, this member should contains a value whose
Rem                       type is specified by the fType member.
Rem  @cch             :   Length of the menu item text when information is received about a menu item of the
Rem                       MFT_STRING type. This member is used only if the MIIM_TYPE flag is set in the fMask
Rem                       member and is zero otherwise. This member is ignored when the content of a menu item
Rem                       is set by calling SetMenuItemInfo.
Rem                       Before calling GetMenuItemInfo, the application must set this member to the length of the
Rem                       buffer pointed to by the dwTypeData member. If the retrieved menu item is of type MFT_STRING
Rem                       (as indicated by the fType member), then GetMenuItemInfo sets cch to the length of the retrieved
Rem                       string. If the retrieved menu item is of some other type, GetMenuItemInfo sets the cch
Rem                       member to zero.
Rem @hbmpItem         :   Handle to the bitmap to be displayed
Public Type MENUITEMINFOEX
  cbSize                  As Long
  fMask                   As MENU_ITEM_INFO_FLAGS
  fType                   As MENU_ITEM_TYPE_FLAGS
  fState                  As MENU_ITEM_TYPE_FLAGS
  wID                     As Long
  hSubMenu                As stdole.OLE_HANDLE
  hbmpChecked             As stdole.OLE_HANDLE
  hbmpUnchecked           As stdole.OLE_HANDLE
  dwItemData              As Long
  dwTypeData              As String
  cch                     As Long
#If (WIN32_IE >= &H600) Then
  hbmpItem                As stdole.OLE_HANDLE
#End If
End Type

Rem The DRAWITEMSTRUCT structure provides information the owner window must have to determine how to paint
Rem an owner-drawn control or menu item. The owner window of the owner-drawn control or menu item receives
Rem a pointer to this structure as the lParam parameter of the WM_DRAWITEM message.
Rem @CtlType          :   Specifies the control type.
Rem @CtlID            :   Specifies the identifier of the combo box, list box, button, or static control.
Rem                       This member is not used for a menu item.
Rem @itemID           :   Specifies the menu item identifier for a menu item or the index of the item in
Rem                       a list box or combo box.
Rem @itemAction       :   Specifies the drawing action required.
Rem @itemState        :   Specifies the visual state of the item after the current drawing action takes place.
Rem @hwndItem         :   Handle to the control for combo boxes, list boxes, buttons, and static controls.
Rem                       For menus, this member is a handle to the menu containing the item.
Rem @hDC              :   Handle to a device context; this device context must be used when performing
Rem                       drawing operations on the control.
Rem @rcItem           :   Specifies a rectangle that defines the boundaries of the control to be drawn.
Rem @itemData         :   Specifies the application-defined value associated with the menu item.
Public Type DRAWITEMSTRUCT
  CtlType                 As OWNER_DRAWN_CONTROL_TYPE
  CtlID                   As Long
  itemID                  As Long
  itemAction              As OWNER_DRAWN_CONTROL_ACTION
  itemState               As OWNER_DRAWN_CONTROL_STATE
  hwndItem                As stdole.OLE_HANDLE
  hDC                     As stdole.OLE_HANDLE
  rcItem                  As RECT
  itemData                As Long
End Type

Rem The MEASUREITEMSTRUCT structure informs the system of the dimensions of an owner-drawn control or
Rem menu item. This allows the system to process user interaction with the control correctly.
Rem @CtlType          :   Specifies the control type.
Rem @CtlID            :   Specifies the identifier of the combo box, list box, button, or static control.
Rem                       This member is not used for a menu item.
Rem @itemID           :   Specifies the menu item identifier for a menu item or the index of the item in
Rem                       a list box or combo box.
Rem @itemWidth        :   Specifies the width, in pixels, of a menu item. Before returning from the
Rem                       message, the owner of the owner-drawn menu item must fill this member.
Rem @itemHeight       :   Specifies the height, in pixels, of an individual item in a list box or a menu
Rem                       . Before returning from the message, the owner of the owner-drawn combo box,
Rem                       list box, or menu item must fill out this member.
Rem @itemData         :   Specifies the application-defined value associated with the menu item.
Public Type MEASUREITEMSTRUCT
  CtlType                 As OWNER_DRAWN_CONTROL_TYPE
  CtlID                   As Long
  itemID                  As Long
  itemWidth               As stdole.OLE_XSIZE_PIXELS
  itemHeight              As stdole.OLE_YSIZE_PIXELS
  itemData                As Long
End Type

Rem This structure defines the attributes of a font.
Public Type LOGFONT
  lfHeight                As Long ' Specifies the height, in logical units, of the
                                  ' font's character cell or character.
  lfWidth                 As Long ' Specifies the average width, in logical units, of characters in
                                  ' the font.
  lfEscapement            As Long ' Specifies the angle, in tenths of degrees, between the escapement
                                  ' vector and the x-axis of the device.
                                  ' The lfEscapement member specifies both the escapement and orientation.
                                  ' You should set lfEscapement and lfOrientation to the same value.
  lfOrientation           As Long ' Specifies the angle, in tenths of degrees, between each character's
                                  ' base line and the x-axis of the device.
  lfWeight                As LOGICAL_FONT_WEIGHT ' Specifies the weight of the font in the range 0 through 1000.
  lfItalic                As Byte ' Specifies an italic font if set to TRUE.
  lfUnderline             As Byte ' Specifies an underlined font if set to TRUE.
  lfStrikeOut             As Byte ' Specifies a strikeout font if set to TRUE.
  lfCharSet               As Byte ' Specifies the character set.
  lfOutPrecision          As Byte ' Specifies the output precision. The output
                                  ' precision defines how closely the output must match the requested
                                  ' font's height, width, character orientation, escapement, pitch, and
                                  ' font type.
  lfClipPrecision         As Byte ' Specifies the clipping precision. The clipping
                                  ' precision defines how to clip characters that are partially outside
                                  ' the clipping region.
  lfQuality               As Byte ' Specifies the output quality. The output
                                  ' quality defines how carefully the graphics device interface (GDI)
                                  ' must attempt to match the logical-font attributes to those of an
                                  ' actual physical font.
  lfPitchAndFamily        As Byte ' Specifies the pitch and family of the font.
                                  ' The proper value can be obtained by using the Boolean OR operator to
                                  ' join one pitch constant with one family constant.
  lfFaceName              As String * LF_FACESIZE ' Specifies a null-terminated string that specifies
                                  ' the typeface name of the font. The length of this string must not
                                  ' exceed 32 characters, including the terminating null character.
End Type

Rem This structure contains basic information about a physical font.
Public Type TEXTMETRIC
  tmHeight                As Long ' Specifies the height (ascent descent) of characters.
  tmAscent                As Long ' Specifies the ascent (units above the base line) of characters.
  tmDescent               As Long ' Specifies the descent (units below the base line) of characters.
  tmInternalLeading       As Long ' Specifies the amount of leading (space) inside the bounds set by
                                  ' the tmHeight member. Accent marks and other diacritical characters
                                  ' may occur in this area. The designer may set this member to zero.
  tmExternalLeading       As Long ' Specifies the amount of extra leading (space) that the application adds between rows.
  tmAveCharWidth          As Long ' Specifies the average width of characters in the font
                                  ' (generally defined as the width of the letter x).
  tmMaxCharWidth          As Long ' Specifies the width of the widest character in the font.
  tmWeight                As Long ' Specifies the weight of the font.
  tmOverhang              As Long ' Specifies the extra width per string that may be added to some
                                  ' synthesized fonts. When synthesizing some attributes, such as bold
                                  ' or italic, graphics device interface (GDI) or a device may have to
                                  ' add width to a string on both a per-character and per-string basis.
  tmDigitizedAspectX      As Long ' Specifies the horizontal aspect of the device for which the font was
                                  ' designed.
  tmDigitizedAspectY      As Long ' Specifies the vertical aspect of the device for which the font was
                                  ' designed. The ratio of the tmDigitizedAspectX and tmDigitizedAspectY
                                  ' members is the aspect ratio of the device for which the font was designed.
  tmFirstChar             As Byte ' Specifies the value of the first character defined in the font.
  tmLastChar              As Byte ' Specifies the value of the last character defined in the font.
  tmDefaultChar           As Byte ' Specifies the value of the character to be substituted for characters
                                  ' not in the font.
  tmBreakChar             As Byte ' Specifies the value of the character that will be used to define
                                  ' word breaks for text justification.
  tmItalic                As Byte ' Specifies an italic font if it is nonzero.
  tmUnderlined            As Byte ' Specifies an underlined font if it is nonzero.
  tmStruckOut             As Byte ' Specifies a strikeout font if it is nonzero.
  tmPitchAndFamily        As TEXTMETRIC_PITCH ' Specifies information about the pitch, the technology,
                                  ' and the family of a physical font.
                                  ' Four low-order bits of this member specify information about the pitch and
                                  ' the technology of the font
                                  ' The four high-order bits of tmPitchAndFamily designate the font's font
                                  ' family. An application can use the value 0xF0 and the bitwise AND operator
                                  ' to mask out the four low-order bits of tmPitchAndFamily, thus obtaining a
                                  ' value that can be directly compared with font family names to find an
                                  ' identical match.
  tmCharSet               As LOGICAL_FONT_CHARSET ' Specifies the character set of the font.
End Type

Rem In the TRIVERTEX structure, x and y indicate position in the same manner as in the POINTL structure
Rem contained in the wtypes.h header file. Red, Green, Blue, and Alpha members indicate color
Rem information at the point x, y. The color information of each channel is specified as a value
Rem from 0x0000 to 0xff00. This allows higher color resolution for an object that has been split into
Rem small triangles for display. The TRIVERTEX structure contains information needed by the pVertex
Rem parameter of GradientFill.
Public Type TRIVERTEX   ' /* The TRIVERTEX structure contains color information and
                        '  * position information.
                        '  */
  x           As Long   ' /* Specifies the x-coordinate, in logical units, of the upper-left corner
                        '  * of the rectangle.
                        '  */

  y           As Long   ' /* Specifies the y-coordinate, in logical units, of the upper-left corner
                        '  * of the rectangle.
                        '  */
  Red         As Long   ' /* Indicates color information at the point of x, y. */
  Green       As Long   ' /* Indicates color information at the point of x, y. */
  Blue        As Long   ' /* Indicates color information at the point of x, y. */
  Alpha       As Long   ' /* Indicates color information at the point of x, y. */
End Type

Rem The GRADIENT_RECT structure specifies the values of the pVertex array that are used when the dwMode
Rem parameter of the GradientFill function is GRADIENT_FILL_RECT_H or GRADIENT_FILL_RECT_V. For related
Rem GradientFill structures, see GRADIENT_TRIANGLE and TRIVERTEX.
Public Type GRADIENT_RECT ' /* The GRADIENT_RECT structure specifies the index of two vertices in the
                          '  * pVertex array in the GradientFill function. These two vertices form the
                          '  * upper-left and lower-right boundaries of a rectangle.
                          '  */
  UpperLeft   As stdole.OLE_XPOS_PIXELS ' /* Specifies the upper-left corner of a rectangle. */
  LowerRight  As stdole.OLE_YPOS_PIXELS ' /* Specifies the lower-right corner of a rectangle. */
End Type

Public Type BITMAPINFOHEADER
  biSize          As Long ' /* Specifies the number of bytes required by the structure. */
  biWidth         As stdole.OLE_XSIZE_PIXELS ' /* Specifies the width of the bitmap, in pixels.
                          '  * Windows 98/Me, Windows 2000/XP: If biCompression is BI_JPEG or BI_PNG,
                          '  * the biWidth member specifies the width of the decompressed JPEG or PNG
                          '  * image file, respectively.
                          '  */
  biHeight        As stdole.OLE_YSIZE_PIXELS ' /* Specifies the height of the bitmap, in pixels. If biHeight is positive,
                          '  * the bitmap is a bottom-up DIB and its origin is the lower-left corner.
                          '  * If biHeight is negative, the bitmap is a top-down DIB and its origin is
                          '  * the upper-left corner.
                          '  * If biHeight is negative, indicating a top-down DIB, biCompression must
                          '  * be either BI_RGB or BI_BITFIELDS. Top-down DIBs cannot be compressed.
                          '  * Windows 98/Me, Windows 2000/XP: If biCompression is BI_JPEG or BI_PNG,
                          '  * the biHeight member specifies the height of the decompressed JPEG or
                          '  * PNG image file, respectively.
                          '  */
  biPlanes        As Long ' /* Specifies the number of planes for the target device. This value must be
                          '  * set to 1.
                          '  */
  biBitCount      As Long ' /* Specifies the number of bits-per-pixel. The biBitCount member of the
                          '  * BITMAPINFOHEADER structure determines the number of bits that define
                          '  * each pixel and the maximum number of colors in the bitmap.
                          '  */
  biCompression   As BMPCOMPRESSION ' /* Specifies the type of compression for a compressed bottom-up bitmap
                          '  * (top-down DIBs cannot be compressed).
                          '  */
  biSizeImage     As Long ' /* Specifies the size, in bytes, of the image. This may be set to zero for
                          '  * BI_RGB bitmaps.
                          '  * Windows 98/Me, Windows 2000/XP: If biCompression is BI_JPEG or BI_PNG,
                          '  * biSizeImage indicates the size of the JPEG or PNG image buffer,
                          '  * respectively.
                          '  */
  biXPelsPerMeter As stdole.OLE_XSIZE_PIXELS ' /* Specifies the horizontal resolution, in pixels-per-meter, of the target
                          '  * device for the bitmap. An application can use this value to select a
                          '  * bitmap from a resource group that best matches the characteristics of
                          '  * the current device.
                          '  */
  biYPelsPerMeter As stdole.OLE_YSIZE_PIXELS ' /* Specifies the vertical resolution, in pixels-per-meter, of the target
                          '  * device for the bitmap.
                          '  */
  biClrUsed       As stdole.OLE_COLOR ' /* Specifies the number of color indexes in the color table that are
                          '  * actually used by the bitmap. If this value is zero, the bitmap uses the
                          '  * maximum number of colors corresponding to the value of the biBitCount
                          '  * member for the compression mode specified by biCompression.
                          '  * If biClrUsed is nonzero and the biBitCount member is less than 16, the
                          '  * biClrUsed member specifies the actual number of colors the graphics
                          '  * engine or device driver accesses. If biBitCount is 16 or greater, the
                          '  * biClrUsed member specifies the size of the color table used to optimize
                          '  * performance of the system color palettes. If biBitCount equals 16 or 32,
                          '  * the optimal color palette starts immediately following the three DWORD masks.
                          '  * When the bitmap array immediately follows the BITMAPINFO structure, it
                          '  * is a packed bitmap. Packed bitmaps are referenced by a single pointer.
                          '  * Packed bitmaps require that the biClrUsed member must be either zero or
                          '  * the actual size of the color table.
                          '  */
  biClrImportant  As stdole.OLE_COLOR ' /* Specifies the number of color indexes that are required for displaying
                          '  * the bitmap. If this value is zero, all colors are required.
                          '  */
End Type

Rem /* The CWPSTRUCT structure defines the message parameters passed to a WH_CALLWNDPROC hook procedure,
Rem  * CallWndProc. */
Public Type CWPSTRUCT
  lParam          As Long ' /* Specifies additional information about the message. The exact meaning
                          '  * depends on the message value.
                          '  */
  wParam          As Long ' /* Specifies additional information about the message. The exact meaning
                          '  * depends on the message value.
                          '  */
  message         As WINDOW_MESSAGES ' /* Specifies the message. */
  hWnd            As stdole.OLE_HANDLE ' /* Handle to the window to receive the message. */
End Type

Rem /* The MSG structure contains message information from a thread's message queue. */
Public Type MSG
  hWnd            As stdole.OLE_HANDLE ' /* Handle to the window whose window procedure receives the message. */
  message         As Long ' /* Specifies the message identifier. Applications can only use the low
                          '  * word; the high word is reserved by the system.
                          '  */
  wParam          As Long ' /* Specifies additional information about the message. The exact meaning
                          '  * depends on the value of the message member.
                          '  */
  lParam          As Long ' /* Specifies additional information about the message. The exact meaning
                          '  * depends on the value of the message member.
                          '  */
  time            As Long ' /* Specifies the time at which the message was posted. */
  pt              As POINTAPI ' /* Specifies the cursor position, in screen coordinates, when the
                          '   * message was posted.
                          '  */
End Type

Public Type WNDCLASSEX    ' /* The WNDCLASSEX structure contains window class information. It is
                          '  * used with the RegisterClassEx and GetClassInfoEx functions.
                          '  * The WNDCLASSEX structure is similar to the WNDCLASS structure.
                          '  * There are two differences. WNDCLASSEX includes the cbSize member,
                          '  * which specifies the size of the structure, and the hIconSm member,
                          '  * which contains a handle to a small icon associated with the window
                          '  * class */
  cbSize          As Long ' /* Specifies the size, in bytes, of this structure. Set this member to
                          '  * sizeof(WNDCLASSEX). Be sure to set this member before calling the
                          '  * GetClassInfoEx function.
                          '  */
  style           As Long ' /* Specifies the class style(s). This member can be any combination of
                          '  * the class styles.
                          '  */
  lpfnWndProc     As stdole.OLE_HANDLE ' /* Pointer to the window procedure. */
  cbClsExtra      As Integer ' /* Specifies the number of extra bytes to allocate following the
                          '  * window-class structure. The system initializes the bytes to zero.
                          '  */
  cbWndExtra      As Integer ' /* Specifies the number of extra bytes to allocate following the window
                          '  * instance. The system initializes the bytes to zero. If an
                          '  * application uses WNDCLASSEX to register a dialog box created by using
                          '  * the CLASS directive in the resource file, it must set this member to
                          '  * DLGWINDOWEXTRA.
                          '  */
  hInstance       As stdole.OLE_HANDLE ' /* Handle to the instance that contains the window procedure for the
                          '  * class.
                          '  */
  hIcon           As stdole.OLE_HANDLE ' /* Handle to the class icon. This member must be a handle to an icon
                          '  * resource.
                          '  */
  hCursor         As stdole.OLE_HANDLE ' /* Handle to the class cursor. This member must be a handle to a cursor
                          '  * resource.
                          '  */
  hbrBackground   As stdole.OLE_COLOR ' /* Handle to the class background brush. This member can be a handle to
                          '  * the physical brush to be used for painting the background, or it can
                          '  * be a color value. A color value must be one of the following
                          '  * standard system colors (the value 1 must be added to the chosen
                          '  * color). If a color value is given, you must convert it to one of the
                          '  * following HBRUSH types:
                          '  *  COLOR_ACTIVEBORDER
                          '  *  COLOR_ACTIVECAPTION
                          '  *  COLOR_APPWORKSPACE
                          '  *  COLOR_BACKGROUND
                          '  *  COLOR_BTNFACE
                          '  *  COLOR_BTNSHADOW
                          '  *  COLOR_BTNTEXT
                          '  *  COLOR_CAPTIONTEXT
                          '  *  COLOR_GRAYTEXT
                          '  *  COLOR_HIGHLIGHT
                          '  *  COLOR_HIGHLIGHTTEXT
                          '  *  COLOR_INACTIVEBORDER
                          '  *  COLOR_INACTIVECAPTION
                          '  *  COLOR_MENU
                          '  *  COLOR_MENUTEXT
                          '  *  COLOR_SCROLLBAR
                          '  *  COLOR_WINDOW
                          '  *  COLOR_WINDOWFRAME
                          '  *  COLOR_WINDOWTEXT
                          '  * The system automatically deletes class background brushes when the
                          '  * class is unregistered by using UnregisterClass. An application should
                          '  * not delete these brushes.
                          '  * When this member is NULL, an application must paint its own background
                          '  * whenever it is requested to paint in its client area. To determine
                          '  * whether the background must be painted, an application can either process
                          '  * the WM_ERASEBKGND message or test the fErase member of the PAINTSTRUCT
                          '  * structure filled by the BeginPaint function
                          '  */
  lpszMenuName    As String ' /* Pointer to a null-terminated character string that specifies the
                          '  *  resource name of the class menu, as the name appears in the resource
                          '  *  file.
                          '  */
  lpszClassName   As String ' /* Pointer to a null-terminated string or is an atom. If this
                          '  * parameter is an atom, it must be a class atom created by a previous
                          '  * call to the RegisterClass or RegisterClassEx function. The atom must
                          '  * be in the low-order word of lpszClassName; the high-order word must
                          '  * be zero.
                          '  * If lpszClassName is a string, it specifies the window class name.
                          '  * The class name can be any name registered with RegisterClass or
                          '  * RegisterClassEx, or any of the predefined control-class names.
                          '  */
  hIconSm         As stdole.OLE_HANDLE ' /* Handle to a small icon that is associated with the window class. If
                          '  * this member is NULL, the system searches the icon resource specified
                          '  * by the hIcon member for an icon of the appropriate size to use as
                          '  * the small icon.
                          '  */
End Type

Public Type WINDOWPOS     ' /* The WINDOWPOS structure contains information about the size and
                          '  * position of a window. */
  hWnd            As stdole.OLE_HANDLE ' /* Handle to the window. */
  hwndInsertAfter As stdole.OLE_HANDLE ' /* Specifies the position of the window in Z order (front-to-back
                          '  * position). This member can be a handle to the window behind which
                          '  * this window is placed, or can be one of the special values listed
                          '  * with the SetWindowPos function.
                          '  */
  x               As stdole.OLE_XPOS_PIXELS ' /* Specifies the position of the left edge of the window. */
  y               As stdole.OLE_YPOS_PIXELS ' /* Specifies the position of the top edge of the window. */
  cx              As stdole.OLE_XSIZE_PIXELS ' /* Specifies the window width, in pixels. */
  cy              As stdole.OLE_YSIZE_PIXELS ' /* Specifies the window height, in pixels. */
  flags           As WINDOWPOSFLAGS ' /* Specifies the window position. */
End Type

Public Type NCCALCSIZE_PARAMS ' /* The NCCALCSIZE_PARAMS structure contains information that an
                          '  * application can use while processing the WM_NCCALCSIZE message to
                          '  * calculate the size, position, and valid contents of the client area
                          '  * of a window.
                          '  */
  rgrc(1 To 3)    As RECT
                          ' /* Specifies an array of rectangles. The first contains the new
                          '  * coordinates of a window that has been moved or resized, that is, it
                          '  * is the proposed new window coordinates. The second contains the
                          '  * coordinates of the window before it was moved or resized. The third
                          '  * contains the coordinates of the window's client area before the
                          '  * window was moved or resized. If the window is a child window, the
                          '  * coordinates are relative to the client area of the parent window.
                          '  * If the window is a top-level window, the coordinates are relative to
                          '  * the screen origin.
                          '  */
  lppos           As WINDOWPOS ' /* Pointer to a WINDOWPOS structure that contains the size and
                          '  * position values specified in the operation that moved or resized the
                          '  * window.
                          '  */
End Type

Public Type CREATESTRUCT
  lpCreateParams  As Long ' /* Contains additional data which may be used to create the window. If
                          '  * the window is being created as a result of a call to the
                          '  * CreateWindow or CreateWindowEx function, this member contains the
                          '  * value of the lpParam parameter specified in the function call.
                          '  * If the window being created is an MDI window, this member contains
                          '  * a pointer to an MDICREATESTRUCT structure.
                          '  * Windows NT/2000/XP: If the window is being created from a dialog
                          '  * template, this member is the address of a SHORT value that specifies
                          '  * the size, in bytes, of the window creation data. The value is
                          '  * immediately followed by the creation data.
                          '  */
  hInstance       As stdole.OLE_HANDLE ' /* Handle to the module that owns the new window. */
  hMenu           As stdole.OLE_HANDLE ' /* Handle to the menu to be used by the new window. */
  hWndParent      As stdole.OLE_HANDLE ' /* Handle to the parent window, if the window is a child window. If the
                          '  * window is owned, this member identifies the owner window. If the
                          '  * window is not a child or owned window, this member is NULL.
                          '  */
  cy              As stdole.OLE_YSIZE_PIXELS ' /* Specifies the height of the new window, in pixels. */
  cx              As stdole.OLE_XSIZE_PIXELS ' /* Specifies the width of the new window, in pixels. */
  y               As stdole.OLE_YPOS_PIXELS  ' /* Specifies the y-coordinate of the upper left corner of the new
                              '  * window. If the new window is a child window, coordinates are
                              '  * relative to the parent window. Otherwise, the coordinates are
                              '  * relative to the screen origin.
                              '  */
  x               As stdole.OLE_XPOS_PIXELS ' /* Specifies the x-coordinate of the upper left corner of the new
                              '  * window. If the new window is a child window, coordinates are
                              '  * relative to the parent window. Otherwise, the coordinates are
                              '  * relative to the screen origin.
                              '  */
  style           As Long     ' /* Specifies the style for the new window. */
  lpszName        As String   ' /* Pointer to a null-terminated string that specifies the name of
                              '  * the new window.
                              '  */
  lpszClass       As String   ' /* Pointer to a null-terminated string that specifies the class
                              '  * name of the new window.
                              '  */
  dwExStyle       As Long     ' /* Specifies the extended window style for the new window. */
End Type

Public Type MOUSEHOOKSTRUCT   ' /* The MOUSEHOOKSTRUCT structure contains information about a mouse
                              '  * event passed to a WH_MOUSE hook procedure, MouseProc.
                              '  */
  pt              As POINTAPI ' /* Specifies a POINT structure that contains the x- and
                              '  * y-coordinates of the cursor, in screen coordinates.
                              '  */
  hWnd            As stdole.OLE_HANDLE ' /* Handle to the window that will receive the mouse message
                              '  * corresponding to the mouse event.
                              '  */
  wHitTestCode    As Long     ' /* Specifies the hit-test value. For a list of hit-test values,
                              '  * see the description of the WM_NCHITTEST message.
                              '  */
  dwExtraInfo     As Long     ' /* Specifies extra information associated with the message. */
End Type

Public Type CBT_CREATEWND         ' /* The CBT_CREATEWND structure contains information passed to
                                  '  * a WH_CBT hook procedure, CBTProc, before a window is
                                  '  * created.
                                  '  */
  lpcs            As CREATESTRUCT ' /* Pointer to a CREATESTRUCT structure that contains initialization
                                  '  * parameters for the window about to be created.
                                  '  */
  hwndInsertAfter As stdole.OLE_HANDLE ' /* Handle to the window whose position in the Z order precedes
                                  '  * that of the window being created.
                                  '  */
End Type

Public Type Luid    ' /* An LUID is a 64-bit value guaranteed to be unique only on the system on which
                    '  * it was generated. The uniqueness of a locally unique identifier (LUID) is
                    '  * guaranteed only until the system is restarted.
                    '  * An LUID is not for direct manipulation. Applications are to use functions and
                    '  * structures to manipulate LUID values
                    '  */
  LowPart As Long   ' /* Low order bits. */
  HighPart As Long  ' /* High order bits */
End Type

Public Type LUID_AND_ATTRIBUTES ' /* The LUID_AND_ATTRIBUTES structure represents a locally unique
                                '  * identifier (LUID) and its attributes.
                                '  */
  Luid        As Luid ' /* Specifies an LUID value. */
  Attributes  As PRIVILEGE_ATTR ' /* Specifies attributes of the LUID. This value contains up to 32 one-bit
                      '  * flags. Its meaning is dependent on the definition and use of the LUID.
                      '  */
End Type

Public Type TOKEN_PRIVILEGES    ' /* The TOKEN_PRIVILEGES structure contains information about a set
                                '  * of privileges for an access token.
                                '  */
  PrivilegeCount As Long        ' /* Specifies the number of entries in the Privileges array. */
  Privileges(1 To 1) As LUID_AND_ATTRIBUTES ' /* Specifies an array of LUID_AND_ATTRIBUTES structures.
                                '  * Each structure contains the LUID and attributes of a privilege.
                                '  */
End Type

Public Type PRIVILEGE_SET ' /* The PRIVILEGE_SET structure specifies a set of privileges. It is also
                          '  * used to indicate which, if any, privileges are held by a user or group
                          '  * requesting access to an object.
                          '  */
  PrivilegeCount  As Long ' /* Specifies the number of privileges in the privilege set. */
  Control         As PRIVILEGE_CONTROL_ATTR ' /* Specifies a control flag related to the privileges. The
                          '  * PRIVILEGE_SET_ALL_NECESSARY control flag is currently defined. It
                          '  * indicates that all of the specified privileges must be held by the
                          '  * process requesting access. If this flag is not set, the presence of
                          '  * any privileges in the user's access token grants the access.
                          '  */
  Privilege(1 To 1) As LUID_AND_ATTRIBUTES ' /* Specifies an array of LUID_AND_ATTRIBUTES structures
                          '  * describing the set's privileges.
                          '  */

  Rem /* Remarks
  Rem  * A privilege is used to control access to an object or service more strictly than is typical
  Rem  * with discretionary access control. A system manager uses privileges to control which users
  Rem  * are able to manipulate system resources. An application uses privileges when it changes a
  Rem  * system-wide resource, such as when it changes the system time or shuts down the system.
  Rem  */
End Type


Public Function findWindowHandle(ByVal title As String) As Long
    findWindowHandle = FindWindowEx(0, 0, "", title)
End Function


Rem /* This function simply masks the portion of the Long integer it wants to return using the
Rem  * bitwise "And" operator.
Rem  * Visual Basic Integers are signed, any low word value that has its high bit set must be
Rem  * converted back into a negative value using the "Or" operator and a mask of &HFFFF0000.
Rem */
Public Function LoWord(ByRef DWORD As Long) As Integer
  Rem /* Mask out the 32bit i.e.32764 Integer value. If there is any Integer 32bit value present
  Rem  * then return the upper 32 bit portion else return the lower portion.
  Rem  */
  If (DWORD And &H8000&) Then       ' &H8000& = &H00008000
    Let LoWord = DWORD Or &HFFFF0000
  Else
    Let LoWord = DWORD And &HFFFF&
  End If
End Function

Rem /* This function simply masks the portion of the Long integer it wants to return using the
Rem  * bitwise "And" operator.
Rem  * The HiWord function shifts this value right 16-bits by dividing it by &H10000.
Rem */
Public Function HiWord(ByRef DWORD As Long) As Integer
  Let HiWord = ((DWORD And &HFFFF0000) \ &H10000)
End Function

Rem /* The GetRValue function returns the R-component of an RGB value. */
Public Function GetRValue(ByVal color As Long) As Integer
  Let GetRValue = (color \ (&H100 ^ 0) And &HFF)
End Function

Rem /* The GetGValue function returns the G-component of an RGB value. */
Public Function GetGValue(ByVal color As Long) As Integer
  Let GetGValue = (color \ (&H100 ^ 1) And &HFF)
End Function

Rem /* The GetBValue function returns the B-component of an RGB value. */
Public Function GetBValue(ByVal color As Long) As Integer
  Let GetBValue = (color \ (&H100 ^ 2) And &HFF)
End Function

Public Function GetClassNameEx(ByVal hWnd As Long) As String
  Const buffer               As Long = &H80&      ' 128
  Dim classBuffer            As String
  Let classBuffer = VBA.Strings.String$(buffer, VBA.Constants.vbNullChar)
  Call GetClassName(hWnd, classBuffer, buffer)
  Let GetClassNameEx = ExtractString(classBuffer)
End Function

Public Function getGlobalAccleratorKey(ByVal InString As String) As String
  Dim mGetGlobalAccleratorKey             As String: Let mGetGlobalAccleratorKey = VBA.Constants.vbNullString
  
  Rem /* Menu global accelerator keys are shortcut keys like CTRL + A, F5 etc. Normally these
  Rem  * are seperated from the menu caption with a tab character. So our first task is to find
  Rem  * the tab position. Next the string will be null terminated. So the second task is to find
  Rem  * the null char position. That's that. Once these two positions are known, the string in these
  Rem  * locations is the accelerator key string.
  Rem  */
  Dim tabStart                          As Long: Let tabStart = &H0&
  Dim tabEnd                            As Long: Let tabEnd = &H0&
  Rem /* Get the positions. */
  Let tabStart = VBA.Strings.InStr(1, InString, VBA.Constants.vbTab, vbTextCompare)
  Let tabEnd = VBA.Strings.InStr(1, InString, VBA.Constants.vbNullChar, vbTextCompare)
  Rem /* Now move the positions to their exact places, so that no extra characters are retrieved. This
  Rem  * the position is returned 'where' the character occurs. So we must move to the next location to
  Rem  * to indicate that we want to extract a substring from the next character.
  Rem  */
  Let tabStart = (tabStart + 1)         ' /* Move a place forward. */
  Let tabEnd = tabEnd                   ' /* No need. This place is correct. */
  Rem /* Extract the string we are looking forward to. */
  Let mGetGlobalAccleratorKey = VBA.Strings.Mid$(String:=InString, _
                                                 start:=tabStart, _
                                                 Length:=(tabEnd - tabStart))
  Rem /* Thus we retrieve the global accelerator key. */
  Let getGlobalAccleratorKey = mGetGlobalAccleratorKey
End Function


Rem For some API functions, such as SendMessage or PostMessage, you may need to package two short
Rem Integer values into a Long variable to pass them as a single parameter.
Rem The trick to packing values is bit shifting. Because Visual Basic does not provide bit shift
Rem operators to use, you need to do things the old fashioned way; through multiplication. To make
Rem an Integer the high word for a Long value, you need to multiply it by &H10000. This has the effect
Rem of shifting the bit values 16-bits (2-bytes) to the left, making room for the low word value you
Rem want to add.
Rem Before you can add the low word value, however, you need to make an adjustment. Remember that Visual
Rem Basic Integer types are signed values, but the low word value needs to be unsigned if you plan to add
Rem it to your high word value. To make sure Visual Basic treats the low word as an unsigned integer, you
Rem need to perform a bitwise "And" on the value using &HFFFF& as a mask. In effect, this saves the value
Rem as a Long integer with the high (signed) bit cleared but keeps the original Integer's bit value preserved.
Public Property Get MAKELONG(ByVal HiWord As Integer, _
                             ByVal LoWord As Integer) _
As Long
  Let MAKELONG = (HiWord * DWORD.HIWORD_MASK) Or (LoWord And DWORD.LOWORD_MASK)
End Property

Public Function ExtractString(ByVal SearchString As String, _
                              Optional ByVal StringToSearch As String = VBA.Constants.vbNullChar) As String
  Dim extractLength           As Long: Let extractLength = VBA.Strings.InStr(start:=1, _
                                                                             string1:=SearchString, _
                                                                             string2:=StringToSearch, _
                                                                             Compare:=vbTextCompare)
  
  Rem Get rid of the very last character...
  If (extractLength >= 1) Then Let extractLength = (extractLength - 1)
  
  Let ExtractString = VBA.Strings.Left$(String:=SearchString, _
                                        Length:=extractLength)
  
  Rem Trim off any spaces present around the string...
  Let ExtractString = VBA.Strings.Trim$(ExtractString)
End Function
    
'Retrieve current windows font used by the system
'gets stock GUI
'Based on VBTK: p138
Public Function getDefaultFont() As StdFont
    On Error Resume Next
    Dim guiFont As Long, oldFont As Long
    Dim ret As Long, hDC As Long
    Dim TYP_METRICS As TEXTMETRIC
    Dim fontFaceName As String
    Dim hWnd As Long
    Dim dwExStyle, lpClassName, lpWindowName
    Dim dwStyle, x, y, nWidth, nHeight, hWndParent, hMenu
    Dim hInstance, lpParam
    'instantiate font
    Set getDefaultFont = New StdFont
    'create win
    hWnd = CreateWindowEx(dwExStyle, "STATIC", "localizer_win", _
            dwStyle, x, y, nWidth, nHeight, hWndParent, hMenu, _
            hInstance, lpParam)
    'get dc of new win
    hDC = GetDC(hWnd)
    'get font handle for DEFAULT_GUI_FONT
    guiFont = GetStockObject(DEFAULT_GUI_FONT)
    oldFont = SelectObject(hDC, guiFont)
    fontFaceName = Space$(255)
    ret = GetTextFace(hDC, 255, fontFaceName)
    fontFaceName = Left$(fontFaceName, InStr(fontFaceName, Chr$(0)) - 1)
    'get font metrics
    ret = GetTextMetrics(hDC, TYP_METRICS)
    'assign font face name
    getDefaultFont.Name = fontFaceName
    'tmInternalLeading is used to reduce the cell sie to point size.
    getDefaultFont.Size = ((TYP_METRICS.tmHeight - _
        TYP_METRICS.tmInternalLeading) * 72) / GetDeviceCaps(hDC, LOGPIXELSY)
    ret = SelectObject(hDC, oldFont)
    ret = ReleaseDC(hWnd, hDC)
    DestroyWindow hWnd
    
End Function

'given a GUI controls, change the font to match the default font of this
'window. Useful for internationalizing.
'based on VBTK: p 139
Public Sub setDefaultFont(obj As Object)
    On Local Error Resume Next
    Dim ctl As Control
    Dim f As New StdFont
    Set f = getDefaultFont()
    For Each ctl In obj.Controls
        ctl.Font.Name = f.Name
        ctl.Font.Size = f.Size
        ctl.Font.Charset = f.Charset
        ctl.Font.Weight = f.Weight
        ctl.Font.Bold = f.Bold
        ctl.Font.Italic = f.Italic
        ctl.Font.Strikethrough = f.Strikethrough
        ctl.Font.Underline = f.Underline
    Next
    Set f = Nothing
End Sub

'get username in network sense.
Public Function getNetUserName() As String
    Dim i As Long
    Dim username As String
    username = Space$(255)
    i = WNetGetUser("", username, Len(username))
    i = InStr(username, Chr$(0))
    If i > 0 Then getNetUserName = Left$(username, i - 1)
    
End Function
