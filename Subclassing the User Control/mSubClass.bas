Attribute VB_Name = "mSubClass"
'===========================================================================
'
' Module Name:  mSubClass
' Author:       Slider
' Date:         11/11/2001
' Version:      01.00.00
' Description:  Windows message handler.
' Edit History: 01.00.00 11/11/01 Initial Release
'
'===========================================================================

Option Explicit

Private moSubClass As cSubClass

Public defWindowProc As Long

Public Const GWL_WNDPROC As Long = (-4)

Public Const WM_NULL = &H0
Public Const WM_CREATE = &H1
Public Const WM_DESTROY = &H2
Public Const WM_MOVE = &H3
Public Const WM_SIZE = &H5
Public Const WM_ACTIVATE = &H6
Public Const WM_SETFOCUS = &H7
Public Const WM_KILLFOCUS = &H8
Public Const WM_ENABLE = &HA
Public Const WM_SETREDRAW = &HB
Public Const WM_SETTEXT = &HC
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_PAINT = &HF
Public Const WM_CLOSE = &H10
Public Const WM_QUERYENDSESSION = &H11
Public Const WM_QUIT = &H12
Public Const WM_QUERYOPEN = &H13
Public Const WM_ERASEBKGND = &H14
Public Const WM_SYSCOLORCHANGE = &H15
Public Const WM_ENDSESSION = &H16
Public Const WM_SHOWWINDOW = &H18
Public Const WM_SETTINGCHANGE = &H1A
Public Const WM_DEVMODECHANGE = &H1B
Public Const WM_ACTIVATEAPP = &H1C
Public Const WM_FONTCHANGE = &H1D
Public Const WM_TIMECHANGE = &H1E
Public Const WM_CANCELMODE = &H1F
Public Const WM_SETCURSOR = &H20
Public Const WM_MOUSEACTIVATE = &H21
Public Const WM_CHILDACTIVATE = &H22
Public Const WM_QUEUESYNC = &H23
Public Const WM_GETMINMAXINFO = &H24
Public Const WM_PAINTICON = &H26
Public Const WM_ICONERASEBKGND = &H27
Public Const WM_NEXTDLGCTL = &H28
Public Const WM_SPOOLERSTATUS = &H2A
Public Const WM_DRAWITEM = &H2B
Public Const WM_MEASUREITEM = &H2C
Public Const WM_DELETEITEM = &H2D
Public Const WM_VKEYTOITEM = &H2E
Public Const WM_CHARTOITEM = &H2F
Public Const WM_SETFONT = &H30
Public Const WM_GETFONT = &H31
Public Const WM_SETHOTKEY = &H32
Public Const WM_GETHOTKEY = &H33
Public Const WM_QUERYDRAGICON = &H37
Public Const WM_COMPAREITEM = &H39
Public Const WM_COMPACTING = &H41
Public Const WM_WINDOWPOSCHANGING = &H46
Public Const WM_WINDOWPOSCHANGED = &H47
Public Const WM_POWER = &H48
Public Const WM_COPYDATA = &H4A
Public Const WM_CANCELJOURNAL = &H4B
Public Const WM_NCCREATE = &H81
Public Const WM_NCDESTROY = &H82
Public Const WM_NCCALCSIZE = &H83
Public Const WM_NCHITTEST = &H84
Public Const WM_NCPAINT = &H85
Public Const WM_NCACTIVATE = &H86
Public Const WM_GETDLGCODE = &H87
Public Const WM_NCMOUSEMOVE = &HA0
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const WM_NCLBUTTONUP = &HA2
Public Const WM_NCLBUTTONDBLCLK = &HA3
Public Const WM_NCRBUTTONDOWN = &HA4
Public Const WM_NCRBUTTONUP = &HA5
Public Const WM_NCRBUTTONDBLCLK = &HA6
Public Const WM_NCMBUTTONDOWN = &HA7
Public Const WM_NCMBUTTONUP = &HA8
Public Const WM_NCMBUTTONDBLCLK = &HA9
Public Const WM_KEYFIRST = &H100
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_CHAR = &H102
Public Const WM_DEADCHAR = &H103
Public Const WM_SYSKEYDOWN = &H104
Public Const WM_SYSKEYUP = &H105
Public Const WM_SYSCHAR = &H106
Public Const WM_SYSDEADCHAR = &H107
Public Const WM_KEYLAST = &H108
Public Const WM_INITDIALOG = &H110
Public Const WM_COMMAND = &H111
Public Const WM_SYSCOMMAND = &H112
Public Const WM_TIMER = &H113
Public Const WM_HSCROLL = &H114
Public Const WM_VSCROLL = &H115
Public Const WM_INITMENU = &H116
Public Const WM_INITMENUPOPUP = &H117
Public Const WM_MENUSELECT = &H11F
Public Const WM_MENUCHAR = &H120
Public Const WM_ENTERIDLE = &H121
Public Const WM_CTLCOLORMSGBOX = &H132
Public Const WM_CTLCOLOREDIT = &H133
Public Const WM_CTLCOLORLISTBOX = &H134
Public Const WM_CTLCOLORBTN = &H135
Public Const WM_CTLCOLORDLG = &H136
Public Const WM_CTLCOLORSCROLLBAR = &H137
Public Const WM_CTLCOLORSTATIC = &H138
Public Const WM_MOUSEFIRST = &H200
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_MBUTTONDOWN = &H207
Public Const WM_MBUTTONUP = &H208
Public Const WM_MBUTTONDBLCLK = &H209
Public Const WM_MOUSELAST = &H209
Public Const WM_PARENTNOTIFY = &H210
Public Const WM_ENTERMENULOOP = &H211
Public Const WM_EXITMENULOOP = &H212
Public Const WM_MDICREATE = &H220
Public Const WM_MDIDESTROY = &H221
Public Const WM_MDIACTIVATE = &H222
Public Const WM_MDIRESTORE = &H223
Public Const WM_MDINEXT = &H224
Public Const WM_MDIMAXIMIZE = &H225
Public Const WM_MDITILE = &H226
Public Const WM_MDICASCADE = &H227
Public Const WM_MDIICONARRANGE = &H228
Public Const WM_MDIGETACTIVE = &H229
Public Const WM_MDISETMENU = &H230
Public Const WM_DROPFILES = &H233
Public Const WM_MDIREFRESHMENU = &H234
Public Const WM_CUT = &H300
Public Const WM_COPY = &H301
Public Const WM_PASTE = &H302
Public Const WM_CLEAR = &H303
Public Const WM_UNDO = &H304
Public Const WM_RENDERFORMAT = &H305
Public Const WM_RENDERALLFORMATS = &H306
Public Const WM_DESTROYCLIPBOARD = &H307
Public Const WM_DRAWCLIPBOARD = &H308
Public Const WM_PAINTCLIPBOARD = &H309
Public Const WM_VSCROLLCLIPBOARD = &H30A
Public Const WM_SIZECLIPBOARD = &H30B
Public Const WM_ASKCBFORMATNAME = &H30C
Public Const WM_CHANGECBCHAIN = &H30D
Public Const WM_HSCROLLCLIPBOARD = &H30E
Public Const WM_QUERYNEWPALETTE = &H30F
Public Const WM_PALETTEISCHANGING = &H310
Public Const WM_PALETTECHANGED = &H311
Public Const WM_HOTKEY = &H312
Public Const WM_PENWINFIRST = &H380
Public Const WM_PENWINLAST = &H38F
Public Const WM_USER = &H400

Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, _
                                                                           ByVal nIndex As Long, _
                                                                           ByVal dwNewLong As Long) As Long

Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, _
                                                                             ByVal hWnd As Long, _
                                                                             ByVal uMsg As Long, _
                                                                             ByVal wParam As Long, _
                                                                             ByVal lParam As Long) As Long

'===========================================================================
'## Public Detecting and Notifying Changes to the System Date/Time
'
Public Sub Hook(oSubClass As cSubClass)
    '
    '## assign window message procedure (WindowProc)
    '
    On Error Resume Next
    Set moSubClass = oSubClass
    defWindowProc = SetWindowLong(moSubClass.hWnd, GWL_WNDPROC, AddressOf WindowProc)
   
End Sub

'===========================================================================
' UnSubClass - must be called when terminating UserControl
'
Public Sub UnHook()
    '
    '## Restore the default message handling before exiting
    '
    If defWindowProc Then
        SetWindowLong moSubClass.hWnd, GWL_WNDPROC, defWindowProc
        defWindowProc = 0
    End If
    Set moSubClass = Nothing    '## Release UserControl class reference

End Sub

'===========================================================================
' Windows Message Proceedure
'
Public Function WindowProc(ByVal hWnd As Long, _
                           ByVal uMsg As Long, _
                           ByVal wParam As Long, _
                           ByVal lParam As Long) As Long

    '!! DON'T PAUSE THE PROGRAM IN HERE!

    On Error Resume Next

    Debug.Print "WP:Hwnd = "; CStr(hWnd); _
                "  SC:Hwnd = "; CStr(moSubClass.hWnd); _
                "  uMsg = &h"; Hex$(uMsg)

    Select Case hWnd

        '## If the handle returned is to our User Control, perform form-specific message
        '   handling to deal with the notifications.
        Case moSubClass.hWnd

            '## UserControl & Parent Form-specific handler
            Select Case uMsg
                '!! NOTE: These messages require parameters
                Case WM_NCHITTEST
                Case Else
                    moSubClass.fRaiseEvent uMsg, wParam, lParam
                    '## Message processed so return 0
                    WindowProc = 0
            End Select

            Select Case uMsg
                '!! NOTE: These messages require parameters.
                Case WM_PAINT, WM_NCPAINT, WM_NCCALCSIZE, WM_GETTEXT, WM_ERASEBKGND, _
                     WM_SETCURSOR, WM_NCHITTEST, WM_MOUSEFIRST, WM_MOUSEMOVE, WM_MOUSEACTIVATE, _
                     WM_NCMOUSEMOVE, WM_MOUSELAST, WM_COMMAND, WM_NCLBUTTONDOWN, WM_NCLBUTTONUP, _
                     WM_NCLBUTTONDBLCLK, WM_QUERYOPEN, WM_SYSCOMMAND, WM_WINDOWPOSCHANGED, _
                     WM_SIZE, WM_NCACTIVATE, WM_ACTIVATE

                            '## This takes care of messages when the handle specified hasn't changed.
                            WindowProc = CallWindowProc(defWindowProc, hWnd, uMsg, wParam, lParam)
                            Exit Function
            End Select

        Case Else
            '## This takes care of messages when the handle specified is not that of the
            '   UserControl.
            WindowProc = CallWindowProc(defWindowProc, hWnd, uMsg, wParam, lParam)

    End Select

End Function
