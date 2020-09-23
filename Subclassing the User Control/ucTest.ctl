VERSION 5.00
Begin VB.UserControl ucTest 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1050
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2610
   FillColor       =   &H000000C0&
   KeyPreview      =   -1  'True
   ScaleHeight     =   1050
   ScaleWidth      =   2610
   Begin VB.ListBox lstEvents 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   2430
   End
End
Attribute VB_Name = "ucTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'===========================================================================
'
' UserControl:  UCTest
' Author:       Slider
' Date:         11/11/2001
' Version:      01.00.00
' Description:  Test UserControl used as a learning tool to understand how
'               to use subclassing with a User Control.
' Edit History: 01.00.00 11/11/01 Initial Release
'
'===========================================================================

Option Explicit

Private WithEvents moSubClass As cSubClass
Attribute moSubClass.VB_VarHelpID = -1

'===========================================================================
'## Widows Message Events

Private Sub moSubClass_WinMessage(ByVal EventID As Long, ByVal wParam As Long, ByVal lParam As Long)

    '## NOTE: hWnd = moSubClass.hWnd

    Select Case EventID
        Case WM_NULL: pAddItem "WM_NULL"
        Case WM_CREATE: pAddItem "WM_CREATE"
        Case WM_DESTROY: pAddItem "WM_DESTROY"
        Case WM_MOVE: pAddItem "WM_MOVE"
        Case WM_SIZE: pAddItem "WM_SIZE"
        Case WM_ACTIVATE: pAddItem "WM_ACTIVATE"
        Case WM_SETFOCUS: pAddItem "WM_SETFOCUS"
        Case WM_KILLFOCUS: pAddItem "WM_KILLFOCUS"
        Case WM_ENABLE: pAddItem "WM_ENABLE"
        Case WM_SETREDRAW: pAddItem "WM_SETREDRAW"
        Case WM_SETTEXT: pAddItem "WM_SETTEXT"
        Case WM_GETTEXT: pAddItem "WM_GETTEXT"
        Case WM_GETTEXTLENGTH: pAddItem "WM_GETTEXTLENGTH"
        Case WM_PAINT: pAddItem "WM_PAINT"
        Case WM_CLOSE: pAddItem "WM_CLOSE"
        Case WM_QUERYENDSESSION: pAddItem "WM_QUERYENDSESSION"
        Case WM_QUIT: pAddItem "WM_QUIT"
        Case WM_QUERYOPEN: pAddItem "WM_QUERYOPEN"
        Case WM_ERASEBKGND: pAddItem "WM_ERASEBKGND"
        Case WM_SYSCOLORCHANGE: pAddItem "*** WM_SYSCOLORCHANGE ***"
        Case WM_ENDSESSION: pAddItem "WM_ENDSESSION"
        Case WM_SHOWWINDOW: pAddItem "WM_SHOWWINDOW"
        Case WM_SETTINGCHANGE: pAddItem "*** WM_SETTINGCHANGE ***"
        Case WM_DEVMODECHANGE: pAddItem "WM_DEVMODECHANGE"
        Case WM_ACTIVATEAPP: pAddItem "WM_ACTIVATEAPP"
        Case WM_FONTCHANGE: pAddItem "WM_FONTCHANGE"
        Case WM_TIMECHANGE: pAddItem "*** WM_TIMECHANGE (& Date) ***"
        Case WM_CANCELMODE: pAddItem "WM_CANCELMODE"
        Case WM_SETCURSOR: pAddItem "WM_SETCURSOR"
        Case WM_MOUSEACTIVATE: pAddItem "WM_MOUSEACTIVATE"
        Case WM_CHILDACTIVATE: pAddItem "WM_CHILDACTIVATE"
        Case WM_QUEUESYNC: pAddItem "WM_QUEUESYNC"
        Case WM_GETMINMAXINFO: pAddItem "WM_GETMINMAXINFO"
        Case WM_PAINTICON: pAddItem "WM_PAINTICON"
        Case WM_ICONERASEBKGND: pAddItem "WM_ICONERASEBKGND"
        Case WM_NEXTDLGCTL: pAddItem "WM_NEXTDLGCTL"
        Case WM_SPOOLERSTATUS: pAddItem "WM_SPOOLERSTATUS"
        Case WM_DRAWITEM: pAddItem "WM_DRAWITEM"
        Case WM_MEASUREITEM: pAddItem "WM_MEASUREITEM"
        Case WM_DELETEITEM: pAddItem "WM_DELETEITEM"
        Case WM_VKEYTOITEM: pAddItem "WM_VKEYTOITEM"
        Case WM_CHARTOITEM: pAddItem "WM_CHARTOITEM"
        Case WM_SETFONT: pAddItem "WM_SETFONT"
        Case WM_GETFONT: pAddItem "WM_GETFONT"
        Case WM_SETHOTKEY: pAddItem "WM_SETHOTKEY"
        Case WM_GETHOTKEY: pAddItem "WM_GETHOTKEY"
        Case WM_QUERYDRAGICON: pAddItem "WM_QUERYDRAGICON"
        Case WM_COMPAREITEM: pAddItem "WM_COMPAREITEM"
        Case WM_COMPACTING: pAddItem "WM_COMPACTING"
        Case WM_WINDOWPOSCHANGING: pAddItem "WM_WINDOWPOSCHANGING"
        Case WM_WINDOWPOSCHANGED: pAddItem "WM_WINDOWPOSCHANGED"
        Case WM_POWER: pAddItem "WM_POWER"
        Case WM_COPYDATA: pAddItem "WM_COPYDATA"
        Case WM_CANCELJOURNAL: pAddItem "WM_CANCELJOURNAL"
        Case WM_NCCREATE: pAddItem "WM_NCCREATE"
        Case WM_NCDESTROY: pAddItem "WM_NCDESTROY"
        Case WM_NCCALCSIZE: pAddItem "WM_NCCALCSIZE"
        Case WM_NCHITTEST: pAddItem "WM_NCHITTEST"
        Case WM_NCPAINT: pAddItem "WM_NCPAINT"
        Case WM_NCACTIVATE: pAddItem "WM_NCACTIVATE"
        Case WM_GETDLGCODE: pAddItem "WM_GETDLGCODE"
        Case WM_NCMOUSEMOVE: pAddItem "WM_NCMOUSEMOVE"
        Case WM_NCLBUTTONDOWN: pAddItem "WM_NCLBUTTONDOWN"
        Case WM_NCLBUTTONUP: pAddItem "WM_NCLBUTTONUP"
        Case WM_NCLBUTTONDBLCLK: pAddItem "WM_NCLBUTTONDBLCLK"
        Case WM_NCRBUTTONDOWN: pAddItem "WM_NCRBUTTONDOWN"
        Case WM_NCRBUTTONUP: pAddItem "WM_NCRBUTTONUP"
        Case WM_NCRBUTTONDBLCLK: pAddItem "WM_NCRBUTTONDBLCLK"
        Case WM_NCMBUTTONDOWN: pAddItem "WM_NCMBUTTONDOWN"
        Case WM_NCMBUTTONUP: pAddItem "WM_NCMBUTTONUP"
        Case WM_NCMBUTTONDBLCLK: pAddItem "WM_NCMBUTTONDBLCLK"
        Case WM_KEYFIRST: pAddItem "WM_KEYFIRST"
        Case WM_KEYDOWN: pAddItem "WM_KEYDOWN"
        Case WM_KEYUP: pAddItem "WM_KEYUP"
        Case WM_CHAR: pAddItem "WM_CHAR"
        Case WM_DEADCHAR: pAddItem "WM_DEADCHAR"
        Case WM_SYSKEYDOWN: pAddItem "WM_SYSKEYDOWN"
        Case WM_SYSKEYUP: pAddItem "WM_SYSKEYUP"
        Case WM_SYSCHAR: pAddItem "WM_SYSCHAR"
        Case WM_SYSDEADCHAR: pAddItem "WM_SYSDEADCHAR"
        Case WM_KEYLAST: pAddItem "WM_KEYLAST"
        Case WM_INITDIALOG: pAddItem "WM_INITDIALOG"
        Case WM_COMMAND: pAddItem "WM_COMMAND"
        Case WM_SYSCOMMAND: pAddItem "WM_SYSCOMMAND"
        Case WM_TIMER: pAddItem "WM_TIMER"
        Case WM_HSCROLL: pAddItem "WM_HSCROLL"
        Case WM_VSCROLL: pAddItem "WM_VSCROLL"
        Case WM_INITMENU: pAddItem "WM_INITMENU"
        Case WM_INITMENUPOPUP: pAddItem "WM_INITMENUPOPUP"
        Case WM_MENUSELECT: pAddItem "WM_MENUSELECT"
        Case WM_MENUCHAR: pAddItem "WM_MENUCHAR"
        Case WM_ENTERIDLE: pAddItem "WM_ENTERIDLE"
        Case WM_CTLCOLORMSGBOX: pAddItem "WM_CTLCOLORMSGBOX"
        Case WM_CTLCOLOREDIT: pAddItem "WM_CTLCOLOREDIT"
        Case WM_CTLCOLORLISTBOX: pAddItem "WM_CTLCOLORLISTBOX"
        Case WM_CTLCOLORBTN: pAddItem "WM_CTLCOLORBTN"
        Case WM_CTLCOLORDLG: pAddItem "WM_CTLCOLORDLG"
        Case WM_CTLCOLORSCROLLBAR: pAddItem "WM_CTLCOLORSCROLLBAR"
        Case WM_CTLCOLORSTATIC: pAddItem "WM_CTLCOLORSTATIC"
        Case WM_MOUSEFIRST: pAddItem "WM_MOUSEFIRST"
        Case WM_MOUSEMOVE: pAddItem "WM_MOUSEMOVE"
        Case WM_LBUTTONDOWN: pAddItem "WM_LBUTTONDOWN"
        Case WM_LBUTTONUP: pAddItem "WM_LBUTTONUP"
        Case WM_LBUTTONDBLCLK: pAddItem "WM_LBUTTONDBLCLK"
        Case WM_RBUTTONDOWN: pAddItem "WM_RBUTTONDOWN"
        Case WM_RBUTTONUP: pAddItem "WM_RBUTTONUP"
        Case WM_RBUTTONDBLCLK: pAddItem "WM_RBUTTONDBLCLK"
        Case WM_MBUTTONDOWN: pAddItem "WM_MBUTTONDOWN"
        Case WM_MBUTTONUP: pAddItem "WM_MBUTTONUP"
        Case WM_MBUTTONDBLCLK: pAddItem "WM_MBUTTONDBLCLK"
        Case WM_MOUSELAST: pAddItem "WM_MOUSELAST"
        Case WM_PARENTNOTIFY: pAddItem "WM_PARENTNOTIFY"
        Case WM_ENTERMENULOOP: pAddItem "WM_ENTERMENULOOP"
        Case WM_EXITMENULOOP: pAddItem "WM_EXITMENULOOP"
        Case WM_MDICREATE: pAddItem "WM_MDICREATE"
        Case WM_MDIDESTROY: pAddItem "WM_MDIDESTROY"
        Case WM_MDIACTIVATE: pAddItem "WM_MDIACTIVATE"
        Case WM_MDIRESTORE: pAddItem "WM_MDIRESTORE"
        Case WM_MDINEXT: pAddItem "WM_MDINEXT"
        Case WM_MDIMAXIMIZE: pAddItem "WM_MDIMAXIMIZE"
        Case WM_MDITILE: pAddItem "WM_MDITILE"
        Case WM_MDICASCADE: pAddItem "WM_MDICASCADE"
        Case WM_MDIICONARRANGE: pAddItem "WM_MDIICONARRANGE"
        Case WM_MDIGETACTIVE: pAddItem "WM_MDIGETACTIVE"
        Case WM_MDISETMENU: pAddItem "WM_MDISETMENU"
        Case WM_DROPFILES: pAddItem "WM_DROPFILES"
        Case WM_MDIREFRESHMENU: pAddItem "WM_MDIREFRESHMENU"
        Case WM_CUT: pAddItem "WM_CUT"
        Case WM_COPY: pAddItem "WM_COPY"
        Case WM_PASTE: pAddItem "WM_PASTE"
        Case WM_CLEAR: pAddItem "WM_CLEAR"
        Case WM_UNDO: pAddItem "WM_UNDO"
        Case WM_RENDERFORMAT: pAddItem "WM_RENDERFORMAT"
        Case WM_RENDERALLFORMATS: pAddItem "WM_RENDERALLFORMATS"
        Case WM_DESTROYCLIPBOARD: pAddItem "WM_DESTROYCLIPBOARD"
        Case WM_DRAWCLIPBOARD: pAddItem "WM_DRAWCLIPBOARD"
        Case WM_PAINTCLIPBOARD: pAddItem "WM_PAINTCLIPBOARD"
        Case WM_VSCROLLCLIPBOARD: pAddItem "WM_VSCROLLCLIPBOARD"
        Case WM_SIZECLIPBOARD: pAddItem "WM_SIZECLIPBOARD"
        Case WM_ASKCBFORMATNAME: pAddItem "WM_ASKCBFORMATNAME"
        Case WM_CHANGECBCHAIN: pAddItem "WM_CHANGECBCHAIN"
        Case WM_HSCROLLCLIPBOARD: pAddItem "WM_HSCROLLCLIPBOARD"
        Case WM_QUERYNEWPALETTE: pAddItem "WM_QUERYNEWPALETTE"
        Case WM_PALETTEISCHANGING: pAddItem "WM_PALETTEISCHANGING"
        Case WM_PALETTECHANGED: pAddItem "WM_PALETTECHANGED"
        Case WM_HOTKEY: pAddItem "WM_HOTKEY"
        Case WM_PENWINFIRST: pAddItem "WM_PENWINFIRST"
        Case WM_PENWINLAST: pAddItem "WM_PENWINLAST"
        Case WM_USER: pAddItem "WM_USER"
    End Select

End Sub

'===========================================================================
'## Private Control Events
'

'===========================================================================
'## Private Subroutines and Functions
'
Private Sub pAddItem(ItemText As String)
    '## Display Windows & User Control messages
    With lstEvents
        .AddItem ItemText
        .ListIndex = .ListCount - 1
    End With
    Debug.Print "Message = " + ItemText
End Sub

'========================================================================================
'## Private User Control Events
'
Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
    pAddItem ">>> UserControl_AccessKeyPress"
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
    pAddItem ">>> UserControl_AmbientChanged"
End Sub

Private Sub UserControl_AsyncReadComplete(AsyncProp As AsyncProperty)
    pAddItem ">>> UserControl_AsyncReadComplete"
End Sub

Private Sub UserControl_AsyncReadProgress(AsyncProp As AsyncProperty)
    pAddItem ">>> UserControl_AsyncReadProgress"
End Sub

Private Sub UserControl_Click()
    pAddItem ">>> UserControl_Click"
End Sub

Private Sub UserControl_DblClick()
    pAddItem ">>> UserControl_DblClick"
End Sub

Private Sub UserControl_EnterFocus()
    pAddItem ">>> UserControl_EnterFocus"
End Sub

Private Sub UserControl_ExitFocus()
    pAddItem ">>> UserControl_ExitFocus"
End Sub

Private Sub UserControl_GotFocus()
    pAddItem ">>> UserControl_GotFocus"
End Sub

Private Sub UserControl_Hide()
    '## Disengage subclassing
    pAddItem ">>> UserControl_Hide*"
    If Ambient.UserMode = False Then
        Exit Sub    '!! IDE MODE
    End If
    mSubClass.UnHook
    Set moSubClass = Nothing
End Sub

Private Sub UserControl_Initialize()
    pAddItem ">>> UserControl_Initialize"
End Sub

Private Sub UserControl_InitProperties()
    pAddItem ">>> UserControl_InitProperties"
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    pAddItem ">>> UserControl_KeyDown"
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    pAddItem ">>> UserControl_KeyPress"
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    pAddItem ">>> UserControl_KeyUp"
End Sub

Private Sub UserControl_LostFocus()
    pAddItem ">>> UserControl_LostFocus"
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    pAddItem ">>> UserControl_MouseDown"
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    pAddItem ">>> UserControl_MouseMove"
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    pAddItem ">>> UserControl_MouseUp"
End Sub

Private Sub UserControl_Paint()
    pAddItem ">>> UserControl_Paint"
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    pAddItem ">>> UserControl_ReadProperties"
End Sub

Private Sub UserControl_Resize()
    With UserControl
        lstEvents.Move .ScaleLeft + 50, .ScaleTop + 50, .ScaleWidth - 100, .ScaleHeight - 100
    End With
End Sub

Private Sub UserControl_Show()
    pAddItem ">>> UserControl_Show*"
    If Ambient.UserMode = False Then
        Exit Sub    '!! IDE MODE
    End If
    '## Engage subclassing
    Set moSubClass = New cSubClass
    moSubClass.hWnd = UserControl.Parent.hWnd
    mSubClass.Hook moSubClass
End Sub

Private Sub UserControl_Terminate()
    pAddItem ">>> UserControl_Terminate"
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    pAddItem ">>> UserControl_WriteProperties"
End Sub
