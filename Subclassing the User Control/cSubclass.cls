VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSubClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===========================================================================
'
' Class Name:   cSubClass
' Author:       Slider
' Date:         11/11/2001
' Version:      01.00.00
' Description:  Subclassing event Mediator. Used to pass events to the
'               UserControl from the hooked Windows message.
' Edit History: 01.00.00 11/11/01 Initial Release
'
'===========================================================================

Option Explicit

Private mhWnd As Long
'
'## Add the events you want to raise here.
'
Event WinMessage(ByVal EventID As Long, ByVal wParam As Long, ByVal lParam As Long)

'===========================================================================
' Called by mSubClass.WindowProc - Event raised will be passed to the
' User Control.
'
Friend Sub fRaiseEvent(ByVal EventID As Long, ByVal wParam As Long, ByVal lParam As Long)
    RaiseEvent WinMessage(EventID, wParam, lParam)
End Sub

'===========================================================================
' Used by the mSubClass.SubClass. Holds the UserControl parent form's hWnd
'
Friend Property Get hWnd() As Long
    hWnd = mhWnd
End Property

Friend Property Let hWnd(ByVal vNewValue As Long)
    mhWnd = vNewValue
End Property

'===========================================================================
' Private Class Events
'
Private Sub Class_Initialize()
    Debug.Print "+++ Class_Initialize"
End Sub

Private Sub Class_Terminate()
    Debug.Print "+++ Class_Terminate"
End Sub
