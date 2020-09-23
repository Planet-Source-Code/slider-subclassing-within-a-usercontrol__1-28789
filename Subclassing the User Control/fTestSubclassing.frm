VERSION 5.00
Object = "*\ApTestSubclassUC.vbp"
Begin VB.Form fTestSubclassing 
   Caption         =   "Test - User Control Subclassing Windows Messages"
   ClientHeight    =   4950
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5250
   LinkTopic       =   "Form1"
   ScaleHeight     =   4950
   ScaleWidth      =   5250
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraDialog 
      Caption         =   "User Control Event List: "
      Height          =   4740
      Left            =   105
      TabIndex        =   1
      Top             =   105
      Width           =   3375
      Begin pTestSubclassUC.ucTest ucTest 
         Height          =   4320
         Left            =   105
         TabIndex        =   4
         Top             =   315
         Width           =   3165
         _ExtentX        =   5583
         _ExtentY        =   7620
      End
   End
   Begin VB.CommandButton cmdDialog 
      Caption         =   "E&xit"
      Height          =   435
      Index           =   2
      Left            =   3570
      TabIndex        =   3
      Top             =   4410
      Width           =   1590
   End
   Begin VB.CommandButton cmdDialog 
      Caption         =   "&Display Settings"
      Height          =   435
      Index           =   1
      Left            =   3570
      TabIndex        =   2
      Top             =   735
      Width           =   1590
   End
   Begin VB.CommandButton cmdDialog 
      Caption         =   "&Set Date/Time"
      Height          =   435
      Index           =   0
      Left            =   3570
      TabIndex        =   0
      Top             =   210
      Width           =   1590
   End
End
Attribute VB_Name = "fTestSubclassing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===========================================================================
'
' Form Name:    fTestSubClassing
' Author:       Slider
' Date:         11/11/2001
' Version:      01.00.00
' Description:  Demonstrates the Subclassing Test User Control.
' Edit History: 01.00.00 11/11/01 Initial Release
'
'===========================================================================

Option Explicit

Private Enum eCommands
    [Set Date/Time] = 0
    [Display Settings] = 1
    [Exit Application] = 2
End Enum

Private mlFrameRight As Long
Private mlCmdLeft    As Long
Private mlCmdTop     As Long
    
Private Sub cmdDialog_Click(Index As Integer)
    Select Case Index
        Case [Set Date/Time]
            Call Shell("rundll32.exe shell32.dll,Control_RunDLL timedate.cpl,,0", vbNormalFocus)
        Case [Display Settings]
            Call Shell("rundll32.exe shell32.dll,Control_RunDLL desk.cpl,,2", vbNormalFocus)
        Case [Exit Application]
            Unload Me
    End Select
End Sub

Private Sub Form_Load()
    mlFrameRight = Me.ScaleWidth - (fraDialog.Width - fraDialog.Left)
    mlCmdLeft = Me.ScaleWidth - cmdDialog([Set Date/Time]).Left
    mlCmdTop = Me.ScaleHeight - cmdDialog([Exit Application]).Top
End Sub

Private Sub Form_Resize()

    Debug.Print "Form_Resize"

    Dim lLoop  As Long
    Dim lWidth As Long
    Dim lHght  As Long

    With fraDialog
        lWidth = Me.ScaleWidth - mlFrameRight
        lHght = Me.ScaleHeight - .Top * 2
        Debug.Print lHght
        If (lWidth > 675) And (lHght > 775) Then
            .Move .Left, .Top, lWidth, lHght
        End If
    End With

    With ucTest
        lWidth = fraDialog.Width - .Left * 2
        lHght = fraDialog.Height - .Top * 1.35
        If (lWidth > 400) And (lHght > 500) Then
            .Move .Left, .Top, lWidth, lHght
        End If
    End With

    For lLoop = [Set Date/Time] To [Exit Application]
        cmdDialog(lLoop).Left = Me.ScaleWidth - mlCmdLeft
    Next
    cmdDialog([Exit Application]).Top = Me.ScaleHeight - mlCmdTop

End Sub
