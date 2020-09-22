VERSION 5.00
Begin VB.Form frmChoose 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "       Choose your API demo"
   ClientHeight    =   7170
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2820
   ControlBox      =   0   'False
   Icon            =   "frmChoose.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7170
   ScaleWidth      =   2820
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdHideCursor 
      Caption         =   "Hide Cursor"
      Height          =   495
      Left            =   360
      TabIndex        =   9
      Top             =   6000
      Width           =   2055
   End
   Begin VB.CommandButton cmdSetCursorPos 
      Caption         =   "Set Cursor Position"
      Height          =   495
      Left            =   360
      TabIndex        =   8
      Top             =   5280
      Width           =   2055
   End
   Begin VB.CommandButton cmdGetCursorPos 
      Caption         =   "Get Cursor Position"
      Height          =   495
      Left            =   360
      TabIndex        =   7
      Top             =   4560
      Width           =   2055
   End
   Begin VB.CommandButton cmdTrapCursor 
      Caption         =   "Trap Cursor"
      Height          =   495
      Left            =   360
      TabIndex        =   6
      Top             =   3840
      Width           =   2055
   End
   Begin VB.CommandButton cmdCtrlAltDel 
      Caption         =   "Enable/Disable Ctrl Alt Del"
      Height          =   495
      Left            =   360
      TabIndex        =   5
      Top             =   3120
      Width           =   2055
   End
   Begin VB.CommandButton cmdOnTop 
      Caption         =   "Form OnTop/NotOnTop"
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   2400
      Width           =   2055
   End
   Begin VB.CommandButton cmdPaintDesktop 
      Caption         =   "Paint Desktop"
      Height          =   495
      Left            =   360
      TabIndex        =   3
      Top             =   1680
      Width           =   2055
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   6720
      Width           =   2055
   End
   Begin VB.CommandButton cmdMenuBitmap 
      Caption         =   "Bitmaps In Menus"
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   960
      Width           =   2055
   End
   Begin VB.CommandButton cmdMessageBox 
      Caption         =   "MessageBox Demo"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "frmChoose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCtrlAltDel_Click()
    frmCtrlAltDel.Show 'Load disable key form
    Unload Me 'Unload this form
End Sub

Private Sub cmdExit_Click()
    End 'Exit
End Sub

Private Sub cmdGetCursorPos_Click()
    frmGetCursorPos.Show 'Load Get Cursor Position form
    Unload Me 'Unload this form
End Sub

Private Sub cmdHideCursor_Click()
    frmHideCursor.Show 'Show Hide cursor form
    Unload Me 'Unload this form
End Sub

Private Sub cmdMenuBitmap_Click()
    frmMenuBitmap.Show 'Load MenuBitmap form
    Unload Me 'Unload this form
End Sub

Private Sub cmdMessageBox_Click()
    frmMsgBox.Show 'Load MessageBox form
    Unload Me 'Unload this form
End Sub

Private Sub cmdOnTop_Click()
    frmOnTop.Show 'Load Ontop Form
    Unload Me 'unload this form
End Sub

Private Sub cmdPaintDesktop_Click()
    frmPaint.Show 'Load Paint Desktop form
    Unload Me 'Unload this form
End Sub

Private Sub cmdSetCursorPos_Click()
    frmMoveCursor.Show 'Load Set Cursor Position Form
    Unload Me 'Unload this form
End Sub

Private Sub cmdTrapCursor_Click()
    frmTrapCursor.Show 'Load Trap Cursor form
    Unload Me 'Unload this form
End Sub
