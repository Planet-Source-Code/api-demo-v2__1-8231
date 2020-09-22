VERSION 5.00
Begin VB.Form frmTrapCursor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "                       Trap the Cursor"
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4230
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   4230
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   615
      Left            =   1320
      TabIndex        =   2
      Top             =   1680
      Width           =   1695
   End
   Begin VB.CommandButton cmdRelease 
      Caption         =   "Release Cursor"
      Height          =   615
      Left            =   2280
      TabIndex        =   1
      Top             =   240
      Width           =   1695
   End
   Begin VB.CommandButton cmdTrap 
      Caption         =   "Trap Cursor"
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "frmTrapCursor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdExit_Click()
    Unload Me 'Unloads this form
    frmChoose.Show 'Loads menu
End Sub

Private Sub cmdRelease_Click()
    ReleaseCursor Me 'Releases cursor
End Sub

Private Sub cmdTrap_Click()
    TrapCursor Me 'Traps cursor
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '//////  IMPORTANT!!!!!!!!!
    '//If you don't call this function the cursor will stay
    '//trapped
    ReleaseCursor Me 'Releases cursor
End Sub
