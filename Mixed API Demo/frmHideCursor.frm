VERSION 5.00
Begin VB.Form frmHideCursor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "                        Hide Cursor"
   ClientHeight    =   1740
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3870
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1740
   ScaleWidth      =   3870
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   1320
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton cmdShow 
      Caption         =   "Show Cursor"
      Height          =   495
      Left            =   2280
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton cmdHide 
      Caption         =   "Hide Cursor"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "frmHideCursor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'//ShowCursor API Declaration
Private Declare Function ShowCursor& Lib "user32" _
(ByVal bShow As Long)


Private Sub cmdExit_Click()
    frmChoose.Show 'Load Menu
    Unload Me 'Unload this form
End Sub

Private Sub cmdHide_Click()
    ShowCursor False 'Hide cursor
End Sub

Private Sub cmdShow_Click()
    ShowCursor True 'ShowCursor
End Sub


Private Sub Form_Unload(Cancel As Integer)
    '//IMPORTANT !!!!!!!!!!!!
    '//Like the trap cursor code you must call the function
    '//to re-display the cursor, otherwise it will stay
    '//invisible
    ShowCursor True 'ShowCursor
End Sub
