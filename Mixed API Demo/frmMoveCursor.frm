VERSION 5.00
Begin VB.Form frmMoveCursor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "                               Move Cursor"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   1320
      TabIndex        =   5
      Top             =   2640
      Width           =   1935
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Move Cursor"
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   1920
      Width           =   1935
   End
   Begin VB.TextBox txtY 
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Text            =   "0"
      Top             =   1200
      Width           =   2895
   End
   Begin VB.TextBox txtX 
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Text            =   "0"
      Top             =   480
      Width           =   2895
   End
   Begin VB.Label lblY 
      Caption         =   "Y Position"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label lblX 
      Caption         =   "X Position"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "frmMoveCursor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdExit_Click()
    Unload Me 'Unloads this form
    frmChoose.Show 'Load menu
End Sub


Private Sub cmdGo_Click()
On Error GoTo Err
    
    '//Calls function providing X & Y positions
    MoveCursor txtX.Text, txtY.Text
    
Exit Sub
Err:
MessageBox hwnd, "You have entered an invalid character" & vbCr & _
    "Enter only numbers", "Error", MB_OK Or _
        MB_ICONCRITICAL Or MB_TASKMODAL
End Sub


