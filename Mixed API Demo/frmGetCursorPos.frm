VERSION 5.00
Begin VB.Form frmGetCursorPos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "                 Get Cursor Position"
   ClientHeight    =   2160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3780
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   3780
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1560
      Width           =   3255
   End
   Begin VB.Timer tmrPos 
      Interval        =   1
      Left            =   3840
      Top             =   600
   End
   Begin VB.Label lblY 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Y"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   3255
   End
   Begin VB.Label lblX 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3255
   End
End
Attribute VB_Name = "frmGetCursorPos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
    Unload Me 'Unload this form
    frmChoose.Show 'Load Menu
End Sub

Private Sub tmrPos_Timer()
    GetCursorPosition 'Call procedure from module
    '//Set Labels captions as X & Y Co-ordinates
    lblX = " X Position : " & XPos
    lblY = " Y Position : " & YPos
End Sub
