VERSION 5.00
Begin VB.Form frmPaint 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "           Drag this form and see what happens"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   2880
      Width           =   4455
   End
   Begin VB.Timer tmrPaint 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   240
      Top             =   240
   End
End
Attribute VB_Name = "frmPaint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'//API call for PaintDesktop(small isn't it)
Private Declare Function PaintDesktop Lib "user32" _
    (ByVal hdc As Long) As Long

Private Sub cmdExit_Click()
    frmChoose.Show 'Show the menu form
    Unload Me 'Unload this form
End Sub

Private Sub Form_Load()
    '//Enables timer
    tmrPaint.Enabled = True
End Sub

Private Sub tmrPaint_Timer()
    '//Paints desktop on thie form
    '//That's it COOL or what!!!!!!!!!!!
    PaintDesktop Me.hdc
End Sub
