VERSION 5.00
Begin VB.Form frmOnTop 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "                       Keep Form OnTop"
   ClientHeight    =   1710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4335
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1710
   ScaleWidth      =   4335
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   1320
      TabIndex        =   2
      Top             =   960
      Width           =   1695
   End
   Begin VB.CommandButton cmdNotOnTop 
      Caption         =   "Not OnTop"
      Height          =   495
      Left            =   2400
      TabIndex        =   1
      Top             =   240
      Width           =   1695
   End
   Begin VB.CommandButton cmdOnTop 
      Caption         =   "OnTop"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "frmOnTop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
    Unload Me 'closes this form
    frmChoose.Show 'Loads Menu
End Sub

Private Sub cmdNotOnTop_Click()
    FormNotOnTop Me 'Returns form to normal
End Sub

Private Sub cmdOnTop_Click()
    FormOnTop Me 'Keeps form ontop
End Sub
