VERSION 5.00
Begin VB.Form frmCtrlAltDel 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "    Disable/Enable Ctrl Alt Del & Windows Keys"
   ClientHeight    =   1950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4245
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   4245
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   1200
      TabIndex        =   2
      Top             =   1200
      Width           =   1695
   End
   Begin VB.CommandButton cmdEnable 
      Caption         =   "Enable"
      Height          =   495
      Left            =   2280
      TabIndex        =   1
      Top             =   360
      Width           =   1695
   End
   Begin VB.CommandButton cmdDisable 
      Caption         =   "Disable"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1695
   End
End
Attribute VB_Name = "frmCtrlAltDel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDisable_Click()
    DisableCtrlAltDel 'Disable Keys
End Sub

Private Sub cmdEnable_Click()
    EnableCtrlAltDel 'Enable keys
End Sub

Private Sub cmdExit_Click()
    EnableCtrlAltDel 'Enables keys (Just incase)
    Unload Me 'unload this form
    frmChoose.Show 'load menu
End Sub
