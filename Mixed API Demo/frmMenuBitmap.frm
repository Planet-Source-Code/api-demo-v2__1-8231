VERSION 5.00
Begin VB.Form frmMenuBitmap 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "           Put bitmaps in the menu"
   ClientHeight    =   1935
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   3495
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   3495
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picExit 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   2640
      Picture         =   "frmMenuBitmap.frx":0000
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   2
      Top             =   1200
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picTest 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   2640
      Picture         =   "frmMenuBitmap.frx":0102
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   0
      Top             =   840
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "These are the pictures that will be copied onto the menu     >>>>>>>>>>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   480
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuHere 
         Caption         =   "<<Here it is"
      End
      Begin VB.Menu mnuSeperator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuSub 
      Caption         =   "SubMenu"
      Begin VB.Menu mnudummy 
         Caption         =   "Dummy"
         Begin VB.Menu mnucool 
            Caption         =   "COOOOL"
         End
      End
   End
End
Attribute VB_Name = "frmMenuBitmap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    BitmapInMenu Me 'Calls function sending this forms name
                    'This can be used to call other forms
                    'e.g BitmapInMenu frmGo etc
End Sub

Private Sub mnuExit_Click()
    frmChoose.Show 'Load the menu form
    Unload Me 'Unload this form
End Sub
