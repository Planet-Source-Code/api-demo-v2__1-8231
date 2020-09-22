VERSION 5.00
Begin VB.Form frmMsgBox 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "             API MessageBox Demo"
   ClientHeight    =   2685
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3600
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   3600
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton cmdTest1 
      Caption         =   "Test1"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "About"
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   840
      Width           =   975
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      Caption         =   $"frmTest.frx":0000
      Height          =   1215
      Left            =   240
      TabIndex        =   3
      Top             =   1440
      Width           =   3135
   End
End
Attribute VB_Name = "frmMsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAbout_Click()
Dim Response As Long
 
'//The line below gives the variable Response the value(Long)
'//that the Function(API) returns
Response = MessageBox(Me.hwnd, "Did you know this MsgBox uses API!!", _
            "Cool huh !!!", MB_YESNO Or _
                MB_ICONQUESTION Or MB_TASKMODAL)


'//We can now manage the 'Response' from the user
Select Case Response
Case Is = IDYES 'The user clicked Yes
        Call MessageBox(Me.hwnd, "Aren't you clever", _
                "Smart Lad Eh!", MB_ICONEXCLAMATION)

Case Is = IDNO 'The user clicked No
        Call MessageBox(Me.hwnd, "Then read the Module", _
                "Information", MB_ICONINFORMATION)
End Select
End Sub


Private Sub cmdTest1_Click()
Dim Response As Long

'//The line below gives the variable Response the value(Long)
'//that the Function(API) returns
Response = MessageBox(Me.hwnd, "Cannot read disk", _
            "Big problem", MB_ABORTRETRYIGNORE Or _
                MB_ICONCRITICAL Or MB_TASKMODAL)

'//We can now manage the 'Response' from the user
Select Case Response
Case Is = IDABORT 'The user clicked Abort
        Call MessageBox(Me.hwnd, "Wuss", _
                "Chicken", MB_ICONEXCLAMATION)

Case Is = IDRETRY 'The user clicked Retry
        Call MessageBox(Me.hwnd, "What a Hero", _
                "Retry", MB_ICONINFORMATION)

Case Is = IDIGNORE 'The user clicked Ignore
        Call MessageBox(Me.hwnd, "Taken the easy way huh", _
                "Leave then", MB_ICONMASK)
End Select
End Sub


Private Sub cmdQuit_Click()
Dim Response As Long
Dim Response1 As Long
Dim Response2 As Long

'//The line below gives the variable Response the value(Long)
'//that the Function(API) returns
Response = MessageBox(Me.hwnd, "Are you sure?", _
            "Well are you?", MB_YESNO Or _
                MB_ICONQUESTION Or MB_TASKMODAL)


'//We can now manage the 'Response' from the user
Select Case Response
Case Is = IDYES 'The user clicked Yes
        Response1 = MessageBox(Me.hwnd, "Positive", _
                    "Don't Go", MB_OK Or MB_ICONINFORMATION)
            
            If Response1 = IDOK Then
                frmChoose.Show 'Load the menu form
                Unload Me 'Unload this form
            End If
            
Case Is = IDNO 'The user clicked No
       Call MessageBox(Me.hwnd, "You still 'ere", _
            "Sod Off", MB_OK Or MB_ICONCRITICAL)
End Select
End Sub

