Attribute VB_Name = "modClipCuror"
Option Explicit

'//Clip Cursor API Declaration
Public Declare Function ClipCursor Lib "user32" (lpRect As _
Any) As Long

'//Clip Cursor Variables
Public lTwipsX As Long
Public lTwipsY As Long

Public RectArea As RECT

'//Clip Cursor Type
Public Type RECT
    Left   As Long
    Top    As Long
    Right  As Long
    Bottom As Long
End Type


'//Trap Cursor Function
Public Function TrapCursor(frm As Form)
    '//Assigns values to variables
    lTwipsX = Screen.TwipsPerPixelX
    lTwipsY = Screen.TwipsPerPixelY

'//Assigns values to Type(This is the forms measurments)
With RectArea
    .Left = frm.Left / lTwipsX
    .Top = frm.Top / lTwipsY
    .Right = .Left + frm.Width / lTwipsX
    .Bottom = .Top + frm.Height / lTwipsY
End With

'//Calls API
Call ClipCursor(RectArea)
End Function

Public Function ReleaseCursor(frm As Form)
    '//Assigns values to variables
    lTwipsX = Screen.TwipsPerPixelX
    lTwipsY = Screen.TwipsPerPixelY

'//Assigns values to Type
With RectArea
    .Left = 0
    .Top = 0
    .Right = Screen.Width / lTwipsX
    .Bottom = Screen.Height / lTwipsY
End With

'//Calls API
Call ClipCursor(RectArea)
End Function
