Attribute VB_Name = "modCursorPosition"
Option Explicit

'//Get Cursor Position API Declaration
Public Declare Function GetCursorPos Lib "user32" (lpPoint _
As POINTAPI) As Long

Public MousePos As POINTAPI
Public XPos As Integer
Public YPos As Integer

'//Get Cursor Position Type
Public Type POINTAPI
  X As Long
  Y As Long
End Type


Public Sub GetCursorPosition()
    GetCursorPos MousePos ' Get Co-ordinates
    XPos = MousePos.X ' Get X co-ordinates
    YPos = MousePos.Y ' Get Y co-ordinates
End Sub
