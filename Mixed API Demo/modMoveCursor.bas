Attribute VB_Name = "modMoveCursor"
Option Explicit

'//Set Cursor Position API Declaration
Public Declare Function SetCursorPos Lib "user32" _
(ByVal X As Long, ByVal Y As Long) As Long


'//MoveCursor Function
Public Function MoveCursor(X As Integer, Y As Integer)
Dim Pos As Integer
    
    Pos = SetCursorPos(X, Y) 'Calls API
    
End Function
