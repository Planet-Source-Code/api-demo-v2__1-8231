Attribute VB_Name = "modFrmOnTop"
Option Explicit



'//FormOnTop/FormNotOnTop API Declaration
Public Declare Function SetWindowPos Lib "user32" _
(ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal _
X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As _
Long, ByVal wFlags As Long) As Long

'//FormOnTop/FormNotOnTop Constants
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const FLAGS = SWP_NOSIZE Or SWP_NOMOVE


'//FormOnTop Function
Public Function FormOnTop(frm As Form)
Dim SetFrmOnTop As Long
    SetFrmOnTop = SetWindowPos(frm.hwnd, HWND_TOPMOST, 0, 0, _
                    0, 0, FLAGS)
End Function

'//FormNotOnTop Function
Public Function FormNotOnTop(frm As Form)
Dim SetFrmNotOnTop As Long
    SetFrmNotOnTop = SetWindowPos(frm.hwnd, HWND_NOTOPMOST, 0, _
                        0, 0, 0, FLAGS)
End Function


