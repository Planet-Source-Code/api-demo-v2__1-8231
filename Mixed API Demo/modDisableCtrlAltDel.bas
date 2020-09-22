Attribute VB_Name = "modDisableCtrlAltDel"
Option Explicit

'//DisableCtrlAltDel/EnableCtrlAltDel API Declaration
Public Declare Function SystemParametersInfo Lib "user32" _
Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal _
uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As _
Long) As Long

'//DisableCtrlAltDel/EnableCtrlAltDel Constant
Public Const SPI_SCREENSAVERRUNNING = 97


'//DisableCtrlAltDel Function
Public Function DisableCtrlAltDel()
Dim dRet As Integer
Dim dOld As Boolean
    dRet = SystemParametersInfo(SPI_SCREENSAVERRUNNING, _
            True, dOld, 0)
End Function

'//EnableCtrlAltDel Function
Public Function EnableCtrlAltDel()
Dim eRet As Integer
Dim eOld As Boolean
    eRet = SystemParametersInfo(SPI_SCREENSAVERRUNNING, _
            False, eOld, 0)
End Function


