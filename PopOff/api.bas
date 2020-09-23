Attribute VB_Name = "api"
Public Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long

Private Declare Function SetForegroundWindow& Lib "user32" (ByVal hwnd As Long)
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function PostMessage& Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any)
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, Y, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Declare Function GetAsyncKeyState Lib "user32" (ByVal dwMessage As Long) As Integer

Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOMOVE = &H2
Const SWP_NOSIZE = &H1
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40
Const TOPMOST_FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Public bCalledFromPopup As Boolean
Public bContextCall As Boolean

Public Active As Boolean

Public Sub MakeNormal(hwnd As Long)
    SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
End Sub

Public Sub MakeTopMost(hwnd As Long)
    SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
End Sub

Public Function CloseAll()
    Dim frm As Form
        
    Call SaveSettings
        
        For Each frm In Forms
            Unload frm
        Next
End Function

Public Function SaveSettings()

    If frmMain.WindowState = vbMinimized Then
        SaveSetting App.Title, "Settings", "WindowStateMin", "True"
    Else
        SaveSetting App.Title, "Settings", "WindowStateMin", "False"
    End If
    
    SaveSetting App.Title, "Settings", "Active", frmMain.chkActive.Value
    SaveSetting App.Title, "Settings", "RunAtStart", frmMain.chkStart.Value
    
End Function

