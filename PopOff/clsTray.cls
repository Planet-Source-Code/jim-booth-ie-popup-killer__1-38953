VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' To use on form:
' Public myTrayIcon As New clsSysTrayIcon
' myTrayIcon.RemoveIcon Me
' myTrayIcon.ShowIcon Me
' myTrayIcon.ChangeIcon Me, (Image, Picture Box, Etc)
' myTrayIcon.ChangeToolTip Me, InputBox("Enter Tool Tip")
'
' ****** Tray Menu Source ****** (Drop In Form)
'
'Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Remember..... The value of X will be different if the icon is minimised
' to the system tray.  The values in this case will be as follows,
'       7680   ' MouseMove
'       7695   ' Left MouseDown
'       7710   ' Left MouseUp
'       7725   ' Left DoubleClick
'       7740   ' Right MouseDown
'       7755   ' Right MouseUp
'       7770   ' Right DoubleClick
'If myTrayIcon.bRunningInTray Then          'Check to see if form is in the system tray
'    Select Case X                           'If it is, use X to get message value
'        Case 7710: PopupMenu Me.mnuBar1, vbPopupMenuRightButton  'Show a menubar
'        Case 7725: myTrayIcon.RemoveIcon Me
'    End Select
'End If
'
'End Sub

Option Explicit

Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type
  
Private FormHandle As Long
Private mvarbRunningInTray As Boolean
Private SysIcon As NOTIFYICONDATA
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

Private Property Let bRunningInTray(ByVal vData As Boolean)
    mvarbRunningInTray = vData
End Property


Property Get bRunningInTray() As Boolean
    bRunningInTray = mvarbRunningInTray
End Property

Public Sub ShowIcon(ByRef sysTrayForm As Form)
    SysIcon.cbSize = Len(SysIcon)
    SysIcon.hwnd = sysTrayForm.hwnd
    SysIcon.uId = vbNull
    SysIcon.uFlags = 7
    SysIcon.ucallbackMessage = 512
    SysIcon.hIcon = sysTrayForm.Icon
    SysIcon.szTip = sysTrayForm.Caption + Chr(0)
    Shell_NotifyIcon 0, SysIcon
    mvarbRunningInTray = True
End Sub

Public Sub RemoveIcon(sysTrayForm As Form)
    SysIcon.cbSize = Len(SysIcon)
    SysIcon.hwnd = sysTrayForm.hwnd
    SysIcon.uId = vbNull
    SysIcon.uFlags = 7
    SysIcon.ucallbackMessage = vbNull
    SysIcon.hIcon = sysTrayForm.Icon
    SysIcon.szTip = Chr(0)
    Shell_NotifyIcon 2, SysIcon
    'If sysTrayForm.Visible = False Then sysTrayForm.Show    'Incase user can't see form
    mvarbRunningInTray = False
End Sub

Public Sub ChangeIcon(sysTrayForm As Form, picNewIcon As ListImage)

If mvarbRunningInTray = True Then   'If running in the tray
    SysIcon.cbSize = Len(SysIcon)
    SysIcon.hwnd = sysTrayForm.hwnd
    'SysIcon.uId = vbNull
    'SysIcon.uFlags = 7
    'SysIcon.ucallbackMessage = 512
    SysIcon.hIcon = picNewIcon.Picture
    'SysIcon.szTip = sysTrayForm.Caption + Chr(0)
    Shell_NotifyIcon 1, SysIcon
End If

End Sub

Public Sub ChangeToolTip(sysTrayForm As Form, strNewTip As String)

If mvarbRunningInTray = True Then   'If running in the tray
    SysIcon.cbSize = Len(SysIcon)
    SysIcon.hwnd = sysTrayForm.hwnd
    SysIcon.szTip = strNewTip & Chr(0)
    Shell_NotifyIcon 1, SysIcon
End If

End Sub

