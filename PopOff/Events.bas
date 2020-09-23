Attribute VB_Name = "Events"

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal Hkey As Long) As Long


Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal Hkey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal Hkey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal Hkey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
    
Const ERROR_SUCCESS = 0&
Const REG_SZ = 1 ' Unicode nul terminated String
Const REG_DWORD = 4 ' 32-bit number


Public Enum HKeyTypes
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    HKEY_PERFORMANCE_DATA = &H80000004
End Enum

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDest As Any, lpSource As Any, ByVal cBytes&)
Public cIEWPtr As Long, bCancel As Boolean
Public Enum IDEVENTS
   ID_BeforeNavigate = 1
   ID_NavigationComplete = 2
   ID_DownloadBegin = 3
   ID_DownloadComplete = 4
   ID_DocumentComplete = 5
   ID_MouseDown = 6
   ID_MouseUp = 7
   ID_ContextMenu = 8
   ID_CommandStateChange = 9
End Enum

Public Function CallEvent(nEvent As IDEVENTS, hwnd As Long, ParamArray EventInfo())
   Select Case nEvent
          Case ID_BeforeNavigate
               ResolvePointer(cIEWPtr).FireEvent nEvent, hwnd, EventInfo(0), EventInfo(1), EventInfo(2), EventInfo(3), EventInfo(4), EventInfo(5), CBool(EventInfo(6))
          Case ID_NavigationComplete, ID_DocumentComplete, ID_CommandStateChange
               ResolvePointer(cIEWPtr).FireEvent nEvent, hwnd, EventInfo(0), EventInfo(1)
          Case ID_MouseDown, ID_MouseUp
               ResolvePointer(cIEWPtr).FireEvent nEvent, hwnd, EventInfo(0), EventInfo(1), EventInfo(2), EventInfo(3)
          Case Else
               ResolvePointer(cIEWPtr).FireEvent nEvent, hwnd
   End Select
End Function

Private Function ResolvePointer(ByVal lpObj&) As cIEWindows
  Dim oIEW As cIEWindows
  CopyMemory oIEW, lpObj, 4&
  Set ResolvePointer = oIEW
  CopyMemory oIEW, 0&, 4&
End Function

Public Sub AddToRun(ProgramName As String, FileToRun As String)
    'Add a program to the 'Run at Startup' r
    '     egistry keys
    Call SaveString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", ProgramName, FileToRun)
End Sub

Public Sub RemoveFromRun(ProgramName As String)
    'Remove a program from the 'Run at Start
    '     up' registry keys
    Call DeleteValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", ProgramName)
End Sub

Public Sub SaveString(Hkey As HKeyTypes, strPath As String, strValue As String, strdata As String)
    'EXAMPLE:
    '
    'Call savestring(HKEY_CURRENT_USER, "Sof
    '     tware\VBW\Registry", "String", text1.tex
    '     t)
    '
    Dim keyhand As Long
    Dim r As Long
    r = RegCreateKey(Hkey, strPath, keyhand)
    r = RegSetValueEx(keyhand, strValue, 0, REG_SZ, ByVal strdata, Len(strdata))
    r = RegCloseKey(keyhand)
End Sub


Public Function DeleteValue(ByVal Hkey As HKeyTypes, ByVal strPath As String, ByVal strValue As String)
    'EXAMPLE:
    '
    'Call DeleteValue(HKEY_CURRENT_USER, "So
    '     ftware\VBW\Registry", "Dword")
    '
    Dim keyhand As Long
    r = RegOpenKey(Hkey, strPath, keyhand)
    r = RegDeleteValue(keyhand, strValue)
    r = RegCloseKey(keyhand)
End Function
