VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pop Off - The IE Popup Killer"
   ClientHeight    =   3240
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   2985
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   2985
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList iList 
      Left            =   840
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0894
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "About"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Frame fraOptions 
      Caption         =   "Options"
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   2775
      Begin VB.Timer tmrFlash 
         Enabled         =   0   'False
         Interval        =   200
         Left            =   240
         Top             =   1320
      End
      Begin VB.CheckBox chkStart 
         Caption         =   "Popoff starts with Windows"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   1815
      End
      Begin VB.CheckBox chkActive 
         Caption         =   "PopOff is active"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label lblNote 
         Caption         =   "Note: To temporarily allow a popup, hold down CTRL when clicking a link"
         Height          =   615
         Left            =   120
         TabIndex        =   4
         Top             =   1920
         Width           =   2535
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents IEWin As cIEWindows
Attribute IEWin.VB_VarHelpID = -1

Public TIcon As New clsTray
Dim lCount As Long

Private Sub chkActive_Click()

    If chkActive = vbChecked Then
        Active = True
        frmTrayMenu.mnuOptionsItems(2).Checked = vbChecked
    Else
        Active = False
        frmTrayMenu.mnuOptionsItems(2).Checked = vbUnchecked
    End If
End Sub

Private Sub chkStart_Click()
    
    If chkStart.Value = vbChecked Then
        AddToRun App.Title, App.Path & "\" & App.EXEName & ".exe"
    Else
        RemoveFromRun App.Title
    End If

End Sub

Private Sub cmdAbout_Click()
    frmAbout.Show vbModal
End Sub

Private Sub cmdExit_Click()
    Me.WindowState = vbMinimized
End Sub

Private Sub Form_Load()

    Dim sTemp As String
    Active = True
    
    Set IEWin = New cIEWindows
    'MakeTopMost Me.hwnd
    
    TIcon.RemoveIcon Me
    TIcon.ShowIcon Me
    TIcon.ChangeToolTip Me, "PopOff Popup Killer"
    
    sTemp = GetSetting(App.Title, "Settings", "RunAtStart", "0")
    chkStart.Value = CInt(sTemp)
    sTemp = GetSetting(App.Title, "Settings", "Active", "0")
    chkActive = CInt(sTemp)
    sTemp = GetSetting(App.Title, "Settings", "WindowStateMin", "")
    If sTemp = "True" Then
        Me.Hide
    End If
    

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If TIcon.bRunningInTray Then

    Select Case X
    
        Case 7725:
            Me.WindowState = vbNormal
            Me.Show
            
        Case 7755:
            PopupMenu frmTrayMenu.mnuOptions

    End Select
    
End If

End Sub

Private Sub Form_Resize()
    If bCalledFromPopup Then
        Me.WindowState = vbNormal
    Else
        If Me.WindowState = vbMinimized Then Me.Hide
    End If
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

If MsgBox("Are you sure you wish to quit PopOff Popup Killer?", vbInformation + vbYesNo, "Really??!!") = vbYes Then
    TIcon.RemoveIcon Me
    Call CloseAll
Else
    Cancel = -1
End If

End Sub

Private Sub IEWin_IEWindowRegistered()
UpdateList
End Sub

Private Sub IEWin_IEWindowRevoked()
UpdateList
End Sub

Private Sub mAbout_Click()
frmAbout.Show 1
End Sub

Private Sub mExit_Click()

If MsgBox("Are you sure you wish to quit?", vbInformation + vbYesNo, "Question") = vbYes Then
    TIcon.RemoveIcon Me
    End
End If

End Sub

Private Sub UpdateList()

On Error Resume Next

Dim temp As IE_Class
Dim itm As ListItem

vindows.ListItems.Clear

For Each temp In IEWin
   Set itm = vindows.ListItems.Add(, "K" & CStr(temp.IEHandle), temp.IEctl.LocationName)
   itm.SubItems(1) = temp.IEctl.LocationURL
Next temp

Set itm = Nothing

End Sub

Private Sub mPopup_Click()

End Sub

Private Sub mKiller_Click()

If mKiller.Checked = True Then
    mKiller.Checked = False
    Active = False
Else
    mKiller.Checked = True
    Active = True
End If

End Sub

Private Sub mRegister_Click()
frmRegister.Show 1
End Sub

Private Sub tmrFlash_Timer()

    Static bOn As Boolean
        
    If bOn Then
        TIcon.ChangeIcon Me, iList.ListImages(1)
        bOn = False
    Else
        TIcon.ChangeIcon Me, iList.ListImages(2)
        bOn = True
    End If

    lCount = lCount + 1
    
    If lCount = 5 Then
        TIcon.ChangeIcon Me, iList.ListImages(1)
        lCount = 0
        tmrFlash.Enabled = False
    End If
    
End Sub
