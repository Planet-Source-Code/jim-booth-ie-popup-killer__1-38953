VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About.."
   ClientHeight    =   1500
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3285
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1500
   ScaleWidth      =   3285
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   2
      Top             =   600
      Width           =   2535
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      Caption         =   "PopOff Popup Killer"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   2535
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOk_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    lblInfo(2) = App.Major & "." & App.Minor & "." & App.Revision
End Sub
