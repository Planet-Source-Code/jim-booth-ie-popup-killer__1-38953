VERSION 5.00
Begin VB.Form frmTrayMenu 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuOptions 
      Caption         =   ""
      Begin VB.Menu mnuOptionsItems 
         Caption         =   "Show"
         Index           =   0
      End
      Begin VB.Menu mnuOptionsItems 
         Caption         =   "About.."
         Index           =   1
      End
      Begin VB.Menu mnuOptionsItems 
         Caption         =   "Active"
         Index           =   2
      End
      Begin VB.Menu mnuOptionsItems 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuOptionsItems 
         Caption         =   "Exit"
         Index           =   4
      End
   End
End
Attribute VB_Name = "frmTrayMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub mnuOptionsItems_Click(Index As Integer)

    Select Case Index
    
    Case 0 'Show
        bCalledFromPopup = True
        frmMain.Show
    Case 1 'About
        frmAbout.Show vbModal
    Case 2 'Active
        mnuOptionsItems(2).Checked = Not mnuOptionsItems(2).Checked
        Active = mnuOptionsItems(2).Checked
        frmMain.chkActive.Value = Not frmMain.chkActive.Value
    Case 3 'Seperator
    
    Case 4 'Exit
        Call CloseAll
    End Select
    
    bCalledFromPopup = False

End Sub
