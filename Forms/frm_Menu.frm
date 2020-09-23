VERSION 5.00
Begin VB.Form frm_Menu 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   15
   ClientLeft      =   30
   ClientTop       =   1050
   ClientWidth     =   3780
   LinkTopic       =   "Form1"
   ScaleHeight     =   15
   ScaleWidth      =   3780
   ShowInTaskbar   =   0   'False
   Begin VB.Menu ppm_File 
      Caption         =   "&File"
      Begin VB.Menu ppi_Welcome 
         Caption         =   "&Welcome"
      End
      Begin VB.Menu Seperator01 
         Caption         =   "-"
      End
      Begin VB.Menu ppi_Exit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu ppm_Forms 
      Caption         =   "&Forms"
      Begin VB.Menu ppi_align 
         Caption         =   "Autoalign"
         Checked         =   -1  'True
      End
      Begin VB.Menu ppi_Form2 
         Caption         =   "Form2"
      End
   End
   Begin VB.Menu ppm_Report 
      Caption         =   "&Report"
      Begin VB.Menu ppi_Out 
         Caption         =   "Items Out"
      End
      Begin VB.Menu ppi_In 
         Caption         =   "Item In"
      End
   End
   Begin VB.Menu ppm_Tools 
      Caption         =   "T&ools"
      Begin VB.Menu ppi_Calculator 
         Caption         =   "Calculator"
      End
      Begin VB.Menu ppi_Calendar 
         Caption         =   "Calendar"
      End
   End
   Begin VB.Menu ppm_Help 
      Caption         =   "&Help"
      Begin VB.Menu ppi_Contents 
         Caption         =   "Contents"
      End
      Begin VB.Menu ppi_Search 
         Caption         =   "Search"
      End
      Begin VB.Menu Seperator02 
         Caption         =   "-"
      End
      Begin VB.Menu ppi_About 
         Caption         =   "About..."
      End
   End
End
Attribute VB_Name = "frm_Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ppi_About_Click()
    'frm_About.Show 1
End Sub

Private Sub ppi_ALPI_Click()
    'Call ChangeSkinToALPI
End Sub

Private Sub ppi_Blue_Click()
    'Call ChangeSkinToBlue
End Sub

Private Sub ppi_ChannelBar_Click()
    'pIndex = 4
    'Call Showtutorials(pIndex)
End Sub

Private Sub ppi_ComingSoon_Click()
    'pIndex = 8
    'Call Showtutorials(pIndex)
End Sub

Private Sub ppi_Coupe_Click()
    'Call ChangeSkinToCoupe
End Sub

Private Sub ppi_Deco_Click()
    'Call ChangeSkinToDeco
End Sub

Private Sub ppi_Default_Click()
    'Call ChangeSkinToDefault
End Sub

Private Sub ppi_Doesnt_Suck_Click()
    'Call ChangeSkinToDoesnt_Suck
End Sub

Private Sub ppi_align_Click()
 
 If ppi_align.Checked = True Then
   ppi_align.Checked = False
 Else
   ppi_align.Checked = True
 End If
 If ppi_align.Checked = True Then
   MainForm.chkAlign = 1
 Else
   MainForm.chkAlign = 0
 End If
  
End Sub

Private Sub ppi_Exit_Click()
    Unload MainForm
    Unload frm_Menu
End Sub

Private Sub ppi_Holograph_Click()
    'Call ChangeSkinToHolograph
End Sub

Private Sub ppi_Introduction_Click()
    'pIndex = 0
    'Call Showtutorials(pIndex)
End Sub

Private Sub ppi_ListObject_Click()
    'pIndex = 5
    'Call Showtutorials(pIndex)
End Sub

Private Sub ppi_Panel_Click()
    'pIndex = 7
    'Call Showtutorials(pIndex)
End Sub

Private Sub ppi_PulldownMenu_Click()
    'pIndex = 3
    'Call Showtutorials(pIndex)
End Sub

Private Sub ppi_ShowChannelBar_Click()
    'If frm_Menu.ppi_ShowChannelBar.Checked = True Then
    '    frm_Menu.ppi_ShowChannelBar.Checked = False
    '    frm_Main.ctrl_PullDownMenu.Visible = True
    '    frm_Main.ctrl_Toolbar.Visible = True
    '    frm_Main.ctrl_ChannelBar.Visible = False
    'Else
    '    frm_Menu.ppi_ShowChannelBar.Checked = True
    '    frm_Main.ctrl_PullDownMenu.Visible = False
    '    frm_Main.ctrl_Toolbar.Visible = False
    '    frm_Main.ctrl_ChannelBar.Visible = True
    'End If
End Sub

Private Sub ppi_ShowPulldownMenu_Click()
    'If frm_Menu.ppi_ShowPulldownMenu.Checked = True Then
    '    frm_Menu.ppi_ShowPulldownMenu.Checked = False
    '    frm_Main.ctrl_PullDownMenu.Visible = False
    'Else
    '    frm_Menu.ppi_ShowPulldownMenu.Checked = True
    '    frm_Main.ctrl_PullDownMenu.Visible = True
    'End If
End Sub

Private Sub ppi_ShowStatus_Click()
    'If frm_Menu.ppi_ShowStatus.Checked = True Then
    '    frm_Menu.ppi_ShowStatus.Checked = False
    '    frm_Main.Line1.Visible = False
    '    frm_Main.lbl_Statusbar.Visible = False
    'Else
    '    frm_Menu.ppi_ShowStatus.Checked = True
    '    frm_Main.Line1.Visible = True
    '    frm_Main.lbl_Statusbar.Visible = True
    'End If
End Sub

Private Sub ppi_ShowToolbar_Click()
    'If frm_Menu.ppi_ShowToolbar.Checked = True Then
    '    frm_Menu.ppi_ShowToolbar.Checked = False
    '    frm_Main.ctrl_Toolbar.Visible = False
    'Else
     '   frm_Menu.ppi_ShowToolbar.Checked = True
     '   frm_Main.ctrl_Toolbar.Visible = True
     'end If
End Sub

Private Sub ppi_SkinableButton_Click()
    'pIndex = 2
    'Call Showtutorials(pIndex)
End Sub

Private Sub ppi_SkinableForm_Click()
    'pIndex = 1
    'Call Showtutorials(pIndex)
End Sub

Private Sub ppi_SteelBlade_Click()
    'Call ChangeSkinToSteelBlade
End Sub

Private Sub ppi_SteelRain_Click()
    'Call ChangeSkinToSteelRain
End Sub

Private Sub ppi_Titanium_Click()
    'Call ChangeSkinToTitanium
End Sub

Private Sub ppi_Toolbar_Click()
    'pIndex = 6
    'Call Showtutorials(pIndex)
End Sub

Private Sub ppi_TreasureChest_Click()
    'Call ChangeSkinToTreasureChest
End Sub

Private Sub ppi_Wazoo_Click()

End Sub


Private Sub ppi_Welcome_Click()
    'Open App.Path & "\Welcome.txt" For Input As #1
    '    frm_Main.tbx_Text.Text = Input$(LOF(1), #1)
    'Close #1
End Sub
