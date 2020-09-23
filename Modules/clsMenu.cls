VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'// coded by edwin delos santos   ... no cheating
Option Explicit
Private WithEvents m_Form1 As MainForm
Attribute m_Form1.VB_VarHelpID = -1
Private menu_Name As String
Public Property Get MainForm() As MainForm

   Set MainForm = m_Form1
End Property

Public Property Set MainForm(ByVal NewForm1 As MainForm)
   Set m_Form1 = NewForm1
End Property


Private Sub m_Form1_TakboNa()   'escape
     MainForm.mnu_ListOption.Visible = False
     MainForm.mnu_Shadow.Visible = False
     MainForm.shadow_mnu.Visible = False
End Sub
Private Sub m_form1_SaKaliwa()  'move to left
   Select Case menu_Name
      Case Is = "File"
          m_form1_HelpClick
      Case Is = "Menu"
          m_form1_FileClick
      Case Is = "Tools"
         m_form1_MenuClick
      Case Is = "Help"
         m_form1_ToolsClick
      End Select
End Sub
Private Sub m_form1_SaKanan()  'move to right
   Select Case menu_Name
      Case Is = "File"
          m_form1_MenuClick
      Case Is = "Menu"
          m_form1_ToolsClick
      Case Is = "Tools"
         m_form1_HelpClick
      Case Is = "Help"
         m_form1_FileClick
      End Select
End Sub

Private Sub m_form1_FileClick()
With m_Form1
   menu_File .mnu_File, .mnu_ListOption
   menu_shadow .mnu_ListOption, .mnu_File, .shadow_mnu, .mnu_Shadow
End With
 menu_Name = "File"
End Sub
Private Sub m_form1_MenuClick()
With m_Form1
   menu_Menu .mnu_Menu, .mnu_ListOption
   menu_shadow .mnu_ListOption, .mnu_Menu, .shadow_mnu, .mnu_Shadow
End With
 menu_Name = "Menu"
End Sub

Private Sub m_form1_ToolsClick()
With m_Form1
   menu_Tools .mnu_Tools, .mnu_ListOption
   menu_shadow .mnu_ListOption, .mnu_Tools, .shadow_mnu, .mnu_Shadow
End With
menu_Name = "Tools"
End Sub
Private Sub m_form1_HelpClick()
With m_Form1
   menu_Help .mnu_Help, .mnu_ListOption
   menu_shadow .mnu_ListOption, .mnu_Help, .shadow_mnu, .mnu_Shadow
End With
menu_Name = "Help"
End Sub
Public Sub menu_File(ByRef mnu_Menu As Label, _
                  ByVal mnu_List As ListBox)
With mnu_List
  .Clear
  .AddItem "Exit"
  .AddItem "Back-Up"
End With
With m_Form1
 mnuList_Height .mnu_ListOption, .mnu_File
End With
End Sub
Public Sub menu_Menu(ByRef mnu_Menu As Label, _
                  ByVal mnu_List As ListBox)
 With mnu_List
  .Clear
  .AddItem "<< Main Menu >>"
  .AddItem "Instant Report"
  .AddItem "Product List"
  .AddItem "Stocked Received"
  .AddItem "Payroll System"
  .AddItem "---------------------------------"
  .AddItem "Sub Menu 1"
  .AddItem "Sub Menu 2"
  .AddItem "Sub Menu 3"
  .AddItem "Sub Menu 4"
End With
With m_Form1
 mnuList_Height .mnu_ListOption, .mnu_Menu
End With
End Sub
Public Sub menu_Tools(ByRef mnu_Menu As Label, _
                  ByVal mnu_List As ListBox)
With mnu_List
  .Clear
  .AddItem "Calculator"
  .AddItem "Calendar"
End With
With m_Form1
 mnuList_Height .mnu_ListOption, .mnu_Tools
End With
End Sub

Public Sub menu_Help(ByRef mnu_Menu As Label, _
                  ByVal mnu_List As ListBox)
With mnu_List
  .Clear
  .AddItem "Contents                  F1"
  .AddItem "Index"
  .AddItem "Search"
  .AddItem "---------------------------------"
  .AddItem "Contact Us"
End With
With m_Form1
 mnuList_Height .mnu_ListOption, .mnu_Help
End With
End Sub
Private Sub mnuList_Height(ByRef lst As ListBox, _
                           ByRef lbl As Label, _
                           Optional ByVal lstText As Integer = 260)
Dim xlist As Integer
    xlist = lst.ListCount
    lst.Height = xlist * lstText
    lst.Width = 2100
    lst.Visible = True

    With lbl
      lst.Top = .Top + .Height + 50
      lst.Left = .Left
      lst.SetFocus
    End With
    lst.BackColor = vbWhite
    lst.ForeColor = vbBlack
End Sub

Private Sub menu_shadow(ByVal lst As ListBox, _
                       ByVal lbl As Label, _
                       ByVal sha_mnu As Shape, _
                       ByVal mnu_sha As PictureBox)
Dim sl As Integer, st As Integer, sh As Integer, sw As Integer
Dim mnu_l As Integer, mnu_h As Integer, mnu_w As Integer, mnu_t As Integer
With lbl
   sl = .Left
   st = .Top
   sw = .Width
   sh = .Height
   With sha_mnu
     .Visible = True
     .Left = sl - 50
     .Top = st - 40
     .Height = sh + 80
     .Width = sw + 100
   End With
End With
With lst
     mnu_l = lst.Left
     mnu_w = lst.Width
     mnu_h = lst.Height
     mnu_t = lst.Top
    With mnu_sha
       .Visible = True
       .Left = mnu_l + 50
       .Top = mnu_t + 50
       .Width = mnu_w
       .Height = mnu_h
    End With
End With  'lst
End Sub




