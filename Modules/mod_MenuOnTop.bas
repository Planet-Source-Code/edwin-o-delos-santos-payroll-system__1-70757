Attribute VB_Name = "mod_Menu"
Option Explicit

Public mnu_Name As String  '//handle current menu name (see mnu_listOption  keydown events)

Public Sub x_File(ByRef mnu_Form As Form, _
                  ByRef mnu_Menu As Label, _
                  ByVal mnu_List As ListBox)
mnu_Name = mnu_Menu
With mnu_Form
 mnu_List.Clear
 mnu_List.AddItem "Exit"
 mnu_List.AddItem "Back-Up"
End With
mnuList_Height mnu_List, mnu_Menu
End Sub
Public Sub x_Menu(ByRef mnu_Form As Form, _
                  ByRef mnu_Menu As Label, _
                  ByVal mnu_List As ListBox)
mnu_Name = mnu_Menu
 With mnu_Form
 mnu_List.Clear
 mnu_List.AddItem "<< Main Menu >>"
 mnu_List.AddItem "Instant Report"
 mnu_List.AddItem "Product List"
 mnu_List.AddItem "Stocked Received"
 mnu_List.AddItem "Payroll System"
 mnu_List.AddItem "---------------------------------"
 mnu_List.AddItem "Sub Menu 1"
 mnu_List.AddItem "Sub Menu 2"
 mnu_List.AddItem "Sub Menu 3"
 mnu_List.AddItem "Sub Menu 4"
End With
mnuList_Height mnu_List, mnu_Menu
End Sub
Public Sub x_Tools(ByRef mnu_Form As Form, _
                  ByRef mnu_Menu As Label, _
                  ByVal mnu_List As ListBox)
mnu_Name = mnu_Menu
With mnu_Form
 mnu_List.Clear
 mnu_List.AddItem "Calculator"
 mnu_List.AddItem "Calendar"
End With
mnuList_Height mnu_List, mnu_Menu
End Sub

Public Sub x_Help(ByRef mnu_Form As Form, _
                  ByRef mnu_Menu As Label, _
                  ByVal mnu_List As ListBox)
mnu_Name = mnu_Menu
With mnu_Form
 mnu_List.Clear
 mnu_List.AddItem "Contents                  F1"
 mnu_List.AddItem "Index"
 mnu_List.AddItem "Search"
 mnu_List.AddItem "---------------------------------"
 mnu_List.AddItem "Contact Us"
End With
mnuList_Height mnu_List, mnu_Menu
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

Public Sub menu_shadow(ByRef frm As Form, _
                       ByVal lst As ListBox, _
                       ByVal lbl As Label, _
                       ByVal sha_mnu As Shape, _
                       ByVal mnu_sha As Shape)
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

