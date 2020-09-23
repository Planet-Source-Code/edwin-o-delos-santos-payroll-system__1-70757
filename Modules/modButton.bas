Attribute VB_Name = "modButton"
Option Explicit


'// enabled or disabled command button//
Public Sub showButton(currBtn As String, frm As Form, Optional btnDel As Boolean, Optional btnRef As Boolean)
'where currBtn is current Button
   Select Case currBtn
    Case Is = "A"
      addRec = True
      editRec = False
      frm.CmdEdit.Enabled = False
      frm.CmdAdd.Enabled = False
      frm.cmdSave.Enabled = True
      frm.cmdCancel.Enabled = True
      If btnDel = True Then
         frm.CmdDelete.Enabled = False
      End If
      If btnRef = True Then
        frm.cmdREFRESH.Enabled = False
      End If
    Case Is = "S"
      addRec = False
      frm.CmdAdd.Enabled = True
      frm.cmdSave.Enabled = False
      frm.CmdEdit.Enabled = True
      frm.cmdCancel.Enabled = False
      If btnDel = True Then
         frm.CmdDelete.Enabled = True
      End If
      If btnRef = True Then
        frm.cmdREFRESH.Enabled = True
      End If
      
    Case Is = "E"
      editRec = True
      addRec = False
      frm.CmdAdd.Enabled = False
      frm.CmdEdit.Enabled = False
      frm.CmdUpdate.Enabled = True
      frm.cmdCancel.Enabled = True
      If btnDel = True Then
         frm.CmdDelete.Enabled = False
      End If
      If btnRef = True Then
        frm.cmdREFRESH.Enabled = False
      End If
      
    Case Is = "U"
      editRec = False
      frm.CmdEdit.Enabled = True
      frm.CmdAdd.Enabled = True
      frm.CmdUpdate.Enabled = False
      frm.cmdCancel.Enabled = False
      If btnDel = True Then
         frm.CmdDelete.Enabled = True
      End If
      If btnRef = True Then
        frm.cmdREFRESH.Enabled = True
      End If
    Case Is = "C"
      addRec = False
      editRec = False
      frm.CmdAdd.Enabled = True
      frm.cmdSave.Enabled = False
      frm.CmdEdit.Enabled = True
      frm.CmdUpdate.Enabled = False
      frm.cmdCancel.Enabled = False
      If btnDel = True Then
         frm.CmdDelete.Enabled = True
      End If
      If btnRef = True Then
        frm.cmdREFRESH.Enabled = True
      End If
   End Select
End Sub

'// show hotkey on commandbutton
Public Sub Btn_Focus(ByRef Btn As Object, ByRef hkey As Object, ByRef hkey2 As Object)
 hkey.Visible = True
 hkey.Left = Btn.Left
 hkey.Top = Btn.Top
'//hkey2 right marker
 hkey2.Visible = True
 hkey2.Left = Btn.Width
 hkey2.Top = Btn.Top
End Sub

