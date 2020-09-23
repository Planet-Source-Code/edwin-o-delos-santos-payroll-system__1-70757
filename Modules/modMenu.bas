Attribute VB_Name = "modMenu"
Option Explicit
Public sMENU(2008) As String 'handle MENU SELECTION
Public sAdmin(2008) As String 'handle whether user is admin or not
Public pIndex As Integer 'to handle menu index
'//to handle menu caption source tbl_menu
Public iMenuCap(30) As String   'do not make the length fixed
Public iAccess As Boolean

Public Sub allowACCESS(ByRef currMENU As String, frm As Form, Optional ByVal Max As Boolean = False)
'//currMenu = current menu loaded
'// frm = form
'// max = maximize (true/false)
Dim i As Integer
'//initial value /declared public in this module
iAccess = False
For i = 0 To 2008
    If UCase(currMENU) = UCase(sMENU(i)) Then
      iAccess = True
      i = 0
      GoTo openSESAME
     End If
  Next i
openSESAME:
If iAccess = False Then
     MsgBox "Access Denied..." & vbCrLf & _
     "You are not authorized !" & vbCrLf & _
     "Keep out! ", vbCritical, "Warning!"
     Exit Sub
Else
    Load frm
    frm.show 'vbModal, MainForm
    If Max Then
       frm.WindowState = vbMaximized
    Else
      frm.WindowState = vbNormal
    End If
    frm.SetFocus
End If

End Sub
Public Sub allowSHOW(ByRef currMENU As String, ByRef ctl As Control)
'//currMenu = current menu loaded
'// frm = form
'// max = maximize (true/false)
Dim i As Integer
'//initial value /declared public in this module
iAccess = False
For i = 0 To 2008
    If UCase(currMENU) = UCase(sMENU(i)) Then
      iAccess = True
      i = 0
      GoTo openSESAME
     End If
  Next i
openSESAME:
If iAccess = False Then
     MsgBox "Access Denied..." & vbCrLf & _
     "You are not authorized !" & vbCrLf & _
     "Keep out! ", vbCritical, "Warning!"
     Exit Sub
Else
   ctl.Visible = True
   ctl.SetFocus
End If

End Sub

Public Sub Menu_List(ByRef srcRS As Recordset, ByVal sPASS As String)
'//USE ARRAY SMENU(100) TO STORE MENU / DECLARED AS PUBLIC
'//sPASS = the current user password upon login
Dim iCount As Integer
If srcRS.State = adStateOpen Then
   srcRS.Close
End If
iCount = 0
'MainForm.List1.Clear

srcRS.Open "select * from users where password like '" & Trim$(sPASS) & "'order by menu", Con, 1, 1, 1
With srcRS
 .MoveFirst
 Do While Not .EOF
     sMENU(iCount) = .Fields("MENU")
     '//TEST
     'MainForm.List1.AddItem sMENU(iCount)
     .MoveNext
     iCount = iCount + 1
 Loop
  
iCount = 0
End With

End Sub

Public Sub Menu_Caption(ByRef srcRSmenu As Recordset)
'//USE ARRAY iMenuCap(100) TO STORE MENU / DECLARED AS PUBLIC global
Dim iCount As Integer
If srcRSmenu.State = adStateOpen Then
   srcRSmenu.Close
End If
iCount = 0
srcRSmenu.Open "select * from tbl_menu order by PK", CN, adOpenStatic, adLockReadOnly

With srcRSmenu
 .MoveFirst
 Do While Not .EOF
     iMenuCap(iCount) = .Fields("MENU")
    .MoveNext
     iCount = iCount + 1
 Loop
  
iCount = 0
Set srcRSmenu = Nothing

End With

End Sub


Public Function user_List(ByRef recset As Recordset, _
                          ByRef flds As String, _
                          ByVal cbo As Control) As String
Dim txt1 As String
txt1 = "!@#$%^&*()"
If recset.RecordCount = 0 Then Exit Function
If recset.RecordCount > 0 Then
    recset.MoveFirst
    cbo.Clear
   While Not recset.EOF
        If Not IsNull(recset.Fields(flds)) Then
          If recset.Fields(flds) <> txt1 Then
            cbo.AddItem recset.Fields(flds)
          End If
        End If
        If Not IsNull(recset.Fields(flds)) Then
           txt1 = recset.Fields(flds)
        End If
        recset.MoveNext
   Wend
End If
End Function



'// show hotkey on button
'// use by lblmenu()
Public Sub HK(ByRef Btn As Label, lbl As Label)
 lbl.Visible = True
 lbl.Left = Btn.Left
 lbl.Top = Btn.Top
 lbl.Caption = Btn.Caption
 lbl.ZOrder
End Sub

