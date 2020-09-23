Attribute VB_Name = "ModLockControl"
Option Explicit

Public Sub lock_me_not(pfx As String, tf As Boolean)
'/======================================================/
'pfx = prefix
'TXT - textbox
'LBL - label
'TAB - xtab
'SYNTAX : lock_me_not("TAB", False)
'/=======================================================/
Dim ctl As Control
pfx = UCase$(pfx)
Select Case pfx
Case Is = "TXT"
For Each ctl In Form1.Controls
  If TypeOf ctl Is TextBox Then
   If tf = False Then
    ctl.Enabled = False
   Else
    ctl.Enabled = True
   End If
  End If
  Next ctl
Case Is = "TAB"
For Each ctl In Form1.Controls
  If TypeOf ctl Is XTab Then
   If tf = False Then
    ctl.Enabled = False
   Else
    ctl.Enabled = True
   End If
  End If
  Next ctl
  
End Select
End Sub

