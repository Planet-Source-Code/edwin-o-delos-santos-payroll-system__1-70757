Attribute VB_Name = "ModProcedure"
Option Explicit

'// procedure to autocomplete textbox in listview search//
Public Sub Autocomplete(Lvw As ListView, sFind, Mytextbox As TextBox)
Dim Lvfindtm As ListItem
Dim TempSelStart As Integer
Dim strTemp As String

Set Lvfindtm = Lvw.FindItem(sFind, lvwText, , lvwPartial)
If Not Lvfindtm Is Nothing Then
Lvfindtm.EnsureVisible
Lvfindtm.Selected = True

If Execute Then
TempSelStart = Mytextbox.SelStart
Mytextbox.text = CStr(Lvfindtm)
If Not Mytextbox.text = "" Then
Mytextbox.SelStart = TempSelStart
Mytextbox.SelLength = Len(Mytextbox.text) - TempSelStart
    End If
        End If
            End If
End Sub

'// enabled or disabled command button//
Public Sub showBUTTON(hkey As String, frm As Form, Optional btnDel As Boolean, Optional btnRef As Boolean)
'where hkey is hotkey
   Select Case hkey
    Case Is = "A"
      frm.CmdEdit.Enabled = False
      frm.CmdAdd.Enabled = False
      frm.CmdSave.Enabled = True
      frm.CmdCancel.Enabled = True
      If btnDel = True Then
         frm.CmdDelete.Enabled = False
      End If
    Case Is = "S"
      frm.CmdAdd.Enabled = True
      frm.CmdSave.Enabled = False
      frm.CmdEdit.Enabled = True
      frm.CmdCancel.Enabled = False
      If btnDel = True Then
         frm.CmdDelete.Enabled = True
      End If
    Case Is = "E"
      frm.CmdAdd.Enabled = False
      frm.CmdEdit.Enabled = False
      frm.CmdUpdate.Enabled = True
      frm.CmdCancel.Enabled = True
      If btnDel = True Then
         frm.CmdDelete.Enabled = False
      End If
    Case Is = "U"
      frm.CmdEdit.Enabled = True
      frm.CmdAdd.Enabled = True
      frm.CmdUpdate.Enabled = False
      frm.CmdCancel.Enabled = False
      If btnDel = True Then
         frm.CmdDelete.Enabled = True
      End If
    Case Is = "C"
      frm.CmdAdd.Enabled = True
      frm.CmdSave.Enabled = False
      frm.CmdEdit.Enabled = True
      frm.CmdUpdate.Enabled = False
      frm.CmdCancel.Enabled = False
      If btnDel = True Then
         frm.CmdDelete.Enabled = False
      End If
   End Select
End Sub

'//Procedure used to search in listview//
Public Sub search_in_listview(ByRef sListView As ListView, ByVal sFindText As String)
    Dim tmp_listtview As ListItem
    Set tmp_listtview = sListView.FindItem(sFindText, lvwSubItem)
    If Not tmp_listtview Is Nothing Then
        tmp_listtview.EnsureVisible
        tmp_listtview.Selected = True
    End If
End Sub

'// procedure to resize controls to their permanent position//
Public Sub ctrl_Rsize(ctl As Control, LF As Long, ht As Long, tp As Long, wd As Long)
 If TypeOf ctl Is Frame Then GoTo align
 If TypeOf ctl Is ListBox Then GoTo align
 If TypeOf ctl Is PictureBox Then GoTo align
 If TypeOf ctl Is Label Then GoTo align
 If TypeOf ctl Is Frame Then GoTo align
 If TypeOf ctl Is ListView Then GoTo align
 If TypeOf ctl Is TreeView Then GoTo align
 If TypeOf ctl Is ProgressBar Then GoTo align
align:
With ctl
 .Left = LF
 .Height = ht
 .Top = tp
 .Width = wd
End With
End Sub
'// procedure to access menu //
Public Sub menuAccess(currUSR As String, myMNU As String, usr As String, mnu As String, fm As Form, Max As Boolean)
Dim currUSER As String
                'purpose  : to limit user to show all form
                'parameter:
                'currUSER      = CURRENT USER
                'myMNU         = menu
                'mnu           = menu storage
                'fm            = form to load
currUSER = UCase$(mnu) & UCase$(usr)

If TrimSpaces(currUSER) = TrimSpaces(UCase$(myMNU) & UCase$(currUSR)) Then
   accss = True
Else
  accss = False
End If
If accss = False Then
     MsgBox "Access Denied..." & vbCrLf & _
     "You are not authorized !" & vbCrLf & _
     "Keep out! ", vbCritical, "Warning!"
     Exit Sub
Else
    Load fm
    fm.show
    If Max Then
       fm.WindowState = vbMaximized
    Else
      fm.WindowState = vbNormal
    End If
    fm.SetFocus
'    Main.WindowState = 1
End If
End Sub

'//load records to combobox at log in user form //
Public Sub SelUser(recset As Recordset, flds As String, sql As String, cbo As Object)

If recset.State = adStateOpen Then
  recset.Close
End If
recset.Open sql
If recset.RecordCount = 0 Then Exit Sub
If recset.RecordCount > 0 Then
    recset.MoveFirst
    cbo.Clear
   While Not recset.EOF
        If Not IsNull(recset.Fields(flds)) Then
          cbo.AddItem recset.Fields(flds)
        End If
        recset.MoveNext
   Wend
End If
End Sub

Public Sub DisplayErrorMsg()
 If err.Number = 0 Then
   Exit Sub
 ElseIf err.Number = 13 Then
   MsgBox "Data type mismatch !", vbCritical, "DATE"
   Exit Sub
 ElseIf err.Number = 3021 Then
   MsgBox "Current Record has been deleted !"
   Exit Sub
 ElseIf err.Number = 9 Then
   MsgBox "Subscrip out of range !", vbInformation, "DblClick"
   Exit Sub
 ElseIf err.Number = 3265 Then
   Exit Sub
 ElseIf err.Number = 7005 Then
   MsgBox "RowSet not available! click refresh!", vbInformation, "Click"
   Exit Sub
  Else
    MsgBox "Error Code: " & err.Number & vbCrLf & _
        "Description: " & err.Description & vbCrLf & _
        "Source: " & err.Source, vbOKOnly + vbCritical
 End If
End Sub
Public Sub Shadow(ctl As Control, sh As Frame)
 If TypeOf ctl Is Frame Then GoTo align
 If TypeOf ctl Is ListBox Then GoTo align
 If TypeOf ctl Is PictureBox Then GoTo align
 If TypeOf ctl Is Frame Then GoTo align
 If TypeOf ctl Is ListView Then GoTo align
 If TypeOf ctl Is TreeView Then GoTo align
align:
With ctl
 sh.Left = .Left = -20
 sh.Height = .Height
 sh.Top = .Top + 50
 sh.Width = .Width
End With
End Sub



Public Function unlockd(frm As Form)
Dim ctl As Control
For Each ctl In frm.Controls
   If TypeOf ctl Is TextBox Then ctl.Locked = False
   Next ctl
End Function


' =====================================================================================================================
'
' Function:     EnhListView_Find
'
' Imputs:
'               Variable Name       Type        Optional    Description
'               lstListName         ListView    No          Name of the ListView to find in
'               strStringToFind     String      No          What to find in the list
'               bolWholeWordOnly    Boolean     Yes         Only 'find' if the 'found' is exactly like the 'search'
'               bolCaseSensitive    Boolean     Yes         Only 'find' if the 'found' is the same case as the 'search'
'
' Returns:      Integer of the 'found' item
'               Also selects the item and makes sure the item is visible
'
' =====================================================================================================================
Public Function EnhListView_Find(lstListName As ListView, _
                                strStringToFind As String, _
                                Optional bolWholeWordOnly As Boolean, _
                                Optional bolCaseSensitive As Boolean) _
                                As Integer
    ' setup variables
    Dim lngIndex As Long        ' used for the current index of the parent items
    Dim lngIndexSub As Long     ' used for the current index of the subitems
    Dim strCurrItem As String   ' used to store the text of the currently selected item for compare
    
    ' if we want to be sensitive about the case then make the 'search' all upper case
    If bolCaseSensitive = True Then strStringToFind = UCase(strStringToFind)
    
    ' set the return to the default zero
    EnhListView_Find = 0
    
    ' if there is nothing to search then exit
    If lstListName.ListItems.Count < 1 Then Exit Function
    
    ' if no item is currently selected then select the first item
    If lstListName.SelectedItem.Index = -1 Then lstListName.SelectedItem.Index = 1
    
    ' move through the rows
    For lngIndex = lstListName.SelectedItem.Index - -1 To lstListName.ListItems.Count
        
        ' if we want to be sensitive about the case then...
        If bolCaseSensitive = True Then
            ' fill our variable with the uppercase version of the current text
            strCurrItem = UCase(lstListName.ListItems.Item(lngIndex).text)
        Else
            ' otherwise, fill our variable with the current text
            strCurrItem = lstListName.ListItems.Item(lngIndex).text
        End If
        
        If bolWholeWordOnly = True Then
            ' if the current item and the 'search' is an exact match then finalize
            If strCurrItem = strStringToFind Then GoTo Finalize
        Else
            ' if the current item contains the 'search' then finalize
            If InStr(strCurrItem, strStringToFind) > 0 Then GoTo Finalize
        End If
        
        ' if we have subitems...
        If lstListName.ColumnHeaders.Count > 1 Then
            
            ' move through the subitems of the current row
            For lngIndexSub = 1 To lstListName.ColumnHeaders.Count - 1
                ' if we want to be sensitive about the case then...
                If bolCaseSensitive = True Then
                    ' fill our variable with the uppercase version of the current text
                    strCurrItem = UCase(lstListName.ListItems.Item(lngIndex).SubItems(lngIndexSub))
                Else
                    ' otherwise, fill our variable with the current text
                    strCurrItem = lstListName.ListItems.Item(lngIndex).SubItems(lngIndexSub)
                End If
                
                If bolWholeWordOnly = True Then
                    ' if the current item and the 'search' is an exact match then finalize
                    If strCurrItem = strStringToFind Then GoTo Finalize
                Else
                    ' if the current item contains the 'search' then finalize
                    If InStr(strCurrItem, strStringToFind) > 0 Then GoTo Finalize
                End If
            ' move to next subitem
            Next lngIndexSub
        
        End If
        
    ' move to next row
    Next lngIndex
    
    Exit Function
    
Finalize:
    EnhListView_Find = lngIndex                             ' send back the index of the found item
    lstListName.ListItems.Item(lngIndex).EnsureVisible      ' make sure the item is visible
    lstListName.ListItems.Item(lngIndex).Selected = True    ' make sure the item is selected
End Function
' =====================================================================================================================
'// form resize
Public Sub frm_rsize(ByRef frm As Form, ByVal frm_ht As Long, ByVal frm_wd As Long)
 With frm
     .Height = frm_ht
     .Width = frm_wd
 End With
End Sub

'Procedure used to center vertical /BY PHILIP NAPARAN
Public Sub center_obj_vertical(ByVal sParentObj As Variant, ByRef sMoveObj As Variant)
    sMoveObj.Top = (sParentObj.Height - sMoveObj.Height) / 2
End Sub
Public Sub center_obj_horizontal(ByVal sParentObj As Variant, ByRef sMoveObj As Variant)
    sMoveObj.Left = (sParentObj.Width - sMoveObj.Width) / 2
End Sub

'Used to locate the key in opened form /BY PHILIP NAPARAM
Public Sub HighlightInWin(ByVal srcKey As String)
    With MAIN.lvWin
        If .ListItems.Count > 0 Then
            If .SelectedItem.Key <> srcKey Then
                Dim c As Integer
                For c = 1 To .ListItems.Count
                    If .ListItems(c).Key = srcKey Then
                        .ListItems(c).Selected = True
                        .ListItems(c).EnsureVisible
                        Exit For
                    End If
                Next c
            End If
        End If
    End With
End Sub
'by philip naparan (modified)
Public Sub loadForm(ByRef srcForm As Form, Max As Boolean)
    srcForm.show
    If Max Then
       srcForm.WindowState = vbMaximized
    Else
      srcForm.WindowState = vbNormal
    End If
    srcForm.SetFocus
End Sub

'// show label hotkey button
Public Sub lblhotK(ByRef lblBtn As Object, hotK As Object)
 hotK.Visible = True
 hotK.Left = lblBtn.Left
 hotK.Width = lblBtn.Width
 hotK.Top = (lblBtn.Top + lblBtn.Height) - 20
 hotK.ZOrder
End Sub

'// show hotkey on button
Public Sub HK(ByRef Btn As Object, ByRef hkey As Object)
 hkey.Visible = True
 hkey.Left = Btn.Left
 hkey.Top = (Btn.Top + Btn.Height) - 40
 hkey.ZOrder
End Sub



