Attribute VB_Name = "modAutoComplete"
Option Explicit
'Public Execute As Boolean
'Public Mytextbox As TextBox



Public Sub strCOMPLETE(Lvw As ListView, sFind, Mytextbox As String)
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

