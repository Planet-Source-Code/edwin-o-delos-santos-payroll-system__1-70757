Attribute VB_Name = "modCmdButton"
Option Explicit
Public Enum cmdButtons
  BtnAdd = 0
  BtnSave = 1
  BtnEdit = 2
  BtnUpdate = 3
  BtnCancel = 4
  BtnDelete = 5
  BtnRefresh = 6
End Enum


Public Sub cmdButtonShow(ByRef buttonString As String, ByRef srcForm As Form)
'< syntax:  cmdButtonShow ("0001111")
''--------------------------------------------------
''-- This routine handles setting the enabled --
''-- to true / false on the buttons.                --
''-------------------------------------------------
''-- A string of 0101 passed. If 0, disabled   --
''-------------------------------------------------
Dim indx As Integer
buttonString = Trim$(buttonString)
For indx = 1 To Len(buttonString)
  If (Mid$(buttonString, indx, 1) = "1") Then
     srcForm.cmdButton(indx - 1).Enabled = True
  Else
    srcForm.cmdButton(indx - 1).Enabled = False
  End If
Next

End Sub

