Attribute VB_Name = "modDATe"
Option Explicit
Public Sub ListCalendar(ByRef lst As ListBox, ByRef srcTxt As String)
Dim D As Integer
Dim m As Integer
Dim Y As Integer
If IsDate(srcTxt) Then
  m = CDate(Month(srcTxt))
  Y = CDate(Year(srcTxt))
Else
  m = CDate(Month(Now))
  Y = CDate(Year(Now))
End If
For D = 1 To 31 Step 1
lst.AddItem m & "-" & D & "-" & Y
  Next D
End Sub
Public Sub AlignObj(ByRef currObj As Control, _
                     objToAlign As Control, _
                     CtlType As Integer, _
                     Optional show As Boolean = True)
  '1 listbox,listview
  '2. dtpdate, combobox
      With currObj
       If CtlType = 1 Then
         objToAlign.Top = .Top + .Height
         objToAlign.ZOrder (0)  'send to front
       ElseIf CtlType = 2 Then
         objToAlign.Top = .Top
         objToAlign.Width = .Width + 300
         objToAlign.ZOrder (1)  'send to back
       End If
        objToAlign.Left = .Left
        If show = True Then
          objToAlign.Visible = True
        Else
          objToAlign.Visible = False
        End If
      End With
End Sub
  

