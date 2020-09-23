Attribute VB_Name = "Modules_Global"
Option Explicit
Public entries() As TextBox   'dynamic array for entries
Public TextWd As Long  '//used by textExtend
Public Expand As Boolean  '//use by textextend

'//procedure to write_data to a current table
'// can be called by any updatable table
'//be sure that textbox is in dynamic array
'//coded by edwin delos santos
'<<syntax>>  call writedata(me,rs,true) or writedata(me,rs,true,15) 15 is the upperbound indeces
Public Sub WriteData(ByRef frm As Form, ByRef srcRS As Recordset, _
                      ByVal addNEW As Boolean, _
                      Optional ByVal srcNumFlds As Integer = 0)
'//addnew = true for new record else false > forced
'//srcnumflds = number of fields loaded in textbox  > optional
                'if not all fields are loaded, srcnumflds is equal to text upperbound indeces
                'based on the numbers of textbox showed in the form (see enabled textbox procedures)
Dim i As Integer
Dim NOF As Integer 'Number Of Feilds
If srcNumFlds > 0 Then
   NOF = srcNumFlds
Else
   NOF = (srcRS.Fields.Count - 1)  'remember that indeces are zero based
End If
ReDim entries(NOF) As TextBox
For i = 0 To NOF
    Set entries(i) = frm.txtEntry(i)  'm tired of using frm, set number of elements allowed
    Next i
i = 0
With srcRS
  If addNEW = True Then
      .addNEW
  End If
      For i = 0 To NOF
      Select Case srcRS.Fields.Item(i).Type
       Case Is = 3   'integer
           If IsNumeric(entries(i).text) Then
              srcRS.Fields(i) = toNumber(entries(i).text)
           End If
      Case Is = 5, 6  'currency or double
           If IsNumeric(entries(i).text) Then
             srcRS.Fields(i) = toMoney(entries(i).text)
           End If
       Case Is = 7   'date
           If IsDate(entries(i).text) Then
               srcRS.Fields(i) = CDate(entries(i).text)
           Else '//save empty entry
               srcRS.Fields(i) = Null
           End If
       Case Is = 202, 203    'text, memo
             srcRS.Fields(i) = CStr(entries(i).text)
      End Select
      Next i
      .Update
End With
End Sub

'//procedure to bind data into textbox control
'//coded by edwin delos santos
'//can be called by any table to view data widely
'//be sure that textbox is in dynamic array
'<<syntax>> BindDatasource(me,rs,lvname,true) or BindDatasource(me,rs,lvname,true,15) 15 is nof
Public Sub BindDatasource(ByRef frm As Form, _
                          ByRef srcRS As Recordset, _
                          ByRef lv As ListView, _
                          Optional ByVal findFirst As Boolean = True, _
                          Optional ByVal numOfFlds As Integer = 0)
'//findFIRST - optional/false when use for next,previous,last,first
Dim abPos As Boolean   'absolutePosition
Dim i As Integer
Dim strFind As String
Dim strMatch As String
Dim NOF As Integer 'Number Of Feilds
'// initialized
If numOfFlds > 0 Then
   NOF = numOfFlds
Else
   NOF = (srcRS.Fields.Count - 1)  'remember that indeces are zero based
End If
For i = 0 To NOF
   frm.txtEntry(i) = Empty
   Next i
If IsNumeric(TrimSpaces(CStr(lv.SelectedItem.text))) Then
    strFind = TrimSpaces(CStr(lv.SelectedItem.text))
    abPos = False
Else
    strFind = lv.SelectedItem.Index
    abPos = True
End If
If findFirst = True Then
 With srcRS
 .MoveFirst
   Do Until srcRS.EOF
   If abPos = False Then
        lv.MousePointer = vbHourglass
       strMatch = TrimSpaces(CStr(toNumber(srcRS.Fields(0))))
     Else
       lv.MousePointer = vbHourglass
       'slower//i use only on alpha type// so you can show the value one
       'row even if there is duplicate reference for viewing record
       'remember that reference must be a unique key
       strMatch = srcRS.Bookmark '// .AbsolutePosition
   End If
   If strMatch = strFind Then
         lv.MousePointer = vbDefault

         GoTo iFound
   Else
     .MoveNext
   End If
   Loop
 End With
 lv.MousePointer = vbDefault
End If 'findFirst
iFound:
With srcRS
         If srcRS.EOF = True Or srcRS.BOF = True Then Exit Sub
         For i = 0 To NOF
          If Not IsNull(srcRS.Fields(i)) Then
             frm.txtEntry(i) = FormatRS(srcRS.Fields(i))
              If srcRS.Fields(i).Type = 6 Or srcRS.Fields(i).Type = 5 Then
                frm.txtEntry(i).Alignment = 1
                 If Val(frm.txtEntry(i)) = 0 Then
                   frm.txtEntry(i).ForeColor = &HD38545
                 ElseIf Val(frm.txtEntry(i)) < 0 Then
                   frm.txtEntry(i).ForeColor = vbRed      ' if the value is negative
                 Else
                   frm.txtEntry(i).ForeColor = vbBlack
                End If
             Else                                          'string value
                 frm.txtEntry(i).ForeColor = vbBlack
            End If
          Else
              frm.txtEntry(i) = Empty
          End If
         Next i
    '//end of Search
End With

End Sub
'//procedure to show field label
'//coded by edwin delos santos
'//be sure that LABEL is in dynamic array
'<<syntax>> showfldslabel(me,rs) or showfldslabel(me,rs,15) 15 is nof
Public Sub ShowFldsLabel(ByRef frm As Form, _
                         ByRef srcRS As Recordset, _
                         Optional ByVal numOfFlds As Integer = 0)
Dim i As Integer
Dim splitCHR As String
Dim NOF As Integer 'Number Of Feilds
'// initialized
If numOfFlds > 0 Then
   NOF = numOfFlds
Else
   NOF = (srcRS.Fields.Count - 1)  'remember that indeces are zero based
End If
For i = 0 To NOF
      frm.lblFLDi(i) = ""              'caption for each field
   Next i
i = 0
For i = 0 To NOF
        If Not IsNull(srcRS.Fields(i)) Then
              splitCHR = splitChar(srcRS.Fields(i).Name, "_")
              frm.lblFLDi(i) = splitCHR & " :"
              'frm.lblFLDi(i) = srcRS.Fields(i).Name & " :"
        Else
               frm.lblFLDi(i) = srcRS.Fields(i).Name & " :"
        End If
        Next i
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

'//Procedure used to locked selected textbox
'// requirement: listbox <> input: numeric
Public Sub TextBox_Locked(frm As Form, ByRef idxList As ListBox)
    Dim i As Integer
    Dim idx As Integer
    On Error Resume Next
    For i = 0 To idxList.ListCount - 1
           idx = Val(idxList.List(i))                   'get the value from listbox, you can use array if you want
           frm.txtEntry(idx).Locked = True
           frm.txtEntry(idx).BackColor = &HE6FFFF
         Next i
End Sub

'// procedure to make textbox available
Public Sub TextBox_Visible(ByRef frm As Form, ByVal rs As Recordset)
  Dim i As Integer
  Dim numba As Integer
  i = 0
      For i = 0 To frm.txtEntry.UBound             'make all textbox visible
               frm.txtEntry(i).Visible = False
               frm.lblFLDi(i).Visible = False
          Next i
  i = 0
  numba = (rs.Fields.Count - 1)
      For i = 0 To numba                           'make number of textbox available
               frm.txtEntry(i).Visible = True
               frm.lblFLDi(i).Visible = True
          Next i
End Sub

Public Sub center_obj(ByVal sParentObj As Variant, ByRef sMoveObj As Variant)
    sMoveObj.Left = (sParentObj.Width - sMoveObj.Width) / 2
    sMoveObj.Top = (sParentObj.Height - sMoveObj.Height) / 2
End Sub

Public Sub CenterObjt(ByRef objt As Object)
Dim X As Integer, Y As Integer
Y = (Screen.Height - objt.Height) / 2
X = (Screen.Width - objt.Width) / 2
objt.Move X, Y
End Sub
Public Sub selectAllText()
'// select all text in textbox on gotfocus
 SendKeys "{home}+{end}"
End Sub
Public Sub center_obj_horizontal(ByVal sParentObj As Variant, ByRef sMoveObj As Variant)
    sMoveObj.Left = (sParentObj.Width - sMoveObj.Width) / 2
End Sub

'//Procedure used to checked selected item in listbox
Public Sub DefaultList(ByRef lst As ListBox, ByVal lstPrn As ListBox, ByRef idxList As ListBox)
'//rem: idxlist > list content must be numeric
    Dim i As Integer
    Dim idx As Integer
    '// initialize
    On Error Resume Next
    For i = 0 To idxList.ListCount - 1
           idx = Val(idxList.List(i))
           lst.Selected(idx) = True
           lstPrn.AddItem lst.List(idx)
         Next i
End Sub

Public Sub TextExtend(cText As TextBox)
  Dim txt As Variant
  Dim tayp As String
  Dim str As String
  If Expand = False Then
     TextWd = cText.Width
  End If
  txt = cText.text
  If IsNumeric(txt) Then
     txt = Val(txt)
  End If
  If IsDate(txt) Then
     txt = CDate(txt)  'result mm/dd/yyyy format
  End If
  str = "String"
  tayp = TypeName(txt)   'result  String
  If tayp = str Then
     cText.Width = 5000
     cText.BackColor = &HEFF4D7
     cText.SelStart = 0
     cText.SelStart = Len(cText)
     Expand = True
  Else
     cText.Width = TextWd
     cText.BackColor = vbWhite
     Expand = False
  End If
End Sub


Public Sub Bind_Data(ByRef frm As Form, _
                          ByRef srcRS As Recordset, _
                          Optional ByVal numOfFlds As Integer = 0)
'//findFIRST - optional/false when use for next,previous,last,first
Dim abPos As Boolean   'absolutePosition
Dim i As Integer
Dim NOF As Integer 'Number Of Feilds
'// initialized
If numOfFlds > 0 Then
   NOF = numOfFlds
Else
   NOF = (srcRS.Fields.Count - 1)  'remember that indeces are zero based
End If
For i = 0 To NOF
   frm.txtEntry(i) = Empty
   Next i
 
 With srcRS
         If srcRS.EOF = True Or srcRS.BOF = True Then Exit Sub
         For i = 0 To NOF
          If Not IsNull(srcRS.Fields(i)) Then
             frm.txtEntry(i) = FormatRS(srcRS.Fields(i))
              If srcRS.Fields(i).Type = 6 Or srcRS.Fields(i).Type = 5 Then
                frm.txtEntry(i).Alignment = 1
                 If Val(frm.txtEntry(i)) = 0 Then
                   frm.txtEntry(i).ForeColor = &HD38545
                 ElseIf Val(frm.txtEntry(i)) < 0 Then
                   frm.txtEntry(i).ForeColor = vbRed      ' if the value is negative
                 Else
                   frm.txtEntry(i).ForeColor = vbBlack
                End If
             Else                                          'string value
                 frm.txtEntry(i).ForeColor = vbBlack
            End If
          Else
              frm.txtEntry(i) = Empty
          End If
         Next i
         Exit Sub
 End With
End Sub


