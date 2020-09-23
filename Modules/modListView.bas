Attribute VB_Name = "modListView"
Option Explicit
'//FUNCTION TO HANDLE UNIQUE KEY FOR NEW ENTRY (numeric value)
'// this will determine the largest reference number even if not in order
'//coded by edwin delos santos
Public Function Last_Recc(ByRef rs As Recordset) As Long
Dim maxRecc As Long, maxExist As Boolean
Dim LastRec As Long
Dim recExist As Long
'// initialize
  maxRecc = 0
  LastRec = 0
  maxExist = False
With rs
  If .RecordCount = 0 Then
      Last_Recc = 1
      Exit Function
  End If
   LastRec = .RecordCount + 1
  .MoveFirst
  If Not IsNumeric(.Fields(0)) Then Exit Function    '// determine if the first field is numeric
  While Not .EOF()
     If maxRecc > recExist Then
        maxRecc = maxRecc
        maxExist = True
     End If
     recExist = .Fields(0)       '//current number encountered
     If recExist > LastRec Then
        Last_Recc = recExist + 1
        If maxExist = False Then
          maxRecc = recExist     '// Rem determine the largest number >
        End If
: Rem debug.print maxrecc
     ElseIf recExist = LastRec Then
       Last_Recc = LastRec + 1
     ElseIf recExist < LastRec Then
       Last_Recc = LastRec
    End If
    .MoveNext
  Wend
  If maxExist = True Then
    Last_Recc = maxRecc + 1
  End If
End With
End Function

'Procedure used to fill listview with data
'coded by edwin delos santos
Public Sub FillListView(ByRef sListView As ListView, ByRef sRecSource As Recordset, ByVal sIcoNdx As Byte)
'//set details
    Dim X As ListItem
    Dim i As Byte
    Dim sFieldsNum As Integer
    On Error Resume Next
    '//initialize
    sListView.ListItems.Clear
    sFieldsNum = (sRecSource.Fields.Count - 1)
    If sRecSource.RecordCount < 1 Then Exit Sub
    sRecSource.MoveFirst
    Do While Not sRecSource.EOF
         Set X = sListView.ListItems.Add(, , sRecSource.Fields(0), sIcoNdx, sIcoNdx)
         For i = 1 To sFieldsNum
               X.SubItems(i) = FormatRS(sRecSource.Fields(CInt(i)))
          Next i
        sRecSource.MoveNext
    Loop
End Sub

'Procedure used to show in listview what you have already added
'need not refresh every time you add new record
'coded by edwin delos santos
'<< syntax >> PopulateLvw lvname, rs, 2
Public Sub lvwPopulateData(ByRef sListView As ListView, ByRef sRecSource As Recordset, ByVal sIcoNdx As Byte)
'//set details
    Dim X As ListItem
    Dim i As Byte
    Dim sFieldsNum As Integer
    On Error Resume Next
    '//initialize
    sFieldsNum = (sRecSource.Fields.Count - 1)  '(sRecSource.Fields.Count - 1)
With sRecSource
    .Requery    '//Use this method to make sure that a Recordset contains the most recent data
                  'first record becomes the current record
    .MoveLast
End With
    If sRecSource.RecordCount < 1 Then Exit Sub
         Set X = sListView.ListItems.Add(, , sRecSource.Fields(0), sIcoNdx, sIcoNdx)
         For i = 1 To sFieldsNum   'sub items must start from 1 not zero / cause of error Invalid Property Value err.number 380
            If Not IsNull(sRecSource.Fields(CInt(i))) Then
               X.SubItems(i) = FormatRS(sRecSource.Fields(CInt(i)))
            End If
          Next i
End Sub

'//procedure to show in listview what you have edited
'//need not refresh every time you edit record
'//<< syntax >> Call LvwReplaceData(Me, rs, lvname)
Public Sub LvwReplaceData(ByRef frm As Form, _
                      ByRef rs As Recordset, _
                      ByRef lv As ListView, _
                      Optional ByVal numOfFlds As Integer = 0)
Dim i As Integer
Dim NOF As Integer  'number of fields
If numOfFlds > 0 Then
   NOF = numOfFlds
Else
   NOF = (rs.Fields.Count - 1)  'remember that indeces are zero based
End If
For i = 1 To NOF
  lv.SelectedItem.ListSubItems(i).text = frm.txtEntry(i).text
  Next i
End Sub

'//Procedure used to Insert column
Public Sub InsertColumn(ByRef lv As ListView, _
                        ByVal sRecordSource As Recordset, _
                        Optional ByVal sNumFields As Integer, _
                        Optional CH_clear As Boolean = True)
    Dim X As String
    Dim i As Integer
    Dim idx As Integer 'index use to align column right
    Dim sNumOfColumn As Integer  'number of fields
    Dim clmHead As ColumnHeader
If sNumFields > 0 Then
   sNumOfColumn = sNumFields
Else
    sNumOfColumn = (sRecordSource.Fields.Count - 1)
End If
    '// initialize value
    If CH_clear = True Then
       lv.ColumnHeaders.Clear
    End If
    On Error Resume Next
         For i = 0 To sNumOfColumn
            X = splitChar(sRecordSource.Fields.Item(i).Name)
             Set clmHead = lv.ColumnHeaders.Add(, , X)
             '// align column data to right if it is currency or double
             '// no modify below this commented lines
             If sRecordSource.Fields.Item(i).Type = 6 Or _
                sRecordSource.Fields.Item(i).Type = 5 Then
                 idx = i + 1
                    lv.ColumnHeaders(idx).Alignment = lvwColumnRight
             End If
         Next i
End Sub

'// procedure to Search in listview//
'//<< syntax >>
Public Sub ListView_Search(ByRef Lvw As ListView, _
                           ByVal sFind As String, _
                           Optional ByVal valSetting = 1)
Rem valSeeting :>> 0 = lvwtext ; 1 = lvwsubitem
'//input exact string ...
Dim itmFound As ListItem
If valSetting = 0 Then
  Set itmFound = Lvw.FindItem(sFind, 0, 1, 1)
Else
  Set itmFound = Lvw.FindItem(sFind, 1, 1, 1)
End If
  If Not itmFound Is Nothing Then
    itmFound.EnsureVisible
    itmFound.Selected = True
 End If
End Sub
'//procedure to delete record
'//reference for deletion is the  unique key
Public Sub Delete_Record(ByRef srcRS As Recordset, ByRef lvName As ListView)
Dim abPos As Boolean
Dim itemStr As Variant
Dim ans As Integer
Dim strMatch As String
Dim toDelete As String
'// INTIALIZE
toDelete = ""
strMatch = ""
If srcRS.RecordCount = 0 Then Exit Sub
If srcRS.EOF = True Or srcRS.BOF = True Then Exit Sub
itemStr = lvName.SelectedItem.text
If IsNumeric(TrimSpaces(CStr(lvName.SelectedItem.text))) Then
   toDelete = TrimSpaces(CStr(lvName.SelectedItem.text))
   abPos = False
Else
    toDelete = lvName.SelectedItem.Index
    abPos = True
End If
ans = MsgBox("Are you Sure you want to delete selected item#:" & "( " & itemStr & ")" & "?", vbYesNo, "Delete")
If ans = vbYes Then
  With srcRS
     If .RecordCount = 0 Then Exit Sub
    .MoveFirst
     While Not .EOF
     If abPos = False Then
        lvName.MousePointer = vbHourglass
        strMatch = TrimSpaces(CStr(toNumber(srcRS.Fields(0))))
     Else
       lvName.MousePointer = vbHourglass
       'slower//i use only on alpha type// so you can show the value one
       'row even if there is duplicate reference for viewing record
       'remember that reference must be a unique key
       strMatch = srcRS.AbsolutePosition
   End If
        '// if record found
        If toDelete = strMatch Then
            '//delete current record
            srcRS.Delete
            lvName.ListItems.Remove lvName.SelectedItem.Index
            lvName.SetFocus
            lvName.MousePointer = vbDefault
            Exit Sub
        Else
          .MoveNext '//if record not found
       End If
     Wend
  End With  'rsprod
  
ElseIf ans = vbNo Then
  MsgBox "Deletion Cancelled!", , "Delete!"
End If
  lvName.MousePointer = vbDefault
End Sub


Public Sub SortListView(ByVal Lvw As MSComctlLib.ListView, _
                        ByVal colHdr As MSComctlLib.ColumnHeader)
'//Sort/ReSort ListView by the clicked column
'//<< syntax >>  SortListView ListView1, ColumnHeader

'//Sort by clicked ListView Column
'--set the sortkey to the column header's index - 1
Lvw.SortKey = colHdr.Index - 1
Lvw.Sorted = True

'--toggle the sort order between ascending & descending
Lvw.SortOrder = 1 Xor Lvw.SortOrder
End Sub
