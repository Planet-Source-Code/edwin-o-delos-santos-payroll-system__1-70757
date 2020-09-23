Attribute VB_Name = "modListCombo"
Option Explicit
'dynamic array used to store data
Public item_added() As Variant  'used by add_item, already_added proc
'//Procedure used to print fields name in listbox
Public Sub Insert_Fields(ByRef ctlLIST As Control, ByRef sRecordSource As Recordset)
    Dim X As String
    Dim i As Integer
    Dim sNumOfFields As Integer
    '// initialize value
    sNumOfFields = (sRecordSource.Fields.Count - 1)
    ctlLIST.Clear
    On Error Resume Next
         For i = 0 To sNumOfFields
             X = sRecordSource.Fields.Item(i).Name
             ctlLIST.AddItem X
             Next i
        i = i + 1
End Sub
'// additem without duplication
Public Sub Add_Item(ByRef recset As Recordset, _
                     ByRef fld As String, _
                     ByRef ctl As Control, _
                     Optional TrapDup As Boolean = False)
Dim uBnd As Long   'upperbound indeces/number of elements
Dim icnt As Integer 'count number of records
Dim txt1 As String
Dim i As Integer
'initialize
If recset.RecordCount = 0 Then Exit Sub
uBnd = (recset.RecordCount - 1)
ReDim item_added(uBnd)
For i = 0 To uBnd     'recset.RecordCount - 1
    item_added(i) = Empty
    Next i
txt1 = "!@#$%^&*()"
icnt = 0
On Error Resume Next
If recset.RecordCount = 0 Then Exit Sub
If recset.RecordCount > 0 Then
    recset.MoveFirst
    ctl.Clear
   While Not recset.EOF
      If Not IsNull(recset.Fields(fld)) Then
         If recset.Fields(fld) <> txt1 Then
           If TrapDup = True Then
              Dim X As Boolean
              X = alReady_Added(recset, recset.Fields(fld))
              If X = False Then
                ctl.AddItem recset.Fields(fld) 'add only one record to listbox
              End If
              item_added(icnt) = recset.Fields(fld)  'continue add to array
            Else
                ctl.AddItem recset.Fields(fld)
            End If
         End If
        If Not IsNull(recset.Fields(fld)) Then
           txt1 = recset.Fields(fld)
        End If
      End If  'not isnull
        icnt = icnt + 1
        recset.MoveNext
   Wend
End If
End Sub
'function to check if record already added to listbox
'//coded by: edwin delos santos
Private Function alReady_Added(ByRef rs As Recordset, ByRef srcStr As String) As Boolean
  Dim i As Integer
On Error Resume Next
For i = 0 To rs.RecordCount - 1
    'find and match record from array; if found added = true
    If srcStr = item_added(i) Then
      alReady_Added = True
      i = 0
      Exit Function
    Else
      alReady_Added = False
    End If
  Next i
End Function
'function to check if record already exist
'//coded by: edwin delos santos
Public Function isExist(ByRef rs As Recordset, ByRef whatfld As String, ByRef findStr As String) As Boolean
Dim matchValue As String
Dim i As Integer
On Error Resume Next
'initialize
matchValue = Empty
With rs
    .MoveFirst
 While Not rs.EOF = True
   matchValue = rs.Fields(whatfld)
   If TrimSpaces(findStr) = TrimSpaces(matchValue) Then
      isExist = True
     Exit Function
   Else
     isExist = False
   End If
  rs.MoveNext
 Wend
End With
End Function

Public Sub lstCalendar(ByRef lst As ListBox, ByRef srcTxt As String)
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
lst.AddItem m & "/" & D & "/" & Y
  Next D
End Sub


'//Call prnCENTERTEXT(x, 150)
Public Function prnCenterText(ByRef txt As String, ByVal b_line As Integer) As String
Dim s_col As Integer
s_col = (b_line - Len(txt)) / 2
Printer.Print Tab(s_col); txt
End Function






