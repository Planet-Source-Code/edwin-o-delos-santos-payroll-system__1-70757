Attribute VB_Name = "ModPrint"
'//coder : Edwin O. Delos Santos
Option Explicit

Public Sub Print_Details(ByRef srcRS As Recordset, ByVal tabStart As Integer, ByRef lst As ListBox)
Dim strvalue As String
Dim fldName As String
Dim currentTab As Integer
Dim i As Integer
On Error Resume Next
currentTab = tabStart
  For i = 0 To lst.ListCount - 1
     If lst.Selected(i) = True Then
       fldName = Empty
       strvalue = Empty
       fldName = printIndex(i)
       strvalue = srcRS.Fields(fldName) '//srcRS.Fields(i) *replaced
       If srcRS.Fields(fldName).Type = 6 Or srcRS.Fields(fldName).Type = 5 Then
            Call PrintValue(currentTab, strvalue, 12)
            currentTab = currentTab + 15
        ElseIf srcRS.Fields(fldName).Type = 7 Then
            Printer.Print Tab(currentTab); strvalue;
           currentTab = currentTab + 15
        ElseIf srcRS.Fields(fldName).Type = 3 Then
            Printer.Print Tab(currentTab); strvalue;
            currentTab = currentTab + 13
        Else
            Printer.Print Tab(currentTab); strvalue;
            currentTab = currentTab + 35
        End If 'isnumeric
            End If  ' selected = true
    Next i
         currentTab = tabStart
End Sub

'//procedure to align value to the right: as in  31220.00
'//                                                200.00
'//original coding by myself
'//no modify within this proc
'---------------------------------------------------------
Public Sub PrintValue(ByRef iTab As Integer, ByVal srcValue As String, maxLEN As Integer)
'//Remarks: maxlen must be equal to maxlen with Print_Headings
'//currLen(15) declared public
 Dim intLEN As Integer, currtab As Integer
 Dim strvalue As Double
 Dim i As Integer
  intLEN = Len(Format(srcValue, "#,###,##0.00"))
  For i = 0 To maxLEN
       currLen(i) = i
          If currLen(i) = intLEN Then
            currtab = iTab + (maxLEN - intLEN)
            GoSub printerP
          End If
      Next i
        i = i + 1
'//sub
printerP:
  If Val(srcValue) > 0 Then
   Printer.Print Tab(currtab); Format(srcValue, "#,###,##0.00");
  Else
     Printer.Print Tab(currtab); "--";
  End If
End Sub
Public Function Print_Amt(ByRef iTab As Integer, ByVal amt As Double, maxLEN As Integer) As Double
 Dim intLEN As Integer, currtab As Integer

 intLEN = Len(Trim(Format(amt, "Standard")))
 Dim i As Integer
  intLEN = Len(Format(amt, "#,###,##0.00"))
  For i = 0 To maxLEN
       currLen(i) = i
          If currLen(i) = intLEN Then
            currtab = iTab + (maxLEN - intLEN)
            GoSub printerP
          End If
      Next i
        i = i + 1
'//sub
printerP:
  If Val(amt) > 0 Then
   Printer.Print Tab(currtab); Format(amt, "#,###,##0.00");
  Else
     Printer.Print Tab(currtab); "--";
  End If
End Function

Public Sub Print_Headings(ByRef srcRS As Recordset, _
                         ByVal tabStart As Integer, _
                         ByRef lst As ListBox, _
                         ByVal maxLEN As Integer)
'//Remarks: maxlen must be equal to maxlen with PrintValue
Dim strvalue As String
Dim currentTab As Integer
Dim fldName As String  'hanlde feidlName
Dim heading As String
currentTab = tabStart
Dim hdng As String
On Error Resume Next
'// recordset required
Dim i As Integer
  For i = 0 To lst.ListCount - 1
     If lst.Selected(i) = True Then
       fldName = Empty
       heading = Empty
       fldName = printIndex(i)
       heading = Mid(printIndex(i), 1, maxLEN)
       strvalue = Empty
       strvalue = srcRS.Fields(fldName)
       If srcRS.Fields(fldName).Type = 6 Or srcRS.Fields(fldName).Type = 5 _
          Or srcRS.Fields(fldName).Type = 7 Then      '//If IsNumeric(strvalue) Then *replaced
            Printer.Print Tab(currentTab); heading;
            currentTab = currentTab + 15
       ElseIf srcRS.Fields(fldName).Type = 3 Then
            Printer.Print Tab(currentTab); heading;
            currentTab = currentTab + 13
       Else
           Printer.Print Tab(currentTab); heading;
            currentTab = currentTab + 35
        End If
            End If
    Next i
         currentTab = tabStart

End Sub

Public Function Screen_Total(ByRef srcRS As Recordset, whichField As Integer) As String
'// which index value to compute Total
Dim strvalue As String
Dim iCount As Integer
Dim i As Integer
Dim NumOfFields As Integer
'//initialize value
NumOfFields = (srcRS.Fields.Count - 1)
ReDim dblTotal(NumOfFields) As Double
For i = 0 To NumOfFields
   dblTotal(i) = 0
   Next i
iCount = 0
With srcRS
  .MoveFirst
If srcRS.EOF = True Or srcRS.BOF = True Then Exit Function
While Not .EOF = True
iCount = iCount + 1
  For i = 0 To NumOfFields
       strvalue = Empty
      If Not IsNull(srcRS.Fields(i)) Then
        strvalue = srcRS.Fields(i)
        If IsNumeric(strvalue) Then
           dblTotal(i) = dblTotal(i) + Val(strvalue)   'array/declare public
            If iCount = .RecordCount Then
              If i = whichField Then
                  Screen_Total = dblTotal(i)
              End If
          End If
        End If 'isnumeric
       End If 'not isnull
    Next i
.MoveNext
Wend
End With
End Function


'//function to initialized printing ...
'//coded by: edwin delos santos
Public Function print_Init(ByRef lst As ListBox) As Boolean
    Dim i As Integer
    Dim ii As Integer
    On Error Resume Next
    '// initialize
     For ii = 0 To 50
         printIndex(ii) = Empty
         Next ii
    For i = 0 To lst.ListCount - 1
           lst.Selected(i) = False
         Next i
    '//
    If lst.ListCount = 0 Then Exit Function
    For i = 0 To lst.ListCount - 1
            lst.Selected(i) = True
            printIndex(i) = lst.text
         Next i
      print_Init = True
End Function


