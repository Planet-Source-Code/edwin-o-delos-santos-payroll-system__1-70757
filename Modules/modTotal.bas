Attribute VB_Name = "modTotal"
'//coded by EDWIN DELOS SANTOS
Option Explicit
Public dblTotal() As Double 'dynamic array to store total

Public Sub Print_Total(ByRef srcRS As Recordset, ByVal tabStart As Integer, ByRef lst As ListBox)
'// i prefer to use Tab() function rather that currentX/Y
Dim X As ListItem
Dim fldName As String
Dim strvalue As String
Dim currentTab As Integer
Dim iCount As Integer
Dim i As Integer
ReDim dblTotal(srcRS.Fields.Count - 1)
'//initialize value
For i = 0 To (srcRS.Fields.Count - 1)
   dblTotal(i) = 0
   Next i
   i = i + 1
iCount = 0
With srcRS
  .MoveFirst
While Not .EOF = True
iCount = iCount + 1
currentTab = tabStart
  For i = 0 To lst.ListCount - 1
     If lst.Selected(i) = True Then
       fldName = Empty
       strvalue = Empty
       fldName = printIndex(i)
       strvalue = srcRS.Fields(fldName)
       If srcRS.Fields(fldName).Type = 6 Or srcRS.Fields(fldName).Type = 5 Then
               dblTotal(i) = dblTotal(i) + Val(strvalue)   'array/declared public
               If iCount = .RecordCount Then
                  Call PrintValue(currentTab, dblTotal(i), 12) 'print total if EOF reached
               End If
           currentTab = currentTab + 15
        ElseIf srcRS.Fields(fldName).Type = 7 Then
           currentTab = currentTab + 15
        ElseIf srcRS.Fields(fldName).Type = 3 Then
           currentTab = currentTab + 13
        Else
            currentTab = currentTab + 35
        End If 'isnumeric
      End If  'selected = true
    Next i
         currentTab = tabStart
         i = i + 1
.MoveNext

Wend
End With
End Sub

'//function to view listview total
'//coder: edwin delos santos
Public Sub Listview_Total(ByRef Lvw As ListView, ByRef srcRS As Recordset)
Dim rec_Count As Long
Dim isCurr As Boolean    'flag for currency or double
Dim X As ListItem
Dim strvalue As String
Dim iCount As Long  'to determine the last record
Dim i As Integer
Dim NumOfFields As Integer
On Error Resume Next
'//initialize value
isCurr = False
NumOfFields = (srcRS.Fields.Count - 1)
ReDim dblTotal(NumOfFields) As Double    'number of elements
For i = 0 To NumOfFields
   dblTotal(i) = 0
   Next i
   i = i + 1
iCount = 0
With srcRS
   rec_Count = CStr(srcRS.RecordCount)
   Set X = Lvw.ListItems.Add(, , "(" & rec_Count & ")" & "Record")
       X.Bold = True
       X.ForeColor = vbBlue
  .MoveFirst
While Not .EOF = True
iCount = iCount + 1
  For i = 1 To NumOfFields
       strvalue = Empty
      If Not IsNull(srcRS.Fields(i)) Then
             If srcRS.Fields(i).Type = 6 Or srcRS.Fields(i).Type = 5 Then
                 strvalue = toMoney(srcRS.Fields(i))
                 isCurr = True
'             ElseIf srcRS.Fields(i).Type = 3 Then
'                 strvalue = toNumber(srcRS.Fields(i))
'                 isCurr = False
'             ElseIf srcRS.Fields(i).Type = 202 Or srcRS.Fields(i).Type = 203 Then
'                 If IsNumeric(srcRS.Fields(i).Value) Then
'                     strvalue = toNumber(srcRS.Fields(i))
'                     isCurr = False
'                  End If
             Else
                 strvalue = ""
             End If
             '// dblTotal(0),dblTotal(1),dblTotal(2) and so on ...
             If isCurr = True Then
                 dblTotal(i) = dblTotal(i) + Val(strvalue)
             End If
            If iCount = .RecordCount Then
                   If dblTotal(i) > 0 Then
                      If isCurr = True Then
                        With X.ListSubItems.Add(, , Format(dblTotal(i), "standard"))
                           X.ListSubItems(i).Bold = True
                           X.ListSubItems(i).ForeColor = vbRed
                        End With
                      Else
                        With X.ListSubItems.Add(, , toNumber(dblTotal(i)))
                           X.ListSubItems(i).Bold = True
                           X.ListSubItems(i).ForeColor = vbRed
                        End With
                      End If
                    Else  '//dbltotal() = 0
                      With X.ListSubItems.Add(, , " - ")
                      End With
                    End If
           End If   '//icount
       Else
        If iCount = .RecordCount Then
          With X.ListSubItems.Add(, , " - ") 'if null string
          End With
        End If
       End If 'not isnull
    Next i
.MoveNext
Wend
End With
Set X = Nothing
End Sub

