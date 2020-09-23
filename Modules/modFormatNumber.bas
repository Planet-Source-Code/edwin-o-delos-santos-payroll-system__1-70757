Attribute VB_Name = "modNumber"
'[=========================]
'< * toNumber - founction  >
'< * toMoney  - function   >
'< * isAmount - function   >
'[=========================]
Option Explicit
'// convert string to number / return the current value
'* Test   :  1,222
'* result :  1222
Public Function toNumber(ByVal srcNum As String) As Long
Dim numba As Long
If IsNumeric(srcNum) Then
  If Val(srcNum) = 0 Then srcNum = 0
  numba = Val(CLng(srcNum))
  toNumber = numba
  numba = 0
End If
End Function

'Function that will return a current format
'* Test  :  1,222.45
'* result:  1222.45
Public Function toMoney(ByVal srcCurr As String) As Double
 Dim sdbl As Double
 If IsNumeric(srcCurr) Then
   If Val(srcCurr) = 0 Then srcCurr = 0
   sdbl = Val(CDbl(srcCurr))
   toMoney = sdbl
   sdbl = 0
 End If
End Function

Public Function IsAmount(ByVal txt As String) As Boolean
    Dim ch As String
    Dim isamountentry As Boolean
    Dim i As Integer, j As Integer
    isamountentry = False
    
    If Len(LTrim(RTrim(txt))) = 0 Then
        IsAmount = False
        Exit Function
    End If
    j = 0
    For i = 1 To Len(txt)
        ch = Mid$(txt, i, 1)
        If ch < "0" Or ch > "9" Then
            If ch <> "." Then
                IsAmount = False
                Exit Function
            Else
                j = j + 1
            End If
        End If
    Next i
    If j > 1 Then
        Exit Function
    End If
    IsAmount = True
End Function
