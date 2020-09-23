Attribute VB_Name = "ModFunction"
Option Explicit
Public valTB() As Double   'used by sum

Public Function TrimSpaces(text As String) As String
    Dim Loop1 As Long, SpaceCheck As String
    Dim FullString As String
    For Loop1 = 1 To Len(text)
        SpaceCheck = Mid(text, Loop1, 1)
        If SpaceCheck <> " " Then
            FullString = FullString & SpaceCheck
        End If
    Next Loop1
    TrimSpaces = FullString
End Function

'Function used to format recordset
'/coded by edwin delos santos
Public Function FormatRS(ByVal srcField As Field) As String
    Dim strRet As String
     With srcField
        If srcField.Type = adCurrency Or srcField.Type = adDouble Then
            strRet = Format$(srcField, "#,###,##0.00")
        ElseIf srcField.Type = 7 Then
            strRet = Format$(srcField, "MMM-dd-yyyy")
        ElseIf srcField.Type = 3 Then
           If IsNumeric(srcField) Then
             strRet = Format$(srcField, "###,##0")
           End If
        ElseIf srcField.Type = 202 Or srcField.Type = 203 Then
            strRet = CStr(srcField)
        End If
    End With
    FormatRS = strRet
    strRet = vbNullString
End Function
'//Function used to display MONTHNAME
Public Function Month_Name(ByVal srcdate As Date) As String
Dim MonthNames As Variant
Dim moName As String
If Not IsDate(srcdate) Then Exit Function
MonthNames = Array("January", "February", "March", "April", "May", "June", _
                  "July", "August", "September", "October", "November", "December")
moName = Month(Format(srcdate, "MM/DD/YYYY"))
Month_Name = MonthNames(moName - 1)
End Function


Public Function WeekDay_Name(ByVal srcdate As Date) As String
Dim daynames() As Variant
If Not IsDate(srcdate) Then Exit Function
daynames = Array("Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday")
Dim wkDay As String
wkDay = Weekday(Format(srcdate, "MM/DD/YYYY"))
WeekDay_Name = daynames(wkDay - 1)
End Function

Public Function splitChar(ByRef str As String, Optional ByRef chr As String = "_") As String
Dim iChar As Integer
Dim mystring
Dim sResult As String
iChar = InStr(1, str, chr, 1)  'search for the char "_"
mystring = Split(str, chr, -1, 1)
If iChar > 0 Then
  sResult = mystring(0) & Space(1) & mystring(1)
Else
  sResult = str
End If
  splitChar = sResult
End Function



