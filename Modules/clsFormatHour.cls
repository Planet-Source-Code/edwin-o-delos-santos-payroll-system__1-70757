VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private dHour As Double   'hour/s
Private dTHour As Double  'Totalhour/s converted to value

Private Sub Class_Initialize()
  dHour = 0
End Sub
Public Function HourToDbl(ByRef TmIn As String, ByRef TmOut As String) As Double
dHour = 0
On Error Resume Next
If Val(TmIn) = 0 Or Val(TmOut) = 0 Then
  Exit Function
End If
dHour = ((DateDiff("n", CDate(TmIn), CDate(TmOut)) / 60))
HourToDbl = Format(dHour, "fixed")
End Function
Public Function toDate(ByRef sDate As String, _
                       Optional ByVal sSep As String = "/") As String
'//TEST string 01102007  - RESULT 01/10/2007

Dim iyear As String, imonth As String, iday As String
If Mid(sDate, 3, 1) <> "/" Then
 imonth = Mid(sDate, 1, 2)
 iday = Mid(sDate, 3, 2)
 iyear = Mid(sDate, 5, 4)
      If Len(sDate) = 8 Then
           toDate = imonth & sSep & iday & sSep & iyear
      ElseIf Len(sDate) >= 11 Then
        If Mid(sDate, 4, 1) = "-" Then
          imonth = Mid(sDate, 1, 3)
          iday = Mid(sDate, 5, 2)
          iyear = Mid(sDate, 8, 4)
          sSep = "-"
          toDate = imonth & sSep & iday & sSep & iyear   'jan-10-2008
        End If
     End If
ElseIf Mid(sDate, 3, 1) = "/" And Mid(sDate, 6, 1) = "/" Then
   imonth = Mid(sDate, 1, 2)
   iday = Mid(sDate, 4, 2)
   iyear = Mid(sDate, 7, 4)
   If Len(sDate) = 10 Then
         toDate = imonth & sSep & iday & sSep & iyear
    End If
End If
  If Not IsDate(toDate) Then
     MsgBox "Invalid Date", vbCritical, "Warning!"
  End If

End Function

Public Function ToHour(ByRef sHrs As String, Optional ByVal Sptr As String = ":") As String
Dim N1 As String, N2 As String
'/test string  07:30
If Mid(sHrs, 3, 1) = Sptr Then ':
  N1 = Mid(sHrs, 1, 2)  '07
  N2 = Mid(sHrs, 4, 2)  '30
ElseIf Len(sHrs) = 4 Then
  N1 = Mid(sHrs, 1, 2)  '07
  N2 = Mid(sHrs, 3, 2)  '30
End If
 ToHour = N1 & Sptr & N2
 If Not IsDate(ToHour) Then
    MsgBox "Invalid!", vbCritical, "Time!"
    Exit Function
 End If
End Function



'Funtion to check if the time is correct
Public Function isTime(ByRef sTime As String) As Boolean
If Not IsDate(sTime) Then
     isTime = False
ElseIf Len(sTime) = 0 Then
     isTime = False
ElseIf sTime = "00:00" Then
     isTime = False
     Exit Function
ElseIf Val(Mid(sTime, 1, 2)) > 12 Then
     isTime = False
ElseIf Val(Mid(sTime, 4, 2)) > 59 Then
     isTime = False
ElseIf Mid(sTime, 3, 1) = ":" And Len(sTime) = 5 Then
    isTime = True
End If
If isTime = False Then
   MsgBox "Invalid Time!", vbCritical, "Enter the correct Time"
  Exit Function
End If
End Function


