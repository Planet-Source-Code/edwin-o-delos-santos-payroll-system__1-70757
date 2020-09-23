VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFormatStr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private dHour As Double   'hour/s
Private dTHour As Double  'Totalhour/s converted to value
Private Sub Class_Initialize()
  dHour = 0
  dTHour = 0
End Sub
Public Function ToValue(ByRef tm1 As String, ByRef tm2 As String) As Double
 dTHour = CDbl(tm1)
 dTHour = dTHour + CDbl(tm2)
 ToValue = Format(Val(Thour), "fixed")
End Function
Public Function TotalHour(ByRef TmIn As String, ByRef TmOut As String) As Double
Dim h As String
h = is_time(TmIn)
h = is_time(TmOut)
dHour = ((DateDiff("n", CDate(TmIn), CDate(TmOut)) / 60))
TotalHour = dHour
End Function

Public Function ToHour(ByRef sHrs As String, Optional ByVal Sptr As String = ":") As String
Dim N1 As String, N2 As String
Dim h As String
h = is_time(sHrs)
'/test string  07:30
N1 = Mid(sHrs, 1, 2)  '07
N2 = Mid(sHrs, 3, 2)  '30
If Len(sHrs) = 4 Then
   ToHour = N1 & Sptr & N2
End If
End Function
'Funtion to check if the time is correct
Private Function is_time(ByRef sTime As String) As Boolean
If sTime = 0 Then  'string if coerced to number ' thread as double
    is_time = False
    Exit Function
ElseIf sTime = "00:00" Then
     is_time = False
     MsgBox "Not a valid time!", vbCritical, "Warning!"
     Exit Function
ElseIf Mid(sTime, 3, 1) = ":" And Len(sTime) = 5 Then
    is_time = True
Else
   is_time = False
   MsgBox "Not a valid time!", vbCritical, "Warning!"
   Exit Function
End If
End Function


