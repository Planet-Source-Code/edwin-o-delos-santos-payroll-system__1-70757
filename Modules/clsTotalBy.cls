VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTotalBy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private str_Value  As Double
Private idx As Integer
Private Sub Class_Initialize()
  str_Value = 0
  idx = 0
End Sub
Public Function Plus(ByVal value1 As String, ByVal value2 As String) As Double
    str_Value = 0
    str_Value = Val(toMoney(value1)) + Val(toMoney(value2))
    Plus = Format(str_Value, "standard")
End Function
Public Function Devide(ByVal value1 As String, ByVal value2 As String) As Double
    str_Value = 0
    str_Value = Val(toMoney(value1)) / Val(toMoney(value2))
    Devide = Format(str_Value, "standard")
End Function
Public Function Minus(ByVal value1 As String, ByVal value2 As String) As Double
    str_Value = 0
    str_Value = Val(toMoney(value1)) - Val(toMoney(value2))
     Minus = Format(str_Value, "standard")
End Function
Public Function Times(ByVal value1 As String, ByVal value2 As String) As Double
    str_Value = 0
    str_Value = Val(toMoney(value1)) * Val(toMoney(value2))
    Times = Format(str_Value, "standard")
End Function

Public Function Sum(ByRef frm As Form, _
                ByVal ndx1 As Integer, ndx2 As Integer) As Double
   idx = 0
   str_Value = 0
   For idx = ndx1 To ndx2
     str_Value = str_Value + Val(toMoney(frm.txtEntry(idx)))
    Next idx
    Sum = Format(str_Value, "standard")
End Function


Private Sub Class_Terminate()
idx = 0
str_Value = 0
End Sub
