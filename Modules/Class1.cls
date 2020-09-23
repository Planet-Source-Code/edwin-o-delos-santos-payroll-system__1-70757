VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CNumConvert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Private Data(9, 3) As String

Private Sub Class_Initialize()
    'Data for conversion
    Data(0, 0) = "ONE": Data(1, 0) = "TWO": Data(2, 0) = "THREE"
    Data(3, 0) = "FOUR": Data(4, 0) = "FIVE": Data(5, 0) = "SIX"
    Data(6, 0) = "SEVEN": Data(7, 0) = "EIGHT": Data(8, 0) = "NINE"
    Data(9, 0) = "TEN"
    Data(0, 1) = "HUNDRED": Data(1, 1) = "TEN": Data(2, 1) = "TWENTY"
    Data(3, 1) = "THIRTY": Data(4, 1) = "FORTY": Data(5, 1) = "FIFTY"
    Data(6, 1) = "SIXTY": Data(7, 1) = "SEVENTY": Data(8, 1) = "EIGHTY"
    Data(9, 1) = "NINETY"
    Data(0, 3) = "TEN": Data(1, 3) = "ELEVEN": Data(2, 3) = "TWELVE"
    Data(3, 3) = "THIRTEEN": Data(4, 3) = "FOURTEEN": Data(5, 3) = "FIFTEEN"
    Data(6, 3) = "SIXTEEN": Data(7, 3) = "SEVENTEEN": Data(8, 3) = "EIGHTEEN"
    Data(9, 3) = "NINETEEN"

End Sub
Public Function ToWords(ByVal NumberStr As String) As String
    Dim z As String, x As String, Temp As String, c As String
    Dim a As Integer, b As Integer, i As Integer
    Dim iPos As Integer
    On Error GoTo errhdl:
    'remove redundant spaces
    NumberStr = Trim(Replace(NumberStr, ",", ""))
    a = Len(NumberStr)
    Temp = NumberStr
    If Val(NumberStr) = 0 Then
        ToWords = "zero!"
        Exit Function
    End If
    
    'get rid of any decimals
    iPos = InStr(Temp, ".")
    If iPos > 0 Then Temp = Left(Temp, iPos - 1)
    
    
    While ((a Mod 3) <> 0)
        Temp = "0" & Temp
        a = Len(Temp)
    Wend
    NumberStr = Temp
    For i = a - 2 To 1 Step -3
        b = b + 1
        Temp = Mid(NumberStr, i, 3)
        z = ""
        '  "Intelligent" routines
        '------------------------
        If Temp <> "000" Then
            c = Left(Temp, 1)
            If c <> "0" Then z = " " & Data(Val(c) - 1, 0) & "HUNDRED"
            c = Mid(Temp, 2, 1)
            If c <> "0" Then
                If c <> "1" Then
                    z = z & " " & Data(Val(c), 1)
                Else
                    z = z & " " & Data(Val(Right(Temp, 2)) - 10, 3)
                End If
            End If
            If Right(Temp, 1) <> "0" And Mid(Temp, 2, 1) <> "1" Then z = z & " " & Data(Val(Right(Temp, 1)) - 1, 0)
        End If
        '------------------------
        If z <> "" Then
            Select Case b
                Case 1:
                    x = z
                Case 2:
                    x = z & " THOUSAND" & x
                Case 3:
                    x = z & " MILLION" & x
                Case 4:
                    x = z & " BILLION" & x
                Case 5:
                    x = z & " TRILLION" & x
                Case Else:
                    Exit Function
                'you can easily add more range
                'like Case 6: can be "zillion"? :) (whatever)
            End Select
        End If
    Next
    ToWords = x
errhdl:
   If err.Number = 9 Then
     MsgBox "Please check the amount!", 64, "Error!!!"
     Exit Function
   End If
End Function
Private Function Replace(ByVal sInput As String, _
   sFind As String, sReplace As String) As String
 
'USED HERE INSTEAD OF BUILT-IN REPLACE FUNCTION
'SO THAT CLASS WILL WORK WITH VB5
Dim lPos As Long
Dim sAns As String
Dim sWkg As String

sAns = ""
sWkg = sInput

lPos = InStr(sWkg, sFind)


If lPos <> 0 Then
  
  
    Do
        If lPos >= Len(sWkg) Then
           sAns = sAns & Left(sWkg, Len(sWkg) - 1) & sReplace
        Else
            sAns = sAns & Left(sWkg, lPos - 1) & sReplace
       End If
        sWkg = Mid(sWkg, lPos + 1)
        lPos = InStr(sWkg, sFind)
        DoEvents
    Loop While lPos > 0
    sAns = sAns & sWkg
Else
    sAns = sInput
End If

Replace = sAns

End Function

