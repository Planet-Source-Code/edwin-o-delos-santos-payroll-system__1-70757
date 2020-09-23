VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "numTOword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'set up two arrays to convert numbers to words
Option Explicit
Private BigOnes(9) As String
Private SmallOnes(19) As String

Private Sub Class_Initialize()
'populate
  BigOnes(1) = "TEN"
  BigOnes(2) = "TWENTY"
  BigOnes(3) = "THIRTY"
  BigOnes(4) = "FORTY"
  BigOnes(5) = "FIFTY"
  BigOnes(6) = "SIXTY"
  BigOnes(7) = "SEVENTY"
  BigOnes(8) = "EIGHTY"
  BigOnes(9) = "NINETY"
  SmallOnes(1) = "ONE"
  SmallOnes(2) = "TWO"
  SmallOnes(3) = "THREE"
  SmallOnes(4) = "FOUR"
  SmallOnes(5) = "FIVE"
  SmallOnes(6) = "SIX"
  SmallOnes(7) = "SEVEN"
  SmallOnes(8) = "EIGHT"
  SmallOnes(9) = "NINE"
  SmallOnes(10) = "TEN"
  SmallOnes(11) = "ELEVEN"
  SmallOnes(12) = "TWELVE"
  SmallOnes(13) = "THIRTEEN"
  SmallOnes(14) = "FOURTEEN"
  SmallOnes(15) = "FIFTEEN"
  SmallOnes(16) = "SIXTEEN"
  SmallOnes(17) = "SEVENTEEN"
  SmallOnes(18) = "EIGHTEEN"
  SmallOnes(19) = "NINETEEN"
End Sub
Public Function TOword(ByVal amt As String) As String
Dim x As String
Dim Pesos As String, Cents As String, Words As String, Chunk As String
Dim digits As String, leftdigit As String, rightdigit As String
'where x  =  words
'format the incoming number to guarantee six digits STANDARD
amt = Format(amt, "000000.00")

Pesos = Left(amt, 6)
Cents = Right(amt, 2)

Words = ""

'make sure INCOMING number is not too large
If Pesos > 999999 Then
   x = "Pesos amount is too large"
Exit Function
End If

'separate pesos into chunks
If Pesos = 0 Then
  Words = "Zero"
Else
''//the thousands
Chunk = Left(Pesos, 3)
If Chunk > 0 Then
GoSub ParseChunk
  Words = Words & " THOUSAND"
End If

'//the pesos
Chunk = Right(Pesos, 3)

If Chunk > 0 Then
GoSub ParseChunk
End If
End If

'concatenate cents
If Cents = 0 Then Cents = "xx"
   Words = Words & " AND " & Cents & "/100"
x = Words
TOword = x
Exit Function

ParseChunk:
digits = Mid(Chunk, 1, 1)
If digits > 0 Then
    Words = Words & " " & SmallOnes(digits) & " HUNDRED"
End If

digits = Mid(Chunk, 2, 2)

If digits > 19 Then
leftdigit = Mid(Chunk, 2, 1)
rightdigit = Mid(Chunk, 3, 1)
Words = Words & " " & BigOnes(leftdigit)
If rightdigit > 0 Then
Words = Words & " " & SmallOnes(rightdigit)
End If

Else

 If digits > 0 Then
    Words = Words & " " & SmallOnes(digits)
 End If
End If
Return
 
End Function

