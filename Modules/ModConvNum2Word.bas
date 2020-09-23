Attribute VB_Name = "ModFunctionNum2Word"

Option Explicit
Public conv As numTOword


Public Function ConvNum2Word(ByVal amt As String, lblW As Label)
Dim mAMOUNT As String
Dim iLEN As Integer
Dim potal As String
Dim dot As String
  iLEN = Len(amt)
  potal = Right(amt, 2)
  dot = Right(amt, 3)
  dot = Left(dot, 1)
  If dot = "." Then
     mAMOUNT = Mid(amt, 1, iLEN - 3)
     lblW = convert.ToWords(mAMOUNT)
     lblW = lblW & " & " & potal & "/100 ONLY"
  Else
     lblW = convert.ToWords(amt)
     lblW = lblW & " ONLY"
  End If

End Function



Public Function num_TO_word(ByVal amt As String, Ucase As Boolean) As String
Dim lbl As String
Dim Pesos As String, Cents As String, Words As String, Chunk As String
Dim digits As String, leftdigit As String, rightdigit As String
'first set up two arrays to convert numbers to words
Dim BigOnes(9) As String
Dim SmallOnes(19) As String
'and populate them
  If Ucase = True Then
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
Else
  BigOnes(1) = "Ten"
  BigOnes(2) = "Twenty"
  BigOnes(3) = "Thirty"
  BigOnes(4) = "Forty"
  BigOnes(5) = "Fifty"
  BigOnes(6) = "Sixty"
  BigOnes(7) = "Seventy"
  BigOnes(8) = "Eighty"
  BigOnes(9) = "Ninety"
  SmallOnes(1) = "One"
  SmallOnes(2) = "Two"
  SmallOnes(3) = "Three"
  SmallOnes(4) = "Four"
  SmallOnes(5) = "Five"
  SmallOnes(6) = "Six"
  SmallOnes(7) = "Seven"
  SmallOnes(8) = "Eight"
  SmallOnes(9) = "Nine"
  SmallOnes(10) = "Ten"
  SmallOnes(11) = "Eleven"
  SmallOnes(12) = "Twelve"
  SmallOnes(13) = "Thirteen"
  SmallOnes(14) = "Fourteen"
  SmallOnes(15) = "Fifteen"
  SmallOnes(16) = "Sixteen"
  SmallOnes(17) = "Seventeen"
  SmallOnes(18) = "Eighteen"
  SmallOnes(19) = "Nineteen"
End If
'format the incoming number to guarantee six digits
'to the left of the decimal point and two to the right
'and then separate the pesos from the cents
amt = Format(amt, "000000.00")

Pesos = Left(amt, 6)
Cents = Right(amt, 2)

Words = ""

'check to make sure incoming number is not too large
If Pesos > 999999 Then

lbl = "Dollar amount is too large"
Exit Function
End If

'separate the dollars into chunks
If Pesos = 0 Then
Words = "Zero"
Else
'first do the thousands
Chunk = Left(Pesos, 3)
If Chunk > 0 Then
GoSub ParseChunk
If Ucase = True Then
  Words = Words & " THOUSAND"
Else
 Words = Words & " Thousand"
End If
End If

'do the rest of the pesos
Chunk = Right(Pesos, 3)

If Chunk > 0 Then
GoSub ParseChunk
End If
End If

'concatenate the cents and display
If Cents = 0 Then Cents = "xx"
If Ucase = True Then
   Words = Words & " AND " & Cents & "/100"
Else
   Words = Words & " and " & Cents & "/100"
End If
lbl = Words
num_TO_word = lbl
Exit Function

ParseChunk:
digits = Mid(Chunk, 1, 1)
If digits > 0 Then
  If Ucase = True Then
    Words = Words & " " & SmallOnes(digits) & " HUNDRED"
  Else
   Words = Words & " " & SmallOnes(digits) & " Hundred"
  End If
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

