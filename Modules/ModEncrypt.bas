Attribute VB_Name = "ModEncrypt"
Option Explicit
Public xPASS
Public tmpPASS As String     'temporary passowrd storage

'Variable structure for user
'//curr_user (see modPublic)
Public Type USER_INFO
    USER_NAME As String
    USER_ID As String
    USER_isADMIN As String '//[Y]/[N]
    USER_PASS As String
    user_MENU As String
End Type

Public Sub encrypt(passCHR As String)
On Error Resume Next
Dim maxLEN
Dim mPASS
Dim P_word
P_word = passCHR
maxLEN = Len(P_word)
Dim x
x = 1
xPASS = ""
Do While x <= maxLEN
  mPASS = Asc(Mid(P_word, x, 1))
  mPASS = mPASS - 64
  mPASS = 128 + mPASS + (x * 2)
  xPASS = xPASS + Chr(mPASS)
  x = x + 1
Loop
End Sub

Public Sub decrypt(passWRD As String)
On Error Resume Next
Dim P_word
Dim maxLEN
Dim mPASS
Dim xuserPASS
Dim x
xuserPASS = passWRD
maxLEN = Len(xuserPASS)
x = 1
xPASS = ""
Do While x <= maxLEN
  mPASS = Asc(Mid(xuserPASS, x, 1))
  mPASS = mPASS - 128 - (x * 2) + 64
  xPASS = xPASS + Chr(mPASS)
  x = x + 1
Loop
End Sub


