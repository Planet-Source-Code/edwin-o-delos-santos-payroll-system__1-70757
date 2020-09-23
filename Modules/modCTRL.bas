Attribute VB_Name = "modSHOWMENU"
Option Explicit

'// procedure to show or hide rightside menu
'// original code by myself pls. vote !!
Public Sub Show_menu(Pic As PictureBox, xlf As Long, wd As Long, ByVal show As Boolean)
Dim i As Long
 '// wd - width when show or hide
 '// xlf - xtreme left
 Dim culf As Long 'current left position
 Dim cuwd As Long 'current width position
 Dim strt As Long 'starting position
 
 If show = False Then
   With Pic
     cuwd = .Width
     culf = .Left
     strt = culf - cuwd
     For i = strt To xlf Step -1
       .Left = i
       .Width = wd
       Next i
     End With
 Else   ' show = false
     With Pic
       culf = .Left
       For i = culf To xlf Step 1
       .Left = i
       .Width = wd
        Next i
      End With
End If
End Sub


'// procedure to show or hide POPUP CONTROL
'// original code by myself pls. vote !!
Public Sub POPUP_ME(Pic As PictureBox, xlf As Long, wd As Long, ByVal show As Boolean)
Dim i As Long
 '// wd - width when show or hide
 '// xlf - xtreme left
 Dim culf As Long 'current left position
 Dim cuwd As Long 'current width position
 Dim strt As Long 'starting position
 
 If show = False Then
   With Pic
     cuwd = .Width
     culf = .Left
     strt = culf - cuwd
     For i = strt To xlf Step 1
       .Left = i
       .Width = wd
       Next i
     End With
 Else   ' show = false
     With Pic
       culf = .Left
       For i = culf To xlf Step -1
       .Left = i
       .Width = wd
        Next i
      End With
End If
End Sub

Public Sub POPUP_ME_DNUP(Pic As PictureBox, xtp As Long, ht As Long, ByVal show As Boolean)
Dim i As Long
 '// ht - heighth when show or hide
 '// xtp - xtreme TOP
 Dim cuht As Long 'current height position
 Dim cuTP As Long 'current TOP position
 Dim strt As Long 'starting position
 
 If show = False Then
   With Pic
     cuht = .Height
     cuTP = .Top
     strt = cuTP - cuht
     For i = strt To xtp Step 1
       .Top = i
       .Height = ht
       Next i
     End With
 Else   ' show = false
     With Pic
       cuTP = .Top
       For i = cuTP To xtp Step -1
       .Top = i
       .Height = ht
        Next i
      End With
End If
End Sub

