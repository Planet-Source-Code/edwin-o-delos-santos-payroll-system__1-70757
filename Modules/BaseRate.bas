Attribute VB_Name = "BaseRate"
Option Explicit
Public LstMainLV As ListItem

Private Sub ListRate(sours As String, destin As String, rate As Double, loc As String)
'Dim LstMainLV As ListItem
 Set LstMainLV = Form1.ListView2.ListItems.Add(, , sours)
     LstMainLV.ForeColor = vbBlack
     LstMainLV.SubItems(1) = destin
     LstMainLV.ListSubItems(1).ForeColor = vbBlack
     LstMainLV.SubItems(2) = rate
     LstMainLV.ListSubItems(2).ForeColor = vbBlack
     LstMainLV.SubItems(3) = loc
     LstMainLV.ListSubItems(3).ForeColor = vbBlack
End Sub

Public Sub Ratelist()
'//  Temporary loading ...
'//  procedure under construction
 Call ListRate("DOLORES", "NIING", "200.00", "SAN ANTONIO")
 Call ListRate("DOLORES", "AYUSAN", "200.00", "TIAONG")
 Call ListRate("PALSABANGON", "BOCOHAN", "150.00", "LUCENA")
 Call ListRate("PALSABANGON", "BALUBAL", "250.00", "SARIAYA")
 Call ListRate("PALSABANGON", "ILA-DUPAY", "150.00", "LUCENA")
 Call ListRate("SARIAYA", "MASALUKOT", "200.00", "CANDELARIA")
 Call ListRate("SARIAYA", "MANGGALANG", "200.00", "CANDELARIA")
 Call ListRate("SARIAYA", "CATANAUAN", "600.00", "CATANAUAN")
 Call ListRate("SARIAYA", "DALAHICAN", "250.00", "LUCENA")
 Call ListRate("SARIAYA", "AVIDA", "200.00", "LUCENA")
 Call ListRate("SARIAYA", "NIING", "300.00", "SAN ANTONIO")
 Call ListRate("SARIAYA", "ARAWAN", "300.00", "SAN ANTONIO")
 Call ListRate("SARIAYA", "MULANAY", "900.00", "MULANAY")
Call ListRate("SARIAYA", "SAN PEDRO", "200.00", "SAN PEDRO")
 Call ListRate("SARIAYA", "AYUSAN", "280.00", "TIAONG")
 Call ListRate("SARIAYA", "CABATANG", "280.00", "TIAONG")
 Call ListRate("SARIAYA", "TULO-TULO", "200.00", "SJ BATANGAS")
 Call ListRate("BOCOHAN", "LUCBAN", "200.00", "LUCBAN")
 Call ListRate("COTA B", "BANTIGUE", "170.00", "PAGBILAO")
 Call ListRate("COTA B", "BOCOHAN", "150.00", "LUCENA-BOC")
 Call ListRate("LUCBAN", "LUCBAN", "40.00", "HAULING")
 Call ListRate("METROPOLIS", "METROPOLIS", "20.00", "HAULING")
 Call ListRate("METROPOLIST", "CALMAR", "50.00", "CALMAR-HAU")
End Sub




