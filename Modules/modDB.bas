Attribute VB_Name = "modDB"

Option Explicit
Public CN    As New Connection 'user by INVENTORY.MDB
Public Con   As New Connection 'used by USERS.MDB
Public CnPay As New Connection 'used by payroll

Public Sub OpenDB(ByRef MDB As String, newConn As Connection, Optional ByVal needPASS As Boolean, Optional ByVal mdbPASS As String)
'// each MDB requires new connection
 Set newConn = New ADODB.Connection
 newConn.CursorLocation = adUseClient
 If needPASS = False Then
   newConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source= " & App.Path & "\DB\" & MDB
 Else
   newConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=false;Data Source= " & App.Path & "\DB\" & MDB & ";Jet OLEDB:Database Password=" & mdbPASS
 End If
End Sub

Public Sub CloseDB()
 CN.Close
 Con.Close
 Set CN = Nothing
 Set Con = Nothing
End Sub

Sub Main()
Dim dbPass As String
'[============]
'< Initialize >
'[============]
  CurrUser.USER_isADMIN = "N"
  dbPass = "Üáúêãóöëîù¢®™ö©±Æ¥"
'[===============================]
'<OPEN OTHER CONNECTION , others >
'[===============================]
  Call OpenDB("PAYROLL.MDB", CnPay, True, dbPass)
  Call OpenDB("INVENTORY.MDB", CN)
  Load MDIPayroll
  MDIPayroll.show
End Sub
