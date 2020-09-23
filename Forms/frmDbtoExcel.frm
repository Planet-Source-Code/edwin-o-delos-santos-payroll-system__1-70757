VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmDbtoExcel 
   Caption         =   "Convert Access Database to Excel"
   ClientHeight    =   3285
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7350
   LinkTopic       =   "Form1"
   ScaleHeight     =   3285
   ScaleWidth      =   7350
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   1620
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   5295
   End
   Begin MSComDlg.CommonDialog Comm 
      Left            =   360
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdBrow 
      Caption         =   "&Browse"
      Height          =   375
      Left            =   5520
      TabIndex        =   2
      Top             =   240
      Width           =   1455
   End
   Begin VB.TextBox txtDbname 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   5295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Convert"
      Height          =   495
      Left            =   2880
      TabIndex        =   0
      Top             =   2640
      Width           =   1215
   End
End
Attribute VB_Name = "frmDbtoExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database
'Dim CN As ADODB.Connection
Dim rs As DAO.Recordset

Private Sub cmdBrow_Click()
Comm.Filter = "Msaccess Database|*.mdb"
Comm.ShowOpen
txtDbname.Text = Comm.FileName
'Set CN = New ADODB.Connection
'CN.Provider = "Microsoft.Jet.OLEDB.3.51"
'CN.ConnectionString = txtDbname.Text '.Path & "\db4.mdb"
'CN.Open
Set db = OpenDatabase(Comm.FileName)
List1.Clear
    ' List the table names.

    For Each TD In db.TableDefs
        ' Do not allow the system tables.
        If Left$(TD.Name, 4) <> "MSys" Then _
            List1.AddItem TD.Name
    Next TD
db.Close
End Sub

Private Sub Command1_Click()
Dim I, J, rtot, m
Dim db As Database, rs As DAO.Recordset
Dim ctot(1 To 4)
rtot = 0

Dim objExcl As Excel.Application
Set db = OpenDatabase(txtDbname.Text)
Set rs = db.OpenRecordset("select *from " & List1.Text)
'Set rs = New Recordset
'rs.Open "select *from " & List1.Text, CN, adOpenKeyset

Set objExcl = New Excel.Application
objExcl.Visible = True
objExcl.SheetsInNewWorkbook = 1
objExcl.Workbooks.Add
For I = 0 To rs.Fields.Count - 1
    objExcl.ActiveSheet.Cells(1, I + 1).Value = rs.Fields(I).Name '"SUPP-CODE"
'objExcl.ActiveSheet.Cells(1, 2).Value = "DATE"
'objExcl.ActiveSheet.Cells(1, 3).Value = "ITEM_CODE"
'objExcl.ActiveSheet.Cells(1, 4).Value = "BLACK"
'objExcl.ActiveSheet.Cells(1, 5).Value = "RED"
'objExcl.ActiveSheet.Cells(1, 6).Value = "BLUE"
'objExcl.ActiveSheet.Cells(1, 7).Value = "GREEN"
'objExcl.ActiveSheet.Cells(1, 8).Value = "TOTAL"
Next

'rs.Open "select *from excl1 ORDER BY ITEM_CODE", CN, adOpenKeyset
J = 3
Do Until rs.EOF
For I = 0 To rs.Fields.Count - 1
objExcl.ActiveSheet.Cells(J, I + 1).Value = rs.Fields(I)
'If I > 2 Then
'rtot = rtot + rs.Fields(I)
'End If
Next
objExcl.ActiveSheet.Cells(J, I + 1).Value = rtot
rs.MoveNext
J = J + 1
Loop
Dim k
k = 1
rs.MoveFirst
'Do Until rs.EOF
'For k = 1 To 4
'ctot(k) = ctot(k) + rs.Fields(k + 2)

'Next
'rs.MoveNext
'Loop
'objExcl.ActiveSheet.Cells(J + 1, 2).Value = "TOTAL"
'objExcl.ActiveSheet.Cells(J + 1, 4).Value = ctot(1)
'objExcl.ActiveSheet.Cells(J + 1, 5).Value = ctot(2)
'objExcl.ActiveSheet.Cells(J + 1, 6).Value = ctot(3)
'objExcl.ActiveSheet.Cells(J + 1, 7).Value = ctot(4)

End Sub

