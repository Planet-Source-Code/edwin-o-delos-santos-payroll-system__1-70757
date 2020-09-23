VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmExcelDB 
   Caption         =   "Convert Excel to DB"
   ClientHeight    =   2685
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7425
   Icon            =   "frmExcelDB.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   7425
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox listDate 
      Appearance      =   0  'Flat
      BackColor       =   &H00E8FBFB&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1710
      Left            =   960
      TabIndex        =   8
      Top             =   840
      Width           =   2055
   End
   Begin MSComDlg.CommonDialog Comm 
      Left            =   120
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtExcel 
      Height          =   315
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   5895
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "&Browse"
      Height          =   375
      Left            =   6240
      TabIndex        =   5
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdConvert 
      Caption         =   "&Convert"
      Height          =   375
      Left            =   5400
      TabIndex        =   4
      Top             =   2040
      Width           =   1935
   End
   Begin VB.ListBox List1 
      Height          =   1035
      ItemData        =   "frmExcelDB.frx":0CCA
      Left            =   960
      List            =   "frmExcelDB.frx":0CD1
      TabIndex        =   3
      Top             =   960
      Width           =   3855
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   2160
      Width           =   4815
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "---"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   225
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   180
      End
   End
   Begin VB.TextBox TextFN 
      Height          =   315
      Left            =   960
      TabIndex        =   0
      Text            =   "EDWIN"
      Top             =   480
      Width           =   3375
   End
   Begin VB.Image imgHelp 
      Height          =   360
      Left            =   4440
      MouseIcon       =   "frmExcelDB.frx":0CE4
      MousePointer    =   99  'Custom
      Picture         =   "frmExcelDB.frx":15AE
      Top             =   480
      Width           =   360
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "File Name:"
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   480
      Width           =   750
   End
End
Attribute VB_Name = "frmExcelDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_FN As String
Private curDate As String

Private Sub cmdBrowse_Click()
Comm.Filter = "Microsoft Excel|*.xls"
Comm.ShowOpen
txtExcel.Text = Comm.FileName

End Sub
Private Sub cmdConvert_Click()
 If Not ValidFN(txtExcel.Text, ".XLS") Then
   MsgBox "Invalid XLS File name > " & txtExcel.Text, vbInformation, "Convert"
   Exit Sub
 End If
 If Not ValidFN(TextFN.Text, ".MDB") Then
   MsgBox "Invalid MDB File name > " & TextFN.Text, vbInformation, "Convert"
   Exit Sub
 End If
Dim FSys As New FileSystemObject
Dim c As Integer, r As Integer
Dim b_FileExist As Boolean     'determine if file destination exists
Dim objExcl As Excel.Application
Dim msg As String
Dim NewDB As Database
Dim NewTable As TableDef
Dim DBName As String
Dim fldNew As Field
Dim fldLoop As Field
Dim dbnm As String
Dim fld As String
Dim db As Database
Dim rs As DAO.Recordset, nofld As Integer, norecd As Integer
m_FN = App.Path & "\" & TextFN.Text
b_FileExist = FSys.FileExists(m_FN)
If b_FileExist = True Then
  msg = m_FN
  msg = msg & vbCrLf & "Already Exists!"
  MsgBox msg, vbInformation, "Convert"
  Exit Sub
End If

Label1.Caption = "Reading Excel Sheet. Wait.............."
Set objExcl = New Excel.Application
objExcl.Workbooks.Open (txtExcel.Text)
Screen.MousePointer = 11

c = 1    'where C is column
dbnm = TextFN.Text
Label1.Caption = "Creating Database. Wait.............."
Set NewDB = CreateDatabase(App.Path & "\" & dbnm, dbLangGeneral)
Set NewTable = NewDB.CreateTableDef("sheet1")
With NewTable
Do While Len(objExcl.ActiveSheet.Cells(1, c).Value) <> 0
    fld = objExcl.ActiveSheet.Cells(1, c).Value
    Set fldNew = .CreateField(fld, dbText, 50)
    fldNew.AllowZeroLength = True
    .Fields.Append fldNew
    List1.AddItem objExcl.ActiveSheet.Cells(1, c).Value
    c = c + 1
Loop
End With
NewDB.TableDefs.Append NewTable
NewDB.Close
r = 2  'where R is Row
Label1.Caption = "Writing Data in Database. Wait............"
Set db = OpenDatabase(App.Path & "\" & dbnm)
Set rs = db.OpenRecordset("sheet1")
Do While Len(objExcl.ActiveSheet.Cells(r, 1).Value) <> 0
    c = 1
    rs.AddNew
    Do While Len(objExcl.ActiveSheet.Cells(r, c).Value) <> 0
        rs.Fields(c - 1).Value = objExcl.ActiveSheet.Cells(r, c).Value
Rem Debug.Print objExcl.ActiveSheet.Cells(r, c).Value
        c = c + 1
    Loop
    rs.Update
    r = r + 1
Loop
Screen.MousePointer = 0
objExcl.Workbooks.Close
Label1.Caption = "Successfully Converted"
End Sub


Private Function ValidFN(ByVal fn As String, ByVal xt As String) As Boolean
  ValidFN = False
  If Right$(UCase$(fn), 4) <> xt Or Right$(UCase$(fn), 4) <> xt Then
      ValidFN = False
  Else
     ValidFN = True
  End If
End Function

Private Sub Form_Load()
 curDate = Format(Now(), "mm-dd-yyyy")
 TextFN.Text = CStr(curDate) & ".MDB"
 ListCalendar listDate, TextFN.Text
End Sub

Private Sub imgHelp_Click()
  Dim msg As String
  msg = "Use date as FileName"
  msg = msg & vbCrLf & "Arrow Down to view list!"
  MsgBox msg, vbInformation, "Database FileName!"
End Sub

Private Sub listDate_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 13 Then
    TextFN.Text = listDate.Text & ".MDB"
    listDate.Visible = False
 ElseIf KeyCode = 27 Then
    listDate.Visible = False
    TextFN.SetFocus
 End If
End Sub

Private Sub TextFN_GotFocus()
AlignObj TextFN, listDate, 1, False
End Sub

Private Sub TextFN_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 40 Then
     listDate.Visible = True
     listDate.SetFocus
  End If
End Sub
