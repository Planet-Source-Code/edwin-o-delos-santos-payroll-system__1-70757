VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmExcelToMDB 
   Caption         =   "Form1"
   ClientHeight    =   2625
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7545
   LinkTopic       =   "Form1"
   ScaleHeight     =   2625
   ScaleWidth      =   7545
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtExcel 
      Height          =   315
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   5895
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "&Browse"
      Height          =   375
      Left            =   6240
      TabIndex        =   6
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdConvert 
      Caption         =   "&Convert"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   1935
   End
   Begin VB.ListBox List1 
      Height          =   1230
      Left            =   3960
      TabIndex        =   4
      Top             =   600
      Width           =   3375
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   2160
      Width           =   7335
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
         TabIndex        =   3
         Top             =   120
         Width           =   180
      End
   End
   Begin VB.TextBox TextFN 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Text            =   "EDWIN"
      Top             =   960
      Width           =   2175
   End
   Begin VB.TextBox TextExt 
      Appearance      =   0  'Flat
      BackColor       =   &H00862229&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "MDB"
      Top             =   960
      Width           =   855
   End
   Begin MSComDlg.CommonDialog Comm 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "File Name:"
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   720
      Width           =   750
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Extension:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   2400
      TabIndex        =   8
      Top             =   720
      Width           =   900
   End
End
Attribute VB_Name = "frmExcelToMDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBrowse_Click()
Comm.Filter = "Microsoft Excel|*.xls"
Comm.ShowOpen
txtExcel.text = Comm.FileName

End Sub
Private Sub cmdConvert_Click()
Dim FSys As New FileSystemObject
Dim m_FN As String
m_FN = App.Path & "\" & TextFN.text & "." & TextExt.text
Dim b_FileExist As Boolean     'determine if folder distination exists
b_FileExist = FSys.FileExists(m_FN)
If b_FileExist = True Then
  MsgBox "Already Exists!"
  Exit Sub
End If
Dim objExcl As Excel.Application
Dim NewDB As Database
Dim NewTable As TableDef
Dim DBName As String
Dim fldNew As Field
Dim fldLoop As Field
Dim dbnm As String
Dim fld As String
Dim db As Database
Dim rs As DAO.Recordset, nofld As Integer, norecd As Integer
Label1.Caption = "Reading Excel Sheet. Wait.............."
Set objExcl = New Excel.Application
objExcl.Workbooks.Open (txtExcel.text)
Screen.MousePointer = 11

i = 1
dbnm = TextFN.text
Label1.Caption = "Creating Database. Wait.............."
'frmMsg.Show 1
Set NewDB = CreateDatabase(App.Path & "\" & dbnm, dbLangGeneral)
Set NewTable = NewDB.CreateTableDef("sheet1")
With NewTable
Do While Len(objExcl.ActiveSheet.Cells(1, i).Value) <> 0
    fld = objExcl.ActiveSheet.Cells(1, i).Value
    Set fldNew = .CreateField(fld, dbText, 50)
    fldNew.AllowZeroLength = True
    .Fields.Append fldNew
    List1.AddItem objExcl.ActiveSheet.Cells(1, i).Value
    i = i + 1
Loop
End With
NewDB.TableDefs.Append NewTable
NewDB.Close
i = 2
Label1.Caption = "Writing Data in Database. Wait............"
Set db = OpenDatabase(App.Path & "\" & dbnm)
Set rs = db.OpenRecordset("sheet1")
Do While Len(objExcl.ActiveSheet.Cells(i, 1).Value) <> 0
    j = 1
    rs.addNEW
    Do While Len(objExcl.ActiveSheet.Cells(i, j).Value) <> 0
        rs.Fields(j - 1).Value = objExcl.ActiveSheet.Cells(i, j).Value
        j = j + 1
    Loop
    rs.Update
    i = i + 1
Loop
Screen.MousePointer = 0
objExcl.Workbooks.Close
Label1.Caption = "Successfully Converted"
End Sub

Private Sub Form_Load()

End Sub
