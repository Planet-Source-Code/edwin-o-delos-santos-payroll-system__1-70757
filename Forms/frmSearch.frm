VERSION 5.00
Begin VB.Form frmSearch 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Search  / Filter"
   ClientHeight    =   3390
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9435
   Icon            =   "frmSearch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   9435
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdRefresh 
      BackColor       =   &H00808080&
      Caption         =   "Refresh"
      Height          =   315
      Left            =   240
      TabIndex        =   15
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Caption         =   $"frmSearch.frx":0CCA
      ForeColor       =   &H00C00000&
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   9255
      Begin VB.TextBox TextFieldValue 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4920
         TabIndex        =   23
         Top             =   2040
         Width           =   2535
      End
      Begin VB.CheckBox chkFieldValue 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Caption         =   "Field Value ( By Input )"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   2880
         TabIndex        =   24
         Top             =   2040
         Width           =   2055
      End
      Begin VB.CommandButton CmdBuildSQL 
         BackColor       =   &H00808080&
         Caption         =   "Build SQL Statment"
         Height          =   435
         Left            =   7920
         TabIndex        =   22
         Top             =   2400
         Width           =   1215
      End
      Begin VB.CommandButton CmdExecuteSql 
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   7920
         Picture         =   "frmSearch.frx":0D60
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   2880
         Width           =   1215
      End
      Begin InstantReport.Hline Hline1 
         Height          =   30
         Left            =   2880
         TabIndex        =   16
         Top             =   1800
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   53
      End
      Begin VB.OptionButton optAny 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Match"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   4440
         TabIndex        =   14
         Top             =   960
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton optExact 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Exact"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   4440
         TabIndex        =   13
         Top             =   720
         Width           =   975
      End
      Begin VB.CheckBox chkLIKE 
         BackColor       =   &H00C0E0FF&
         Caption         =   "CONTAINS"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2880
         TabIndex        =   12
         Top             =   840
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CheckBox chkBETWEEN 
         BackColor       =   &H00C0E0FF&
         Caption         =   "BETWEEN  >"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2880
         TabIndex        =   11
         Top             =   1440
         Width           =   1335
      End
      Begin VB.CheckBox chkCOND 
         BackColor       =   &H00C0E0FF&
         Caption         =   "CONDITIONS"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2880
         TabIndex        =   10
         Top             =   240
         Width           =   1455
      End
      Begin VB.ComboBox ListCond 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmSearch.frx":170A
         Left            =   4440
         List            =   "frmSearch.frx":1726
         TabIndex        =   9
         Text            =   "Select"
         Top             =   240
         Width           =   975
      End
      Begin VB.ListBox ListFields 
         Appearance      =   0  'Flat
         Height          =   2175
         ItemData        =   "frmSearch.frx":1749
         Left            =   120
         List            =   "frmSearch.frx":1750
         Sorted          =   -1  'True
         TabIndex        =   8
         Top             =   240
         Width           =   2535
      End
      Begin VB.ListBox ListValue 
         Appearance      =   0  'Flat
         Height          =   1005
         ItemData        =   "frmSearch.frx":1760
         Left            =   6360
         List            =   "frmSearch.frx":1767
         Sorted          =   -1  'True
         TabIndex        =   7
         Top             =   240
         Width           =   2775
      End
      Begin VB.TextBox TextSQL 
         BackColor       =   &H00FFFFFF&
         Height          =   765
         HideSelection   =   0   'False
         Left            =   2880
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Text            =   "frmSearch.frx":1776
         Top             =   2400
         Width           =   4935
      End
      Begin VB.TextBox txtDATE 
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   6360
         TabIndex        =   2
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox txtDATE 
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   7800
         TabIndex        =   1
         Top             =   1440
         Width           =   1335
      End
      Begin InstantReport.Hline ctrlLiner3 
         Height          =   30
         Left            =   960
         TabIndex        =   3
         Top             =   120
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   53
      End
      Begin InstantReport.Hline ctrlLiner2 
         Height          =   30
         Left            =   2880
         TabIndex        =   4
         Top             =   1320
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   53
      End
      Begin InstantReport.Hline ctrlLiner1 
         Height          =   30
         Left            =   2880
         TabIndex        =   5
         Top             =   645
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   53
      End
      Begin VB.Image imgHelp 
         Height          =   360
         Left            =   8640
         MouseIcon       =   "frmSearch.frx":1789
         MousePointer    =   99  'Custom
         Picture         =   "frmSearch.frx":2053
         Top             =   1920
         Width           =   360
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0E0FF&
         Caption         =   "( Input Date )"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4440
         TabIndex        =   20
         Top             =   1440
         Width           =   930
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0E0FF&
         Caption         =   ">>"
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
         Left            =   5760
         TabIndex        =   19
         Top             =   1440
         Width           =   225
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0E0FF&
         Caption         =   ">>"
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
         Left            =   5760
         TabIndex        =   18
         Top             =   840
         Width           =   225
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0E0FF&
         Caption         =   ">>"
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
         Left            =   5760
         TabIndex        =   17
         Top             =   240
         Width           =   225
      End
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'[---------------------------------]
'< SQL Builder                     >
'< designed for Global purpose     >
'< ... can be called by any table  >
'< coded by:edwin delos santos     >
'[---------------------------------]
                
Option Explicit

Public pFindForm As Form  'source Form
Public pFindTABLE As String
Public pFindCon As New ADODB.Connection
Public pFindRecset As New ADODB.Recordset
Private sqlStatement As String   'store sql statement
Private operand As String        'sql criteria operator
Private wStatement As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdBuildSQL_Click()
  BuildSQL
End Sub

Private Sub BuildSQL()
On Error GoTo errMsg
Dim vlid As Boolean
vlid = isValid(ListFields.text)

sqlStatement = ""
wStatement = ""

'Then build our SQL statement according from the input strings and combo/Listboxes boxes
wStatement = " WHERE " & "[" & ListFields.text & "]"
sqlStatement = "SELECT * FROM [" & pFindTABLE & "]"
sqlStatement = sqlStatement & wStatement

If chkLIKE.Value = 1 Then
 If optExact.Value = True Then   'Exact or not exact ?
   If chkFieldValue.Value = 1 Then '// BY INPUT
      sqlStatement = sqlStatement & "LIKE '" & TextFieldValue.text & "'"
    Else
      sqlStatement = sqlStatement & "LIKE '" & ListValue.text & "'"
   End If
 ElseIf optAny.Value = True Then
    If chkFieldValue.Value = 1 Then '// BY INPUT
      sqlStatement = sqlStatement & "LIKE '%" & TextFieldValue.text & "%'"
    Else
       sqlStatement = sqlStatement & "LIKE '%" & ListValue.text & "%'"
    End If
 End If
ElseIf chkBETWEEN.Value = 1 Then
   operand = "BETWEEN #" & CDate(txtDATE(0).text) & "# AND #" & CDate(txtDATE(1).text) & "#"
  sqlStatement = sqlStatement & operand
ElseIf chkCOND.Value = 1 Then
  operand = ListCond.text
    If chkFieldValue.Value = 1 Then '// BY INPUT
      sqlStatement = sqlStatement & operand & TextFieldValue.text
    Else
      sqlStatement = sqlStatement & operand & ListValue.text
    End If
End If

TextSQL.text = sqlStatement

errMsg:
    errorMsg Err, Me.Name, "Build Sql"

End Sub

Private Sub chkFieldValue_Click()
  ListValue.Enabled = (chkFieldValue.Value = 0)
  TextFieldValue.Enabled = (chkFieldValue.Value = 1)
End Sub


Private Sub cmdRefresh_Click()
    Dim sqlSTR As String
    If pFindRecset.State = adStateOpen Then
      pFindRecset.Close
    End If
    sqlSTR = "SELECT * FROM [" & pFindTABLE & "]"
    pFindRecset.Open sqlSTR, pFindCon  ', adOpenStatic, adLockOptimistic
End Sub





Private Sub Form_Load()
    FormRndCorner Me, 640, 260
    '//initialize
    show
    ListFields.SetFocus
    ListCond.ListIndex = 0
    '// end
    Dim sqlSTR As String
    If pFindRecset.State = adStateOpen Then
      pFindRecset.Close
    End If
    sqlSTR = "SELECT * FROM [" & pFindTABLE & "]"
    pFindRecset.CursorLocation = adUseClient
    pFindRecset.Open sqlSTR, pFindCon, adOpenStatic, adLockOptimistic
    Call Insert_Fields(ListFields, pFindRecset)

End Sub

Private Sub Form_Resize()
With Me
  If .WindowState = 0 Then
   .Height = 3900
   .Width = 9570
  End If
End With
   SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub

Private Sub Form_Unload(Cancel As Integer)
    isFilter = False
    Set frmSearch = Nothing
End Sub
Private Sub chkCOND_Click()
ListValue.ListIndex = 0
If chkFieldValue.Value = 0 Then
   If Not IsNumeric(ListValue.text) Then
      chkCOND.Value = 0
      myMsg "Fields Value must be numeric!", "Search/Filter", 1, True
      Exit Sub
   End If
Else
   If Not IsNumeric(TextFieldValue.text) Then
      chkCOND.Value = 0
      myMsg "Input Value must be numeric!", "Search/Filter", 1, True
      Exit Sub
   End If
End If
If chkCOND.Value = 1 Then
   chkBETWEEN.Value = 0
   chkLIKE.Value = 0
End If
 EnableControl
 
End Sub

Private Sub chkLIKE_Click()
If chkLIKE.Value = 1 Then
   chkBETWEEN.Value = 0
   chkCOND.Value = 0
End If
EnableControl
End Sub

Private Sub chkBETWEEN_Click()
On Error GoTo ERRORHANDLE
ListValue.ListIndex = 0
If Not IsDate(ListValue.text) Then
   chkBETWEEN.Value = 0
   Exit Sub
End If
If chkBETWEEN.Value = 1 Then
   chkLIKE.Value = 0
   chkCOND.Value = 0
End If
chkFieldValue.Enabled = (chkBETWEEN.Value = 0)
If txtDATE(0).text = Empty Or txtDATE(1).text = Empty Then Exit Sub
operand = "BETWEEN #" & CDate(txtDATE(0).text) & "# AND #" & CDate(txtDATE(1).text) & "#"
EnableControl
ERRORHANDLE:
   errorMsg Err, Me.Name
End Sub

Private Sub EnableControl()
  '[/] CONDISTIONS
    ListCond.Enabled = (chkCOND.Value = 1)
  '[/] LIKE
   optExact.Enabled = (chkLIKE.Value = 1)
   optAny.Enabled = (chkLIKE.Value = 1)
   optAny.Value = (chkLIKE.Value = 1)        '//default
   '[/] BETWEEN
   txtDATE(0).Enabled = (chkBETWEEN.Value = 1)
   txtDATE(1).Enabled = (chkBETWEEN.Value = 1)
End Sub


Private Sub imgHelp_Click()
  Dim msg As String
  msg = "DblClick FieldList to get the value. If you feel that "
  msg = msg & "getting the value takes longer time than usual then, "
  msg = msg & vbCrLf & "You may just input the expression on the "
  msg = msg & "open box. Check [/] Field Value ( By Input ) to enable "
  msg = msg & "open box."
  myMsg msg, "Field Value ( By Input )", 2, True
End Sub

Private Sub ListCond_Click()
   If ListValue.ListIndex < 0 Or ListFields.ListIndex < 0 Then Exit Sub
   operand = ""
   operand = ListCond.text
   chkLIKE.Value = 0
   chkBETWEEN.Value = 0
End Sub
Private Sub ListFields_DblClick()
On Error GoTo errMsg
  Call Add_Item(pFindRecset, ListFields.text, ListValue)
errMsg:
    errorMsg Err, Me.Name, "ListFields_DblClick"
End Sub
Private Function isValid(ByRef srcStr As String) As Boolean
   If srcStr = Empty Then
     isValid = False
     GoSub invalid
   ElseIf Len(srcStr) = 0 Then
     isValid = False
     GoSub invalid
   Else
     isValid = True
     Exit Function
   End If
invalid:
   MsgBox "Invalid!", vbCritical, "Build SQL"
End Function

Private Sub CmdExecuteSql_Click()
If TextSQL.text = "SQL Statement..." Then Exit Sub
If chkFieldValue.Value = 0 Then
   If ListValue.ListIndex < 0 Or ListFields.ListIndex < 0 Then Exit Sub
End If
isFilter = True
On Error GoTo errMsg
  If pFindRecset.State = adStateOpen Then
    pFindRecset.Close
  End If
  pFindRecset.Open sqlStatement, pFindCon
  If pFindRecset.RecordCount > 0 Then
      '// reminder:  lvList cannot be set as listview
      '// be sure that listview name on all form is lvList
      '// else procedure to execute sql must be put on form where listview is located
      Call InsertColumn(pFindForm.lvList, pFindRecset)
      Call FillListView(pFindForm.lvList, pFindRecset, 2)
      Call Listview_Total(pFindForm.lvList, pFindRecset)
  Else
    MsgBox "No record to load!", vbInformation, "Rebuild SQL Please!"
  End If

errMsg:
  errorMsg Err, Me.Name, "Execute Sql"
End Sub

Private Sub ListValue_Click()
If Not IsDate(ListValue.text) Then Exit Sub
If chkBETWEEN.Value = 0 Then Exit Sub
If Len(txtDATE(0).text) = 0 Then
   txtDATE(0).text = ListValue.text
ElseIf Len(txtDATE(0).text) > 0 Then
   txtDATE(1).text = ListValue.text
End If
End Sub

Private Sub TxtDate_Change(Index As Integer)
If txtDATE(0).text = Empty Or txtDATE(1).text = Empty Then Exit Sub
operand = "BETWEEN #" & CDate(txtDATE(0).text) & "# AND #" & CDate(txtDATE(1).text) & "#"
End Sub
