VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDeduction 
   BackColor       =   &H00F7EBD0&
   Caption         =   "Employee's Deduction "
   ClientHeight    =   3480
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6555
   Icon            =   "frmDeduction.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   6555
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ListView lvList 
      Height          =   1875
      Left            =   240
      TabIndex        =   12
      Top             =   1320
      Width           =   4425
      _ExtentX        =   7805
      _ExtentY        =   3307
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "i16x16"
      SmallIcons      =   "i16x16"
      ForeColor       =   -2147483640
      BackColor       =   16777215
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.TextBox txtEntry 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   285
      Index           =   2
      Left            =   1800
      TabIndex        =   13
      Top             =   840
      Width           =   2835
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H00F7EBD0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2715
      Left            =   4920
      ScaleHeight     =   2715
      ScaleWidth      =   1455
      TabIndex        =   4
      Top             =   120
      Width           =   1455
      Begin VB.CommandButton cmdButton 
         Caption         =   "&Edit"
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   1215
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "&Add"
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "&Update"
         Height          =   315
         Index           =   3
         Left            =   120
         TabIndex        =   9
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "&Save"
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "&Cancel"
         Height          =   315
         Index           =   4
         Left            =   120
         TabIndex        =   7
         Top             =   1560
         Width           =   1215
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "&Refresh"
         Height          =   315
         Index           =   6
         Left            =   120
         TabIndex        =   6
         Top             =   2280
         Width           =   1215
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "&Delete"
         Height          =   315
         Index           =   5
         Left            =   120
         TabIndex        =   5
         Top             =   1920
         Width           =   1215
      End
   End
   Begin VB.TextBox txtEntry 
      BackColor       =   &H00E8FBFB&
      Height          =   285
      Index           =   0
      Left            =   1800
      TabIndex        =   2
      Top             =   120
      Width           =   2835
   End
   Begin VB.TextBox txtEntry 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   1
      Left            =   1800
      TabIndex        =   0
      Top             =   480
      Width           =   2835
   End
   Begin MSComctlLib.ImageList i16x16 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   18
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeduction.frx":109A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeduction.frx":1AAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeduction.frx":1E46
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeduction.frx":21E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeduction.frx":2BF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeduction.frx":2C46
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeduction.frx":2FE0
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeduction.frx":337A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeduction.frx":3714
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeduction.frx":3AAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeduction.frx":44C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeduction.frx":4ED2
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeduction.frx":58E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeduction.frx":62F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeduction.frx":6D08
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeduction.frx":771A
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeduction.frx":812C
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeduction.frx":86C8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblFLDi 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   240
      TabIndex        =   14
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label lblFLDi 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label lblFLDi 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   1575
   End
End
Attribute VB_Name = "frmDeduction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset


Private Sub cmdButton_Click(Index As Integer)
'//                  A S E U C D R
On Error GoTo ERRORHANDLE
Select Case Index
   Case BtnAdd                       '<------ add new record ------->'
'     addRec = True
'     cmdButtonShow ("0100100"), Me
'     If isFilter = True Then
'        MsgBox "Data is Filtered", vbCritical, "Refresh Record First!"
'        Exit Sub
'     End If
'     Dim NextNo As Long
'     '//initialize//
'     txtEntry(29).text = Format(Now(), "Short Date")
'     txtEntry(30).text = CurrUser.user_id
'     '//assign next number//
'     NextNo = Last_Recc(rsPAY)
'     If NextNo > 0 Then
'       txtEntry(0).text = NextNo
'       txtEntry(1).SetFocus
'     Else
'       txtEntry(0).Locked = False
'       txtEntry(0).SetFocus
'    End If
Exit Sub
   Case BtnSave                       '<------ save new record ------>'
'        cmdButtonShow ("1010011"), Me
'        Call WriteData(Me, rs, True)
'        Call lvwPopulateData(lvList, rs, 2)
'        addRec = False
Exit Sub
   Case BtnEdit                       '<------ edit record ---------->'
        editRec = True
        cmdButtonShow ("0001100"), Me
        txtEntry(1).SetFocus
   Case BtnUpdate                     '<------ update record -------->'
        cmdButtonShow ("0010001"), Me
        Call WriteData(Me, rs, False)
        LvwReplaceData Me, rs, lvList
        editRec = False
   Case BtnCancel                     '<------ cancel update -------->'
        cmdButtonShow ("0010001"), Me
        addRec = False
        editRec = False
   Case BtnDelete                     '<------ delete record -------->'
        '// no delete here please !
        'Call Delete_Record(rs, lvList)
   Case BtnRefresh                    '<------ Refresh record ------->'
        addRec = False
        edirec = False
       If rs.State = adStateOpen Then
          rs.Close
        End If
        rs.Open "SELECT * From DEDUCTION order by SN", CnPay, adOpenStatic, adLockOptimistic
        Load_DATA
        isFilter = False
        lvList.SetFocus
End Select
ERRORHANDLE:
 errorMsg Err, Me.Name, "Command Button"

End Sub


Private Sub Form_Load()
cmdButtonShow ("0010001"), Me
Set rs = New ADODB.Recordset
rs.Open "SELECT * From DEDUCTION order by SN", CnPay, adOpenStatic, adLockOptimistic
Load_DATA

Call ShowFldsLabel(Me, rs)
End Sub
Private Sub Load_DATA()
On Error GoTo ERRORHANDLE
'// set columnheaders
'Insert_ExtraCol lvList, rsDed

Call InsertColumn(lvList, rs)
'//set details
 Call FillListView(lvList, rs, 2)
ERRORHANDLE:
    errorMsg Err, Me.Name
End Sub

Private Sub Form_Resize()
With Me
  If .WindowState = 0 Then
   .Height = 3990
   .Width = 6675
  End If
End With
 SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub

Private Sub lvList_Click()
On Error GoTo ERRORHANDLE
If addRec = True Or editRec = True Then Exit Sub
Call BindDatasource(Me, rs, lvList, True)
ERRORHANDLE:
 errorMsg Err, Me.Name
End Sub

Private Sub lvList_KeyUp(KeyCode As Integer, Shift As Integer)
 lvList_Click
End Sub
