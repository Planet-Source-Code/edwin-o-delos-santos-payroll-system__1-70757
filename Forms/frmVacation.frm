VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmVacation 
   BackColor       =   &H00F7EBD0&
   Caption         =   "Vacation Schedule"
   ClientHeight    =   6705
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   9870
   Icon            =   "frmVacation.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   9870
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox List1Type 
      Appearance      =   0  'Flat
      BackColor       =   &H00E6F1FD&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2430
      ItemData        =   "frmVacation.frx":109A
      Left            =   4440
      List            =   "frmVacation.frx":10B3
      Sorted          =   -1  'True
      TabIndex        =   29
      Top             =   2760
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.CommandButton CmdViewName 
      BackColor       =   &H00C0FFFF&
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   1080
      Width           =   375
   End
   Begin VB.PictureBox PicNameList 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3180
      Left            =   5520
      Picture         =   "frmVacation.frx":1106
      ScaleHeight     =   3150
      ScaleWidth      =   3885
      TabIndex        =   24
      Top             =   3000
      Visible         =   0   'False
      Width           =   3915
      Begin VB.PictureBox PicNameClose 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   3600
         Picture         =   "frmVacation.frx":46AEA
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   25
         Top             =   0
         Width           =   270
      End
      Begin MSComctlLib.ListView lvName 
         Height          =   2595
         Left            =   0
         TabIndex        =   26
         Top             =   480
         Width           =   3825
         _ExtentX        =   6747
         _ExtentY        =   4577
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "i16x16"
         SmallIcons      =   "i16x16"
         ForeColor       =   12582912
         BackColor       =   15268859
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Name"
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
         Height          =   195
         Left            =   120
         TabIndex        =   27
         Top             =   120
         Width           =   1365
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   23
      Top             =   6330
      Width           =   9870
      _ExtentX        =   17410
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9948
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   1764
            MinWidth        =   1764
            Picture         =   "frmVacation.frx":47074
            Text            =   "Print"
            TextSave        =   "Print"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "5:39 PM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "7/3/2008"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H00F7EBD0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2715
      Left            =   8160
      ScaleHeight     =   2715
      ScaleWidth      =   1455
      TabIndex        =   15
      Top             =   240
      Width           =   1455
      Begin VB.CommandButton cmdButton 
         Caption         =   "&Delete"
         Height          =   315
         Index           =   5
         Left            =   120
         TabIndex        =   22
         Top             =   1920
         Width           =   1215
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "&Refresh"
         Height          =   315
         Index           =   6
         Left            =   120
         TabIndex        =   21
         Top             =   2280
         Width           =   1215
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "&Cancel"
         Height          =   315
         Index           =   4
         Left            =   120
         TabIndex        =   20
         Top             =   1560
         Width           =   1215
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "&Save"
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   19
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "&Update"
         Height          =   315
         Index           =   3
         Left            =   120
         TabIndex        =   18
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "&Add"
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   17
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "&Edit"
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   16
         Top             =   840
         Width           =   1215
      End
   End
   Begin VB.TextBox txtEntry 
      BackColor       =   &H00FFFFFF&
      Height          =   2205
      Index           =   6
      Left            =   5640
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   600
      Width           =   2355
   End
   Begin VB.TextBox txtEntry 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   4
      Left            =   2040
      TabIndex        =   5
      Top             =   1800
      Width           =   2835
   End
   Begin VB.TextBox txtEntry 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   2
      Left            =   2040
      TabIndex        =   4
      Top             =   1080
      Width           =   2835
   End
   Begin VB.TextBox txtEntry 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   3
      Left            =   2040
      TabIndex        =   3
      Top             =   1440
      Width           =   2835
   End
   Begin VB.TextBox txtEntry 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   5
      Left            =   2040
      TabIndex        =   2
      Top             =   2160
      Width           =   2835
   End
   Begin VB.TextBox txtEntry 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00008080&
      Height          =   285
      Index           =   0
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   360
      Width           =   1155
   End
   Begin VB.TextBox txtEntry 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   1
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   720
      Width           =   1155
   End
   Begin MSComctlLib.ListView lvList 
      Height          =   2955
      Left            =   240
      TabIndex        =   7
      Top             =   3240
      Width           =   9345
      _ExtentX        =   16484
      _ExtentY        =   5212
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
      BackColor       =   13235143
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
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
         NumListImages   =   24
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVacation.frx":4740E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVacation.frx":479A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVacation.frx":47F42
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVacation.frx":484DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVacation.frx":48876
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVacation.frx":48C10
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVacation.frx":48FAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVacation.frx":49344
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVacation.frx":496DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVacation.frx":4A0F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVacation.frx":4A144
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVacation.frx":4A4DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVacation.frx":4A878
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVacation.frx":4AC12
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVacation.frx":4AFAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVacation.frx":4B9BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVacation.frx":4C3D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVacation.frx":4CDE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVacation.frx":4D7F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVacation.frx":4E206
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVacation.frx":4EC18
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVacation.frx":4F62A
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVacation.frx":4FBC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVacation.frx":50162
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "[F2]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Left            =   4920
      TabIndex        =   30
      Top             =   1440
      Width           =   420
   End
   Begin VB.Label lblFLDi 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
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
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   5
      Left            =   480
      TabIndex        =   14
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label lblFLDi 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
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
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   4
      Left            =   480
      TabIndex        =   13
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label lblFLDi 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
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
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   1
      Left            =   480
      TabIndex        =   12
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label lblFLDi 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
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
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   2
      Left            =   480
      TabIndex        =   11
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label lblFLDi 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
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
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   3
      Left            =   480
      TabIndex        =   10
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label lblFLDi 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
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
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   6
      Left            =   5640
      TabIndex        =   9
      Top             =   360
      Width           =   2330
   End
   Begin VB.Label lblFLDi 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
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
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   0
      Left            =   480
      TabIndex        =   8
      Top             =   360
      Width           =   1575
   End
End
Attribute VB_Name = "frmVacation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsEmp As ADODB.Recordset
Dim rs As ADODB.Recordset


Private Sub cmdButton_Click(Index As Integer)
'//                  A S E U C D R
'On Error GoTo ERRORHANDLE
Select Case Index
   Case BtnAdd                       '<------ add new record ------->'
     addRec = True
     cmdButtonShow ("0100100"), Me
'     If isFilter = True Then
'        MsgBox "Data is Filtered", vbCritical, "Refresh Record First!"
'        Exit Sub
'     End If
'     Dim NextNo As Long
'     '//initialize//
'     txtEntry(29).text = Format(Now(), "Short Date")
'     txtEntry(30).text = CurrUser.user_id
'     '//assign next number//
      NextNo = Last_Recc(rs)
      If NextNo > 0 Then
       txtEntry(0).text = NextNo
       txtEntry(2).SetFocus
      Else
       txtEntry(0).Locked = False
       txtEntry(0).SetFocus
      End If
   Case BtnSave                       '<------ save new record ------>'
        cmdButtonShow ("1010011"), Me
        Call WriteData(Me, rs, True)
        Call lvwPopulateData(lvList, rs, 2)
        addRec = False
   Case BtnEdit                       '<------ edit record ---------->'
        editRec = True
        cmdButtonShow ("0001100"), Me
        txtEntry(2).SetFocus
   Case BtnUpdate                     '<------ update record -------->'
        cmdButtonShow ("1010001"), Me
        Call WriteData(Me, rs, False)
        LvwReplaceData Me, rs, lvList
        editRec = False
   Case BtnCancel                     '<------ cancel update -------->'
        cmdButtonShow ("1010001"), Me
        addRec = False
        editRec = False
   Case BtnDelete                     '<------ delete record -------->'
        '// no delete here please !
        Call Delete_Record(rs, lvList)
   Case BtnRefresh                    '<------ Refresh record ------->'
        addRec = False
        edirec = False
       If rs.State = adStateOpen Then
          rs.Close
        End If
        rs.Open "SELECT * From VACATION order by SN", CnPay, adOpenStatic, adLockOptimistic
        Load_DATA
        isFilter = False
        lvList.SetFocus
End Select
'ERRORHANDLE:
' errorMsg Err, Me.Name, "Command Button"

End Sub


Private Sub CmdViewName_Click()
 If addRec = True Or editRec = True Then
   PicNameList.Visible = True
   lvName.SetFocus
 End If
End Sub

Private Sub Form_Load()
'// initialized
cmdButtonShow ("1010011"), Me
AlignObj txtEntry(2), PicNameList, 1, False
AlignObj txtEntry(3), List1Type, 1, False
'// set focus
show
lvList.SetFocus
      
Set rs = New ADODB.Recordset
rs.Open "SELECT * From VACATION order by SN", CnPay, adOpenStatic, adLockOptimistic
Load_DATA
Call ShowFldsLabel(Me, rs)

Set rsEmp = New ADODB.Recordset
Dim SQL As String
SQL = "SELECT Employee_Name,ID_Code "
SQL = SQL & "From PAYROLL order by Employee_name"
rsEmp.Open SQL, CnPay, adOpenStatic, adLockOptimistic
Load_Employee


End Sub
Private Sub Load_DATA()
'On Error GoTo ERRORHANDLE
'// set columnheaders
'Insert_ExtraCol lvList, rsDed

Call InsertColumn(lvList, rs)
'//set details
 Call FillListView(lvList, rs, 2)
'ERRORHANDLE:
'    errorMsg Err, Me.Name
End Sub

Private Sub Load_Employee()
On Error GoTo ERRORHANDLE
'// set columnheaders
'Insert_ExtraCol lvList, rsDed
Call InsertColumn(lvName, rsEmp)
'//set details
Call FillListView(lvName, rsEmp, 1)
autoAlignCol lvName
ERRORHANDLE:
    errorMsg Err, Me.Name, "Load_Employee proc"
End Sub

Private Sub Form_Resize()
With Me
  If .WindowState = 0 Then
   .Height = 7185
   .Width = 9990
  End If
End With
 SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub


Private Sub List1Type_KeyUp(KeyCode As Integer, Shift As Integer)
 If KeyCode = 13 Then
   txtEntry(3).text = List1Type.text
   txtEntry(3).SetFocus
   List1Type.Visible = False
 End If
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

Private Sub lvName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   txtEntry(2).text = lvName.SelectedItem.text
   txtEntry(1).text = lvName.SelectedItem.ListSubItems(1).text  'ID
  txtEntry(2).SetFocus
  PicNameList.Visible = False
End If
End Sub

Private Sub PicNameClose_Click()
 PicNameList.Visible = False
 txtEntry(2).SetFocus
End Sub

Private Sub txtEntry_GotFocus(Index As Integer)
On Error GoTo errorMsg
nxTab = Index
txtEntry(nxTab).SelStart = 0
txtEntry(nxTab).SelLength = Len(txtEntry(nxTab).text)
'Select Case nxTab
'  Case Is = 2
'     If addRec = True Or editRec = True Then
'       AlignObj txtEntry(2), PicNameList, 1, False
'     End If
'End Select
errorMsg:
 errorMsg Err, Me.Name, "txtEntry_GotFocus"

End Sub

Private Sub txtEntry_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim lastTab As Integer
On Error GoTo ERRORHANDLE
lastTab = 5
If KeyCode = 13 Then
    If nxTab = lastTab Then Exit Sub
    nxTab = nxTab + 1
ElseIf KeyCode = 38 Then  'up arrow key
     If nxTab = 0 Or nxTab = 1 Then Exit Sub
     nxTab = nxTab - 1
End If
txtEntry(nxTab).SetFocus
ERRORHANDLE:
 errorMsg Err, Me.Name
End Sub

Private Sub txtEntry_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case Index
Case Is = 2
  If addRec = True Or editRec = True Then
    If KeyCode = 113 Then 'F2
       PicNameList.Visible = True
       lvName.SetFocus
    End If
  End If
Case Is = 3
  If addRec = True Or editRec = True Then
    If KeyCode = 113 Then 'F2
       List1Type.Visible = True
       List1Type.SetFocus
    End If
  End If
Case Is = 27
  lvList.SetFocus
End Select
End Sub
