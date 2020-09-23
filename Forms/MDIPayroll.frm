VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIPayroll 
   BackColor       =   &H00008000&
   Caption         =   "Payroll System   (o_o)  Registered to edwinSoftware"
   ClientHeight    =   9600
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13050
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIPayroll.frx":0000
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   4
      Top             =   9330
      Width           =   13050
      _ExtentX        =   23019
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   19923
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "7/3/2008"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox PicTop 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00C62B74&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      Picture         =   "MDIPayroll.frx":2A60F4
      ScaleHeight     =   855
      ScaleWidth      =   13050
      TabIndex        =   1
      Top             =   0
      Width           =   13050
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   0
         Top             =   0
      End
      Begin VB.Label lblContacts 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contacts"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   14160
         MouseIcon       =   "MDIPayroll.frx":2B296B
         MousePointer    =   99  'Custom
         TabIndex        =   6
         Top             =   600
         Width           =   765
      End
      Begin VB.Label lblServices 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Services"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   12840
         MouseIcon       =   "MDIPayroll.frx":2B3235
         MousePointer    =   99  'Custom
         TabIndex        =   5
         Top             =   600
         Width           =   750
      End
      Begin VB.Label LblTime 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00:00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C000&
         Height          =   480
         Left            =   13800
         TabIndex        =   2
         Top             =   0
         Width           =   1050
      End
   End
   Begin VB.PictureBox PicLeft 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8475
      Left            =   0
      ScaleHeight     =   8475
      ScaleWidth      =   3255
      TabIndex        =   0
      Top             =   855
      Width           =   3255
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   0
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   36
         ImageHeight     =   36
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   27
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIPayroll.frx":2B3AFF
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIPayroll.frx":2B4BA9
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIPayroll.frx":2B5C53
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIPayroll.frx":2B6CFD
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIPayroll.frx":2B7DA7
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIPayroll.frx":2B8E51
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIPayroll.frx":2B9EFB
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIPayroll.frx":2BAFA5
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIPayroll.frx":2BC04F
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIPayroll.frx":2BD0F9
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIPayroll.frx":2BE1A3
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIPayroll.frx":2BF24D
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIPayroll.frx":2C02F7
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIPayroll.frx":2C13A1
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIPayroll.frx":2C244B
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIPayroll.frx":2C34F5
               Key             =   ""
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIPayroll.frx":2C459F
               Key             =   ""
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIPayroll.frx":2C5279
               Key             =   ""
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIPayroll.frx":2C6323
               Key             =   ""
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIPayroll.frx":2C73CD
               Key             =   ""
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIPayroll.frx":2C8477
               Key             =   ""
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIPayroll.frx":2C9521
               Key             =   ""
            EndProperty
            BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIPayroll.frx":2CA5CB
               Key             =   ""
            EndProperty
            BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIPayroll.frx":2CB675
               Key             =   ""
            EndProperty
            BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIPayroll.frx":2CC71F
               Key             =   ""
            EndProperty
            BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIPayroll.frx":2CD7C9
               Key             =   ""
            EndProperty
            BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIPayroll.frx":2CE873
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.TreeView TreeView1 
         Height          =   4455
         Left            =   0
         TabIndex        =   3
         Top             =   240
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   7858
         _Version        =   393217
         Indentation     =   0
         LabelEdit       =   1
         Style           =   5
         FullRowSelect   =   -1  'True
         SingleSel       =   -1  'True
         Appearance      =   0
         MousePointer    =   99
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "MDIPayroll.frx":2CF91D
      End
   End
End
Attribute VB_Name = "MDIPayroll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private ndx As Integer 'menu index

Private Sub lblContacts_Click()
 Dim msg As String
 msg = "Edwin O. delos Santos" & vbCrLf
 msg = msg & "cyber_edu2005@yahoo.com"
 MsgBox msg, vbInformation, "Contact Me"
End Sub

Private Sub lblServices_Click()
 Dim msg As String
 msg = "Database System Solution" & vbCrLf
 MsgBox msg, vbInformation, "Services"
End Sub

Private Sub MDIForm_Activate()
 PicLeft.SetFocus
End Sub

Private Sub MDIForm_Load()
 ndx = 0
 Dim nd As Node

'setting imagelist1 which contain list of
'images which will be displayed on the tree
Set Me.TreeView1.ImageList = Me.ImageList1
Me.TreeView1.LabelEdit = tvwManual
' adding the main node 1 with key "mnu"
Set nd = Me.TreeView1.Nodes.Add(, , "mnu", "Main", 14)
'expanding the inserted node
nd.Expanded = True
'inserting the sub nodes to the node which contains key as "mnu"
Set nd = Me.TreeView1.Nodes.Add("mnu", tvwChild, "mnu1", "Payroll", 21)
Set nd = Me.TreeView1.Nodes.Add("mnu", tvwChild, "mnu2", "Daily Time Record", 19)
Set nd = Me.TreeView1.Nodes.Add("mnu", tvwChild, "mnu3", "SS Contribution", 9)
Set nd = Me.TreeView1.Nodes.Add("mnu", tvwChild, "mnu4", "Employee Deduction", 22)
Set nd = Me.TreeView1.Nodes.Add("mnu", tvwChild, "mnu5", "Department", 24)
Set nd = Me.TreeView1.Nodes.Add("mnu", tvwChild, "mnu6", "Sick Leave", 7)
Set nd = Me.TreeView1.Nodes.Add("mnu", tvwChild, "mnu7", "Vacation Leave", 8)
'//
' adding the main node 1 with key "tools"
Set nd = Me.TreeView1.Nodes.Add(, , "tools", "Tools", 15)
'expanding the inserted node
nd.Expanded = False
'inserting the sub nodes to the node which contains key as "tools"
Set nd = Me.TreeView1.Nodes.Add("tools", tvwChild, "tool1", "Calculator", 13)
Set nd = Me.TreeView1.Nodes.Add("tools", tvwChild, "tool2", "Back-Up Database", 12)
Set nd = Me.TreeView1.Nodes.Add("tools", tvwChild, "tool3", "Instant Report", 23)
'//
' adding the main node 1 with key "user"
Set nd = Me.TreeView1.Nodes.Add(, , "user", "Users", 16)
'expanding the inserted node
nd.Expanded = False
'inserting the sub nodes to the node which contains key as "user"
Set nd = Me.TreeView1.Nodes.Add("user", tvwChild, "user1", "Log In User", 2)
Set nd = Me.TreeView1.Nodes.Add("user", tvwChild, "user2", "User Maintenance", 27)
'//
' adding the main node 1 with key "help"
Set nd = Me.TreeView1.Nodes.Add(, , "help", "Help", 20)
'expanding the inserted node
nd.Expanded = False
'inserting the sub nodes to the node which contains key as "help"
Set nd = Me.TreeView1.Nodes.Add("help", tvwChild, "hlp1", "Contents", 25)
Set nd = Me.TreeView1.Nodes.Add("help", tvwChild, "hlp2", "Index", 26)


End Sub



Private Sub PicLeft_Resize()
 TreeView1.Width = PicLeft.ScaleWidth
 TreeView1.Height = PicLeft.ScaleHeight
 TreeView1.Top = 240
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)

Dim menuKey As String
Dim menuName As String
menuKey = TreeView1.SelectedItem.Key
menuName = TreeView1.SelectedItem

Select Case menuKey
Case "mnu1"
 If FormLoadedByName("Frmpayroll") = True Then
    MsgBox "The Form is loaded", vbInformation, "Payroll"
    Exit Sub
 Else
   FrmPayroll.show
 End If
  'Call allowACCESS(menuName, FrmPayroll, True)
Case "mnu2"
 If FormLoadedByName("FrmDTR") = True Then
    MsgBox "The Form is loaded", vbInformation, "DTR"
    Exit Sub
 Else
   FrmDTR.show
 End If
Case "mnu3"
   FrmSS.show
Case "mnu4"
  frmDeduction.show
Case "mnu7"
  frmVacation.show
  
Case "tool1"
    FrmCalcu.show
Case "tool2"
    FrmBackUp.show
Case "tool3"
    FrmSQL.show
Case "user1"
  Load frmLogIn
  frmLogIn.show
Case "user2"
   Load FrmUser
   FrmUser.show
End Select

End Sub
Private Sub PicLeft_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 69 And Shift = 4 Then  'alt + e
  FrmPayroll.show
ElseIf KeyCode = 68 And Shift = 4 Then 'alt + d
   FrmDTR.show
'ElseIf KeyCode = 80 And Shift = 4 Then 'alt + p
ElseIf KeyCode = 83 And Shift = 4 Then 'alt + s
ElseIf KeyCode = 70 And Shift = 4 Then 'alt + f
End If
End Sub

Private Sub Timer1_Timer()
LblTime.Caption = Format(Now, "hh:mm:ss AMPM")
End Sub
