VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{9ACEED45-5983-4474-BF17-55AA24019736}#1.0#0"; "esGuard.ocx"
Object = "{49B6E90E-8237-4A84-B038-1CC97F259065}#1.0#0"; "esLabel.ocx"
Begin VB.Form FrmUser 
   Appearance      =   0  'Flat
   BackColor       =   &H00D38545&
   Caption         =   "User File Maintenance"
   ClientHeight    =   6735
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   8790
   HelpContextID   =   240
   Icon            =   "FrmUser.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   Picture         =   "FrmUser.frx":08CA
   ScaleHeight     =   6735
   ScaleWidth      =   8790
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox PicLv 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   360
      ScaleHeight     =   195
      ScaleWidth      =   165
      TabIndex        =   35
      Top             =   3360
      Visible         =   0   'False
      Width           =   195
   End
   Begin eSRedAlert.esGuard Redalert 
      Height          =   375
      Left            =   8160
      TabIndex        =   34
      Top             =   120
      Visible         =   0   'False
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
   End
   Begin VB.TextBox txtEntry 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00D6C5A9&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   0
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   25
      Top             =   370
      Width           =   735
   End
   Begin VB.PictureBox PicEntry 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   240
      ScaleHeight     =   495
      ScaleWidth      =   6975
      TabIndex        =   17
      Top             =   5280
      Width           =   6975
      Begin VB.CommandButton CmdDelete 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Delete"
         Height          =   315
         Left            =   4920
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton CmdRefresh 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Refresh"
         Height          =   315
         Left            =   5880
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Cancel"
         Height          =   315
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Save"
         Height          =   315
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton CmdUpdate 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Update"
         Height          =   315
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton CmdAdd 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Add"
         Height          =   315
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton CmdEdit 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Edit"
         Height          =   315
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   120
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   360
      TabIndex        =   6
      Top             =   720
      Width           =   8175
      Begin VB.ListBox CboMenu 
         Appearance      =   0  'Flat
         BackColor       =   &H00E8FBFB&
         Height          =   1200
         ItemData        =   "FrmUser.frx":C1EF0
         Left            =   1440
         List            =   "FrmUser.frx":C1EF2
         TabIndex        =   32
         Top             =   720
         Visible         =   0   'False
         Width           =   3855
      End
      Begin VB.TextBox txtEntry 
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Index           =   7
         Left            =   5280
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   1440
         Width           =   2055
      End
      Begin VB.TextBox txtPassword 
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
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   5280
         PasswordChar    =   "*"
         TabIndex        =   16
         Text            =   "MACALEN"
         Top             =   720
         Width           =   2655
      End
      Begin VB.TextBox txtEntry 
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Index           =   6
         Left            =   5280
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   1080
         Width           =   2055
      End
      Begin VB.TextBox txtEntry 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   4
         Left            =   1440
         TabIndex        =   27
         Top             =   1080
         Width           =   2535
      End
      Begin VB.TextBox txtEntry 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   5
         Left            =   1440
         TabIndex        =   26
         Top             =   1440
         Width           =   2535
      End
      Begin VB.CheckBox chkAdmin 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Check If User Is Admin"
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   6000
         TabIndex        =   9
         Top             =   360
         Width           =   2055
      End
      Begin VB.TextBox txtEntry 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   1440
         TabIndex        =   8
         Top             =   360
         Width           =   3855
      End
      Begin VB.TextBox txtEntry 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   2
         Left            =   1440
         TabIndex        =   7
         Top             =   720
         Width           =   2535
      End
      Begin VB.TextBox txtEntry 
         BackColor       =   &H000000C0&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   3
         Left            =   5520
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   720
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "F7"
         Height          =   195
         Left            =   5400
         TabIndex        =   33
         Top             =   360
         Width           =   180
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         Caption         =   "DATE EDITED:"
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   4170
         TabIndex        =   30
         Top             =   1440
         Width           =   1125
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         Caption         =   "DATE ADDED:"
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   4200
         TabIndex        =   29
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   "MENU:"
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   240
         TabIndex        =   15
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   "FIRSTNAME:"
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   240
         TabIndex        =   14
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         Caption         =   "PASSWORD:"
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   4200
         TabIndex        =   13
         Top             =   720
         Width           =   1110
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   "LAST NAME:"
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   240
         TabIndex        =   12
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   "USER ID:"
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   240
         TabIndex        =   11
         Top             =   720
         Width           =   1215
      End
   End
   Begin VB.PictureBox hline1 
      Height          =   30
      Left            =   240
      ScaleHeight     =   30
      ScaleWidth      =   8415
      TabIndex        =   5
      Top             =   5760
      Width           =   8415
   End
   Begin VB.PictureBox Picture1 
      Height          =   375
      Left            =   200
      ScaleHeight     =   315
      ScaleWidth      =   8475
      TabIndex        =   3
      Top             =   6240
      Width           =   8535
      Begin MSComctlLib.StatusBar StatusBar1 
         Height          =   300
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   8790
         _ExtentX        =   15505
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
            NumPanels       =   6
            BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               AutoSize        =   1
               Bevel           =   0
               Object.Width           =   6121
               MinWidth        =   882
            EndProperty
            BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Text            =   "User Security >>"
               TextSave        =   "User Security >>"
            EndProperty
            BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Object.Width           =   1058
               MinWidth        =   1058
               Text            =   "Admin:"
               TextSave        =   "Admin:"
            EndProperty
            BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Object.Width           =   529
               MinWidth        =   529
            EndProperty
            BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Style           =   6
               TextSave        =   "6/30/2008"
            EndProperty
            BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Style           =   5
               TextSave        =   "4:37 PM"
            EndProperty
         EndProperty
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   200
      ScaleHeight     =   375
      ScaleWidth      =   8535
      TabIndex        =   1
      Top             =   5880
      Width           =   8535
      Begin esHotKeyLabel.HotKeylabel lblFilter 
         Height          =   255
         Left            =   2160
         TabIndex        =   37
         Top             =   45
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         Caption         =   "Filter"
         CaptionHotkey   =   "[Enter]-"
         ForeColor       =   16777215
         BackColor       =   8421504
         BackColor       =   8421504
         BackColor       =   8421504
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         WidthHotkey     =   645
         Object.Left            =   660
      End
      Begin esHotKeyLabel.HotKeylabel KeyIn 
         Height          =   255
         Left            =   240
         TabIndex        =   36
         Top             =   45
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         Caption         =   "Action"
         CaptionHotkey   =   "KeyIn-"
         ForeColor       =   16777215
         BackColor       =   8421504
         BackColor       =   8421504
         BackColor       =   8421504
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         WidthHotkey     =   555
         Object.Left            =   570
      End
   End
   Begin MSComctlLib.ListView lvUSERNAME 
      Height          =   2535
      Left            =   360
      TabIndex        =   0
      Top             =   2640
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   4471
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      PictureAlignment=   5
      _Version        =   393217
      Icons           =   "i16x16"
      SmallIcons      =   "i16x16"
      ForeColor       =   0
      BackColor       =   14075305
      Appearance      =   0
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
         NumListImages   =   18
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmUser.frx":C1EF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmUser.frx":C2906
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmUser.frx":C2CA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmUser.frx":C303A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmUser.frx":C3A4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmUser.frx":C3AA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmUser.frx":C3E3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmUser.frx":C41D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmUser.frx":C456E
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmUser.frx":C4908
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmUser.frx":C531A
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmUser.frx":C5D2C
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmUser.frx":C673E
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmUser.frx":C7150
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmUser.frx":C7B62
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmUser.frx":C8574
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmUser.frx":C8F86
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmUser.frx":C9522
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Reference Number:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   5160
      TabIndex        =   2
      Top             =   360
      Width           =   1755
   End
End
Attribute VB_Name = "FrmUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rsPASS As ADODB.Recordset
Private rsMENU As ADODB.Recordset
Private s_Admin As Integer
Private dbPass  As String    '//Database Password container


Private Sub CboMenu_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
  txtEntry(1).text = CboMenu.text
  CboMenu.Visible = False
  txtEntry(1).SetFocus
ElseIf KeyCode = 27 Then
  CboMenu.Visible = False
  txtEntry(1).SetFocus
End If
End Sub

Private Sub chkAdmin_Click()
Dim chk As Integer
chk = chkAdmin.Value
If StatusBar1.Panels(4) = "N" Or StatusBar1.Panels(4) = "" Then
   chkAdmin.Value = 0
   Exit Sub
End If
If chk = 0 Then
   s_Admin = 0
Else
   s_Admin = -1
End If
End Sub

Private Sub CmdAdd_Click()
Dim NextNo As Long
On Error GoTo ERRORHANDLE
'//initialize
If CurrUser.USER_isADMIN <> "Y" Then
    MsgBox "You are not authorized", vbCritical, "Warning! Admin Only!"
    Exit Sub
End If

StatusBar1.Panels(4).text = ""
txtEntry(6).text = Format(Now(), "Short Date")
'//
NextNo = Last_Recc(rsPASS)
showButton "A", Me, True, True
If NextNo > 0 Then
 txtEntry(0).text = NextNo
 txtEntry(1).SetFocus
Else
 txtEntry(0).Locked = False
 txtEntry(0).SetFocus
End If

ERRORHANDLE:
  errorMsg Err, Me.Name
End Sub

Private Sub cmdCancel_Click()
showButton "C", Me, True, True
End Sub

Private Sub CmdDelete_Click()
On Error GoTo ERRORHANDLE
If CurrUser.USER_isADMIN <> "Y" Then
    MsgBox "You are not authorized", vbCritical, "Warning"
    Exit Sub
End If
 Call Delete_Record(rsPASS, lvUSERNAME)
ERRORHANDLE:
 errorMsg Err, Me.Name
End Sub

Private Sub CmdEdit_Click()
On Error GoTo ERRORHANDLE
If CurrUser.USER_PASS <> lvUSERNAME.SelectedItem.ListSubItems(3).text Then
  MsgBox "You can not edit other password", vbCritical, "Warning!"
  Exit Sub
End If
txtEntry(7).text = Format(Now(), "Short Date")
  showButton "E", Me, True, True
  txtEntry(1).SetFocus
ERRORHANDLE:
 errorMsg Err, Me.Name

End Sub

Private Sub cmdRefresh_Click()
 If rsPASS.State = adStateOpen Then
   rsPASS.Close
 End If
 rsPASS.Open "select * from users order by SN,userID, menu ", Con, adOpenStatic, adLockOptimistic
 Load_DATA
 lvUSERNAME.SetFocus
End Sub

Private Sub cmdSave_Click()
On Error GoTo errMsg
 showButton "S", Me, True, True
 Call WriteData(Me, rsPASS, True)
 Call lvwPopulateData(lvUSERNAME, rsPASS, 2)
errMsg:
 errorMsg Err, Me.Name, "save"
End Sub

Private Sub CmdUpdate_Click()
On Error GoTo errMsg
showButton "U", Me, True, True
Call WriteData(Me, rsPASS, False)
'LvwReplaceData Me, rsPAY, lvList
errMsg:
  errorMsg Err, Me.Name, "uPDATE"
End Sub

Private Sub Form_Load()
showButton "C", Me, True, True
s_Admin = 0

'[====================]
'<database password   >
'[====================]
RedAlert.xEncrypt RedAlert.xdbPassword
dbPass = RedAlert.xPASS
'[============================]
'< open connection user menu  >
'[============================]
Call OpenDB("USERS.MDB", Con, True, dbPass)
Set rsPASS = New ADODB.Recordset
rsPASS.Open "select * from users order by SN,userID, menu ", Con, adOpenStatic, adLockOptimistic
Load_DATA

'// List BackColour Formatting
Call SetListViewColor(lvUSERNAME, PicLv, &HD6C5A9, vbWhite)

Set rsMENU = New ADODB.Recordset
rsMENU.Open "select * from tbl_menu order by menu", CN, adOpenStatic, adLockReadOnly
Add_Item rsMENU, "MENU", CboMenu

End Sub
Private Sub Load_DATA()
On Error GoTo ERRORHANDLE
'// set columnheaders
Call InsertColumn(lvUSERNAME, rsPASS)
'//set details
 Call FillListView(lvUSERNAME, rsPASS, 2)

ERRORHANDLE:
    errorMsg Err, Me.Name
End Sub

Private Sub Form_Resize()
With FrmUser
  If .WindowState = 0 Then
   .Height = 7095
   .Width = 8880
  End If
End With
SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub

Private Sub lvUSERNAME_Click()
On Error GoTo ERRORHANDLE
If addRec = True Or editRec = True Then Exit Sub
Call BindDatasource(Me, rsPASS, lvUSERNAME, True)
ERRORHANDLE:
 errorMsg Err, Me.Name
End Sub
Private Sub lvUSERNAME_KeyDown(KeyCode As Integer, Shift As Integer)
Dim my_pass As String
Dim user_id As String
my_pass = lvUSERNAME.SelectedItem.ListSubItems(3).text
user_id = lvUSERNAME.SelectedItem.ListSubItems(2).text
If KeyCode = 13 Then
    If rsPASS.State = adStateOpen Then
      rsPASS.Close
    End If
   rsPASS.Open "select * from users where userID like '" & user_id & "' order by menu", Con
   Load_DATA
ElseIf KeyCode = 86 Then  'V-iew
   If CurrUser.USER_isADMIN = "Y" Then  'ONLY ADMIN CAN VEIW PASSWORD
      RedAlert.xText = my_pass 'txtEntry(3).text
      RedAlert.xDecrypt RedAlert.xText
      StatusBar1.Panels(4).text = RedAlert.xADMIN
      StatusBar1.Panels(1).text = RedAlert.xPASS
   End If
   If my_pass = CurrUser.USER_PASS Then
     RedAlert.xText = txtEntry(3).text
     RedAlert.xDecrypt RedAlert.xText
     StatusBar1.Panels(1).text = RedAlert.xPASS
     StatusBar1.Panels(4).text = RedAlert.xADMIN
     If StatusBar1.Panels(4) = "N" Then
        chkAdmin.Value = 0
     End If
  Else
    StatusBar1.Panels(1).text = "You are not autorized to view other password!"
    StatusBar1.Panels(4).text = ""
    chkAdmin.Value = 0
  End If
 End If

End Sub

Private Sub lvUSERNAME_KeyUp(KeyCode As Integer, Shift As Integer)
lvUSERNAME_Click
End Sub




Private Sub txtEntry_GotFocus(Index As Integer)
nxTab = Index
End Sub

Private Sub txtEntry_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim lastTab As Integer

lastTab = 5
If KeyCode = 13 Then
    If nxTab = lastTab Then Exit Sub
    nxTab = nxTab + 1
    If nxTab = 3 Then nxTab = 4 ''// remapping nxtab , passed 3 GOTO 4
    
ElseIf KeyCode = 38 Then  'up arrow key
     If nxTab = 0 Or nxTab = 1 Then Exit Sub
     nxTab = nxTab - 1
     If nxTab = 3 Then nxTab = 2 ''// remapping nxtab , passed 3 BACK TO 2
End If
txtEntry(nxTab).SetFocus

End Sub

Private Sub txtEntry_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 118 Then 'F7
 If Index = 1 Then
   AlignObj txtEntry(1), CboMenu, 1
   CboMenu.SetFocus
 End If
End If
End Sub

Private Sub txtPassword_Change()
If addRec = True Or editRec = True Then

  RedAlert.xText = txtPassword.text
  RedAlert.xEncrypt RedAlert.xText, s_Admin
  txtEntry(3).text = RedAlert.xPASS
End If
End Sub
