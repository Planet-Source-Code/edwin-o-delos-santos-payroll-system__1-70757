VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{9ACEED45-5983-4474-BF17-55AA24019736}#1.0#0"; "esGuard.ocx"
Begin VB.Form MainForm 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "edwinSotware"
   ClientHeight    =   9420
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14055
   Icon            =   "MainForm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "MainForm.frx":0CCE
   ScaleHeight     =   9420
   ScaleWidth      =   14055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin eSRedAlert.esGuard RedAlert 
      Height          =   375
      Left            =   240
      TabIndex        =   48
      Top             =   120
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
   End
   Begin VB.Timer Timer3 
      Interval        =   100
      Left            =   10080
      Top             =   6840
   End
   Begin VB.CheckBox chkAdmin 
      Appearance      =   0  'Flat
      BackColor       =   &H00CFE081&
      Caption         =   "Admin"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   12360
      TabIndex        =   46
      Top             =   2880
      Width           =   855
   End
   Begin VB.ListBox mnu_ListOption 
      Appearance      =   0  'Flat
      BackColor       =   &H00DE9A72&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   510
      Left            =   3960
      TabIndex        =   39
      ToolTipText     =   "Use [<-]  or [->]  arrow key to move to next menu !"
      Top             =   720
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.PictureBox mnu_Shadow 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   4200
      ScaleHeight     =   1185
      ScaleWidth      =   2745
      TabIndex        =   44
      Top             =   1320
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.PictureBox picUpDwn 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   0
      Left            =   1350
      Picture         =   "MainForm.frx":327E5
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   38
      Top             =   1680
      Width           =   240
   End
   Begin VB.PictureBox picUpDwn 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   1
      Left            =   1350
      Picture         =   "MainForm.frx":32856
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   37
      Top             =   1920
      Width           =   240
   End
   Begin VB.ListBox CboUser 
      Appearance      =   0  'Flat
      BackColor       =   &H00E8FBFB&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1710
      Left            =   5760
      TabIndex        =   26
      Top             =   3000
      Visible         =   0   'False
      Width           =   3135
   End
   Begin InstantReport.Hline ctrlLiner3 
      Height          =   30
      Left            =   9240
      TabIndex        =   32
      Top             =   2760
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   53
   End
   Begin InstantReport.Hline ctrlLiner2 
      Height          =   30
      Left            =   9240
      TabIndex        =   31
      Top             =   3840
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   53
   End
   Begin VB.CommandButton CmdLogIn 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   12240
      Picture         =   "MainForm.frx":328D4
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   3360
      Width           =   1095
   End
   Begin VB.TextBox txtname 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   10200
      MousePointer    =   99  'Custom
      TabIndex        =   28
      ToolTipText     =   "Down arrow key to view user list!"
      Top             =   2880
      Width           =   1935
   End
   Begin VB.TextBox txtPass 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   10200
      MaxLength       =   15
      MousePointer    =   99  'Custom
      PasswordChar    =   "*"
      TabIndex        =   27
      Top             =   3360
      Width           =   1935
   End
   Begin VB.Timer Timer2 
      Interval        =   500
      Left            =   8160
      Top             =   2160
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   7680
      Top             =   2160
   End
   Begin VB.PictureBox PicClose 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   13440
      MouseIcon       =   "MainForm.frx":334DE
      MousePointer    =   99  'Custom
      Picture         =   "MainForm.frx":33DA8
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   23
      Top             =   200
      Width           =   240
   End
   Begin VB.PictureBox PicMinimize 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   12885
      MouseIcon       =   "MainForm.frx":34332
      MousePointer    =   99  'Custom
      Picture         =   "MainForm.frx":34BFC
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   22
      Top             =   200
      Width           =   240
   End
   Begin VB.PictureBox PicRestore 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   13170
      MouseIcon       =   "MainForm.frx":35186
      MousePointer    =   99  'Custom
      Picture         =   "MainForm.frx":35A50
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   21
      Top             =   200
      Width           =   240
   End
   Begin VB.PictureBox PicMenuPointer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   0
      Left            =   1680
      Picture         =   "MainForm.frx":35FDA
      ScaleHeight     =   180
      ScaleWidth      =   225
      TabIndex        =   9
      Top             =   2805
      Width           =   225
   End
   Begin VB.PictureBox PicMenuPointer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   1
      Left            =   1680
      Picture         =   "MainForm.frx":36294
      ScaleHeight     =   180
      ScaleWidth      =   225
      TabIndex        =   8
      Top             =   3045
      Width           =   225
   End
   Begin VB.PictureBox PicMenuPointer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   2
      Left            =   1680
      Picture         =   "MainForm.frx":3654E
      ScaleHeight     =   180
      ScaleWidth      =   225
      TabIndex        =   7
      Top             =   3285
      Width           =   225
   End
   Begin VB.PictureBox PicMenuPointer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   3
      Left            =   1680
      Picture         =   "MainForm.frx":36808
      ScaleHeight     =   180
      ScaleWidth      =   225
      TabIndex        =   6
      Top             =   3525
      Width           =   225
   End
   Begin VB.PictureBox PicMenuPointer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   4
      Left            =   1680
      Picture         =   "MainForm.frx":36AC2
      ScaleHeight     =   180
      ScaleWidth      =   225
      TabIndex        =   5
      Top             =   3765
      Width           =   225
   End
   Begin VB.PictureBox PicMenuPointer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   5
      Left            =   1680
      Picture         =   "MainForm.frx":36D7C
      ScaleHeight     =   180
      ScaleWidth      =   225
      TabIndex        =   4
      Top             =   4005
      Width           =   225
   End
   Begin VB.PictureBox PicMenuPointer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   6
      Left            =   1680
      Picture         =   "MainForm.frx":37036
      ScaleHeight     =   180
      ScaleWidth      =   225
      TabIndex        =   3
      Top             =   4245
      Width           =   225
   End
   Begin VB.PictureBox PicMenuPointer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   7
      Left            =   1680
      Picture         =   "MainForm.frx":372F0
      ScaleHeight     =   180
      ScaleWidth      =   225
      TabIndex        =   2
      Top             =   4485
      Width           =   225
   End
   Begin VB.PictureBox PicMenuPointer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   8
      Left            =   1680
      Picture         =   "MainForm.frx":375AA
      ScaleHeight     =   180
      ScaleWidth      =   225
      TabIndex        =   1
      Top             =   4725
      Width           =   225
   End
   Begin VB.PictureBox PicMenuPointer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   9
      Left            =   1680
      Picture         =   "MainForm.frx":37864
      ScaleHeight     =   180
      ScaleWidth      =   225
      TabIndex        =   0
      Top             =   5040
      Width           =   225
   End
   Begin MSComctlLib.ImageList i16x16 
      Left            =   120
      Top             =   1200
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
            Picture         =   "MainForm.frx":37B1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":38530
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":388CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":38C64
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":39676
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":396CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":39A64
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":39DFE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":3A198
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":3A532
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":3AF44
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":3B956
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":3C368
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":3CD7A
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":3D78C
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":3E19E
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":3EBB0
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":3F14C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   $"MainForm.frx":3F6E8
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1455
      Left            =   9960
      TabIndex        =   47
      Top             =   9240
      Width           =   3855
   End
   Begin VB.Image imgHelp 
      Height          =   360
      Left            =   12960
      MouseIcon       =   "MainForm.frx":3F7F5
      MousePointer    =   99  'Custom
      Picture         =   "MainForm.frx":400BF
      Top             =   2280
      Width           =   360
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   9360
      TabIndex        =   45
      Top             =   3000
      Width           =   660
   End
   Begin VB.Label mnu_Menu 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Menu"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2760
      TabIndex        =   43
      Top             =   240
      Width           =   495
   End
   Begin VB.Label mnu_Tools 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Tools"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3720
      TabIndex        =   42
      Top             =   240
      Width           =   525
   End
   Begin VB.Label mnu_File 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&File"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2040
      TabIndex        =   41
      Top             =   240
      Width           =   330
   End
   Begin VB.Label mnu_Help 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Help"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4680
      TabIndex        =   40
      Top             =   240
      Width           =   435
   End
   Begin VB.Label lblSelected 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "v"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Left            =   3600
      TabIndex        =   36
      Top             =   2400
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label lblDate 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01/01/1999"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   8880
      TabIndex        =   35
      Top             =   9000
      Width           =   1200
   End
   Begin VB.Label lblUserName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   360
      TabIndex        =   34
      Top             =   8760
      Width           =   1005
   End
   Begin VB.Label lblUserID 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User  ID    : Logged Out"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   360
      TabIndex        =   33
      Top             =   9000
      Width           =   2070
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H000000C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   9240
      TabIndex        =   29
      Top             =   3360
      Width           =   885
   End
   Begin VB.Label lblAdvisory 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "  system designed by: edwin o. delos santos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   10440
      TabIndex        =   25
      Top             =   8640
      Width           =   3240
   End
   Begin VB.Label LblTime 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00:00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   12720
      TabIndex        =   24
      Top             =   9000
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   9240
      Picture         =   "MainForm.frx":40829
      Top             =   195
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Image2 
      Height          =   240
      Left            =   9600
      Picture         =   "MainForm.frx":40DB3
      Top             =   195
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label lblMenu 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MENU0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   2040
      TabIndex        =   20
      Top             =   2760
      Width           =   570
   End
   Begin VB.Label lblMenu 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MENU1"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   2040
      TabIndex        =   19
      Top             =   3000
      Width           =   570
   End
   Begin VB.Label lblMenu 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MENU2"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   2040
      TabIndex        =   18
      Top             =   3240
      Width           =   570
   End
   Begin VB.Label lblMenu 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MENU3"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   2040
      TabIndex        =   17
      Top             =   3510
      Width           =   570
   End
   Begin VB.Label lblMenu 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MENU4"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   4
      Left            =   2040
      TabIndex        =   16
      Top             =   3765
      Width           =   570
   End
   Begin VB.Label lblMenu 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MENU5"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   5
      Left            =   2040
      TabIndex        =   15
      Top             =   4005
      Width           =   570
   End
   Begin VB.Label lblMenu 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MENU6"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   6
      Left            =   2040
      TabIndex        =   14
      Top             =   4245
      Width           =   570
   End
   Begin VB.Label lblMenu 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MENU7"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   7
      Left            =   2040
      TabIndex        =   13
      Top             =   4485
      Width           =   570
   End
   Begin VB.Label lblMenu 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MENU8"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   8
      Left            =   2040
      TabIndex        =   12
      Top             =   4725
      Width           =   570
   End
   Begin VB.Label lblMenu 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MENU9"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   9
      Left            =   2040
      TabIndex        =   11
      Top             =   4965
      Width           =   570
   End
   Begin VB.Label lblMenuOver 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Edwin Delos Santos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   1680
      MouseIcon       =   "MainForm.frx":4133D
      MousePointer    =   99  'Custom
      TabIndex        =   10
      Top             =   2445
      Visible         =   0   'False
      Width           =   1830
   End
   Begin VB.Shape shadow_mnu 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00008000&
      BorderWidth     =   3
      Height          =   255
      Left            =   3000
      Shape           =   4  'Rounded Rectangle
      Top             =   1080
      Visible         =   0   'False
      Width           =   615
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public rsMENU         As Recordset
Private rs            As Recordset
Private spacer        As Long
Private iTRY           As Integer
Private bRestore       As Boolean 'hanlde restore to maximize
Private bMin           As Boolean 'handle minimize state
Private lblMenu_index  As Integer ' handle menu index
Private lf As Integer, tp As Integer, wd As Integer, ht As Integer  'handle save max procedure
'// events class menu
Event TakboNa()      'escape
Event SaKaliwa()     'move to left
Event SaKanan()      ' move to right
Event FileClick()
Event MenuClick()
Event ToolsClick()
Event HelpClick()
Private MyMenu As clsMenu   '//CLASS MENU
Private dbPass As String    '//Database Password container

Private Sub CboUser_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
 txtname.text = CboUser.text
 CboUser.Visible = False
 txtPass.SetFocus
ElseIf KeyCode = 27 Then
 CboUser.Visible = False
 txtPass.SetFocus
End If
End Sub

Private Sub chkAdmin_Click()
  txtname.SetFocus
End Sub

Private Sub LockLogIN()
 CmdLogIn.Enabled = False
End Sub



Private Sub cmdLogin_Click()
On Error GoTo ERRORHANDLE
If iTRY = 3 Then
: Rem lblTry.Caption = "No more try! Exit now!"
   Exit Sub
End If
 iTRY = iTRY + 1

If txtname.text = "" Or txtPass.text = "" Then
    txtname.SetFocus
End If
If rs.State = adStateOpen Then
  rs.Close
End If
rs.Open "select * from users ", Con, 1, 1, 1
rs.MoveFirst
Do Until rs.EOF
   If rs.Fields("userID") = txtname.text And _
      rs.Fields("password") = RedAlert.xPASS Then
      
      CurrUser.user_id = rs.Fields("UserID")
      CurrUser.USER_PASS = rs.Fields("password")
      CurrUser.USER_NAME = rs.Fields("FIRSTNAME") & " " & rs.Fields("LASTNAME")
      CurrUser.USER_isADMIN = RedAlert.xADMIN
      iTRY = 0
      GoTo iFound
   Else
      rs.MoveNext
   End If
   Loop
 
   txtPass.text = Empty
   txtname.SetFocus
If iTRY = 1 Then
: Rem lblTry.Caption = "Invalid Username or pasword! Try again"
ElseIf iTRY = 2 Then
: Rem lblTry.Caption = "Invalid Username or pasword! Try again"
ElseIf iTRY = 3 Then
: Rem lblTry.Caption = "Invalid Username or pasword!"
   End If
   
   Exit Sub

iFound:

With MainForm.lblUserName
  .Caption = "User Name: "
  .Caption = .Caption & CurrUser.USER_NAME
End With
With MainForm.lblUserID
  .Caption = "User  ID    : "
  .Caption = .Caption & CurrUser.user_id
End With
'// determine access
Menu_List rs, CurrUser.USER_PASS

txtname.text = Empty
txtPass = Empty

LockLogIN

ERRORHANDLE:
 errorMsg Err, Me.Name, "CmdLogin"
End Sub




 


Private Sub Form_Load()
On Error GoTo errMsg
App.HelpFile = App.Path & "\instantreport.chm"
'[===============================================================]
'< Not the best way to check                                     >
'< Better to use the FindWindow API, when running an exe file    >
'[===============================================================]
If App.PrevInstance = True Then
  MsgBox "This application is already running.", vbInformation, "Warning!"
  End
End If
'[=======================]
'< class menu settings   >
'[=======================]
   Set MyMenu = New clsMenu
   Set MyMenu.MainForm = Me
'[===========]
'<Initilize  >
'[===========]
RedAlert.xEncrypt RedAlert.xdbPassword
lblDate.Caption = Format(Now(), "dddd,mmmm dd,yyyy")
bRestore = False
bMin = True
FormRndCorner Me, 935, 630
pIndex = 0

txtname.text = Empty
txtPass.text = Empty
show
txtname.SetFocus
'[====================]
'<database password   >
'[====================]
dbPass = RedAlert.xPASS
'[============================]
'< open connection user menu  >
'[============================]
Call OpenDB("USERS.MDB", Con, True, dbPass)
Set rs = New ADODB.Recordset
rs.Open "select * from users order by userID, menu ", Con, 1, 1, 1
Call Add_Item(rs, "userID", CboUser)


 Set rsMENU = New ADODB.Recordset
 Call Menu_Caption(rsMENU)
 
 Call Menu_Position(2000, 365, 3000)
errMsg:
  errorMsg Err, Me.Name, "Form_Load"
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 70 And Shift = 4 Then  'alt + f
   mnu_File_Click
ElseIf KeyCode = 77 And Shift = 4 Then 'alt + m
   mnu_Menu_Click
ElseIf KeyCode = 84 And Shift = 4 Then 'alt + t
   mnu_Tools_Click
ElseIf KeyCode = 72 And Shift = 4 Then 'alt + h
   mnu_Help_Click
End If
End Sub



Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   down = True
    w = X
    t = Y
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If down Then
        Top = Top + Y - t
        Left = Left + X - w
   End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 down = False
End Sub



Private Sub imgHelp_Click()
  myMsg "Check Admin if you want to log-in " _
  & "as admin." & vbCrLf & vbCrLf _
  & "Login button will be locked upon login correctly. " _
  & "To enable Login button type your User ID.", "Password Help", 1, True
End Sub

Private Sub lblMenu_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
pIndex = Index
 Call HK(lblMenu(pIndex), lblMenuOver)
End Sub
Private Sub lblMenuOver_DblClick()
Select Case pIndex
 Case Is = 0
      FrmSQL.show
     'Call allowACCESS(lblMenu(pIndex), FrmSQL)
 Case Is = 1
      FrmProdList.show
     'Call allowACCESS(lblMenu(pIndex), FrmProdList)
Case Is = 2
    Call allowACCESS(lblMenu(pIndex), FrmStockReceive)
 Case Is = 3
    MDIPayroll.show
    'Call allowACCESS(lblMenu(pIndex), MDIPayroll, True)
 Case Is = 4
    If RedAlert.xADMIN = "Y" Then
       Call allowACCESS(lblMenu(pIndex), FrmUser)
    Else
       MsgBox "For Admin Only!"
    End If
 Case Is = 5
     FrmCalcu.show
     'Call allowACCESS(lblMenu(pIndex), FrmCalcu)
 Case Is = 6
     FrmBackUp.show
     'Call allowACCESS(lblMenu(pIndex), frmAbout)
 Case Is = 7
      OpenConvert
     'Call allowACCESS(lblMenu(pIndex), frmAbout)
 Case Is = 8
     Call allowACCESS(lblMenu(pIndex), frmAbout)
 Case Is = 9
     log_out
End Select
    lockMenu True
End Sub

Private Sub lockMenu(ByVal sconfirm As Boolean)
Dim i As Integer
For i = 0 To lblMenu.UBound
    If sconfirm = True Then
      lblMenu(i).Enabled = False
      lblMenuOver.Enabled = False
    Else
      lblMenu(i).Enabled = True
      lblMenuOver.Enabled = True
    End If
    Next i
End Sub
Private Sub log_out()
If MsgBox("Do you want to logout?", vbYesNo + vbQuestion, "edwinSoftware") = vbYes Then
 Dim i As Integer
 For i = 0 To 50
      sMENU(i) = vbNullString
  Next i
  i = i + 1
 lblUserID.Caption = "User  ID    : Logged Out"
 lblUserName.Caption = "User Name:"
 CurrUser.USER_isADMIN = "N"
 CurrUser.USER_PASS = ""
 
End If
End Sub

Private Sub Menu_Position(ByRef Start As Long, ByVal spac As Long, ByVal lft As Integer)
Dim i As Long
For i = 0 To lblMenu.UBound
       PicMenuPointer(i).Left = lft
       spacer = spacer + spac
       lblMenu(i).Top = Start + spacer
       PicMenuPointer(i).Top = Start + spacer
       lblMenu(i).Alignment = 0
       lblMenu(i).Caption = iMenuCap(i)
       lblMenu(i).Left = PicMenuPointer(i).Left + 500
   Next i
End Sub




Private Sub lblMenuOver_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 With lblSelected
  .Visible = True
  .Top = lblMenuOver.Top
  .Left = (lblMenuOver.Left + lblMenuOver.Width + 100)
 End With
 
End Sub

Private Sub mnu_File_Click()
  RaiseEvent FileClick
End Sub


Private Sub mnu_ListOption_DblClick()
Dim iSelect As String
iSelect = TrimSpaces(mnu_ListOption.text)
Select Case iSelect
 Case Is = "Exit"
 '//FILE--------------------------------------------------
   PicClose_Click
'//MENU---------------------------------------------------
 Case Is = "PayrollSystem"
      Call allowACCESS(lblMenu(pIndex), FrmPayroll)
 '//HELP--------------------------------------------------
 Case Is = "ContentsF1"
   HHShowContents Me.hWnd
 Case Is = "Index"
   HHShowIndex Me.hWnd
 Case Is = "Search"
   HHShowSearch Me.hWnd
 Case Is = "ContactUs"
   Dim msg As String
   msg = "09206747545"
   msg = msg & Chr(10) & "cyber_edu2005@yahoo.com"
   MsgBox msg, vbInformation, "edwinSoftware"
End Select
End Sub

Private Sub mnu_Help_Click()
  RaiseEvent HelpClick
End Sub
Private Sub mnu_ListOption_KeyDown(KeyCode As Integer, Shift As Integer)
'// coded by edwin delos santos
If KeyCode = 27 Then 'escape
   RaiseEvent TakboNa
ElseIf KeyCode = 37 Then   '<- arrow key
   RaiseEvent SaKaliwa
ElseIf KeyCode = 39 Then   '-> arrow key
   RaiseEvent SaKanan
End If
End Sub


Private Sub mnu_ListOption_KeyUp(KeyCode As Integer, Shift As Integer)
With mnu_ListOption
  If Mid(.text, 1, 1) = "-" Or Mid(.text, 1, 1) = "<" Then
     If KeyCode = 40 Then
       .ListIndex = .ListIndex + 1
     Else
       .ListIndex = .ListIndex - 1
     End If
  End If
End With
End Sub

Private Sub mnu_Menu_Click()
  RaiseEvent MenuClick
End Sub

Private Sub mnu_Tools_Click()
 RaiseEvent ToolsClick
End Sub


Private Sub PicClose_Click()
        Dim Cancel As Boolean
        If MsgBox("This will close the application.Do you want to proceed?", vbYesNo + vbQuestion, "edwinSoftware") = vbNo Then
           Cancel = True
        Else
'           UnloadAllForms
'           UnloadChilds
           Unload Me
           End
        End If
End Sub

Private Sub PicMinimize_Click()
  FormMinimize
End Sub

Private Sub PicRestore_Click()
 FormRestore
End Sub
Private Sub FormMinimize()
 If bMin = False Then Exit Sub
 Set PicRestore.Picture = Image2
 PicRestore.Enabled = True
 PicMinimize.Enabled = False
 bRestore = True
 bMin = False
 Save_FrmMax
 Me.Move 100, Me.Height + 1000, Me.Width, 600
End Sub
Private Sub FormRestore()
  If bRestore = False Then Exit Sub
  bMin = True
  Set PicRestore.Picture = Image1
  PicRestore.Enabled = False
  PicMinimize.Enabled = True
  Me.Move lf, tp, wd, ht
End Sub
Private Sub Save_FrmMax()
   lf = Me.Left
   tp = Me.Top
   wd = Me.Width
   ht = Me.Height
End Sub
Private Sub picUpDwn_Click(Index As Integer)
lockMenu False
Select Case Index
Case Is = 1
 If lblMenu.Count - pIndex = 1 Then Exit Sub
  pIndex = pIndex + 1
  Call HK(lblMenu(pIndex), lblMenuOver)
Case Is = 0
 If pIndex = 0 Then Exit Sub
  pIndex = pIndex - 1
  Call HK(lblMenu(pIndex), lblMenuOver)
End Select
 With lblSelected
  .Visible = True
  .Top = lblMenuOver.Top
  .Left = (lblMenuOver.Left + lblMenuOver.Width + 100)
 End With
End Sub

Private Sub Timer1_Timer()
LblTime.Caption = Format(Now, "hh:mm:ss AMPM")
End Sub
Private Sub Timer2_Timer()
   marque_2left
End Sub

Private Sub marque_2left()
    Dim str As String
    str = lblAdvisory.Caption
    str = Mid$(str, 2, Len(str)) + Left(str, 1)
    lblAdvisory.Caption = str
End Sub

Private Sub Timer3_Timer()
Label7.Top = Label7.Top - 20
If Label7.Top <= -1320 Then
Label7.Top = 9360
End If
End Sub

Private Sub txtname_Change()
 CboUser.ListIndex = SendMessage(CboUser.hWnd, LB_FINDSTRING, -1, ByVal txtname.text)
  If UCase$(TrimSpaces(txtname)) = UCase$(CboUser.text) Then
     CmdLogIn.Enabled = True
     If iTRY = 3 Then
        iTRY = 0
: Rem       lblTry.Caption = "Access granted!"
     End If
  End If
End Sub

Private Sub txtname_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then
    txtPass.SetFocus
  End If
  If KeyCode = 40 Then  'down arrow
     AlignObj txtname, CboUser, 1
     CboUser.SetFocus
     CboUser.ListIndex = 0
  End If
  If KeyCode = 27 Then
    CboUser.Visible = False
  End If
End Sub


Private Sub txtPass_Change()
On Error GoTo errMsg
Dim isAdmin As Integer
If chkAdmin.Value = 0 Then
   isAdmin = 0
Else
   isAdmin = -1
End If
 RedAlert.xText = txtPass.text
 RedAlert.xEncrypt RedAlert.xText, isAdmin
errMsg:
 errorMsg Err, Me.Name, "TxtPass_change"
End Sub

Private Sub txtPass_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then
   If CmdLogIn.Enabled = True Then
     CmdLogIn.SetFocus
   End If
  ElseIf KeyCode = 38 Then
   txtname.SetFocus
  End If
  If KeyCode = 27 Then
    CboUser.Visible = False
  End If
End Sub

Private Sub OpenConvert()
On Error GoTo errorRoutineErr

Dim hProcess As Long
Dim retval As Long
Dim slAppToRun As String
slAppToRun = App.Path & "\ConvertXDB.exe"

'The next line launches Notepad
hProcess = OpenProcess(PROCESS_QUERY_INFORMATION, 1, _
   Shell(slAppToRun, vbNormalFocus))

errorRoutineResume:
  Exit Sub
errorRoutineErr:
 MsgBox "Open file " & Err & Error
 Resume Next

End Sub
