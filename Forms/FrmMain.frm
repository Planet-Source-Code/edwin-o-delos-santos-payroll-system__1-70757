VERSION 5.00
Begin VB.Form FrmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "edwinSotware"
   ClientHeight    =   9450
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14025
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmMain.frx":0CCE
   ScaleHeight     =   9450
   ScaleWidth      =   14025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox PicMenuPointer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   0
      Left            =   4320
      Picture         =   "FrmMain.frx":2C5FF
      ScaleHeight     =   180
      ScaleWidth      =   195
      TabIndex        =   9
      Top             =   2805
      Width           =   195
   End
   Begin VB.PictureBox PicMenuPointer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   1
      Left            =   4320
      Picture         =   "FrmMain.frx":2C64B
      ScaleHeight     =   180
      ScaleWidth      =   195
      TabIndex        =   8
      Top             =   3045
      Width           =   195
   End
   Begin VB.PictureBox PicMenuPointer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   2
      Left            =   4320
      Picture         =   "FrmMain.frx":2C697
      ScaleHeight     =   180
      ScaleWidth      =   195
      TabIndex        =   7
      Top             =   3285
      Width           =   195
   End
   Begin VB.PictureBox PicMenuPointer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   3
      Left            =   4320
      Picture         =   "FrmMain.frx":2C6E3
      ScaleHeight     =   180
      ScaleWidth      =   195
      TabIndex        =   6
      Top             =   3525
      Width           =   195
   End
   Begin VB.PictureBox PicMenuPointer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   4
      Left            =   4320
      Picture         =   "FrmMain.frx":2C72F
      ScaleHeight     =   180
      ScaleWidth      =   195
      TabIndex        =   5
      Top             =   3765
      Width           =   195
   End
   Begin VB.PictureBox PicMenuPointer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   5
      Left            =   4320
      Picture         =   "FrmMain.frx":2C77B
      ScaleHeight     =   180
      ScaleWidth      =   195
      TabIndex        =   4
      Top             =   4005
      Width           =   195
   End
   Begin VB.PictureBox PicMenuPointer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   6
      Left            =   4320
      Picture         =   "FrmMain.frx":2C7C7
      ScaleHeight     =   180
      ScaleWidth      =   195
      TabIndex        =   3
      Top             =   4245
      Width           =   195
   End
   Begin VB.PictureBox PicMenuPointer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   7
      Left            =   4320
      Picture         =   "FrmMain.frx":2C813
      ScaleHeight     =   180
      ScaleWidth      =   195
      TabIndex        =   2
      Top             =   4485
      Width           =   195
   End
   Begin VB.PictureBox PicMenuPointer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   8
      Left            =   4320
      Picture         =   "FrmMain.frx":2C85F
      ScaleHeight     =   180
      ScaleWidth      =   195
      TabIndex        =   1
      Top             =   4725
      Width           =   195
   End
   Begin VB.PictureBox PicMenuPointer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   9
      Left            =   4320
      Picture         =   "FrmMain.frx":2C8AB
      ScaleHeight     =   180
      ScaleWidth      =   195
      TabIndex        =   0
      Top             =   4965
      Width           =   195
   End
   Begin VB.Label lblMenu 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MENU0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   4680
      TabIndex        =   20
      Top             =   2760
      Width           =   570
   End
   Begin VB.Label lblMenu 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MENU1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   4680
      TabIndex        =   19
      Top             =   3000
      Width           =   570
   End
   Begin VB.Label lblMenu 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MENU2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   4680
      TabIndex        =   18
      Top             =   3240
      Width           =   570
   End
   Begin VB.Label lblMenu 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MENU3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   4680
      TabIndex        =   17
      Top             =   3510
      Width           =   570
   End
   Begin VB.Label lblMenu 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MENU4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   4
      Left            =   4680
      TabIndex        =   16
      Top             =   3765
      Width           =   570
   End
   Begin VB.Label lblMenu 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MENU5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   5
      Left            =   4680
      TabIndex        =   15
      Top             =   4005
      Width           =   570
   End
   Begin VB.Label lblMenu 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MENU6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   6
      Left            =   4680
      TabIndex        =   14
      Top             =   4245
      Width           =   570
   End
   Begin VB.Label lblMenu 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MENU7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   7
      Left            =   4680
      TabIndex        =   13
      Top             =   4485
      Width           =   570
   End
   Begin VB.Label lblMenu 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MENU8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   8
      Left            =   4680
      TabIndex        =   12
      Top             =   4725
      Width           =   570
   End
   Begin VB.Label lblMenu 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MENU9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   9
      Left            =   4680
      TabIndex        =   11
      Top             =   4965
      Width           =   570
   End
   Begin VB.Label lblMenuOver 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Left            =   5640
      MouseIcon       =   "FrmMain.frx":2C8F7
      MousePointer    =   99  'Custom
      TabIndex        =   10
      Top             =   2805
      Visible         =   0   'False
      Width           =   1710
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
pIndex = 0
 Set rsMenu = New ADODB.Recordset
 Call Menu_Caption(rsMenu)
 Call Menu_Position(2505, 365, 840)
End Sub
