VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmDTR 
   BackColor       =   &H00E0E0E0&
   Caption         =   "DTR - Daily Time Record"
   ClientHeight    =   8505
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   12945
   Icon            =   "FrmDTR.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "FrmDTR.frx":109A
   ScaleHeight     =   8505
   ScaleWidth      =   12945
   WindowState     =   2  'Maximized
   Begin VB.PictureBox PicNameList 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3180
      Left            =   3960
      Picture         =   "FrmDTR.frx":2565
      ScaleHeight     =   3150
      ScaleWidth      =   6045
      TabIndex        =   102
      Top             =   3000
      Visible         =   0   'False
      Width           =   6075
      Begin VB.PictureBox PicNameClose 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   5760
         Picture         =   "FrmDTR.frx":47F49
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   103
         Top             =   0
         Width           =   270
      End
      Begin MSComctlLib.ListView lvName 
         Height          =   2595
         Left            =   0
         TabIndex        =   105
         Top             =   480
         Width           =   5985
         _ExtentX        =   10557
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
            Size            =   8.25
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
         TabIndex        =   104
         Top             =   120
         Width           =   1365
      End
   End
   Begin VB.PictureBox PictureUpdAmt 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2820
      Left            =   7440
      Picture         =   "FrmDTR.frx":484D3
      ScaleHeight     =   2790
      ScaleWidth      =   4125
      TabIndex        =   90
      Top             =   5400
      Visible         =   0   'False
      Width           =   4155
      Begin VB.PictureBox PicPercent 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1080
         ScaleHeight     =   315
         ScaleWidth      =   2925
         TabIndex        =   91
         Top             =   2280
         Visible         =   0   'False
         Width           =   2925
         Begin MSComctlLib.ProgressBar ProgressBar1 
            Height          =   255
            Left            =   720
            TabIndex        =   92
            Top             =   15
            Visible         =   0   'False
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   450
            _Version        =   393216
            Appearance      =   0
            Max             =   1
            Scrolling       =   1
         End
         Begin VB.Label lblPercent 
            AutoSize        =   -1  'True
            BackColor       =   &H00E19D86&
            Caption         =   "12%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   0
            TabIndex        =   93
            Top             =   15
            Width           =   690
         End
      End
      Begin VB.PictureBox PictureX 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   3840
         Picture         =   "FrmDTR.frx":8DEB7
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   99
         Top             =   0
         Width           =   270
      End
      Begin VB.CheckBox CheckRefresh 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Caption         =   "Rebuild Index"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2160
         TabIndex        =   97
         Top             =   480
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "GO"
         Height          =   795
         Left            =   120
         MouseIcon       =   "FrmDTR.frx":8E441
         MousePointer    =   99  'Custom
         Picture         =   "FrmDTR.frx":8ED0B
         Style           =   1  'Graphical
         TabIndex        =   95
         Top             =   1800
         Width           =   735
      End
      Begin VB.CommandButton BtnRefresh 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Rebuild Index"
         Enabled         =   0   'False
         Height          =   315
         Left            =   2160
         MouseIcon       =   "FrmDTR.frx":8F5D5
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   94
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label LabelNote 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Note: Be sure that current index must be the employee name!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   2160
         TabIndex        =   101
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "This module will update Employee's Payroll. Take note that it will overwrite existing record."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Index           =   1
         Left            =   120
         TabIndex        =   100
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label lblPayUpdate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Payroll Update"
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
         TabIndex        =   98
         Top             =   120
         Width           =   1260
      End
   End
   Begin VB.PictureBox PictureMenu 
      Appearance      =   0  'Flat
      BackColor       =   &H00E19D86&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4275
      Left            =   10800
      Picture         =   "FrmDTR.frx":8FE9F
      ScaleHeight     =   4275
      ScaleWidth      =   1215
      TabIndex        =   82
      Top             =   480
      Width           =   1215
      Begin VB.CommandButton CmdUpdatePay 
         BackColor       =   &H00C9F3C7&
         Caption         =   "&Payroll Update"
         Height          =   555
         Left            =   120
         MaskColor       =   &H00C9F3C7&
         Style           =   1  'Graphical
         TabIndex        =   96
         Top             =   3480
         Width           =   975
      End
      Begin VB.CommandButton CmdEdit 
         Caption         =   "&Edit"
         Height          =   315
         Left            =   120
         TabIndex        =   89
         Top             =   840
         Width           =   975
      End
      Begin VB.CommandButton CmdAdd 
         Caption         =   "&Add"
         Height          =   315
         Left            =   120
         TabIndex        =   88
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton CmdUpdate 
         Caption         =   "&Update"
         Height          =   315
         Left            =   120
         TabIndex        =   87
         Top             =   1200
         Width           =   975
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   315
         Left            =   120
         TabIndex        =   86
         Top             =   480
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   315
         Left            =   120
         TabIndex        =   85
         Top             =   1560
         Width           =   975
      End
      Begin VB.CommandButton CmdDelete 
         Caption         =   "&Delete"
         Height          =   315
         Left            =   120
         TabIndex        =   84
         Top             =   1920
         Width           =   975
      End
      Begin VB.CommandButton CmdRefresh 
         Caption         =   "&Refresh"
         Height          =   315
         Left            =   120
         TabIndex        =   83
         Top             =   2280
         Width           =   975
      End
   End
   Begin VB.PictureBox PicTop 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   4320
      ScaleHeight     =   345
      ScaleWidth      =   1065
      TabIndex        =   81
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.PictureBox PicLv 
      Appearance      =   0  'Flat
      BackColor       =   &H00ECFEE0&
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
      Left            =   240
      ScaleHeight     =   195
      ScaleWidth      =   225
      TabIndex        =   79
      Top             =   5520
      Width           =   255
   End
   Begin MSComctlLib.ListView lvList 
      Height          =   2115
      Left            =   240
      TabIndex        =   54
      Top             =   5280
      Width           =   4305
      _ExtentX        =   7594
      _ExtentY        =   3731
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
      BackColor       =   -2147483643
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
   Begin VB.Timer Timer2 
      Left            =   120
      Top             =   7320
   End
   Begin VB.PictureBox PicSep 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   0
      ScaleHeight     =   135
      ScaleWidth      =   10455
      TabIndex        =   67
      Top             =   4800
      Width           =   10455
      Begin VB.PictureBox ButtonSep 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   120
         Left            =   240
         MouseIcon       =   "FrmDTR.frx":22D8BB
         MousePointer    =   99  'Custom
         Picture         =   "FrmDTR.frx":22E185
         ScaleHeight     =   120
         ScaleWidth      =   3795
         TabIndex        =   68
         Top             =   0
         Width           =   3795
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   120
         Left            =   120
         Picture         =   "FrmDTR.frx":230167
         Top             =   0
         Visible         =   0   'False
         Width           =   3795
      End
      Begin VB.Image Image2 
         Appearance      =   0  'Flat
         Height          =   120
         Left            =   240
         Picture         =   "FrmDTR.frx":231969
         Top             =   0
         Visible         =   0   'False
         Width           =   3795
      End
   End
   Begin VB.PictureBox picBottom 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   0
      ScaleHeight     =   285
      ScaleWidth      =   12015
      TabIndex        =   58
      Top             =   8160
      Width           =   12015
      Begin VB.TextBox TxtFind 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   275
         Left            =   5880
         TabIndex        =   59
         Text            =   "Find Here !"
         Top             =   0
         Width           =   5415
      End
   End
   Begin VB.PictureBox PicTopBar 
      BackColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      ScaleHeight     =   195
      ScaleWidth      =   555
      TabIndex        =   47
      Top             =   4920
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Picentry 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4335
      Left            =   0
      ScaleHeight     =   4335
      ScaleWidth      =   10815
      TabIndex        =   0
      Top             =   480
      Width           =   10815
      Begin VB.Frame Frame1 
         Height          =   4095
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   10575
         Begin VB.Frame Frame5 
            BackColor       =   &H00DEA576&
            Caption         =   "User Info"
            ForeColor       =   &H00FFFFFF&
            Height          =   1335
            Left            =   8280
            TabIndex        =   71
            Top             =   2640
            Visible         =   0   'False
            Width           =   2175
            Begin VB.TextBox txtEntry 
               Appearance      =   0  'Flat
               BackColor       =   &H00DE9A72&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   22
               Left            =   600
               TabIndex        =   75
               Top             =   240
               Width           =   1215
            End
            Begin VB.TextBox txtEntry 
               Appearance      =   0  'Flat
               BackColor       =   &H00DE9A72&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   23
               Left            =   600
               Locked          =   -1  'True
               TabIndex        =   74
               Top             =   480
               Width           =   1215
            End
            Begin VB.TextBox txtEntry 
               Appearance      =   0  'Flat
               BackColor       =   &H00DE9A72&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   24
               Left            =   600
               Locked          =   -1  'True
               TabIndex        =   73
               Top             =   720
               Width           =   1215
            End
            Begin VB.TextBox txtEntry 
               Appearance      =   0  'Flat
               BackColor       =   &H00DE9A72&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   25
               Left            =   600
               TabIndex        =   72
               Top             =   960
               Width           =   1215
            End
            Begin VB.Label Label17 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Add:"
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   150
               TabIndex        =   69
               Top             =   270
               Width           =   330
            End
            Begin VB.Label Label19 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "User:"
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   105
               TabIndex        =   78
               Top             =   510
               Width           =   375
            End
            Begin VB.Label Label21 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Edit:"
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   165
               TabIndex        =   77
               Top             =   750
               Width           =   315
            End
            Begin VB.Label Label20 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "User:"
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   105
               TabIndex        =   76
               Top             =   990
               Width           =   375
            End
         End
         Begin VB.CheckBox Check1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Caption         =   "User Info"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   9480
            TabIndex        =   70
            Top             =   2400
            Width           =   975
         End
         Begin InstantReport.Hline Hline1 
            Height          =   30
            Left            =   240
            TabIndex        =   64
            Top             =   3120
            Width           =   10215
            _ExtentX        =   18018
            _ExtentY        =   53
         End
         Begin VB.TextBox txtEntry 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   21
            Left            =   6840
            Locked          =   -1  'True
            TabIndex        =   4
            Top             =   3600
            Width           =   1215
         End
         Begin VB.Frame Frame3 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Caption         =   "OVERTIME: "
            ForeColor       =   &H00FF0000&
            Height          =   1215
            Left            =   5640
            TabIndex        =   12
            Top             =   1080
            Width           =   4815
            Begin VB.CheckBox CheckOvertime 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Caption         =   "OVERTIME PAY"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   2640
               TabIndex        =   61
               Top             =   240
               Width           =   1650
            End
            Begin VB.TextBox TxtPERCENT 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   3000
               TabIndex        =   60
               Text            =   "0.00"
               Top             =   600
               Visible         =   0   'False
               Width           =   495
            End
            Begin VB.TextBox txtEntry 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   14
               Left            =   1200
               TabIndex        =   57
               Top             =   840
               Width           =   1215
            End
            Begin VB.TextBox txtEntry 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   11
               Left            =   1440
               TabIndex        =   14
               Top             =   480
               Width           =   975
            End
            Begin VB.TextBox txtEntry 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   10
               Left            =   120
               TabIndex        =   13
               Top             =   480
               Width           =   975
            End
            Begin MSComCtl2.UpDown UpDownOT 
               Height          =   255
               Left            =   2640
               TabIndex        =   62
               Top             =   600
               Visible         =   0   'False
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   450
               _Version        =   393216
               Value           =   1
               Max             =   2
               Enabled         =   -1  'True
            End
            Begin VB.Label Label23 
               AutoSize        =   -1  'True
               Caption         =   "( % )"
               Height          =   195
               Left            =   4320
               TabIndex        =   66
               Top             =   240
               Width           =   300
            End
            Begin VB.Label lblOTNOTE 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "-"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   210
               Left            =   3720
               TabIndex        =   63
               Top             =   600
               Visible         =   0   'False
               Width           =   60
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   " Hours:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Left            =   120
               TabIndex        =   17
               Top             =   840
               Width           =   585
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Start Time"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Left            =   120
               TabIndex        =   16
               Top             =   240
               Width           =   870
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "End Time"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Left            =   1440
               TabIndex        =   15
               Top             =   240
               Width           =   900
            End
         End
         Begin VB.TextBox txtEntry 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   2
            Left            =   3360
            Locked          =   -1  'True
            TabIndex        =   53
            Top             =   240
            Width           =   1455
         End
         Begin InstantReport.Hline ctrlLiner1 
            Height          =   30
            Left            =   7680
            TabIndex        =   52
            Top             =   360
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   53
         End
         Begin VB.TextBox txtEntry 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   1
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   51
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox txtEntry 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   0
            Left            =   9480
            Locked          =   -1  'True
            TabIndex        =   50
            Top             =   120
            Width           =   975
         End
         Begin VB.TextBox txtEntry 
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
            Height          =   285
            Index           =   3
            Left            =   240
            Locked          =   -1  'True
            TabIndex        =   49
            Top             =   720
            Width           =   3735
         End
         Begin VB.TextBox txtEntry 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   5
            Left            =   7560
            TabIndex        =   31
            Top             =   720
            Width           =   1215
         End
         Begin VB.TextBox txtEntry 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   4
            Left            =   5160
            TabIndex        =   30
            Top             =   720
            Width           =   1215
         End
         Begin VB.Frame Frame2 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Caption         =   "MORNING:"
            ForeColor       =   &H00FF0000&
            Height          =   1215
            Left            =   240
            TabIndex        =   24
            Top             =   1080
            Width           =   2535
            Begin VB.TextBox txtEntry 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   12
               Left            =   1200
               TabIndex        =   55
               Top             =   840
               Width           =   1095
            End
            Begin VB.TextBox txtEntry 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   7
               Left            =   1440
               TabIndex        =   26
               Top             =   480
               Width           =   855
            End
            Begin VB.TextBox txtEntry 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   6
               Left            =   120
               TabIndex        =   25
               Top             =   480
               Width           =   975
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Start Time"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Left            =   120
               TabIndex        =   29
               Top             =   240
               Width           =   870
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "End Time"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   0
               Left            =   1440
               TabIndex        =   28
               Top             =   240
               Width           =   780
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   " Hours:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Left            =   120
               TabIndex        =   27
               Top             =   840
               Width           =   585
            End
         End
         Begin VB.Frame Frame4 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Caption         =   "AFTERNOON:"
            ForeColor       =   &H00FF0000&
            Height          =   1215
            Left            =   2880
            TabIndex        =   18
            Top             =   1080
            Width           =   2655
            Begin VB.TextBox txtEntry 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   13
               Left            =   1080
               TabIndex        =   56
               Top             =   840
               Width           =   1215
            End
            Begin VB.TextBox txtEntry 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   8
               Left            =   120
               TabIndex        =   20
               Top             =   480
               Width           =   975
            End
            Begin VB.TextBox txtEntry 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   9
               Left            =   1440
               TabIndex        =   19
               Top             =   480
               Width           =   975
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "End Time"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Left            =   1440
               TabIndex        =   23
               Top             =   240
               Width           =   900
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Start Time"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Left            =   120
               TabIndex        =   22
               Top             =   240
               Width           =   870
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Hours:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Left            =   120
               TabIndex        =   21
               Top             =   840
               Width           =   525
            End
         End
         Begin VB.TextBox TxtHOLIDAY 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   3360
            TabIndex        =   11
            Text            =   "2.00"
            Top             =   3240
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox txtEntry 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   16
            Left            =   3960
            TabIndex        =   10
            Top             =   2400
            Width           =   1215
         End
         Begin VB.TextBox txtEntry 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   17
            Left            =   3960
            Locked          =   -1  'True
            TabIndex        =   9
            Top             =   2760
            Width           =   1215
         End
         Begin VB.TextBox txtEntry 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   15
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   8
            Top             =   2400
            Width           =   1095
         End
         Begin VB.CheckBox CheckSPHOLIDAY 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Caption         =   "SPECIAL HOLIDAY PAY"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   240
            TabIndex        =   7
            Top             =   3600
            Width           =   2250
         End
         Begin VB.TextBox txtEntry 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   19
            Left            =   6840
            Locked          =   -1  'True
            TabIndex        =   6
            Top             =   2400
            Width           =   1215
         End
         Begin VB.TextBox txtEntry 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   20
            Left            =   6840
            Locked          =   -1  'True
            TabIndex        =   5
            Top             =   2760
            Width           =   1215
         End
         Begin VB.TextBox txtEntry 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   18
            Left            =   6840
            TabIndex        =   3
            Top             =   3240
            Width           =   1215
         End
         Begin VB.CheckBox CheckHOLIDAY 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Caption         =   "HOLIDAY PAY"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   240
            TabIndex        =   2
            Top             =   3240
            Width           =   1530
         End
         Begin MSComCtl2.UpDown UpDownSPHOLIDAY 
            Height          =   255
            Left            =   3000
            TabIndex        =   32
            Top             =   3240
            Visible         =   0   'False
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
            _Version        =   393216
            Value           =   1
            Max             =   2
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown UpDownHOLIDAY 
            Height          =   255
            Left            =   3000
            TabIndex        =   33
            Top             =   3240
            Visible         =   0   'False
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
            _Version        =   393216
            Value           =   1
            Max             =   2
            Enabled         =   -1  'True
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "( % )"
            Height          =   195
            Left            =   2520
            TabIndex        =   65
            Top             =   3240
            Width           =   300
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "F2"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   240
            Left            =   4080
            TabIndex        =   48
            Top             =   720
            Width           =   210
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Reference Number:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   7680
            TabIndex        =   46
            Top             =   120
            Width           =   1680
         End
         Begin VB.Label Label27 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "PAY DATE:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   6480
            TabIndex        =   45
            Top             =   720
            Width           =   930
         End
         Begin VB.Label lbl_PAYMODE 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PAYMODE:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   2400
            TabIndex        =   44
            Top             =   240
            Width           =   915
         End
         Begin VB.Label Lbl_IDNO 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ID CODE:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   240
            TabIndex        =   43
            Top             =   240
            Width           =   795
         End
         Begin VB.Label Label18 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "DATE:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   4560
            TabIndex        =   42
            ToolTipText     =   "Click To Get DayName!"
            Top             =   720
            Width           =   525
         End
         Begin VB.Label lblHOLIDAY 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Regular Day"
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   3960
            TabIndex        =   41
            Top             =   3240
            Visible         =   0   'False
            Width           =   885
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Total Hours:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   2880
            TabIndex        =   40
            Top             =   2400
            Width           =   1005
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Basic Pay:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   3000
            TabIndex        =   39
            Top             =   2760
            Width           =   795
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Daily Rate:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   360
            TabIndex        =   38
            Top             =   2400
            Width           =   855
         End
         Begin VB.Label Label13 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Total Hours:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   5760
            TabIndex        =   37
            Top             =   2400
            Width           =   1005
         End
         Begin VB.Label Label14 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Overtime Pay:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   5640
            TabIndex        =   36
            Top             =   2760
            Width           =   1155
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Gross Pay:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   240
            Left            =   5790
            TabIndex        =   35
            Top             =   3600
            Width           =   990
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Adjustment:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   5760
            TabIndex        =   34
            Top             =   3240
            Width           =   1020
         End
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
            NumListImages   =   17
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmDTR.frx":23316B
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmDTR.frx":2331E0
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmDTR.frx":23357A
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmDTR.frx":233914
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmDTR.frx":234326
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmDTR.frx":2346C0
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmDTR.frx":234A5A
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmDTR.frx":234DF4
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmDTR.frx":23518E
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmDTR.frx":235BA0
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmDTR.frx":2365B2
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmDTR.frx":236FC4
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmDTR.frx":2379D6
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmDTR.frx":2383E8
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmDTR.frx":238DFA
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmDTR.frx":23980C
               Key             =   ""
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmDTR.frx":239DA8
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Daily Time Record  (DTR)"
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
      Left            =   120
      TabIndex        =   80
      Top             =   0
      Width           =   2205
   End
End
Attribute VB_Name = "FrmDTR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rs    As ADODB.Recordset
Private rsDTR As ADODB.Recordset
Private rsBUP As ADODB.Recordset
Private showForm As Boolean
Private s4mat As clsFormat
Private totalBy As clsTotalBy
Private Filter As Boolean  '//handle Filter record vital for adding new rec cause the basis for new SN is the last record




Private Sub Check1_Click()
Dim chk As Integer
chk = Check1.Value
If chk = 1 Then
   Check1.Value = 1
Else
   Check1.Value = 0
End If
  Frame5.Visible = (Check1.Value = 1)

End Sub

Private Sub CheckHOLIDAY_Click()
Dim chk As Long
chk = CheckHOLIDAY.Value
If chk = 0 Then
   If CheckSPHOLIDAY.Value = 1 Then
      CheckHOLIDAY.Value = 0
   End If
  '// show hide asst control when unchecked
  lblHOLIDAY.Visible = False
  TxtHOLIDAY.Visible = False
  UpDownHOLIDAY.Visible = False
  UpDownSPHOLIDAY.Visible = False
  TxtHOLIDAY = "1.00"
Else
  CheckHOLIDAY = 1
  CheckSPHOLIDAY = 0
  '// show hide asst control when checked
  TxtHOLIDAY.text = "2.00"
  lblHOLIDAY.Visible = True
  TxtHOLIDAY.Visible = True
  lblHOLIDAY.Enabled = True
  TxtHOLIDAY.Enabled = True
  UpDownHOLIDAY.Visible = True
  UpDownSPHOLIDAY.Visible = False
End If
End Sub

Private Sub CheckOvertime_Click()
Dim chk As Integer
chk = CheckOvertime.Value
If chk = 1 Then
 lblOTNOTE.Visible = True
 UpDownOT.Visible = True
 TxtPERCENT.Visible = True
ElseIf chk = 0 Then
 lblOTNOTE.Visible = False
 UpDownOT.Visible = False
 TxtPERCENT.Visible = False
 TxtPERCENT.text = "0.00"
End If
End Sub

Private Sub CheckRefresh_Click()
BtnRefresh.Enabled = (CheckRefresh.Value = 1)
End Sub

Private Sub CheckSPHOLIDAY_Click()
Dim chk As Long
chk = CheckSPHOLIDAY.Value
If chk = 0 Then
   If CheckHOLIDAY.Value = 1 Then
      CheckSPHOLIDAY.Value = 0
   End If
 '// show hide asst control when unchecked
  lblHOLIDAY.Visible = False
  TxtHOLIDAY.Visible = False
  UpDownHOLIDAY.Visible = False
  UpDownSPHOLIDAY.Visible = False
    TxtHOLIDAY = "1.00"
Else
  CheckSPHOLIDAY = 1
  CheckHOLIDAY.Value = 0
 '// show hide asst control when checked
  TxtHOLIDAY.text = "1.30"
  lblHOLIDAY.Visible = True
  TxtHOLIDAY.Visible = True
  lblHOLIDAY.Enabled = True
  TxtHOLIDAY.Enabled = True
  UpDownHOLIDAY.Visible = False
  UpDownSPHOLIDAY.Visible = True
End If
End Sub


Private Sub CmdAdd_Click()
Dim NextNo As Long
'//initialize
   txtEntry(22).text = Format(Now(), "short date")
   txtEntry(23).text = CurrUser.user_id
'//eS
On Error GoTo ERRORHANDLE
If Filter = True Then
   MsgBox "Record Filtered!, Refresh Record First!", vbCritical, "Warning!"
   Exit Sub
End If
NextNo = Last_Recc(rsDTR)
 showButton "A", Me, True, True
If NextNo > 0 Then
 txtEntry(0).text = NextNo
 txtEntry(3).SetFocus
Else
 txtEntry(0).Locked = False
 txtEntry(0).SetFocus
End If
ERRORHANDLE:
  errorMsg Err, Me.Name, "Add"
End Sub

Private Sub cmdCancel_Click()
showButton "C", Me, True, True
 lvList.SetFocus
End Sub

Private Sub CmdDelete_Click()
 Call Delete_Record(rsDTR, lvList)
End Sub

Private Sub CmdEdit_Click()
'//initialize
   txtEntry(24).text = Format(Now(), "short date")
   txtEntry(25).text = CurrUser.user_id
'//eS
showButton "E", Me, True, True
txtEntry(3).SetFocus
End Sub

Private Sub BtnRefresh_Click()

 If rsDTR.State = adStateOpen Then
    rsDTR.Close
 End If
 rsDTR.Open "SELECT * From DTR order by Employee_Name", CnPay, adOpenStatic, adLockOptimistic
 Load_DATA
 Filter = False
 lvList.SetFocus
End Sub

Private Sub cmdRefresh_Click()
 If rsDTR.State = adStateOpen Then
    rsDTR.Close
 End If
 rsDTR.Open "SELECT * From DTR order by SN", CnPay, adOpenStatic, adLockOptimistic
 Load_DATA
 Filter = False
 lvList.SetFocus

End Sub

Private Sub cmdSave_Click()
On Error GoTo errMsg
 showButton "S", Me, True, True
 Call WriteData(Me, rsDTR, True)
 Call lvwPopulateData(lvList, rsDTR, 2)
 lvList.SetFocus
errMsg:
     errorMsg Err, Me.Name, "Save"
End Sub



Private Sub CmdUpdate_Click()
On Error GoTo errMsg
showButton "U", Me, True, True
Call WriteData(Me, rsDTR, False)
LvwReplaceData Me, rsDTR, lvList, 21
 lvList.SetFocus
errMsg:
  errorMsg Err, Me.Name, "save"
End Sub




Private Sub CmdUpdatePay_Click()
PictureUpdAmt.Visible = True
End Sub

Private Sub Command1_Click()
 Update_Amount
End Sub
Private Sub Update_Amount()
Dim ret
Dim iSEE As Boolean
Dim objFlds As Object
'// DTR
Dim txtname As String
'// Employee
Dim TxtEMPL As String
'//
Dim gPAY As Double
Dim dAYSWRK As Double
Dim basicpay As Double
Dim dAYSOT As Double
Dim OTpay As Double
Dim dpADJUST As Double
Dim netPAY As Double
'// deduction var //
Dim deduct1 As Double
Dim deduct2 As Double
Dim deduct3 As Double
Dim deduct4 As Double
Dim deduct5 As Double
Dim deduct6 As Double
Dim deduct7 As Double
Dim deduct8 As Double
Dim iDEDUCT As Double
'// period var //
Dim dtPayDate As Date
'Dim dto As Date
'//
Dim iDAYS As Double
'----------------------
'//Initialize variables

txtname = "!@#\$%^&*_+"

ret = MsgBox("Proceed with update?", vbYesNo, "Update Amount")
If ret = vbYes Then
Screen.MousePointer = vbHourglass
 Timer2.Enabled = True
 ProgressBar1.Visible = True
 PicPercent.Visible = True
 ProgressBar1.Value = 0

If rsDTR.RecordCount > 0 Then
   ProgressBar1.Max = rsDTR.RecordCount
   rsDTR.MoveFirst
   While Not rsDTR.EOF = True
   ProgressBar1.Value = ProgressBar1.Value + 1
      lblPercent.Caption = ProgressBar1.Value / ProgressBar1.Max * 100
      lblPercent.Caption = Round(Format(Val(lblPercent), "FIXED")) & " %"
      '// initialize variable to search
       txtname = rsDTR.Fields("EMPLOYEE_NAME")
       If rsDTR.Fields("EMPLOYEE_NAME") = txtname Then
      '-------------------------------------
      '// compute days as in  8/8 = 1 day
      iDAYS = rsDTR.Fields("REGULAR_TOTALHRS") / 8
       '// get value of the succeeding records
       '-----------------------------------------
            gPAY = gPAY + rsDTR.Fields("GROSS_PAY")
            dAYSWRK = dAYSWRK + iDAYS
            basicpay = basicpay + rsDTR.Fields("BASIC_PAY")
'            dAYSOT = dAYSOT + rsDTR.Fields("OTHOURS")
            OTpay = OTpay + rsDTR.Fields("OVERTIME_PAY")
            dpADJUST = dpADJUST + rsDTR.Fields("HOLIDAYPAY_ADJUSTMENT")
            '// PAY DATE
            dtPayDate = rsDTR.Fields("PAYROLL_DATE")
           '// get the value of the first record
           '-------------------------------
          If txtname <> TxtEMPL Then
            iDAYS = rsDTR.Fields("REGULAR_TOTALHRS") / 8
            gPAY = rsDTR.Fields("GROSS_PAY")
            dAYSWRK = iDAYS   'rsDTR.Fields("REGULAR_TOTALHRS")
            basicpay = rsDTR.Fields("BASIC_PAY")
'            dAYSOT = rsDTR.Fields("OTHOURS")
            OTpay = rsDTR.Fields("OVERTIME_PAY")
            dpADJUST = rsDTR.Fields("HOLIDAYPAY_ADJUSTMENT")
            '// PAY DATE
            dtPayDate = rsDTR.Fields("PAYROLL_DATE")
            '//======
          End If  'TXTNAME <> TXTEMPL
          
           '// Seek records to update
           '----------------------------------
           If rs.State = adStateOpen Then  'rs  payroll
              rs.Close
           End If
            rs.Open "Select * from PAYROLL where Employee_Name like '" & txtname & "'", CnPay, adOpenStatic, adLockOptimistic
            
                 If rs.RecordCount > 0 Then
                    TxtEMPL = rs!employee_name
                    iSEE = True
                 End If
               If iSEE = True Then
                Set objFlds = rs.Fields
                    If rs.RecordCount > 0 Then
'                     If Not IsNull(objFlds("W_TAX")) Then
'                       deduct1 = objFlds("W_TAX")
'                     Else
'                       deduct1 = 0
'                     End If
                     deduct1 = objFlds("Deduction1")
                     deduct2 = objFlds("Deduction2")
                     deduct3 = objFlds("Deduction3")
                     deduct4 = objFlds("Deduction4")
                     deduct5 = objFlds("Deduction5")
                     deduct6 = objFlds("Deduction6")
                     deduct7 = objFlds("Deduction7")
                     deduct8 = objFlds("Deduction8")
                     iDEDUCT = deduct1 + deduct2
                     iDEDUCT = iDEDUCT + deduct3
                     iDEDUCT = iDEDUCT + deduct4
                     iDEDUCT = iDEDUCT + deduct5
                     iDEDUCT = iDEDUCT + deduct6
                     iDEDUCT = iDEDUCT + deduct7
                     iDEDUCT = iDEDUCT + deduct8
                     objFlds("GROSS_PAY") = gPAY
                     objFlds("Days_Work") = dAYSWRK
'                      objFlds("HRSOT") = dAYSOT
'                      If dAYSOT = 0 Then
'                        objFlds("BASIC") = 0
'                      Else
'                         objFlds("BASIC") = basicpay
'                      End If
                      objFlds("Basic_Pay") = basicpay
                      objFlds("Overtime_Pay") = OTpay
                      objFlds("Adjustment") = dpADJUST
                      objFlds("Total_Deduction") = iDEDUCT
                      objFlds("Net_Pay") = gPAY - iDEDUCT
                      objFlds("Pay_Date") = dtPayDate
'                      objFlds("UPDCODE") = "Y"
                     
                      rs.Update
                    End If  'recordcount > 0
               End If ' isee = true
          End If 'rsDTR.Fields("NAME") = TXTNAME
          iDEDUCT = 0
        rsDTR.MoveNext

   Wend
   
End If 'rs2 > 0

MsgBox "D O N E !", vbOKOnly, "Update"
Else
  MsgBox "Cancelled !", vbOKOnly, "Update"
End If 'ret
 Screen.MousePointer = vbDefault
 Timer2.Enabled = False
 ProgressBar1.Visible = False
 PicPercent.Visible = False
 ProgressBar1.Value = 0

End Sub

Private Sub Form_Load()

center_obj Me, PictureUpdAmt
'//initialize
Set totalBy = New clsTotalBy
Set s4mat = New clsFormat
showForm = True
Filter = False
showButton "C", Me, True, True
show
lvList.SetFocus
'// List BackColour Formatting
Call SetListViewColor(lvList, PicLv, vbWhite, &HECFEE0, 0.1)
'//set controlbox
'pic_controlBox Picture1, picMinimize, picRestore, PicClose
'pic_controlBox Picture2, PicMin, PicRes, PicClose2
'// ope recordset
Set rs = New ADODB.Recordset
Dim SQL As String
SQL = "SELECT Employee_Name,Rate_PerDay,Pay_Mode,ID_Code "
SQL = SQL & "From PAYROLL order by Employee_name"
rs.Open SQL, CnPay, adOpenStatic, adLockOptimistic
Load_Employee


Set rsDTR = New ADODB.Recordset
rsDTR.Open "SELECT * From DTR order by SN", CnPay, adOpenStatic, adLockOptimistic
Load_DATA

End Sub

Private Sub Load_DATA()
On Error GoTo ERRORHANDLE
'// set columnheaders
'Insert_ExtraCol lvList, rsDed
If rsDTR.RecordCount = 0 Then Exit Sub
Call InsertColumn(lvList, rsDTR)
'//set details
 Call FillListView(lvList, rsDTR, 2)
 Call Listview_Total(lvList, rsDTR)
autoAlignCol lvList
ERRORHANDLE:
    errorMsg Err, Me.Name, "Load_Data proc"
End Sub
Private Sub Load_Employee()
On Error GoTo ERRORHANDLE
'// set columnheaders
'Insert_ExtraCol lvList, rsDed
Call InsertColumn(lvName, rs)
'//set details
Call FillListView(lvName, rs, 1)
autoAlignCol lvName
ERRORHANDLE:
    errorMsg Err, Me.Name, "Load_Employee proc"
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'If addRec = True Then
'  myMsg "You have pending record to save...", "Warning!!!", 2, True
'  Exit Sub
'ElseIf editRec = True Then
'  myMsg "You have pending record to update...", "Warning!!!", 2, True
' Exit Sub
'End If
End Sub

Private Sub Form_Resize()
   On Error Resume Next
    If WindowState <> vbMinimized Then
        If Me.Width < 9195 Then Me.Width = 9195
        If Me.Height < 7500 Then Me.Height = 7500
        
        PicTop.Width = ScaleWidth
        Picentry.Width = ScaleWidth
        picBottom.Width = ScaleWidth
        lvList.Left = 0
        lvList.Width = Me.ScaleWidth
        PicSep.Width = ScaleWidth
        
        '//hide show entry form
       If showForm = True Then
          lvList.Top = PicTopBar.Top
          lvList.Height = Me.ScaleHeight - (PicTop.Height + Picentry.Height + PicSep.Height + picBottom.Height + 100)
          picBottom.Top = Me.ScaleHeight - picBottom.Height
        Else
          PicSep.Top = PicTop.Height
          lvList.Top = Picentry.Top
          lvList.Height = Me.ScaleHeight - (PicTop.Height + PicSep.Height + picBottom.Height)
        End If
        
        TxtFind.Left = (picBottom.Width - TxtFind.Width)
    End If
        'PicBottom.Top = (lvList.Height + Pictop.Height + PicTop.Height) + 100
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo errMsg
'MAIN.RemToWin Me.Caption

rsDTR.Close
'rsPAY.Close
'rsBUP.Close

Set rsDTR = Nothing
Set rs = Nothing
'Set rsPAY = Nothing
'Set rsBUP = Nothing
Set FrmDTR = Nothing
Unload Me
errMsg:
  errorMsg Err, Me.Name, "Form_Unload"
End Sub







Private Sub lvName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   txtEntry(3).text = lvName.SelectedItem.text
   txtEntry(15).text = lvName.SelectedItem.ListSubItems(1).text
   txtEntry(2).text = lvName.SelectedItem.ListSubItems(2).text
   txtEntry(1).text = lvName.SelectedItem.ListSubItems(3).text
  txtEntry(3).SetFocus
  PicNameList.Visible = False
End If
End Sub


Private Sub lvList_Click()
On Error GoTo errMsg
If addRec = True Or editRec = True Then Exit Sub
If rsDTR.RecordCount = 0 Then Exit Sub
 Call BindDatasource(Me, rsDTR, lvList, True)
errMsg:
 errorMsg Err, Me.Name, "lv_MouseUP"

End Sub

Private Sub lvList_KeyDown(KeyCode As Integer, Shift As Integer)
Dim sNAME As String
Dim lngID As Long
If addRec = True Or editRec = True Then Exit Sub
On Error GoTo errMsg
Select Case KeyCode
Case Is = 13
 sNAME = lvList.SelectedItem.ListSubItems(3).text
 lngID = Val(lvList.SelectedItem.ListSubItems(1).text)
 If rsDTR.State = adStateOpen Then
    rsDTR.Close
 End If
 'rsDTR.Open "Select * from DTR where employee_Name like '" & sNAME & "' order by date_attend", CnPay
 rsDTR.Open "SELECT * FROM [DTR] WHERE [ID_Code]=" & lngID & " order by date_attend"
 If rsDTR.RecordCount > 0 Then
   Load_DATA
 End If
 Filter = True
End Select
errMsg:
 errorMsg Err, Me.Name, "lvList_KeyDown"
End Sub

Private Sub lvList_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo errMsg
If addRec = True Or editRec = True Then Exit Sub
If rsDTR.RecordCount = 0 Then Exit Sub
Select Case KeyCode
'//-------------------------------
Case Is = 37, 38, 39, 40, 33, 34
  Call BindDatasource(Me, rsDTR, lvList, True)
errMsg:
 errorMsg Err, Me.Name, "lv_MouseUP"
End Select
End Sub

Private Sub lvList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo errMsg
If addRec = True Or editRec = True Then Exit Sub
If rsDTR.RecordCount = 0 Then Exit Sub
If Button = 1 Then
Call BindDatasource(Me, rsDTR, lvList, True, 21)
errMsg:
 errorMsg Err, Me.Name, "lv_MouseUP"
End If

End Sub




Private Sub PicNameClose_Click()
 PicNameList.Visible = False
 txtEntry(3).SetFocus
End Sub

Private Sub PicNameList_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        Call DragIt(PicNameList.hWnd)
    End If
End Sub

Private Sub PicSep_Resize()
Call center_obj_horizontal(PicSep, ButtonSep)
End Sub





Private Sub PictureUpdAmt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        Call DragIt(PictureUpdAmt.hWnd)
    End If
End Sub
Private Sub DragIt(ByVal lngHwnd As Long)
Dim lngReturn As Long
    lngReturn = ReleaseCapture()
    lngReturn = SendMessage(lngHwnd, WM_NCLBUTTONDOWN, HTCAPTION, CLng(0))
End Sub
Private Sub PictureX_Click()
PictureUpdAmt.Visible = False
End Sub

Private Sub txtEntry_Change(Index As Integer)
'On Error GoTo errMsg
If addRec = True Or editRec = True Then txtEntry(Index).ForeColor = vbBlack
Dim perHOUR As Double, HRLY As Double
Dim xHours As String
Dim wkDay As String
Select Case Index
Case Is = 6, 7
 If Len(txtEntry(Index).text) = 4 Then
   xHours = s4mat.ToHour(txtEntry(Index).text)
   txtEntry(Index) = xHours
   xHours = s4mat.isTime(txtEntry(Index).text)
 End If
   txtEntry(12).text = s4mat.HourToDbl(txtEntry(6).text, txtEntry(7).text)
Case Is = 8, 9
 If Len(txtEntry(Index).text) = 4 Then
   xHours = s4mat.ToHour(txtEntry(Index).text)
   txtEntry(Index) = xHours
   xHours = s4mat.isTime(txtEntry(Index).text)
 End If
  txtEntry(13).text = s4mat.HourToDbl(txtEntry(8).text, txtEntry(9).text)
Case Is = 10, 11
 If Len(txtEntry(Index).text) = 4 Then
   xHours = s4mat.ToHour(txtEntry(Index).text)
   txtEntry(Index) = xHours
   xHours = s4mat.isTime(txtEntry(Index).text)
 End If
   txtEntry(14).text = s4mat.HourToDbl(txtEntry(10).text, txtEntry(11).text)
   txtEntry(19).text = toMoney(txtEntry(14))
Case Is = 12, 13
   txtEntry(16).text = totalBy.plus(txtEntry(12).text, txtEntry(13).text)
Case Is = 16
 perHOUR = 0
 HRLY = 0
    perHOUR = Val(txtEntry(15).text) / 8
    HRLY = perHOUR
    txtEntry(17).text = totalBy.times(txtEntry(16).text, HRLY)
Case Is = 17, 18   '//basic pay , adjustment pay
  total_GP
Case Is = 19
 perHOUR = 0
 HRLY = 0
'BASE RATE = P230.00
'If addRec = True Or editRec = True Then
  perHOUR = toMoney(txtEntry(15).text) / 8  '=28.75
  HRLY = totalBy.times(perHOUR, TxtPERCENT.text) ' default(25%) '=7.1875
  HRLY = totalBy.plus(perHOUR, HRLY)     '=35.9375
  txtEntry(20).text = totalBy.times(txtEntry(19).text, HRLY)
'End If
 Case Is = 20   '//overtime pay
   total_GP
 End Select
'errMsg:
 ' errorMsg Err, Me.Name, "Txtetnry_change Events"
End Sub
Private Sub total_GP()
Dim finalPay As Double
finalPay = toMoney(txtEntry(17).text)
finalPay = finalPay + toMoney(txtEntry(18).text)
finalPay = finalPay + toMoney(txtEntry(20).text)
txtEntry(21) = toMoney(finalPay)
End Sub
Private Sub txtEntry_GotFocus(Index As Integer)
On Error GoTo errorMsg
nxTab = Index
txtEntry(nxTab).SelStart = 0
txtEntry(nxTab).SelLength = Len(txtEntry(nxTab).text)
Select Case nxTab
  Case Is = 3
     If addRec = True Or editRec = True Then
       AlignObj txtEntry(3), PicNameList, 1, False
     End If
End Select
errorMsg:
 errorMsg Err, Me.Name, "txtEntry_GotFocus"

End Sub

Private Sub txtEntry_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim lastTab As Integer
On Error GoTo ERRORHANDLE
lastTab = 23
If KeyCode = 13 Then
    If nxTab = lastTab Then Exit Sub
    nxTab = nxTab + 1
    If nxTab = 12 Then nxTab = 10 ''// remapping nxtab , passed 12 back to 10
ElseIf KeyCode = 38 Then  'up arrow key
     If nxTab = 0 Or nxTab = 1 Then Exit Sub
     nxTab = nxTab - 1
End If
txtEntry(nxTab).SetFocus
ERRORHANDLE:
 errorMsg Err, Me.Name

End Sub
Private Sub ButtonSep_Click()
    Separate
End Sub
Private Sub Separate()
   Dim img As Image
    If showForm = True Then
        Set img = Image1
    Else
        Set img = Image2
    End If
 If showForm = False Then
   PicSep.Top = PicTopBar.Top - PicSep.Height
   lvList.Top = PicTopBar.Top
   lvList.Height = Me.ScaleHeight - (PicTop.Height + Picentry.Height + PicSep.Height + picBottom.Height + 100)
   picBottom.Top = Me.ScaleHeight - picBottom.Height
   
   Set ButtonSep.Picture = img.Picture
   showForm = True
 Else
   PicSep.Top = PicTop.Height
   lvList.Top = Picentry.Top
   lvList.Height = Me.ScaleHeight - (PicTop.Height + PicSep.Height + picBottom.Height)
   Set ButtonSep.Picture = img.Picture
   showForm = False
 End If
End Sub
Private Sub txtEntry_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case Index
Case Is = 3
  If addRec = True Or editRec = True Then
    If KeyCode = 113 Then 'F2
       PicNameList.Visible = True
       lvName.SetFocus
    End If
  End If
Case Is = 27
  lvList.SetFocus
End Select
End Sub

Private Sub txtEntry_LostFocus(Index As Integer)
Select Case Index
 Case Is = 4, 5
     txtEntry(Index).text = s4mat.toDate(txtEntry(Index).text)
     If Not IsDate(txtEntry(Index).text) Then
        txtEntry(Index).SetFocus
     End If
 Case 6 To 11
   If Not s4mat.isTime(txtEntry(Index).text) Then
      txtEntry(Index).text = "00:00"
      With txtEntry(Index)
         If .text = "00:00" Then
            .ForeColor = vbRed
         Else
            .ForeColor = vbBlack
         End If
       End With
  End If
End Select
End Sub


Private Sub TxtFind_Change()
    Call ListView_Search(lvList, TxtFind)
End Sub

Private Sub TxtFind_KeyUp(KeyCode As Integer, Shift As Integer)
 If KeyCode = 13 Then lvList.SetFocus
End Sub

Private Sub TxtHOLIDAY_Change()
 Dim perHOUR As Double, HRLY As Double
 Dim pay As Double
If addRec = True Or editRec = True Then
 perHOUR = Val(txtEntry(15).text) / 8
 HRLY = perHOUR
 pay = totalBy.times(txtEntry(16).text, HRLY)
 pay = totalBy.times(pay, TxtHOLIDAY.text)
 txtEntry(18).text = toMoney(totalBy.minus(pay, txtEntry(17).text))
End If
End Sub

Private Sub TxtPERCENT_Change()
Dim perHOUR As Double, HRLY As Double
Dim pay As Double
'//test fire //
'//COMPUTE OVERTIME
'//BASE RATE = P230.00
If addRec = True Or editRec = True Then
  perHOUR = toMoney(txtEntry(15).text) / 8  '=28.75
  HRLY = totalBy.times(perHOUR, TxtPERCENT.text) ' default(25%) '=7.1875
  HRLY = totalBy.plus(perHOUR, HRLY)     '=35.9375
  txtEntry(20).text = totalBy.times(txtEntry(19).text, HRLY)
End If
End Sub

Private Sub UpDownHOLIDAY_Change()
Select Case UpDownHOLIDAY.Value
 Case Is = 0
  TxtHOLIDAY = "1.00"
  lblHOLIDAY.Enabled = False
  TxtHOLIDAY.Enabled = False
 Case Is = 1
  TxtHOLIDAY = "2.00"
  lblHOLIDAY = "Regular Day"
  lblHOLIDAY.Enabled = True
  TxtHOLIDAY.Enabled = True
 Case Is = 2
  TxtHOLIDAY = "2.60"
  lblHOLIDAY = "Rest Day"
  lblHOLIDAY.Enabled = True
  TxtHOLIDAY.Enabled = True
End Select
End Sub

Private Sub UpDownOT_Change()
Dim perHOUR As Double, HRLY As Double

Select Case UpDownOT.Value
 Case Is = 0
  TxtPERCENT.text = "0.00"
  lblOTNOTE = "-"
 Case Is = 1
   TxtPERCENT.text = "0.25"
   lblOTNOTE = "Regular Day"
 Case Is = 2
   TxtPERCENT.text = "0.30"
   lblOTNOTE = "Holiday"
End Select

End Sub

Private Sub UpDownSPHOLIDAY_Change()
Dim perHOUR As Double, HRLY As Double
Dim pay As Double
Select Case UpDownSPHOLIDAY.Value
 Case Is = 0
  TxtHOLIDAY.text = "1.00"
  lblHOLIDAY.Enabled = False
  TxtHOLIDAY.Enabled = False
 Case Is = 1
  TxtHOLIDAY = "1.30"
  lblHOLIDAY = "Regular Day"
  lblHOLIDAY.Enabled = True
  TxtHOLIDAY.Enabled = True
 Case Is = 2
  TxtHOLIDAY = "1.50"
  lblHOLIDAY = "Rest Day"
  lblHOLIDAY.Enabled = True
  TxtHOLIDAY.Enabled = True
End Select

End Sub
