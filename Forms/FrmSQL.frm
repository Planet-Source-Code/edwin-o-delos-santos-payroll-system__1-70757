VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmSQL 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Instant Report ®"
   ClientHeight    =   10710
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   Icon            =   "FrmSQL.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10710
   ScaleWidth      =   15240
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picPrint 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   4095
      Left            =   5880
      Picture         =   "FrmSQL.frx":1CCA
      ScaleHeight     =   4065
      ScaleWidth      =   7185
      TabIndex        =   12
      Top             =   3240
      Visible         =   0   'False
      Width           =   7215
      Begin VB.CommandButton cmdPrint 
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   5760
         MouseIcon       =   "FrmSQL.frx":B7286
         MousePointer    =   99  'Custom
         Picture         =   "FrmSQL.frx":B7B50
         Style           =   1  'Graphical
         TabIndex        =   122
         ToolTipText     =   "Print"
         Top             =   3600
         Width           =   1215
      End
      Begin VB.CommandButton CmdOk 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4320
         Picture         =   "FrmSQL.frx":B869E
         Style           =   1  'Graphical
         TabIndex        =   121
         Top             =   3600
         Width           =   1215
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   0
         Picture         =   "FrmSQL.frx":B91EC
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   104
         Top             =   0
         Width           =   480
      End
      Begin VB.PictureBox PicClose 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   300
         Left            =   6720
         MouseIcon       =   "FrmSQL.frx":BA18E
         Picture         =   "FrmSQL.frx":BAA58
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   19
         Top             =   0
         Width           =   300
      End
      Begin VB.Frame FmePrint 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Print Option ( Print only what you want ! )"
         ForeColor       =   &H00C00000&
         Height          =   2895
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   6975
         Begin VB.ListBox ListPrint 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   2280
            ItemData        =   "FrmSQL.frx":BADE2
            Left            =   3840
            List            =   "FrmSQL.frx":BADE4
            Style           =   1  'Checkbox
            TabIndex        =   17
            Top             =   360
            Width           =   2895
         End
         Begin VB.ListBox List2Print 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   2370
            ItemData        =   "FrmSQL.frx":BADE6
            Left            =   120
            List            =   "FrmSQL.frx":BADED
            Sorted          =   -1  'True
            TabIndex        =   16
            Top             =   360
            Width           =   2895
         End
         Begin VB.CommandButton CmdMove 
            Caption         =   ">"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3120
            TabIndex        =   15
            Top             =   1080
            Width           =   615
         End
         Begin VB.CommandButton CmdMoveBack 
            Caption         =   "<"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3120
            TabIndex        =   14
            Top             =   1560
            Width           =   615
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Validate list,  Click check button. "
            ForeColor       =   &H000000C0&
            Height          =   195
            Left            =   3840
            TabIndex        =   18
            Top             =   0
            Width           =   2355
         End
      End
   End
   Begin VB.PictureBox PicLv 
      Appearance      =   0  'Flat
      BackColor       =   &H00F7EBD0&
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
      Left            =   0
      ScaleHeight     =   195
      ScaleWidth      =   165
      TabIndex        =   119
      Top             =   5520
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox picConvert 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   4095
      Left            =   5040
      Picture         =   "FrmSQL.frx":BAE00
      ScaleHeight     =   4065
      ScaleWidth      =   7185
      TabIndex        =   97
      Top             =   360
      Visible         =   0   'False
      Width           =   7215
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   0
         Picture         =   "FrmSQL.frx":1703BC
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   103
         Top             =   0
         Width           =   480
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Select Field To Convert:"
         ForeColor       =   &H00C00000&
         Height          =   3375
         Left            =   120
         TabIndex        =   99
         Top             =   600
         Width           =   6975
         Begin VB.CheckBox ChkWhere 
            Caption         =   "Get Where Statment From Search ..."
            Height          =   255
            Left            =   240
            TabIndex        =   112
            Top             =   2040
            Width           =   255
         End
         Begin VB.CommandButton BntConvert 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   5520
            Picture         =   "FrmSQL.frx":171086
            Style           =   1  'Graphical
            TabIndex        =   110
            ToolTipText     =   "Convert To Excel"
            Top             =   240
            Width           =   1215
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00E0E0E0&
            Height          =   1455
            Left            =   3720
            TabIndex        =   107
            Top             =   120
            Width           =   1575
            Begin VB.OptionButton OptConvert 
               Caption         =   "All Fields"
               Height          =   255
               Index           =   0
               Left            =   240
               TabIndex        =   109
               Top             =   240
               Value           =   -1  'True
               Width           =   1095
            End
            Begin VB.OptionButton OptConvert 
               Caption         =   "Selective"
               Height          =   255
               Index           =   1
               Left            =   240
               TabIndex        =   108
               Top             =   600
               Width           =   1095
            End
         End
         Begin VB.TextBox TextSQL 
            Height          =   735
            Left            =   240
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   106
            Text            =   "FrmSQL.frx":171BD4
            Top             =   2520
            Width           =   5175
         End
         Begin VB.CommandButton BtnSelect 
            Caption         =   "Select"
            Enabled         =   0   'False
            Height          =   315
            Left            =   5520
            TabIndex        =   105
            Top             =   2520
            Width           =   1215
         End
         Begin InstantReport.Hline Hline1 
            Height          =   30
            Left            =   120
            TabIndex        =   102
            Top             =   1800
            Width           =   6615
            _ExtentX        =   11668
            _ExtentY        =   53
         End
         Begin VB.CommandButton BtnOK 
            BackColor       =   &H00E0E0E0&
            Height          =   315
            Left            =   5520
            Picture         =   "FrmSQL.frx":171BE4
            Style           =   1  'Graphical
            TabIndex        =   101
            Top             =   2880
            Width           =   1215
         End
         Begin VB.ListBox ListConvert 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   1380
            ItemData        =   "FrmSQL.frx":17258E
            Left            =   240
            List            =   "FrmSQL.frx":172595
            Style           =   1  'Checkbox
            TabIndex        =   100
            Top             =   240
            Width           =   3375
         End
         Begin VB.Label lblWhere 
            BackStyle       =   0  'Transparent
            Caption         =   "Get Where statement from search...."
            Height          =   255
            Left            =   720
            MouseIcon       =   "FrmSQL.frx":1725A6
            MousePointer    =   99  'Custom
            TabIndex        =   120
            Top             =   2040
            Width           =   4695
         End
      End
      Begin VB.PictureBox picCloseMe 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   300
         Left            =   6720
         MouseIcon       =   "FrmSQL.frx":172E70
         MousePointer    =   99  'Custom
         Picture         =   "FrmSQL.frx":17373A
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   98
         Top             =   0
         Width           =   300
      End
   End
   Begin VB.PictureBox PicSep 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   0
      ScaleHeight     =   135
      ScaleWidth      =   12975
      TabIndex        =   21
      Top             =   1320
      Width           =   12975
      Begin VB.PictureBox ButtonSep 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   120
         Left            =   6600
         MouseIcon       =   "FrmSQL.frx":173AC4
         MousePointer    =   99  'Custom
         Picture         =   "FrmSQL.frx":17438E
         ScaleHeight     =   120
         ScaleWidth      =   945
         TabIndex        =   22
         Top             =   0
         Width           =   945
      End
      Begin VB.Image Image2 
         Appearance      =   0  'Flat
         Height          =   120
         Left            =   120
         Picture         =   "FrmSQL.frx":174BB0
         Top             =   0
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   120
         Left            =   600
         Picture         =   "FrmSQL.frx":1753D2
         Top             =   0
         Visible         =   0   'False
         Width           =   945
      End
   End
   Begin MSComctlLib.ListView lvList 
      Height          =   1695
      Left            =   0
      TabIndex        =   1
      Top             =   5520
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   2990
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "i16x16"
      SmallIcons      =   "i16x16"
      ForeColor       =   0
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.PictureBox PicEntry 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3495
      Left            =   0
      ScaleHeight     =   3495
      ScaleWidth      =   14895
      TabIndex        =   23
      Top             =   1320
      Width           =   14895
      Begin VB.Frame FmeView 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   3255
         Left            =   120
         TabIndex        =   24
         Top             =   120
         Width           =   15660
         Begin VB.CommandButton CmdLast 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   275
            Left            =   14160
            MousePointer    =   99  'Custom
            Picture         =   "FrmSQL.frx":175BF4
            Style           =   1  'Graphical
            TabIndex        =   118
            ToolTipText     =   "Last"
            Top             =   2880
            Width           =   375
         End
         Begin VB.CommandButton CmdNext 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   275
            Left            =   13800
            MousePointer    =   99  'Custom
            Picture         =   "FrmSQL.frx":175EA9
            Style           =   1  'Graphical
            TabIndex        =   117
            ToolTipText     =   "Next"
            Top             =   2880
            Width           =   375
         End
         Begin VB.CommandButton CmdPrev 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   275
            Left            =   13440
            MousePointer    =   99  'Custom
            Picture         =   "FrmSQL.frx":17615E
            Style           =   1  'Graphical
            TabIndex        =   116
            ToolTipText     =   "Previous"
            Top             =   2880
            Width           =   375
         End
         Begin VB.CommandButton CmdFirst 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   13080
            MaskColor       =   &H00404040&
            MousePointer    =   99  'Custom
            Picture         =   "FrmSQL.frx":176413
            Style           =   1  'Graphical
            TabIndex        =   115
            ToolTipText     =   "First"
            Top             =   2880
            Width           =   375
         End
         Begin VB.TextBox txtEntry 
            Height          =   285
            Index           =   30
            Left            =   13320
            TabIndex        =   55
            Top             =   2400
            Width           =   2000
         End
         Begin VB.TextBox txtEntry 
            Height          =   285
            Index           =   29
            Left            =   13320
            TabIndex        =   54
            Top             =   2040
            Width           =   2000
         End
         Begin VB.TextBox txtEntry 
            Height          =   285
            Index           =   28
            Left            =   13320
            TabIndex        =   53
            Top             =   1680
            Width           =   2000
         End
         Begin VB.TextBox txtEntry 
            Height          =   285
            Index           =   27
            Left            =   13320
            TabIndex        =   52
            Top             =   1320
            Width           =   2000
         End
         Begin VB.TextBox txtEntry 
            Height          =   285
            Index           =   26
            Left            =   13320
            TabIndex        =   51
            Top             =   960
            Width           =   2000
         End
         Begin VB.TextBox txtEntry 
            Height          =   285
            Index           =   25
            Left            =   13320
            TabIndex        =   50
            Top             =   600
            Width           =   2000
         End
         Begin VB.TextBox txtEntry 
            Height          =   285
            Index           =   24
            Left            =   13320
            TabIndex        =   49
            Top             =   240
            Width           =   2000
         End
         Begin VB.TextBox txtEntry 
            Height          =   285
            Index           =   23
            Left            =   9480
            TabIndex        =   48
            Top             =   2760
            Width           =   2000
         End
         Begin VB.TextBox txtEntry 
            Height          =   285
            Index           =   22
            Left            =   9480
            TabIndex        =   47
            Top             =   2400
            Width           =   2000
         End
         Begin VB.TextBox txtEntry 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   21
            Left            =   9480
            TabIndex        =   46
            Top             =   2040
            Width           =   2000
         End
         Begin VB.TextBox txtEntry 
            Height          =   285
            Index           =   20
            Left            =   9480
            TabIndex        =   45
            Top             =   1680
            Width           =   2000
         End
         Begin VB.TextBox txtEntry 
            Height          =   285
            Index           =   19
            Left            =   9480
            TabIndex        =   44
            Top             =   1320
            Width           =   2000
         End
         Begin VB.TextBox txtEntry 
            Height          =   285
            Index           =   18
            Left            =   9480
            TabIndex        =   43
            Top             =   960
            Width           =   2000
         End
         Begin VB.TextBox txtEntry 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   17
            Left            =   9480
            TabIndex        =   42
            Top             =   600
            Width           =   2000
         End
         Begin VB.TextBox txtEntry 
            Height          =   285
            Index           =   1
            Left            =   1800
            TabIndex        =   41
            Top             =   600
            Width           =   2000
         End
         Begin VB.TextBox txtEntry 
            BackColor       =   &H00E8FBFB&
            Height          =   285
            Index           =   0
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   40
            Top             =   240
            Width           =   2000
         End
         Begin VB.TextBox txtEntry 
            Height          =   285
            Index           =   12
            Left            =   5640
            TabIndex        =   39
            Top             =   1680
            Width           =   2000
         End
         Begin VB.TextBox txtEntry 
            Height          =   285
            Index           =   11
            Left            =   5640
            TabIndex        =   38
            Top             =   1320
            Width           =   2000
         End
         Begin VB.TextBox txtEntry 
            Height          =   285
            Index           =   10
            Left            =   5640
            TabIndex        =   37
            Top             =   960
            Width           =   2000
         End
         Begin VB.TextBox txtEntry 
            Height          =   285
            Index           =   9
            Left            =   5640
            TabIndex        =   36
            Top             =   600
            Width           =   2000
         End
         Begin VB.TextBox txtEntry 
            Height          =   285
            Index           =   8
            Left            =   5640
            TabIndex        =   35
            Top             =   240
            Width           =   2000
         End
         Begin VB.TextBox txtEntry 
            Height          =   285
            Index           =   7
            Left            =   1800
            TabIndex        =   34
            Top             =   2760
            Width           =   2000
         End
         Begin VB.TextBox txtEntry 
            Height          =   285
            Index           =   6
            Left            =   1800
            TabIndex        =   33
            Top             =   2400
            Width           =   2000
         End
         Begin VB.TextBox txtEntry 
            Height          =   285
            Index           =   5
            Left            =   1800
            TabIndex        =   32
            Top             =   2040
            Width           =   2000
         End
         Begin VB.TextBox txtEntry 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   3
            Left            =   1800
            TabIndex        =   31
            Top             =   1320
            Width           =   2000
         End
         Begin VB.TextBox txtEntry 
            Height          =   285
            Index           =   2
            Left            =   1800
            TabIndex        =   30
            Top             =   960
            Width           =   2000
         End
         Begin VB.TextBox txtEntry 
            Height          =   285
            Index           =   13
            Left            =   5640
            TabIndex        =   29
            Top             =   2040
            Width           =   2000
         End
         Begin VB.TextBox txtEntry 
            Height          =   285
            Index           =   14
            Left            =   5640
            TabIndex        =   28
            Top             =   2400
            Width           =   2000
         End
         Begin VB.TextBox txtEntry 
            Height          =   285
            Index           =   15
            Left            =   5640
            TabIndex        =   27
            Top             =   2760
            Width           =   2000
         End
         Begin VB.TextBox txtEntry 
            Height          =   285
            Index           =   16
            Left            =   9480
            TabIndex        =   26
            Top             =   240
            Width           =   2000
         End
         Begin VB.TextBox txtEntry 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   4
            Left            =   1800
            TabIndex        =   25
            Top             =   1680
            Width           =   2000
         End
         Begin MSComCtl2.DTPicker dtpDate 
            Height          =   285
            Index           =   2
            Left            =   2040
            TabIndex        =   56
            Top             =   0
            Visible         =   0   'False
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   503
            _Version        =   393216
            CustomFormat    =   "MMM-dd-yyyy"
            Format          =   60096515
            CurrentDate     =   38207
         End
         Begin VB.Image imgHelp 
            Height          =   360
            Left            =   14640
            MouseIcon       =   "FrmSQL.frx":1766C8
            MousePointer    =   99  'Custom
            Picture         =   "FrmSQL.frx":176F92
            Top             =   2760
            Width           =   360
         End
         Begin VB.Label lblFLDi 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   30
            Left            =   11760
            TabIndex        =   87
            Top             =   2400
            Width           =   1575
         End
         Begin VB.Label lblFLDi 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   29
            Left            =   11760
            TabIndex        =   86
            Top             =   2040
            Width           =   1575
         End
         Begin VB.Label lblFLDi 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   28
            Left            =   11760
            TabIndex        =   85
            Top             =   1680
            Width           =   1575
         End
         Begin VB.Label lblFLDi 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   27
            Left            =   11760
            TabIndex        =   84
            Top             =   1320
            Width           =   1575
         End
         Begin VB.Label lblFLDi 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   26
            Left            =   11760
            TabIndex        =   83
            Top             =   960
            Width           =   1575
         End
         Begin VB.Label lblFLDi 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   25
            Left            =   11760
            TabIndex        =   82
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label lblFLDi 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   24
            Left            =   11760
            TabIndex        =   81
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label lblFLDi 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   23
            Left            =   7920
            TabIndex        =   80
            Top             =   2760
            Width           =   1575
         End
         Begin VB.Label lblFLDi 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   22
            Left            =   7920
            TabIndex        =   79
            Top             =   2400
            Width           =   1575
         End
         Begin VB.Label lblFLDi 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   21
            Left            =   7920
            TabIndex        =   78
            Top             =   2040
            Width           =   1575
         End
         Begin VB.Label lblFLDi 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   20
            Left            =   7920
            TabIndex        =   77
            Top             =   1680
            Width           =   1575
         End
         Begin VB.Label lblFLDi 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   19
            Left            =   7920
            TabIndex        =   76
            Top             =   1320
            Width           =   1575
         End
         Begin VB.Label lblFLDi 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   18
            Left            =   7920
            TabIndex        =   75
            Top             =   960
            Width           =   1575
         End
         Begin VB.Label lblFLDi 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   17
            Left            =   7920
            TabIndex        =   74
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label lblFLDi 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   16
            Left            =   7920
            TabIndex        =   73
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label lblFLDi 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   15
            Left            =   4080
            TabIndex        =   72
            Top             =   2760
            Width           =   1575
         End
         Begin VB.Label lblFLDi 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
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
            Height          =   255
            Index           =   14
            Left            =   4080
            TabIndex        =   71
            Top             =   2400
            Width           =   1575
         End
         Begin VB.Label lblFLDi 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
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
            Height          =   255
            Index           =   13
            Left            =   4080
            TabIndex        =   70
            Top             =   2040
            Width           =   1575
         End
         Begin VB.Label lblFLDi 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
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
            Height          =   255
            Index           =   12
            Left            =   4080
            TabIndex        =   69
            Top             =   1680
            Width           =   1575
         End
         Begin VB.Label lblFLDi 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
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
            Height          =   255
            Index           =   11
            Left            =   4080
            TabIndex        =   68
            Top             =   1320
            Width           =   1575
         End
         Begin VB.Label lblFLDi 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
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
            Height          =   255
            Index           =   10
            Left            =   4080
            TabIndex        =   67
            Top             =   960
            Width           =   1575
         End
         Begin VB.Label lblFLDi 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
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
            Height          =   255
            Index           =   9
            Left            =   4080
            TabIndex        =   66
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label lblFLDi 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
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
            Height          =   255
            Index           =   8
            Left            =   4080
            TabIndex        =   65
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label lblFLDi 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
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
            Height          =   255
            Index           =   7
            Left            =   240
            TabIndex        =   64
            Top             =   2760
            Width           =   1575
         End
         Begin VB.Label lblFLDi 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
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
            Height          =   255
            Index           =   6
            Left            =   240
            TabIndex        =   63
            Top             =   2400
            Width           =   1575
         End
         Begin VB.Label lblFLDi 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
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
            Height          =   255
            Index           =   5
            Left            =   240
            TabIndex        =   62
            Top             =   2040
            Width           =   1575
         End
         Begin VB.Label lblFLDi 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
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
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   61
            Top             =   1680
            Width           =   1575
         End
         Begin VB.Label lblFLDi 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
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
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   60
            Top             =   1320
            Width           =   1575
         End
         Begin VB.Label lblFLDi 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
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
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   59
            Top             =   960
            Width           =   1575
         End
         Begin VB.Label lblFLDi 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
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
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   58
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label lblFLDi 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
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
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   57
            Top             =   240
            Width           =   1575
         End
      End
   End
   Begin VB.PictureBox PicTopBar 
      BackColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      ScaleHeight     =   195
      ScaleWidth      =   555
      TabIndex        =   95
      Top             =   4920
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox PicTop 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   345
      ScaleWidth      =   15225
      TabIndex        =   88
      Top             =   960
      Width           =   15255
      Begin VB.ComboBox CboTables 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1200
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   89
         Top             =   10
         Width           =   2940
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   "Tables:"
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
         Height          =   255
         Left            =   120
         TabIndex        =   113
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton CmdBrowse 
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   13200
         MouseIcon       =   "FrmSQL.frx":1776FC
         MousePointer    =   99  'Custom
         Picture         =   "FrmSQL.frx":177FC6
         Style           =   1  'Graphical
         TabIndex        =   91
         ToolTipText     =   "Open Database"
         Top             =   10
         Width           =   1755
      End
      Begin VB.TextBox TxtPathDatabase 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   7680
         Locked          =   -1  'True
         TabIndex        =   92
         Top             =   10
         Width           =   5355
      End
      Begin VB.TextBox TextPassword 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5640
         TabIndex        =   90
         Top             =   30
         Width           =   1095
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DB Path:"
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
         Left            =   6840
         TabIndex        =   94
         Top             =   30
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DB Password:"
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
         Left            =   4320
         TabIndex        =   93
         Top             =   30
         Width           =   1200
      End
   End
   Begin VB.PictureBox picTB 
      BackColor       =   &H00808080&
      Height          =   900
      Left            =   0
      ScaleHeight     =   840
      ScaleWidth      =   13275
      TabIndex        =   2
      Top             =   0
      Width           =   13335
      Begin VB.PictureBox PicBtnInfo 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   720
         Left            =   9720
         Picture         =   "FrmSQL.frx":179294
         ScaleHeight     =   720
         ScaleWidth      =   720
         TabIndex        =   114
         Top             =   20
         Width           =   720
      End
      Begin VB.CommandButton cmdSep 
         Enabled         =   0   'False
         Height          =   855
         Left            =   9480
         TabIndex        =   20
         Top             =   0
         Width           =   3375
      End
      Begin VB.CommandButton CmdCalcu 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Ca&lculator"
         Height          =   855
         Left            =   7680
         Picture         =   "FrmSQL.frx":17AF5E
         Style           =   1  'Graphical
         TabIndex        =   111
         Top             =   0
         Width           =   855
      End
      Begin VB.CommandButton cmdPrintOpt 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Print"
         Height          =   855
         Left            =   8640
         Picture         =   "FrmSQL.frx":17B828
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         Width           =   855
      End
      Begin VB.CommandButton CmdConvert 
         BackColor       =   &H00E0E0E0&
         Caption         =   "E&xport"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   6840
         Picture         =   "FrmSQL.frx":17D4F2
         Style           =   1  'Graphical
         TabIndex        =   96
         ToolTipText     =   "Convert To Excel"
         Top             =   0
         Width           =   855
      End
      Begin VB.CommandButton CmdButton 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Refresh"
         Height          =   855
         Index           =   6
         Left            =   5040
         Picture         =   "FrmSQL.frx":17F1BC
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   0
         Width           =   855
      End
      Begin VB.CommandButton CmdButton 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Delete"
         Height          =   855
         Index           =   5
         Left            =   4200
         Picture         =   "FrmSQL.frx":180E86
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   0
         Width           =   855
      End
      Begin VB.CommandButton CmdButton 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Cancel"
         Height          =   855
         Index           =   4
         Left            =   3360
         Picture         =   "FrmSQL.frx":182B50
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   0
         Width           =   855
      End
      Begin VB.CommandButton cmdFind 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Searc&h"
         Height          =   855
         Left            =   6000
         Picture         =   "FrmSQL.frx":18481A
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   0
         Width           =   855
      End
      Begin VB.CommandButton CmdButton 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Update"
         Height          =   855
         Index           =   3
         Left            =   2520
         Picture         =   "FrmSQL.frx":1864E4
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   0
         Width           =   855
      End
      Begin VB.CommandButton CmdButton 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Edit"
         Height          =   855
         Index           =   2
         Left            =   1680
         Picture         =   "FrmSQL.frx":1873AE
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   0
         Width           =   855
      End
      Begin VB.CommandButton CmdButton 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Save"
         Height          =   855
         Index           =   1
         Left            =   840
         Picture         =   "FrmSQL.frx":189078
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         Width           =   855
      End
      Begin VB.CommandButton CmdButton 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Add"
         Height          =   855
         Index           =   0
         Left            =   0
         Picture         =   "FrmSQL.frx":18AD42
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   0
         Width           =   855
      End
   End
   Begin MSComctlLib.StatusBar SB1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   10335
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   21220
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "7/3/2008"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "12:31 AM"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList i16x16 
      Left            =   1200
      Top             =   4920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   22
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSQL.frx":18CA0C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSQL.frx":18D41E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSQL.frx":18D578
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSQL.frx":18D6D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSQL.frx":18D82C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSQL.frx":18DB26
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSQL.frx":18DEC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSQL.frx":18E25A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSQL.frx":18EC6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSQL.frx":18ECC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSQL.frx":18F05A
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSQL.frx":18F3F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSQL.frx":18F78E
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSQL.frx":18FB28
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSQL.frx":19053A
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSQL.frx":190F4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSQL.frx":19195E
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSQL.frx":192370
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSQL.frx":192D82
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSQL.frx":193794
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSQL.frx":1941A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSQL.frx":194742
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   600
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "FrmSQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Forces the module to declare the variable first
Private cnn As New Connection
Private rcd As New Recordset
Private CurrentTable As String
Private PathDatabase As String   'database pathname
Private sqlStatement As String   'store sql statement
Private conSTR As String         'connection string
Private strPROVIDER As String    'provider
Private m_WhereStatement As String   'handle Where statement from search
Private showForm As Boolean

Private Sub BntConvert_Click()
Dim ans As Integer
ans = MsgBox("Proceed?", vbYesNo + vbQuestion, "Export Data To Excel!")
  If ans = vbYes Then
     ConvertDB2Excel
  End If
End Sub


Private Sub BtnOK_Click()
On Error GoTo errMsg
Dim sqlSTR As String
'//INITIALIZE
If ChkWhere.Value = 1 Then
    If isFilter = True Then
        GetWhereStatement
    End If
Else
  m_WhereStatement = ""
End If

   If rcd.State = adStateOpen Then
       rcd.Close
   End If
   sqlSTR = TextSQL.text
   rcd.CursorLocation = adUseClient
   rcd.Open sqlSTR, cnn, adOpenStatic, adLockOptimistic
   Load_DATA
errMsg:
   errorMsg Err, Me.Name, "OK Click"
End Sub

Private Sub BtnOK_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ChkWhere.Enabled = (isFilter = True)
End Sub

Private Sub BtnSelect_Click()
If ListConvert.ListIndex < 0 Then Exit Sub
Dim listEmpty As Boolean
Dim sqlSTR As String
Dim i As Integer
'//initialize
If ChkWhere.Value = 1 Then
    If isFilter = True Then
        GetWhereStatement
    End If
Else
  m_WhereStatement = ""
End If

TextSQL.text = ""
listEmpty = True
i = 0
   For i = 0 To ListConvert.ListCount - 1
          If ListConvert.Selected(i) = True Then
             listEmpty = False
              If Len(TextSQL.text) <> 0 Then
                   TextSQL.text = TextSQL.text & "," & "[" & ListConvert.List(i) & "]": Rem edwin delos santos
               Else: Rem textsql len = 0 ; so no comma
                   TextSQL.text = "[" & ListConvert.List(i) & "]"
              End If
           End If
         Next i
    sqlSTR = "SELECT " & TextSQL.text & " FROM " & "[" & CboTables.text & "]"
    sqlSTR = sqlSTR & m_WhereStatement
   If listEmpty = True Then
      sqlSTR = "SELECT * FROM [" & CboTables.text & "]"
      sqlSTR = sqlSTR & m_WhereStatement
   End If
    TextSQL.text = sqlSTR
End Sub

Private Sub BtnSelect_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ChkWhere.Enabled = (isFilter = True)
End Sub

Private Sub ButtonSep_Click()
If Len(CboTables.text) = 0 Then Exit Sub
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
 
   PicSep.Top = PicTopBar.Top - PicSep.Height: Rem edwin delos santos
   lvList.Top = PicTopBar.Top: Rem edwin delos santos
   lvList.Height = Me.ScaleHeight - (picTB.Height + PicTop.Height + PicEntry.Height + PicSep.Height + SB1.Height + 50): Rem edwin delos santos
   SB1.Top = Me.ScaleHeight - SB1.Height: Rem edwin delos santos
   Set ButtonSep.Picture = img.Picture: Rem edwin delos santos
   showForm = True: Rem edwin delos santos
 Else
   PicSep.Top = PicTop.Height + picTB.Height + 50: Rem edwin delos santos
   lvList.Top = PicEntry.Top + PicSep.Height: Rem edwin delos santos
   lvList.Height = Me.ScaleHeight - (picTB.Height + PicTop.Height + PicSep.Height + SB1.Height + 50): Rem edwin delos santos
   Set ButtonSep.Picture = img.Picture: Rem edwin delos santos
   showForm = False
 End If
End Sub


Private Sub Check1_Click()
 CboTables.Locked = (Check1.Value = 0)
End Sub



Private Sub CmdCalcu_Click()
 Load FrmCalcu
 FrmCalcu.show
End Sub

Private Sub CmdConvert_Click()
cmdButtonShow ("0000001"), Me
ChkWhere.Enabled = (isFilter = True)
  picConvert.Visible = True
  CenterObjt picConvert
  picConvert.ZOrder
End Sub

Private Sub ConvertDB2Excel()
On Error GoTo errMsg
Dim i As Integer, r As Integer
Dim objExcl As Excel.Application

r = 1
Set objExcl = New Excel.Application
rcd.MoveFirst
objExcl.Visible = True
objExcl.SheetsInNewWorkbook = 1
objExcl.Workbooks.Add
For i = 0 To rcd.Fields.Count - 1
    objExcl.ActiveSheet.Cells(r, i + 1).Value = rcd.Fields(i).Name
Next
r = 2      'write start at row 2
Do Until rcd.EOF
For i = 0 To rcd.Fields.Count - 1
     objExcl.ActiveSheet.Cells(r, i + 1).Value = rcd.Fields(i)
Next
rcd.MoveNext
r = r + 1
Loop

errMsg:
 errorMsg Err, Me.Name, "Convert to Excel"
End Sub


Private Sub GetWhereStatement()
Dim wStatement As String    'WhereStatement
Dim m_where As Integer
On Error GoTo errMsg
    With frmSearch
        If OptConvert(0).Value = True Then
         TextSQL.text = .TextSQL.text
        Else
         m_WhereStatement = .TextSQL.text   '//wherestatement (Private)
        End If
   End With
If UCase$(Mid(m_WhereStatement, 1, 6)) <> "SELECT" Then Exit Sub
m_where = InStr(1, m_WhereStatement, "WHERE", 0)
wStatement = Mid(m_WhereStatement, m_where)
m_WhereStatement = wStatement
errMsg:
   errorMsg Err, Me.Name, "Ger Where Statement"
End Sub












Private Sub dtpDate_CloseUp(Index As Integer)
   txtEntry(nxTab).text = Format(dtpDate(2).Value, "mmm-dd-yyyy")
   txtEntry(nxTab).SetFocus
 End Sub






Private Sub imgHelp_Click()
Dim msg As String
msg = "<< Expand Textbox >>" & vbCrLf
msg = msg & "Your may expand textbox width on focus to view long string!"
msg = msg & vbCrLf & "F3 - to extend"
msg = msg & vbCrLf & "F4 - to restore width:"
myMsg msg, "Expand Textbox", 2, True
End Sub



Private Sub lblWhere_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ChkWhere.Enabled = (isFilter = True)
End Sub

Private Sub lvList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   SortListView lvList, ColumnHeader
End Sub

Private Sub PicBtnInfo_Click()
 frmAbout.show  '1 if modal
End Sub

Private Sub TextSQL_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  TextSQL.ToolTipText = TextSQL.text
End Sub

Private Sub OptConvert_Click(Index As Integer)
Dim sqlSTR As String
sqlSTR = "SELECT * FROM [" & CboTables.text & "]"
TextSQL = sqlSTR
BtnSelect.Enabled = (OptConvert(1).Value = True)
ListConvert.Enabled = (OptConvert(1).Value = True)
End Sub

Private Sub picCloseMe_Click()
picConvert.Visible = False
End Sub

Private Sub picConvert_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        Call DragIt(picConvert.hWnd)
    End If
End Sub

Private Sub PicSep_Resize()
 center_obj_horizontal PicSep, ButtonSep
End Sub

Private Sub CboTables_Click()
    On Error GoTo ERRORHANDLE
    Dim sqlSTR As String
    Dim lblStr As String

    Set rcd = New ADODB.Recordset
    rcd.CursorLocation = adUseClient
    sqlSTR = "SELECT * FROM [" & CboTables.text & "]"
    rcd.Open sqlSTR, cnn, adOpenStatic, adLockOptimistic
    Call Insert_Fields(List2Print, rcd)
    Call Insert_Fields(ListConvert, rcd)
    Load_DATA
    
   TextBox_Visible Me, rcd
   ShowFldsLabel Me, rcd
'   If addRec = False Or editRec = False Then
'     cmdButtonShow ("1010011"), Me
'   End If
Check1.Value = 0
ERRORHANDLE:
    errorMsg Err, Me.Name, "CboTables_Click()"
End Sub





Private Sub CmdBrowse_Click()
    Dim strPass As String
    On Error GoTo ERRORHANDLE
    strPass = TextPassword.text
    CD1.Filter = "Access Database (*.mdb)|*.MDB"
    CD1.DialogTitle = "Open Access Database"
    ' Exit if user presses Cancel.
    CD1.CancelError = True
    CD1.FileName = ""
    CD1.ShowOpen
     
    PathDatabase = CD1.FileName
       

      strPROVIDER = "Provider=Microsoft.Jet.OLEDB.4.0"
      
      conSTR = strPROVIDER & ";Persist Security Info=false "
      conSTR = conSTR & ";Data Source="
      conSTR = conSTR & PathDatabase
      conSTR = conSTR & ";Jet OLEDB:Database Password=" & strPass
       
     If cnn.State = adStateOpen Then cnn.Close
     cnn.CursorLocation = adUseClient
     cnn.Mode = adModeReadWrite
     cnn.Open conSTR
    
    SB1.Panels(1).text = "Status : Connected"
    
    TxtPathDatabase.text = CD1.FileName

ERRORHANDLE:
  errorMsg Err, Me.Name, "Open Database"
End Sub







Private Sub CmdMove_Click()
' Move one item from left to right.
    If List2Print.ListIndex >= 0 Then
        ListPrint.AddItem List2Print.text
        List2Print.RemoveItem List2Print.ListIndex
        initPrint = False
    End If
End Sub

Private Sub CmdMoveBack_Click()
' Move one item from left to right.
Dim idx As Integer
idx = ListPrint.ListIndex
    If ListPrint.ListIndex >= 0 Then
        If ListPrint.Selected(idx) = True Then Exit Sub
        List2Print.AddItem ListPrint.text
        ListPrint.RemoveItem ListPrint.ListIndex
        initPrint = False
    End If
End Sub



Private Sub cmdOK_Click()
  '//validate
  initPrint = print_Init(ListPrint)
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
   Exit Function
End Function






Private Sub CmdFind_Click()
If Len(CboTables.text) = 0 Then Exit Sub
    With frmSearch
            Set .pFindForm = Me
            Set .pFindRecset = rcd
            Set .pFindCon = cnn
                .pFindTABLE = CboTables.text
                .show
   End With
End Sub


Private Sub cmdButton_Click(Index As Integer)
If Len(CboTables.text) = 0 Then Exit Sub
'//                  A S E U C D R
On Error GoTo ERRORHANDLE
Select Case Index
   Case BtnAdd                       '<------ add new record ------->'
'     If Not IsNumeric(lvList.SelectedItem.text) Then Exit Sub
     cmdButtonShow ("0100100"), Me
     If isFilter = True Then
        MsgBox "Data is Filtered", vbCritical, "Refresh Record First!"
        Exit Sub
     End If
     Dim NextNo As Long
     '//initialize//
     txtEntry(29).text = Format(Now(), "Short Date")
     txtEntry(30).text = CurrUser.user_id
     '//assign next number//
     NextNo = Last_Recc(rcd)
     If NextNo > 0 Then
       txtEntry(0).text = NextNo
       txtEntry(1).SetFocus
     Else
       txtEntry(0).Locked = False
       txtEntry(0).SetFocus
    End If
     txtEntry(1).SetFocus
   Case BtnSave                       '<------ save new record ------>'
        cmdButtonShow ("1010011"), Me
        Call WriteData(Me, rcd, True)
        Call lvwPopulateData(lvList, rcd, 2)
   Case BtnEdit                       '<------ edit record ---------->'
'        If Not IsNumeric(lvList.SelectedItem.text) Then Exit Sub
        cmdButtonShow ("0001100"), Me
        txtEntry(1).SetFocus
   Case BtnUpdate                     '<------ update record -------->'
        cmdButtonShow ("1010011"), Me
        Call WriteData(Me, rcd, False)
        LvwReplaceData Me, rcd, lvList
   Case BtnCancel                     '<------ cancel update -------->'
        cmdButtonShow ("1010011"), Me
   Case BtnDelete                     '<------ delete record -------->'
        'If Not IsNumeric(lvList.SelectedItem.text) Then Exit Sub
        Call Delete_Record(rcd, lvList)
   Case BtnRefresh                    '<------ Refresh record ------->'
       CboTables_Click
       isFilter = False
       cmdButtonShow ("1010011"), Me
       Unload frmSearch
      lvList.SetFocus
End Select
ERRORHANDLE:
 errorMsg Err, Me.Name, "Command Button"
End Sub

Private Sub CmdPrintOpt_Click()
If Len(CboTables.text) = 0 Then Exit Sub
 picPrint.Visible = True
 CenterObjt picPrint
 picPrint.ZOrder
End Sub

Private Sub Form_Activate()
' MainForm.PicClose.Enabled = False
End Sub

Private Sub Form_Load()
isFilter = False
showForm = False
show
lvList.SetFocus
cmdButtonShow ("1010001"), Me
'// List BackColour Formatting
Call SetListViewColor(lvList, PicLv, vbWhite, &HF7EBD0, 0.1)
dtpDate(2).Value = Format(Now(), "mmm-dd-yyyy")
SB1.Panels(1).text = "Status : Disconnected Database"
End Sub


Private Sub Load_DATA()
On Error GoTo ERRORHANDLE
'// set columnheaders
 Call InsertColumn(lvList, rcd)
'//set details
 Call FillListView(lvList, rcd, 3)
 Call Listview_Total(lvList, rcd)
ERRORHANDLE:
    errorMsg Err, Me.Name
End Sub

Private Sub Enabled_TBox(ByVal rs As Recordset)
  Dim i As Integer
  Dim numba As Integer
  On Error Resume Next
      For i = 0 To txtEntry.UBound
               txtEntry(i).Visible = False
               lblFLDi(i).Visible = False
          Next i
  i = 0
  numba = (rs.Fields.Count - 1)
      For i = 0 To numba
               If i = 0 Then
                 txtEntry(i).Locked = True
                 txtEntry(i).BackColor = vbCyan
               End If
               txtEntry(i).Visible = True
               lblFLDi(i).Visible = True
          Next i
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim msg As String
Dim ans As Integer

If addRec = True Or editRec = True Then
  MsgBox "Pending operation required action", vbCritical, "<Save><Update><Cancel>"
  Cancel = True
Else
  If MsgBox("This will close the application.Do you want to proceed?", vbYesNo + vbQuestion, "Instant Report") = vbNo Then
         Cancel = True
     Else
         Unload FrmSQL
 End If
End If
End Sub

Private Sub Form_Resize()
   On Error Resume Next
    If WindowState <> vbMinimized Then
        If Me.Width < 9195 Then Me.Width = 9195
        If Me.Height < 4500 Then Me.Height = 4500
        
        cmdSep.Width = Me.ScaleWidth
        picTB.Width = Me.ScaleWidth
        PicTop.Width = Me.ScaleWidth
        PicEntry.Width = Me.ScaleWidth
        lvList.Left = 0: lvList.Width = Me.ScaleWidth
        PicSep.Width = Me.ScaleWidth
        
       If showForm = True Then
          PicSep.Top = PicTopBar.Top - PicSep.Height
          lvList.Top = PicTopBar.Top
          lvList.Height = Me.ScaleHeight - (picTB.Height + PicTop.Height + PicEntry.Height + PicSep.Height + SB1.Height + 50)
          SB1.Top = Me.ScaleHeight - SB1.Height
        Else
         '//showform = false
          PicSep.Top = PicTop.Height + picTB.Height + 50
          lvList.Top = PicEntry.Top + PicSep.Height
          lvList.Height = Me.ScaleHeight - (picTB.Height + PicTop.Height + PicSep.Height + SB1.Height + 50)
        End If
        
   End If
  
'  With Me
'    picTB.Width = .ScaleWidth
'    PicTop.Width = .ScaleWidth
'    lvList.Width = .ScaleWidth
'    PicEntry.Width = .ScaleWidth
'  End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ERRORHANDLE
 
 
 Set rcd = Nothing
 Set cnn = Nothing
 If rcd Is Nothing Then
        rcd.Close
        cnn.Close
 End If
 '*MainForm.PicClose.Enabled = True
ERRORHANDLE:
  errorMsg Err, Me.Name
End Sub
Private Sub List2Print_DblClick()
  CmdMove_Click
End Sub


Private Sub ListPrint_DblClick()
CmdMoveBack_Click
End Sub
Private Sub lvList_Click()
On Error GoTo ERRORHANDLE
If addRec = True Or editRec = True Then Exit Sub
If lvList.ListItems.Count = 0 Then Exit Sub
If showForm = True Then
   Call BindDatasource(Me, rcd, lvList, True)
End If
ERRORHANDLE:
errorMsg Err, Me.Name
End Sub
Private Sub lvList_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo ERRORHANDLE
If addRec = True Or editRec = True Then Exit Sub
Select Case KeyCode
 Case Is = 38, 40 'down,up arrow key
    If lvList.ListItems.Count = 0 Then Exit Sub
    If showForm = True Then
        Call BindDatasource(Me, rcd, lvList, True)
    End If
End Select
ERRORHANDLE:
errorMsg Err, Me.Name
End Sub


Private Sub PicClose_Click()
   picPrint.Visible = False
End Sub

Private Sub PicPRINT_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        Call DragIt(picPrint.hWnd)
    End If
End Sub

Private Sub DragIt(ByVal lngHwnd As Long)
Dim lngReturn As Long
    lngReturn = ReleaseCapture()
    lngReturn = SendMessage(lngHwnd, WM_NCLBUTTONDOWN, HTCAPTION, CLng(0))
End Sub





Private Sub PicTop_Resize()
 CmdBrowse.Left = PicTop.Width - (CmdBrowse.Width + 100)
 CmdBrowse.Top = 40

End Sub

Private Sub txtEntry_Change(Index As Integer)
' If Index = 0 Then
'   Call ListView_Search(lvList, txtEntry(0).text)
' End If
End Sub

Private Sub txtEntry_GotFocus(Index As Integer)
Dim idx As Integer
On Error GoTo ERRORHANDLE
idx = Index
nxTab = idx
txtEntry(idx).SelStart = 0
txtEntry(idx).SelLength = Len(txtEntry(idx).text)
If IsDate(txtEntry(idx).text) Then
 If Len(txtEntry(idx).text) > 8 Then
    AlignObj txtEntry(idx), dtpDate(2), 2
 End If
 txtEntry(idx).SetFocus
End If
ERRORHANDLE:
  errorMsg Err, Me.Name
End Sub



Private Sub txtEntry_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim lastTab As Integer
On Error GoTo ERRORHANDLE
lastTab = (rcd.Fields.Count - 1) 'txtEntry.UBound
If KeyCode = 13 Then
 If nxTab = lastTab Then Exit Sub
     nxTab = nxTab + 1
ElseIf KeyCode = 38 Then  'up arrow key
 If nxTab = 0 Or nxTab = 1 Then Exit Sub
     nxTab = nxTab - 1
End If
If KeyCode = 114 Then   'f3
    TextExtend txtEntry(Index)
ElseIf KeyCode = 115 Then 'f4
  If Expand = True Then
     txtEntry(Index).Width = TextWd
     txtEntry(Index).BackColor = vbWhite
  End If
 End If

txtEntry(nxTab).SetFocus
ERRORHANDLE:
 errorMsg Err, Me.Name
End Sub

Private Sub TxtPathDatabase_Change()
On Error GoTo ERRORHANDLE
    CboTables.Clear
    If PathDatabase <> "" Then
        Call Load_Table(PathDatabase)
    Else
        CboTables.AddItem "- no selected table -"
    End If
ERRORHANDLE:
   errorMsg Err, Me.Name
End Sub

Private Sub Load_Table(ByVal pathname As String)
Dim cat As New Catalog
Dim tbl As Table
    'bind connection
    Set cat.ActiveConnection = cnn
    For Each tbl In cat.Tables
        If LCase(Left(tbl.Name, 4)) <> "msys" Then
           CboTables.AddItem tbl.Name
        End If
    Next tbl
End Sub

Private Sub CmdPrint_Click()
     If initPrint = False Then
       myMsg "Please Validate First", "Instant Report - Print", 1, True
       Exit Sub
     End If
     PrintReport rcd
End Sub

'// PRINTER REPORT PROCEDURE
'//CODED BY EDWIN DELOS SANTOS
Private Sub Headers()
 Dim dat
 Dim co As String 'company
 Dim paydat As String
 co = "EDWIN SOFTWARE "  'your company
 dat = Format(Now(), "long date")
 Printer.Print Tab(6); dat
 Call prnCenterText(co, 180)
 'Call prnCENTERTEXT("My Report", 180)  'type of report
End Sub
Private Sub PrintReport(ByRef srcRS As Recordset)
'// coded by edwin delos santos
'// fixed settings has been made to this report
'// you may adjust settings using variables ...
'// like default:  tab, page orientation, number of pages to print, quality and so on...
'// no modify except for the above said defaul settings ...
On Error GoTo printErr
Dim curr_Rec As Long  'current record looping becomes 0 if max line per page is reaced
Dim rec_Counter As Long 'record  counter
Dim ans 'answer
Dim strFont As String, sngSize As Single
'//initialize
rec_Counter = 0
curr_Rec = 0
ans = MsgBox("Proceed?", vbYesNo + vbQuestion, "Print Report")
  If ans = vbYes Then
'//save current printer settings
         strFont = Printer.Font
         sngSize = Printer.FontSize
         Printer.Orientation = 2   'Landscape
         Printer.Font = "ms sans serif"
         Printer.FontSize = 9
         Printer.Print
'// headers
         Headers
         Printer.Print
         Printer.Print
       With srcRS
            .MoveFirst
              Printer.Font.Underline = True
              Call Print_Headings(srcRS, 6, ListPrint, 12)
              Printer.Font.Underline = False
         While Not .EOF = True
'//print details
               Call Print_Details(srcRS, 6, ListPrint)  'see procedure
               Printer.Print                            'line space
'//if not eof = true
                .MoveNext
                rec_Counter = rec_Counter + 1          'counter
                curr_Rec = curr_Rec + 1                 'store record printed
             If curr_Rec = 25 Then                      'proceed to next page
                Printer.Print Tab(6); ">>>------- next page --->"
                curr_Rec = 0                           'reset current record to 0
'//next page
               If rec_Counter < .RecordCount Then
                   Printer.NewPage
                   Printer.Print
                   Printer.Print
                   Headers   'next page headers
                   Printer.Print
                   Printer.Font.Underline = True
                   Call Print_Headings(srcRS, 6, ListPrint, 12) '2nd page headings
                   Printer.Font.Underline = False
                   Printer.Print
                End If 'rec_count
            End If   'curr_Rec
        Wend  'while not EOF
'//print line
         Printer.Print Tab(6); "___________________________________________________________________________________________________________________________________________________"
'// print Total
          Printer.Print Tab(6); "T O T A L >>";
          Call Print_Total(srcRS, 6, ListPrint)
          Printer.Print
'//Footer
          Footers
      End With
'//send information to the printer
         Printer.EndDoc
'//reset printer setting
         Printer.Font = strFont
         Printer.FontSize = sngSize
         If srcRS.EOF = True Then   'end of file
              MsgBox "D O N E !!", vbInformation, "Printing..."
         End If
   Else
       Exit Sub
   End If  'answer
printErr:
  errorMsg Err, Me.Name

End Sub

Private Sub Footers()
     Printer.FontSize = 5
     Printer.Print Tab(6); "system created by: edwin delos santos" 'userNAME
     Printer.Print Tab(6); "contact us: 0920-6747545 or cyber_edu2005@yahoo.com"
     Printer.Print
End Sub

Private Sub CmdFirst_Click()
rcd.MoveFirst
 Call BindDatasource(Me, rcd, lvList, False)
End Sub

Private Sub CmdLast_Click()
rcd.MoveLast
 Call BindDatasource(Me, rcd, lvList, False)
End Sub

Private Sub CmdNext_Click()
If rcd.EOF = True Then
 Exit Sub
Else
 rcd.MoveNext
Call BindDatasource(Me, rcd, lvList, False)
End If

End Sub

Private Sub CmdPrev_Click()
If rcd.BOF = True Then
 Exit Sub
Else
 rcd.MovePrevious
Call BindDatasource(Me, rcd, lvList, False)
End If

End Sub

