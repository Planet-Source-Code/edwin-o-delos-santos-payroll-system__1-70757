VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmPayroll 
   Caption         =   "Payroll System ver.1.2+"
   ClientHeight    =   8820
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12360
   Icon            =   "FrmPayroll.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "FrmPayroll.frx":109A
   ScaleHeight     =   8820
   ScaleWidth      =   12360
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   3975
      Left            =   1800
      Picture         =   "FrmPayroll.frx":2565
      ScaleHeight     =   3945
      ScaleWidth      =   6945
      TabIndex        =   1
      Top             =   3240
      Visible         =   0   'False
      Width           =   6975
      Begin VB.PictureBox picRestore 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Enabled         =   0   'False
         Height          =   300
         Left            =   4800
         Picture         =   "FrmPayroll.frx":B7B21
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   4
         Top             =   0
         Width           =   300
      End
      Begin VB.PictureBox PicClose 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   300
         Left            =   5160
         MouseIcon       =   "FrmPayroll.frx":B7EAB
         MousePointer    =   99  'Custom
         Picture         =   "FrmPayroll.frx":B8775
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   3
         Top             =   0
         Width           =   300
      End
      Begin VB.PictureBox picMinimize 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Enabled         =   0   'False
         Height          =   300
         Left            =   4440
         Picture         =   "FrmPayroll.frx":B8AFF
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   2
         Top             =   0
         Width           =   300
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   3375
         Left            =   120
         TabIndex        =   59
         Top             =   480
         Width           =   6660
         _ExtentX        =   11748
         _ExtentY        =   5953
         _Version        =   393216
         Tabs            =   4
         Tab             =   3
         TabsPerRow      =   5
         TabHeight       =   794
         BackColor       =   14737632
         TabCaption(0)   =   "Data Fields "
         TabPicture(0)   =   "FrmPayroll.frx":B8E89
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "CmdOk"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "ListPrint"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "List2Print"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "CmdMove"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "CmdMoveBack"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "indexList"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "CmdClear"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).ControlCount=   7
         TabCaption(1)   =   "Headers / Footers"
         TabPicture(1)   =   "FrmPayroll.frx":B8EA5
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "chkLongDate"
         Tab(1).Control(1)=   "ctrlLiner1"
         Tab(1).Control(2)=   "Picture3"
         Tab(1).Control(3)=   "Picture4"
         Tab(1).ControlCount=   4
         TabCaption(2)   =   "Load Data"
         TabPicture(2)   =   "FrmPayroll.frx":B8EC1
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Frame1"
         Tab(2).Control(1)=   "lblSQL"
         Tab(2).ControlCount=   2
         TabCaption(3)   =   "Print"
         TabPicture(3)   =   "FrmPayroll.frx":B8EDD
         Tab(3).ControlEnabled=   -1  'True
         Tab(3).Control(0)=   "lblnetpay"
         Tab(3).Control(0).Enabled=   0   'False
         Tab(3).Control(1)=   "lblAmtInWord"
         Tab(3).Control(1).Enabled=   0   'False
         Tab(3).Control(2)=   "Label10"
         Tab(3).Control(2).Enabled=   0   'False
         Tab(3).Control(3)=   "Label7"
         Tab(3).Control(3).Enabled=   0   'False
         Tab(3).Control(4)=   "BtnTotal"
         Tab(3).Control(4).Enabled=   0   'False
         Tab(3).Control(5)=   "Hline2"
         Tab(3).Control(5).Enabled=   0   'False
         Tab(3).Control(6)=   "txtCutOff(1)"
         Tab(3).Control(6).Enabled=   0   'False
         Tab(3).Control(7)=   "txtCutOff(0)"
         Tab(3).Control(7).Enabled=   0   'False
         Tab(3).Control(8)=   "Frame3"
         Tab(3).Control(8).Enabled=   0   'False
         Tab(3).Control(9)=   "CmdShow"
         Tab(3).Control(9).Enabled=   0   'False
         Tab(3).Control(10)=   "chkBankLetter"
         Tab(3).Control(10).Enabled=   0   'False
         Tab(3).ControlCount=   11
         Begin VB.CheckBox chkBankLetter 
            Caption         =   "Bank Letter "
            Height          =   255
            Left            =   3960
            TabIndex        =   106
            Top             =   2520
            Width           =   1215
         End
         Begin VB.CommandButton CmdShow 
            Caption         =   "Show "
            Enabled         =   0   'False
            Height          =   315
            Left            =   5280
            TabIndex        =   105
            Top             =   2520
            Width           =   1095
         End
         Begin VB.CommandButton CmdClear 
            Caption         =   "Clear >"
            Height          =   315
            Left            =   -72120
            TabIndex        =   102
            Top             =   1680
            Width           =   855
         End
         Begin VB.ListBox indexList 
            Appearance      =   0  'Flat
            BackColor       =   &H00008000&
            ForeColor       =   &H00FFFFFF&
            Height          =   420
            ItemData        =   "FrmPayroll.frx":B8EF9
            Left            =   -74880
            List            =   "FrmPayroll.frx":B8F1B
            TabIndex        =   101
            Top             =   2640
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Frame Frame3 
            Caption         =   "Option:"
            Height          =   1215
            Left            =   2040
            TabIndex        =   96
            Top             =   600
            Width           =   4335
            Begin VB.CommandButton BtnPrint 
               BackColor       =   &H00E0E0E0&
               Height          =   555
               Left            =   3360
               MouseIcon       =   "FrmPayroll.frx":B8F40
               MousePointer    =   99  'Custom
               Picture         =   "FrmPayroll.frx":B980A
               Style           =   1  'Graphical
               TabIndex        =   99
               ToolTipText     =   "Print"
               Top             =   240
               Width           =   735
            End
            Begin VB.OptionButton OptPayMode 
               Caption         =   "PaySlip"
               Height          =   375
               Left            =   360
               TabIndex        =   98
               Top             =   240
               Width           =   1575
            End
            Begin VB.OptionButton OptPaySummary 
               Caption         =   "Payroll Summary"
               Height          =   375
               Left            =   360
               TabIndex        =   97
               Top             =   600
               Width           =   1575
            End
         End
         Begin VB.TextBox txtCutOff 
            Height          =   285
            Index           =   0
            Left            =   240
            TabIndex        =   94
            Top             =   1080
            Width           =   1455
         End
         Begin VB.TextBox txtCutOff 
            Height          =   285
            Index           =   1
            Left            =   240
            TabIndex        =   93
            Top             =   1440
            Width           =   1455
         End
         Begin InstantReport.Hline Hline2 
            Height          =   30
            Left            =   240
            TabIndex        =   90
            Top             =   2040
            Width           =   6255
            _ExtentX        =   11033
            _ExtentY        =   53
         End
         Begin VB.CommandButton BtnTotal 
            Caption         =   "Total"
            Height          =   315
            Left            =   5280
            TabIndex        =   87
            Top             =   2880
            Width           =   1095
         End
         Begin VB.Frame Frame1 
            Caption         =   "Coverage:"
            ForeColor       =   &H00000000&
            Height          =   2055
            Left            =   -74760
            TabIndex        =   82
            Top             =   720
            Width           =   6135
            Begin VB.CheckBox chkALL 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Caption         =   "ALL"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   2040
               TabIndex        =   151
               Top             =   1080
               Width           =   855
            End
            Begin VB.ListBox lstPayDate 
               Appearance      =   0  'Flat
               Columns         =   2
               Enabled         =   0   'False
               ForeColor       =   &H00C00000&
               Height          =   1200
               ItemData        =   "FrmPayroll.frx":BA7AC
               Left            =   3000
               List            =   "FrmPayroll.frx":BA7AE
               Sorted          =   -1  'True
               TabIndex        =   92
               Top             =   600
               Width           =   2895
            End
            Begin VB.CheckBox chkAND 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Caption         =   "AND"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   2040
               TabIndex        =   100
               Top             =   720
               Width           =   855
            End
            Begin VB.CommandButton CmdFilter 
               BackColor       =   &H00E4DBC2&
               Caption         =   "Filter"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   240
               Style           =   1  'Graphical
               TabIndex        =   86
               Top             =   1560
               Width           =   1575
            End
            Begin VB.ListBox lstPayMode 
               Appearance      =   0  'Flat
               ForeColor       =   &H00C00000&
               Height          =   810
               ItemData        =   "FrmPayroll.frx":BA7B0
               Left            =   240
               List            =   "FrmPayroll.frx":BA7C0
               TabIndex        =   83
               Top             =   600
               Width           =   1575
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "Pay Date:"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   3000
               TabIndex        =   85
               Top             =   240
               Width           =   705
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Pay Mode:"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   240
               TabIndex        =   84
               Top             =   240
               Width           =   765
            End
         End
         Begin VB.CheckBox chkLongDate 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Caption         =   "Long Date"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   -69720
            TabIndex        =   73
            Top             =   1560
            Width           =   1095
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
            Left            =   -72120
            TabIndex        =   65
            Top             =   1200
            Width           =   855
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
            Left            =   -72120
            TabIndex        =   64
            Top             =   840
            Width           =   855
         End
         Begin VB.ListBox List2Print 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1950
            ItemData        =   "FrmPayroll.frx":BA7EA
            Left            =   -74880
            List            =   "FrmPayroll.frx":BA7F1
            TabIndex        =   63
            Top             =   600
            Width           =   2655
         End
         Begin VB.ListBox ListPrint 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1920
            ItemData        =   "FrmPayroll.frx":BA804
            Left            =   -71160
            List            =   "FrmPayroll.frx":BA806
            Style           =   1  'Checkbox
            TabIndex        =   62
            Top             =   600
            Width           =   2655
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
            Left            =   -69840
            Picture         =   "FrmPayroll.frx":BA808
            Style           =   1  'Graphical
            TabIndex        =   61
            ToolTipText     =   "Validate List To Print"
            Top             =   2640
            Width           =   1335
         End
         Begin InstantReport.Hline ctrlLiner1 
            Height          =   30
            Left            =   -74760
            TabIndex        =   60
            Top             =   1920
            Width           =   6135
            _ExtentX        =   10821
            _ExtentY        =   53
         End
         Begin VB.PictureBox Picture3 
            BackColor       =   &H00FFFFFF&
            Height          =   1215
            Left            =   -74760
            ScaleHeight     =   1155
            ScaleWidth      =   4875
            TabIndex        =   66
            Top             =   600
            Width           =   4935
            Begin VB.TextBox txtHeader 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   285
               Index           =   2
               Left            =   0
               TabIndex        =   69
               Text            =   "PAY DATE"
               Top             =   840
               Width           =   3135
            End
            Begin VB.TextBox txtHeader 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   285
               Index           =   1
               Left            =   0
               TabIndex        =   68
               Text            =   "PAYROLL SUMMARY"
               Top             =   480
               Width           =   4935
            End
            Begin VB.TextBox txtHeader 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   285
               Index           =   0
               Left            =   0
               TabIndex        =   67
               Text            =   "EDWIN SOFTWARE"
               Top             =   120
               Width           =   4935
            End
         End
         Begin VB.PictureBox Picture4 
            BackColor       =   &H00FFFFFF&
            Height          =   855
            Left            =   -74760
            ScaleHeight     =   795
            ScaleWidth      =   4875
            TabIndex        =   70
            Top             =   2160
            Width           =   4935
            Begin VB.TextBox txtFooter 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   285
               Index           =   1
               Left            =   0
               TabIndex        =   72
               Text            =   "copyright 2008 - edwinSoftware"
               Top             =   480
               Width           =   4935
            End
            Begin VB.TextBox txtFooter 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   285
               Index           =   0
               Left            =   0
               TabIndex        =   71
               Text            =   "EDWIN DELOS SANTOS"
               Top             =   120
               Width           =   4935
            End
         End
         Begin VB.Label lblSQL 
            Caption         =   "SQL Statement"
            Height          =   435
            Left            =   -74760
            TabIndex        =   152
            Top             =   2880
            Width           =   6120
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Cut Off Date"
            ForeColor       =   &H000000C0&
            Height          =   195
            Left            =   240
            TabIndex        =   95
            Top             =   720
            Width           =   885
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "NET PAY:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   91
            Top             =   2640
            Width           =   870
         End
         Begin VB.Label lblAmtInWord 
            BackColor       =   &H0000FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Amount In Word"
            Height          =   435
            Left            =   240
            TabIndex        =   89
            Top             =   2880
            Width           =   4875
         End
         Begin VB.Label lblnetpay 
            AutoSize        =   -1  'True
            BackColor       =   &H0000FF00&
            BackStyle       =   0  'Transparent
            Caption         =   " P0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   1080
            TabIndex        =   88
            Top             =   2640
            Width           =   570
         End
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Payroll Report"
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
         TabIndex        =   140
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   4095
      Left            =   5160
      Picture         =   "FrmPayroll.frx":BB356
      ScaleHeight     =   4065
      ScaleWidth      =   6225
      TabIndex        =   55
      Top             =   3600
      Visible         =   0   'False
      Width           =   6255
      Begin TabDlg.SSTab SSTab2 
         Height          =   3135
         Left            =   120
         TabIndex        =   110
         Top             =   720
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   5530
         _Version        =   393216
         TabHeight       =   520
         TabCaption(0)   =   "Account"
         TabPicture(0)   =   "FrmPayroll.frx":170912
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "lblFLDi(25)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "lblFLDi(24)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "lblFLDi(23)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "lblFLDi(22)"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "lblFLDi(21)"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "lblFLDi(26)"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "txtEntry(25)"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "txtEntry(24)"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "txtEntry(23)"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "txtEntry(22)"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "txtEntry(21)"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "txtEntry(26)"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).ControlCount=   12
         TabCaption(1)   =   "Contact"
         TabPicture(1)   =   "FrmPayroll.frx":17092E
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "txtEntry(30)"
         Tab(1).Control(1)=   "txtEntry(29)"
         Tab(1).Control(2)=   "txtEntry(28)"
         Tab(1).Control(3)=   "txtEntry(27)"
         Tab(1).Control(4)=   "lblFLDi(30)"
         Tab(1).Control(5)=   "lblFLDi(29)"
         Tab(1).Control(6)=   "lblFLDi(28)"
         Tab(1).Control(7)=   "lblFLDi(27)"
         Tab(1).ControlCount=   8
         TabCaption(2)   =   "User Info"
         TabPicture(2)   =   "FrmPayroll.frx":17094A
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "txtEntry(34)"
         Tab(2).Control(1)=   "txtEntry(33)"
         Tab(2).Control(2)=   "txtEntry(31)"
         Tab(2).Control(3)=   "txtEntry(32)"
         Tab(2).Control(4)=   "lblFLDi(34)"
         Tab(2).Control(5)=   "lblFLDi(33)"
         Tab(2).Control(6)=   "lblFLDi(31)"
         Tab(2).Control(7)=   "lblFLDi(32)"
         Tab(2).ControlCount=   8
         Begin VB.TextBox txtEntry 
            Height          =   285
            Index           =   34
            Left            =   -72840
            TabIndex        =   136
            Top             =   1680
            Width           =   1875
         End
         Begin VB.TextBox txtEntry 
            Height          =   285
            Index           =   33
            Left            =   -72840
            TabIndex        =   135
            Top             =   1320
            Width           =   1875
         End
         Begin VB.TextBox txtEntry 
            Height          =   285
            Index           =   30
            Left            =   -72840
            TabIndex        =   132
            Top             =   2160
            Width           =   2715
         End
         Begin VB.TextBox txtEntry 
            Height          =   285
            Index           =   29
            Left            =   -72840
            TabIndex        =   131
            Top             =   1800
            Width           =   2715
         End
         Begin VB.TextBox txtEntry 
            Height          =   285
            Index           =   26
            Left            =   4560
            TabIndex        =   129
            ToolTipText     =   "input: 01012008 -> output:01/01/2008"
            Top             =   600
            Width           =   1155
         End
         Begin VB.TextBox txtEntry 
            Height          =   285
            Index           =   28
            Left            =   -72840
            TabIndex        =   127
            Top             =   1440
            Width           =   2715
         End
         Begin VB.TextBox txtEntry 
            Height          =   285
            Index           =   21
            Left            =   1680
            TabIndex        =   121
            ToolTipText     =   "input: 01012008 -> output:01/01/2008"
            Top             =   600
            Width           =   1215
         End
         Begin VB.TextBox txtEntry 
            Height          =   285
            Index           =   22
            Left            =   2160
            TabIndex        =   120
            Top             =   960
            Width           =   2715
         End
         Begin VB.TextBox txtEntry 
            Height          =   285
            Index           =   23
            Left            =   2160
            TabIndex        =   119
            Top             =   1320
            Width           =   2715
         End
         Begin VB.TextBox txtEntry 
            Height          =   285
            Index           =   24
            Left            =   2160
            TabIndex        =   118
            Top             =   1680
            Width           =   2715
         End
         Begin VB.TextBox txtEntry 
            Height          =   285
            Index           =   25
            Left            =   2160
            TabIndex        =   117
            Top             =   2040
            Width           =   2715
         End
         Begin VB.TextBox txtEntry 
            Height          =   765
            Index           =   27
            Left            =   -72840
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   115
            Top             =   600
            Width           =   2715
         End
         Begin VB.TextBox txtEntry 
            Height          =   285
            Index           =   31
            Left            =   -72840
            TabIndex        =   112
            Top             =   600
            Width           =   1875
         End
         Begin VB.TextBox txtEntry 
            Height          =   285
            Index           =   32
            Left            =   -72840
            TabIndex        =   111
            Top             =   960
            Width           =   1875
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
            Index           =   34
            Left            =   -74640
            TabIndex        =   138
            Top             =   1680
            Width           =   1815
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
            Index           =   33
            Left            =   -74640
            TabIndex        =   137
            Top             =   1320
            Width           =   1815
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
            Index           =   30
            Left            =   -74640
            TabIndex        =   134
            Top             =   2160
            Width           =   1815
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
            Index           =   29
            Left            =   -74640
            TabIndex        =   133
            Top             =   1800
            Width           =   1815
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
            Index           =   26
            Left            =   3000
            TabIndex        =   130
            Top             =   600
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
            Index           =   28
            Left            =   -74640
            TabIndex        =   128
            Top             =   1440
            Width           =   1815
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
            Index           =   21
            Left            =   360
            TabIndex        =   126
            Top             =   600
            Width           =   1335
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
            Index           =   22
            Left            =   360
            TabIndex        =   125
            Top             =   960
            Width           =   1815
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
            Index           =   23
            Left            =   360
            TabIndex        =   124
            Top             =   1320
            Width           =   1815
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
            Index           =   24
            Left            =   360
            TabIndex        =   123
            Top             =   1680
            Width           =   1815
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
            Index           =   25
            Left            =   360
            TabIndex        =   122
            Top             =   2040
            Width           =   1815
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
            Index           =   27
            Left            =   -74640
            TabIndex        =   116
            Top             =   600
            Width           =   1815
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
            Index           =   31
            Left            =   -74640
            TabIndex        =   114
            Top             =   600
            Width           =   1815
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
            Index           =   32
            Left            =   -74640
            TabIndex        =   113
            Top             =   960
            Width           =   1815
         End
      End
      Begin VB.PictureBox PicMin 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Enabled         =   0   'False
         Height          =   300
         Left            =   4680
         Picture         =   "FrmPayroll.frx":170966
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   58
         Top             =   0
         Width           =   300
      End
      Begin VB.PictureBox PicClose2 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   300
         Left            =   5640
         MouseIcon       =   "FrmPayroll.frx":170CF0
         MousePointer    =   99  'Custom
         Picture         =   "FrmPayroll.frx":1715BA
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   57
         Top             =   0
         Width           =   300
      End
      Begin VB.PictureBox PicRes 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Enabled         =   0   'False
         Height          =   300
         Left            =   5160
         Picture         =   "FrmPayroll.frx":171944
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   56
         Top             =   0
         Width           =   300
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Confidential Data"
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
         TabIndex        =   139
         Top             =   120
         Width           =   1485
      End
   End
   Begin VB.PictureBox Picture6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4335
      Left            =   3360
      Picture         =   "FrmPayroll.frx":171CCE
      ScaleHeight     =   4305
      ScaleWidth      =   8370
      TabIndex        =   143
      Top             =   3000
      Visible         =   0   'False
      Width           =   8400
      Begin VB.CommandButton refreshDTR 
         BackColor       =   &H00F1E97E&
         Caption         =   "Refresh"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   325
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   154
         Top             =   3840
         Width           =   1215
      End
      Begin VB.TextBox txtID 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7440
         TabIndex        =   150
         Top             =   960
         Width           =   735
      End
      Begin VB.CommandButton PrintDTR 
         BackColor       =   &H00F1E97E&
         Height          =   525
         Left            =   7440
         Picture         =   "FrmPayroll.frx":22728A
         Style           =   1  'Graphical
         TabIndex        =   148
         Top             =   3240
         Width           =   735
      End
      Begin MSComctlLib.ListView LvDTR 
         Height          =   3255
         Left            =   120
         TabIndex        =   145
         Top             =   480
         Width           =   7080
         _ExtentX        =   12488
         _ExtentY        =   5741
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "i16x16"
         SmallIcons      =   "i16x16"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin VB.PictureBox PicBtnClose 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   300
         Left            =   7920
         MouseIcon       =   "FrmPayroll.frx":22822C
         MousePointer    =   99  'Custom
         Picture         =   "FrmPayroll.frx":228AF6
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   144
         Top             =   0
         Width           =   300
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DTR Summary"
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
         TabIndex        =   153
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Input ID"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   7560
         TabIndex        =   149
         Top             =   600
         Width           =   570
      End
   End
   Begin VB.TextBox TxtFind 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   7440
      TabIndex        =   142
      Text            =   "Find Here !"
      Top             =   8160
      Width           =   4575
   End
   Begin MSComctlLib.ImageList i16x16 
      Left            =   120
      Top             =   3960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   25
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPayroll.frx":229080
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPayroll.frx":229A92
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPayroll.frx":22A02C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPayroll.frx":22A5C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPayroll.frx":22AB60
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPayroll.frx":22AEFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPayroll.frx":22B294
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPayroll.frx":22B62E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPayroll.frx":22B9C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPayroll.frx":22BD62
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPayroll.frx":22C774
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPayroll.frx":22C7C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPayroll.frx":22CB62
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPayroll.frx":22CEFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPayroll.frx":22D296
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPayroll.frx":22D630
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPayroll.frx":22E042
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPayroll.frx":22EA54
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPayroll.frx":22F466
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPayroll.frx":22FE78
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPayroll.frx":23088A
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPayroll.frx":23129C
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPayroll.frx":231CAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPayroll.frx":23224A
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPayroll.frx":2327E6
            Key             =   ""
         EndProperty
      EndProperty
   End
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
      Left            =   0
      ScaleHeight     =   195
      ScaleWidth      =   165
      TabIndex        =   141
      Top             =   4800
      Visible         =   0   'False
      Width           =   195
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   360
      Top             =   7080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   36
      ImageHeight     =   36
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPayroll.frx":232D80
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPayroll.frx":233E2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPayroll.frx":234ED4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPayroll.frx":235F7E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbMenu 
      Align           =   2  'Align Bottom
      Height          =   870
      Left            =   0
      TabIndex        =   109
      Top             =   7950
      Width           =   12360
      _ExtentX        =   21802
      _ExtentY        =   1535
      ButtonWidth     =   2302
      ButtonHeight    =   1482
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Filter / Search"
            Key             =   "find"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Payroll Report"
            Key             =   "report"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Confidential Info"
            Key             =   "info"
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3315
      Left            =   10440
      Picture         =   "FrmPayroll.frx":237028
      ScaleHeight     =   3315
      ScaleWidth      =   1575
      TabIndex        =   74
      Top             =   480
      Width           =   1575
      Begin VB.CommandButton cmdButton 
         Caption         =   "&Delete"
         Height          =   315
         Index           =   5
         Left            =   120
         TabIndex        =   81
         Top             =   1920
         Width           =   1215
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "&Refresh"
         Height          =   315
         Index           =   6
         Left            =   120
         TabIndex        =   80
         Top             =   2280
         Width           =   1215
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "&Cancel"
         Height          =   315
         Index           =   4
         Left            =   120
         TabIndex        =   79
         Top             =   1560
         Width           =   1215
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "&Save"
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   78
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "&Update"
         Height          =   315
         Index           =   3
         Left            =   120
         TabIndex        =   77
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "&Add"
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   76
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "&Edit"
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   75
         Top             =   840
         Width           =   1215
      End
   End
   Begin VB.Frame FraEntry 
      BackColor       =   &H00E0E0E0&
      Caption         =   $"FrmPayroll.frx":3D4A44
      ForeColor       =   &H000000C0&
      Height          =   3375
      Left            =   0
      TabIndex        =   7
      Top             =   480
      Width           =   10380
      Begin VB.CommandButton cmdGO 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Caption         =   "GO >"
         Enabled         =   0   'False
         Height          =   285
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   147
         Top             =   360
         Width           =   615
      End
      Begin VB.CheckBox chkDTR 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "View DTR"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2760
         TabIndex        =   146
         Top             =   360
         Width           =   1815
      End
      Begin InstantReport.Hline Hline4 
         Height          =   30
         Left            =   5760
         TabIndex        =   108
         Top             =   120
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   53
      End
      Begin InstantReport.Hline Hline3 
         Height          =   30
         Left            =   1800
         TabIndex        =   107
         Top             =   120
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   53
      End
      Begin VB.ListBox List2 
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   1005
         ItemData        =   "FrmPayroll.frx":3D4ADB
         Left            =   240
         List            =   "FrmPayroll.frx":3D4AEB
         TabIndex        =   8
         Top             =   2040
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox txtEntry 
         Height          =   285
         Index           =   20
         Left            =   9120
         TabIndex        =   29
         Top             =   2880
         Width           =   1035
      End
      Begin VB.TextBox txtEntry 
         Height          =   285
         Index           =   19
         Left            =   9120
         TabIndex        =   28
         Top             =   2520
         Width           =   1035
      End
      Begin VB.TextBox txtEntry 
         Height          =   285
         Index           =   18
         Left            =   9120
         TabIndex        =   27
         Top             =   2160
         Width           =   1035
      End
      Begin VB.TextBox txtEntry 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   17
         Left            =   9120
         TabIndex        =   26
         Top             =   1800
         Width           =   1035
      End
      Begin VB.TextBox txtEntry 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   1
         Left            =   1800
         TabIndex        =   25
         Top             =   720
         Width           =   2835
      End
      Begin VB.TextBox txtEntry 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00008080&
         Height          =   285
         Index           =   0
         Left            =   1800
         TabIndex        =   24
         Top             =   360
         Width           =   795
      End
      Begin VB.TextBox txtEntry 
         Height          =   285
         Index           =   12
         Left            =   6360
         TabIndex        =   23
         Top             =   2880
         Width           =   1035
      End
      Begin VB.TextBox txtEntry 
         Height          =   285
         Index           =   11
         Left            =   6360
         TabIndex        =   22
         Top             =   2160
         Width           =   1035
      End
      Begin VB.TextBox txtEntry 
         Height          =   285
         Index           =   10
         Left            =   6360
         TabIndex        =   21
         Top             =   1800
         Width           =   1035
      End
      Begin VB.TextBox txtEntry 
         Height          =   285
         Index           =   9
         Left            =   6360
         TabIndex        =   20
         Top             =   1080
         Width           =   1035
      End
      Begin VB.TextBox txtEntry 
         Height          =   285
         Index           =   8
         Left            =   6360
         TabIndex        =   19
         Top             =   720
         Width           =   1035
      End
      Begin VB.TextBox txtEntry 
         Height          =   285
         Index           =   7
         Left            =   6360
         TabIndex        =   18
         Top             =   360
         Width           =   1035
      End
      Begin VB.TextBox txtEntry 
         Height          =   285
         Index           =   6
         Left            =   1800
         TabIndex        =   17
         Top             =   2880
         Width           =   1275
      End
      Begin VB.TextBox txtEntry 
         Height          =   285
         Index           =   5
         Left            =   1800
         TabIndex        =   16
         Top             =   2520
         Width           =   1275
      End
      Begin VB.TextBox txtEntry 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   3
         Left            =   1800
         TabIndex        =   15
         Top             =   1800
         Width           =   1635
      End
      Begin VB.TextBox txtEntry 
         Height          =   285
         Index           =   2
         Left            =   1800
         TabIndex        =   14
         Top             =   1080
         Width           =   2835
      End
      Begin VB.TextBox txtEntry 
         Height          =   285
         Index           =   13
         Left            =   9120
         TabIndex        =   13
         Top             =   360
         Width           =   1035
      End
      Begin VB.TextBox txtEntry 
         Height          =   285
         Index           =   14
         Left            =   9120
         TabIndex        =   12
         Top             =   720
         Width           =   1035
      End
      Begin VB.TextBox txtEntry 
         Height          =   285
         Index           =   15
         Left            =   9120
         TabIndex        =   11
         Top             =   1080
         Width           =   1035
      End
      Begin VB.TextBox txtEntry 
         Height          =   285
         Index           =   16
         Left            =   9120
         TabIndex        =   10
         Top             =   1440
         Width           =   1035
      End
      Begin VB.TextBox txtEntry 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   4
         Left            =   1800
         TabIndex        =   9
         Top             =   2160
         Width           =   1635
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   285
         Index           =   2
         Left            =   3600
         TabIndex        =   30
         Top             =   2160
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "MMM-dd-yyyy"
         Format          =   20643843
         CurrentDate     =   38207
      End
      Begin VB.ListBox List1 
         BackColor       =   &H0080FF80&
         Height          =   255
         ItemData        =   "FrmPayroll.frx":3D4B15
         Left            =   0
         List            =   "FrmPayroll.frx":3D4B37
         TabIndex        =   52
         Top             =   120
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackColor       =   &H00DD686F&
         Caption         =   "TOTALS"
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
         Left            =   4800
         TabIndex        =   104
         Top             =   1440
         Width           =   2535
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackColor       =   &H00DD686F&
         Caption         =   "COMPUTATIONS"
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
         Left            =   240
         TabIndex        =   103
         Top             =   1440
         Width           =   3210
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
         Left            =   240
         TabIndex        =   32
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "F2"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   3480
         TabIndex        =   54
         Top             =   1800
         Width           =   165
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
         Index           =   20
         Left            =   7560
         TabIndex        =   51
         Top             =   2880
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
         Index           =   19
         Left            =   7560
         TabIndex        =   50
         Top             =   2520
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
         Index           =   18
         Left            =   7560
         TabIndex        =   49
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
         Index           =   17
         Left            =   7560
         TabIndex        =   48
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
         Index           =   16
         Left            =   7560
         TabIndex        =   47
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
         Index           =   14
         Left            =   7560
         TabIndex        =   46
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
         Index           =   13
         Left            =   7560
         TabIndex        =   45
         Top             =   360
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
         Index           =   12
         Left            =   4800
         TabIndex        =   44
         Top             =   2880
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
         Index           =   11
         Left            =   4800
         TabIndex        =   43
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
         Index           =   10
         Left            =   4800
         TabIndex        =   42
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
         Index           =   9
         Left            =   4800
         TabIndex        =   41
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
         Index           =   8
         Left            =   4800
         TabIndex        =   40
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
         Index           =   7
         Left            =   4800
         TabIndex        =   39
         Top             =   360
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
         Left            =   240
         TabIndex        =   38
         Top             =   2880
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
         Index           =   5
         Left            =   240
         TabIndex        =   37
         Top             =   2520
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
         Left            =   240
         TabIndex        =   36
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
         Index           =   3
         Left            =   240
         TabIndex        =   35
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
         Index           =   2
         Left            =   240
         TabIndex        =   34
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
         Index           =   1
         Left            =   240
         TabIndex        =   33
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
         Index           =   15
         Left            =   7560
         TabIndex        =   31
         Top             =   1080
         Width           =   1575
      End
   End
   Begin VB.PictureBox PicTopBar 
      BackColor       =   &H000000FF&
      Height          =   375
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   4755
      TabIndex        =   5
      Top             =   3840
      Visible         =   0   'False
      Width           =   4815
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Do not delete"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   3600
         TabIndex        =   6
         Top             =   0
         Width           =   960
      End
   End
   Begin MSComctlLib.ListView lvList 
      Height          =   2115
      Left            =   0
      TabIndex        =   0
      Top             =   4560
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
   Begin MSComctlLib.ImageList ImageList2 
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
            Picture         =   "FrmPayroll.frx":3D4B60
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPayroll.frx":3D5572
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPayroll.frx":3D590C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPayroll.frx":3D5CA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPayroll.frx":3D66B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPayroll.frx":3D670C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPayroll.frx":3D6AA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPayroll.frx":3D6E40
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPayroll.frx":3D71DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPayroll.frx":3D7574
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPayroll.frx":3D7F86
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPayroll.frx":3D8998
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPayroll.frx":3D93AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPayroll.frx":3D9DBC
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPayroll.frx":3DA7CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPayroll.frx":3DB1E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPayroll.frx":3DBBF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPayroll.frx":3DC18E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Employee's Payroll"
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
      TabIndex        =   53
      Top             =   0
      Width           =   1590
   End
End
Attribute VB_Name = "FrmPayroll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsPAY As Recordset
Private rsDed As Recordset
Private rsDTR As Recordset
Dim totalBy As clsTotalBy
Dim s4mat As clsFormat
Dim convert As numTOword
Private prevDate As String 'store prev date format (use by txtheader(2) - see tab)
Private deductLabel() As String

Private Sub CmdPrintOption_Click()
Picture1.Visible = True
End Sub




Private Sub BtnPrint_Click()
If Not IsDate(txtCutOff(0).text) Then
    MsgBox "Invalid date!", vbCritical, "Cut Off"
    Exit Sub
End If
If Not IsDate(txtCutOff(1).text) Then
    MsgBox "Invalid date!", vbCritical, "Cut Off"
    Exit Sub
End If
If OptPayMode.Value = True Then
    PrintPayShort
ElseIf OptPaySummary.Value = True Then
     If initPrint = False Then
       MsgBox "Please Validate!", vbInformation, "Check"
       Exit Sub
     End If
     PrintReport rsPAY

End If

End Sub

Private Sub BtnTotal_Click()
  ScreenTotal
End Sub


Private Sub chkALL_Click()
  chkAND.Value = 0
  lstPayMode.Enabled = (chkALL.Value = 0)
  lstPayDate.Enabled = (chkALL.Value = 1)
  chkAND.Enabled = (chkALL.Value = 0)
End Sub

Private Sub refreshDTR_Click()
Dim SQL As String

       If rsDTR.State = adStateOpen Then
          rsDTR.Close
        End If
        
SQL = "SELECT Employee_Name,Date_Attend,Gross_Pay,ID_Code "
SQL = SQL & "From DTR order by id_code"
rsDTR.Open SQL, CnPay, adOpenStatic, adLockOptimistic

  Load_DTR
End Sub

Private Sub txtId_Change()
Dim SQL As String
Dim lngID As Long
lngID = Val(txtID.text)
If lngID = 0 Then Exit Sub
       If rsDTR.State = adStateOpen Then
          rsDTR.Close
        End If

SQL = "SELECT Employee_Name,Date_Attend,Gross_Pay,ID_Code "
SQL = SQL & "From DTR WHERE [ID_Code]=" & lngID & " order by date_attend"
rsDTR.Open SQL, CnPay, adOpenStatic, adLockOptimistic
'rsDTR.Open "SELECT Employee_Name,Date_Attend,Gross_Pay FROM DTR WHERE [ID_Code]=" & lngID & " order by date_attend"
If rsDTR.RecordCount > 0 Then
   Load_DTR
Else
  MsgBox "No Record Found!", vbInformation, "DTR File"
  Exit Sub
End If
'autoAlignCol LvDTR

End Sub

Private Sub chkAND_Click()
'If chkAND.Value = 1 Then
'  chkALL.Value = 0
'End If
  lstPayDate.Enabled = (chkAND.Value = 1)
End Sub

Private Sub chkDTR_Click()
   Picture6.Visible = (chkDTR.Value = 1)
   CmdGo.Enabled = (chkDTR.Value = 1)
 
End Sub

Private Sub viewDTR()
Dim SQL As String
Dim lngID As Long
lngID = Val(txtEntry(0).text)
If lngID = 0 Then Exit Sub
       If rsDTR.State = adStateOpen Then
          rsDTR.Close
        End If

SQL = "SELECT Employee_Name,Date_Attend,Gross_Pay,ID_Code "
SQL = SQL & "From DTR WHERE [ID_Code]=" & lngID & " order by date_attend"
rsDTR.Open SQL, CnPay, adOpenStatic, adLockOptimistic
'rsDTR.Open "SELECT Employee_Name,Date_Attend,Gross_Pay FROM DTR WHERE [ID_Code]=" & lngID & " order by date_attend"
If rsDTR.RecordCount > 0 Then
   Load_DTR
Else
  MsgBox "No Record Found!", vbInformation, "DTR File"
  Exit Sub
End If
'autoAlignCol LvDTR
End Sub

Private Sub chkLongDate_Click()
Dim chk As Integer
chk = chkLongDate.Value
If chk = 1 Then
   If IsDate(txtHeader(2).text) Then
      prevDate = txtHeader(2).text
      Dim xd As String
       xd = Format(txtHeader(2).text, "Long Date")
        txtHeader(2).text = xd
   Else
      chkLongDate.Value = 0
   End If
ElseIf chk = 0 Then
      txtHeader(2).text = prevDate
End If
End Sub

Private Sub cmdButton_Click(Index As Integer)
'//                  A S E U C D R
On Error GoTo ERRORHANDLE
Select Case Index
   Case BtnAdd                       '<------ add new record ------->'
     addRec = True
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
     NextNo = Last_Recc(rsPAY)
     If NextNo > 0 Then
       txtEntry(0).text = NextNo
       txtEntry(1).SetFocus
     Else
       txtEntry(0).Locked = False
       txtEntry(0).SetFocus
    End If
   Case BtnSave                       '<------ save new record ------>'
        cmdButtonShow ("1010011"), Me
        Call WriteData(Me, rsPAY, True)
        Call lvwPopulateData(lvList, rsPAY, 2)
        addRec = False
   Case BtnEdit                       '<------ edit record ---------->'
        editRec = True
        cmdButtonShow ("0001100"), Me
        txtEntry(31).text = Format(Now(), "Short Date")
        txtEntry(32).text = CurrUser.user_id
        txtEntry(1).SetFocus
   Case BtnUpdate                     '<------ update record -------->'
        cmdButtonShow ("1010011"), Me
        Call WriteData(Me, rsPAY, False)
        LvwReplaceData Me, rsPAY, lvList
        editRec = False
   Case BtnCancel                     '<------ cancel update -------->'
        cmdButtonShow ("1010011"), Me
        addRec = False
        editRec = False
   Case BtnDelete                     '<------ delete record -------->'
        Call Delete_Record(rsPAY, lvList)
   Case BtnRefresh                    '<------ Refresh record ------->'
        addRec = False
        edirec = False
       If rsPAY.State = adStateOpen Then
          rsPAY.Close
        End If
        rsPAY.Open "SELECT * From PAYROLL order by ID_CODE", CnPay, adOpenStatic, adLockOptimistic
        Load_DATA
        isFilter = False
        lvList.SetFocus
End Select
ERRORHANDLE:
 errorMsg Err, Me.Name, "Command Button"
End Sub

Private Sub CmdClear_Click()
Dim ans As Integer
ans = MsgBox("Proceed?", vbYesNo + vbQuestion, "Clear Data Fields")
If ans = vbYes Then
  ListPrint.Clear
End If
End Sub

Private Sub CmdFilter_Click()
      Dim sqlStatement As String
      Dim m_Table As String
      Dim m_field1 As String, m_field2 As String
      Dim m_Value1 As String, m_value2 As String
      m_Table = "PAYROLL"
      m_field1 = "PAY_MODE"
      m_field2 = "PAY_DATE"
      m_Value1 = lstPayMode.text
      m_value2 = "#" & CDate(lstPayDate.text) & "#"
      sqlStatement = "SELECT * FROM [" & m_Table & "] WHERE [" & m_field1 & "]"
      sqlStatement = sqlStatement & "LIKE '" & m_Value1 & "'"
      '// include Date
      If chkAND.Value = 1 Then
        sqlStatement = sqlStatement & "AND" & "[" & m_field2 & "]=" & m_value2
      End If
      '// all record by paydate
       
      If chkALL.Value = 1 Then
        sqlStatement = "SELECT * FROM [" & m_Table & "] WHERE [" & m_field2 & "]=" & m_value2
      End If
       lblSQL.Caption = sqlStatement
       If rsPAY.State = adStateOpen Then
          rsPAY.Close
       End If
         rsPAY.Open sqlStatement, CnPay
       If rsPAY.RecordCount > 0 Then
         Load_DATA
       Else
        MsgBox "No record found", vbInformation, "Print Option"
       End If
       
       isFilter = True
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




Private Sub PrintPayShort()
On Error GoTo errPrint
'//print payslip
Dim amtInWord As String
Dim recc As Integer
Dim rec As Integer
Dim CUT_OFF1 As String
Dim CUT_OFF2 As String
Dim ret
Dim co As String
Dim paydat As String
Dim strFont As String, sngSize As Single
ret = MsgBox("Print PaySlip?", vbYesNo + vbQuestion, "Confirm - paper size 8.5 x 11")
 If ret = vbYes Then
 
         co = txtHeader(0).text
         dat = Format(Now(), "Short Date") 'lblDateNow.Caption
         paydat = txtHeader(2).text
         CUT_OFF1 = txtCutOff(0).text
         CUT_OFF2 = txtCutOff(1).text
'         CmdPrint.Caption = "Printing..."
'//save current printer settings
         strFont = Printer.Font
         sngSize = Printer.FontSize
         '//change printer settings, then print
         Printer.Font = "ms sans serif"
         Printer.FontSize = 8
         Printer.Print
         Printer.Print
       With rsPAY
            .MoveFirst
         While Not .EOF = True
             amtInWord = convert.TOword(!net_pay)
             Printer.FontUnderline = True
             Printer.Print Tab(5); " PAYSLIP " & " [ Company: " & co & " ] " & " [ Pay Date: "; paydat & " ]";
                                                          Printer.Print Tab(85); "                             ACKNOWLEDGEMENT                          "
             Printer.FontUnderline = False
             Printer.Print Tab(5); "NAME: ";
             Printer.FontBold = True
                  Printer.Print ; Tab(15); !employee_name
             Printer.FontBold = False
             Printer.Print Tab(5); "PERIOD COVERED :"; Tab(28); CUT_OFF1 & " - " & CUT_OFF2; Tab(55); "PAY MODE: " & !PAY_MODE;
                                                           Printer.Print Tab(85); !employee_name
             Printer.Print Tab(5); "# DAYS WORKED: "; Tab(26); !days_work; Tab(31); "POS: " & !Position; Tab(61); "RATE: " & !Rate_PerDay
             Printer.Print Tab(5); "*****EARNINGS*****"; Tab(43); "*****DEDUCTION*******"
             Printer.Print Tab(5); "BASIC PAY:";
                  Call Print_Amt(25, rsPAY!basic_pay, 11)
                       Printer.Print Tab(43); deductLabel(0) & ":";
                            Call Print_Amt(65, rsPAY!Deduction1, 11)
                                                          Printer.Print Tab(85); "Received the amount in Pesos " & "{P" & Format(!net_pay, "Standard") & ")"
             Printer.Print Tab(5); "OVERTIME :";
                  Call Print_Amt(25, rsPAY!overtime_pay, 11)
                       Printer.Print Tab(43); deductLabel(1); ":";
                              Call Print_Amt(65, rsPAY!Deduction1, 11)
             Printer.Print Tab(5); "ADJUSTMENT:";
                   Call Print_Amt(25, rsPAY!adjustment, 11)
                        Printer.Print Tab(43); deductLabel(2) & ":";
                             Call Print_Amt(65, rsPAY!Deduction2, 11)
                                                         Printer.Print Tab(85); amtInWord
                   Printer.Print Tab(25); "----------------"; Tab(43); deductLabel(3) & ":";
                               Call Print_Amt(65, rsPAY!Deduction3, 11)
             Printer.Print Tab(5); "GROSS PAY:";
                   Call Print_Amt(25, rsPAY!gross_pay, 11)
                          Printer.Print Tab(43); deductLabel(4) & ":";
                               Call Print_Amt(65, rsPAY!Deduction4, 11)
             Printer.Print Tab(5); "DEDUCTION:";
                    Call Print_Amt(25, rsPAY!Total_Deduction, 11)
                          Printer.Print Tab(43); deductLabel(5) & ":";
                                  Call Print_Amt(65, rsPAY!Deduction5, 11)
                          Printer.Print Tab(43); deductLabel(6) & ":";
                                  Call Print_Amt(65, rsPAY!Deduction6, 11)
                                         Printer.Print Tab(85); "_________________________________________________"
                          Printer.Print Tab(43); deductLabel(7) & ":";
                                  Call Print_Amt(65, rsPAY!Deduction7, 11)
                                         Printer.Print Tab(85); "                    Singnature Over Printed Name"
                    Printer.Print Tab(25); "----------------"; Tab(65); "----------------"
             Printer.Print Tab(5); "NET PAY:";
                    Call Print_Amt(25, rsPAY!net_pay, 11)
'                                   Call Print_Amt(65, rsPAY!Total_Deduction, 11)
            Printer.Print Tab(5); "- - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - -" _
                                   ; " - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - " _
                                   ; "- - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - ><8 cut here"
If Not .EOF = True Then
     .MoveNext
     recc = recc + 1
     rec = rec + 1
End If
'/--
             If rec = 4 Then
                rec = 0
                If recc < .RecordCount Then
                   Printer.NewPage
                   Printer.Print
                   Printer.Print
                End If
             End If
         Wend
       End With
        '//send information to the printer
         Printer.EndDoc
         '//reset printer setting
         Printer.Font = strFont
         Printer.FontSize = sngSize
         If rsPAY.EOF = True Then
  '            CmdPrint.Caption = "Print"
          MsgBox "D O N E !!!", vbInformation, "Printing ..."
         End If
         rsPAY.MoveFirst
'         Call setDATAsource
'         Call bindDATA
    Else
     Exit Sub
   End If
errPrint:
  errorMsg Err, Me.Name, "PrintPayslip"
End Sub



Private Sub CmdShow_Click()
  frmBankLetter.show
End Sub

Private Sub CmdGo_Click()
  viewDTR
End Sub

Private Sub dtpDate_CloseUp(Index As Integer)
   txtEntry(nxTab).text = Format(dtpDate(2).Value, "mmm-dd-yyyy")
   txtEntry(nxTab).SetFocus
End Sub




Private Sub Form_Activate()
          TxtFind.Left = (tbMenu.Width - TxtFind.Width)
          TxtFind.Top = tbMenu.Top
End Sub

Private Sub Form_Load()
On Error GoTo errMsg
'[==============]
'< initialize   >
'< classes      >
'[==============]
Set totalBy = New clsTotalBy
Set s4mat = New clsFormat
Set convert = New numTOword
'[==============]
'< initialize   >
'[==============]
center_obj Me, Picture1
center_obj Me, Picture2
center_obj Me, Picture6

show
lvList.SetFocus

'// List BackColour Formatting
Call SetListViewColor(lvList, PicLv, vbWhite, &HE6F1FD, 0.1)

dtpDate(2).Value = Format(Now(), "mmm-dd-yyyy")
txtHeader(2).text = Format(Now(), "mm/dd/yyyy")
txtCutOff(0).text = Format(Now(), "mm/dd/yyyy")
txtCutOff(1).text = Format(Now(), "mm/dd/yyyy")
bRestore = False
bMin = True
cmdButtonShow ("1010011"), Me
lstPayMode.ListIndex = 0
isFilter = False
'[===============]
'< For listview  >
'[===============]
'With MainForm
'    Set lvList.SmallIcons = .i16x16
'    Set lvList.Icons = .i16x16
'End With
'[================]
'< set controlbox >
'[================]
pic_controlBox Picture1, PicMinimize, PicRestore, PicClose
pic_controlBox Picture2, PicMin, PicRes, PicClose2
'[================]
'< open recordset >
'[================]
Set rsPAY = New ADODB.Recordset
rsPAY.Open "SELECT * From PAYROLL order by ID_CODE", CnPay, adOpenStatic, adLockOptimistic
Load_DATA
Call autoAlignCol(lvList)

TextBox_Locked Me, List1
Insert_Fields List2Print, rsPAY
Call DefaultList(List2Print, ListPrint, indexList)
Call Add_Item(rsPAY, "pay_date", lstPayDate, True)
lstPayDate.ListIndex = 0

Call ShowFldsLabel(Me, rsPAY)
'// deduction
Set rsDed = New ADODB.Recordset
rsDed.Open "SELECT * From DEDUCTION order by SN", CnPay, adOpenStatic, adLockOptimistic
BindData_Label rsDed


Set rsDTR = New ADODB.Recordset
Dim SQL As String
SQL = "SELECT Employee_Name,Date_Attend,Gross_Pay,ID_Code "
SQL = SQL & "From DTR order by date_attend"
rsDTR.Open SQL, CnPay, adOpenStatic, adLockOptimistic
Load_DTR
autoAlignCol LvDTR

errMsg:
  errorMsg Err, Me.Name, "Form Load"
End Sub

Private Sub Load_DATA()
On Error GoTo ERRORHANDLE
'// set columnheaders

Call InsertColumn(lvList, rsPAY)
'//set details
 Call FillListView(lvList, rsPAY, 2)
'//get total
 Call Listview_Total(lvList, rsPAY)
ERRORHANDLE:
    errorMsg Err, Me.Name, "Load_Data"
End Sub
Private Sub Load_DTR()
'// set columnheaders
'Insert_ExtraCol lvList, rsDed
If rsDTR.RecordCount = 0 Then Exit Sub

Call InsertColumn(LvDTR, rsDTR)
'//set details
Call FillListView(LvDTR, rsDTR, 25)
Call Listview_Total(LvDTR, rsDTR)
End Sub

Private Sub BindData_Label(ByRef rs As Recordset)
Dim i As Integer
Dim ii As Integer  '// user for printing payslip
ReDim deductLabel(8) As String
i = 13
ii = 0
With rs
   .MoveFirst
While Not .EOF = True
    lblFLDi(i) = rs.Fields("LABEL")
    deductLabel(ii) = rs.Fields("LABEL")
    .MoveNext
    i = i + 1
    ii = ii + 1
Wend
Set rs = Nothing
End With
End Sub

Private Sub Deduct_Label(ByRef rs As Recordset)
Dim ii As Integer  '// user for printing payslip
ReDim deductLabel(7) As String
ii = 0
With rs
   .MoveFirst
While Not .EOF = True
    deductLabel(ii) = rs.Fields("LABEL")
    .MoveNext
    ii = ii + 1
Wend
Set rs = Nothing
End With
End Sub

Private Sub pic_controlBox(ByVal picForm As PictureBox, _
            min As PictureBox, _
            res As PictureBox, _
            clo As PictureBox)
 clo.Top = 0
 res.Top = 0
 min.Top = 0
 With picForm
  clo.Left = .ScaleWidth - clo.Width
  res.Left = .ScaleWidth - (res.Width * 2)
 min.Left = .ScaleWidth - (min.Width * 3)
 End With
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
          lvList.Width = Me.ScaleWidth
          lvList.Top = PicTopBar.Top
          lvList.Height = Me.ScaleHeight - (FraEntry.Height + 1370)
          TxtFind.Left = (tbMenu.Width - TxtFind.Width)
          TxtFind.Top = tbMenu.Top
  End If
End Sub

Private Sub lblnetpay_Change()
'//
End Sub
Private Sub ScreenTotal()
  lblnetpay.Caption = Format(Val(Screen_Total(rsPAY, 12)), "Standard")
  lblAmtInWord = convert.TOword(lblnetpay)
  lblnetpay.Caption = Format(lblnetpay, "P###,##0.00")
End Sub
Private Sub List2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   txtEntry(nxTab).text = List2.text
   txtEntry(nxTab).SetFocus
   List2.Visible = False
ElseIf KeyCode = 27 Then
   txtEntry(idx).text = List2.text
   txtEntry(nxTab).SetFocus
   List2.Visible = False
End If
End Sub

Private Sub List2Print_DblClick()
   CmdMove_Click
End Sub



Private Sub lvList_Click()
On Error GoTo ERRORHANDLE
If addRec = True Or editRec = True Then Exit Sub
Call BindDatasource(Me, rsPAY, lvList, True)
ERRORHANDLE:
 errorMsg Err, Me.Name
End Sub

Private Sub lvList_KeyUp(KeyCode As Integer, Shift As Integer)
 lvList_Click
End Sub

Private Sub chkBankLetter_Click()
 CmdShow.Enabled = (chkBankLetter.Value = 1)
End Sub

Private Sub PicBtnClose_Click()
  chkDTR.Value = 0
  Picture6.Visible = False
End Sub

Private Sub PicClose_Click()
Picture1.Visible = False
isFilter = False
End Sub

Private Sub PicClose2_Click()
 Picture2.Visible = False
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        Call DragIt(Picture1.hWnd)
    End If
End Sub

Private Sub DragIt(ByVal lngHwnd As Long)
Dim lngReturn As Long
    lngReturn = ReleaseCapture()
    lngReturn = SendMessage(lngHwnd, WM_NCLBUTTONDOWN, HTCAPTION, CLng(0))
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        Call DragIt(Picture2.hWnd)
    End If
End Sub



Private Sub Picture6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        Call DragIt(Picture6.hWnd)
    End If
End Sub

Private Sub tbMenu_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
  Case "find"
   isFilter = True
            With frmSearch
               Set .pFindForm = Me
               Set .pFindRecset = rsPAY
               Set .pFindCon = CnPay
                  .pFindTABLE = "PAYROLL"
                  .show
            End With
  Case "report"
     Picture1.Visible = True
     Picture1.ZOrder
  Case "info"
     Picture2.Visible = True
     Picture2.ZOrder
End Select
End Sub

Private Sub txtCutOff_Change(Index As Integer)
Dim xLen As Integer
 Dim xDat As String
  xLen = Len(txtCutOff(Index).text)
  Select Case xLen
    Case Is = 8, 10, 11
     xDat = txtCutOff(Index).text
     txtCutOff(Index).text = s4mat.toDate(xDat)
 End Select
End Sub

Private Sub txtEntry_Change(Index As Integer)
On Error GoTo errMsg
If addRec = True Or editRec = True Then txtEntry(Index).ForeColor = vbBlack
Select Case Index
   Case Is = 5
      txtEntry(7).text = totalBy.times(txtEntry(5), txtEntry(6))    '//(7)basic pay
   Case Is = 6
    If addRec = True Or editRec = True Then
      If CurrUser.USER_isADMIN = "Y" Then
         txtEntry(6).Locked = False
         txtEntry(7).text = totalBy.times(txtEntry(5), txtEntry(6))    '//(7)basic pay
       End If
    End If
   Case Is = 7, 8, 9 '//basic+overtime+adjustment
      txtEntry(10).text = totalBy.Sum(Me, 7, 9)                     '//(10)grosspay
   Case 13 To 20    '//deduction 1 to 8
      Dim dbl As Double
      dbl = totalBy.Sum(Me, 13, 20)                    '//(11)total deductions
      txtEntry(11).text = Format(dbl, "standard")
   Case Is = 21, 26
      If Mid(txtEntry(Index), 3, 1) <> "/" Then
           If Len(txtEntry(nxTab).text) = 8 Then
               Dim xDat As String
                   xDat = txtEntry(nxTab)
                   txtEntry(nxTab).text = s4mat.toDate(xDat)
           End If
      End If
End Select
     '//basic+grosspay-deduction = netpay
     Dim dbl5 As Double
      dbl5 = totalBy.minus(txtEntry(10), txtEntry(11)) '//(12)netpay
      txtEntry(12).text = Format(dbl5, "Standard")
      If txtEntry(12).text < 0 Then txtEntry(12).ForeColor = vbRed
errMsg:
 errorMsg Err, Me.Name, "Txtentry_change Events"
End Sub

Private Sub txtEntry_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim lastTab As Integer
On Error GoTo ERRORHANDLE
lastTab = 28  ' (rsPAY.Fields.Count - 1) 'or txtEntry Upper Bound if kung limitado lang ang textbox
If KeyCode = 13 Then
     If nxTab = lastTab Then Exit Sub
     If nxTab = 20 Then                            '//txtentry(21)
        nxTab = Index                              'stay foot ka lang
        If Picture2.Visible = False Then Exit Sub '//if confidential entry is hindi mo makita
     End If
     nxTab = nxTab + 1
     If nxTab = 6 Then nxTab = 8                   '//Passed 6 punta ka ng 8
     If nxTab = 10 Then nxTab = 13                   '//Passed 6 punta ka ng 7
ElseIf KeyCode = 38 Then  'up arrow key
     If nxTab = 0 Or nxTab = 1 Then Exit Sub
     nxTab = nxTab - 1
     If nxTab = 12 Then nxTab = 9
     If nxTab = 7 Then nxTab = 5                   '//passed 6 balik ka sa 5
End If
txtEntry(nxTab).SetFocus
ERRORHANDLE:
 errorMsg Err, Me.Name
End Sub
Private Sub txtEntry_GotFocus(Index As Integer)
Dim idx As Integer
On Error GoTo ERRORHANDLE
idx = Index
nxTab = idx
txtEntry(idx).SelStart = 0
txtEntry(idx).SelLength = Len(txtEntry(idx).text)
Select Case nxTab
Case Is = 3  'PAY MODE
     AlignObj txtEntry(idx), List2, 1, False
Case Is = 4
If IsDate(txtEntry(idx).text) Then
 If Len(txtEntry(idx).text) > 8 Then
    AlignObj txtEntry(idx), dtpDate(2), 2
 End If
End If
Case Is = 6
    If addRec = True Or editRec = True Then
      If CurrUser.USER_isADMIN = "Y" Then
         txtEntry(6).Locked = False
         txtEntry(6).BackColor = vbWhite
       End If
    End If

End Select
ERRORHANDLE:
  errorMsg Err, Me.Name
End Sub




Private Sub txtEntry_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = Asc(UCase(chr(KeyAscii)))
End Sub

Private Sub txtEntry_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case nxTab
Case Is = 3
   If KeyCode = 113 Then  'F2
        List2.Visible = True
        List2.SetFocus
   End If
End Select

End Sub

Private Sub txtEntry_LostFocus(Index As Integer)
Select Case Index
 Case Is = 4
     txtEntry(Index).text = s4mat.toDate(txtEntry(Index).text)
 Case Is = 6
      txtEntry(6).Locked = True
      txtEntry(6).BackColor = &HE0E0E0
 End Select
End Sub

Private Sub TxtFind_Change()
   Call ListView_Search(lvList, TxtFind)
End Sub

Private Sub TxtFind_KeyUp(KeyCode As Integer, Shift As Integer)
 If KeyCode = 13 Then lvList.SetFocus
End Sub

Private Sub txtHeader_Change(Index As Integer)
Dim xLen As Integer
 Dim xDat As String
  xLen = Len(txtHeader(2).text)
  Select Case xLen
    Case Is = 8, 10, 11
     xDat = txtHeader(2)
     txtHeader(2).text = s4mat.toDate(xDat)
 End Select
 End Sub


'// PRINTER REPORT PROCEDURE
'//CODED BY EDWIN DELOS SANTOS
Private Sub Headers()
 Dim dat As String
 Dim paydat As String
 Dim cutOff As String
 cutOff = "CutOff : " & txtCutOff(0).text & "-" & txtCutOff(1).text
 dat = Format(Now(), "long date")
 Printer.Print Tab(6); dat
 
 Call prnCenterText(txtHeader(0).text, 180)
 Call prnCenterText(txtHeader(1).text, 180)    'type of report
 Call prnCenterText(cutOff, 180)               'cutoff
 Call prnCenterText(txtHeader(2).text, 180)    'paydate
 
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
ans = MsgBox("Proceed?", vbYesNo + vbQuestion, "Print Payroll Summary")
  If ans = vbYes Then
'//save current printer settings
         strFont = Printer.Font
         sngSize = Printer.FontSize
         Printer.Orientation = 2   'Landscape
         Printer.Font = "ms sans serif"
         Printer.FontSize = 9
         Printer.Print
'// headers
         Headers             '< --------------------  headers ----------------->
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
          Footers             '< --------------------  footers ----------------->
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
'   errorMsg err, Me.name

End Sub

Private Sub Footers()
     Printer.Print Tab(6); txtFooter(0).text
     Printer.Print Tab(6); txtFooter(1).text
     Printer.FontSize = 5
     Printer.Print Tab(6); "copyright 2008-edwinSoftware*all rights reserved"
     Printer.Print
End Sub

