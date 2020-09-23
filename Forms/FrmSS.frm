VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmSS 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   Caption         =   "Social Security System  -  SSS"
   ClientHeight    =   8340
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   13110
   HelpContextID   =   244
   Icon            =   "FrmSS.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "FrmSS.frx":08CA
   ScaleHeight     =   8340
   ScaleWidth      =   13110
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1200
      Top             =   4800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSS.frx":4486
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSS.frx":45E0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox PicLv 
      Appearance      =   0  'Flat
      BackColor       =   &H00FAD9AF&
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
      Left            =   120
      ScaleHeight     =   195
      ScaleWidth      =   2445
      TabIndex        =   57
      Top             =   5520
      Visible         =   0   'False
      Width           =   2475
   End
   Begin VB.PictureBox PicStatusBar 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   10815
      TabIndex        =   56
      Top             =   8040
      Width           =   10815
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2175
      Left            =   0
      TabIndex        =   0
      Top             =   5280
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   3836
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   14
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "SURNAME"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "NAME"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "MI"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "SS NUMBER"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "SSS01"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "MED01"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "EC001"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "SSS02"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "MED02"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Text            =   "EC002"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   10
         Text            =   "SSS03"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   11
         Text            =   "MED03"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   12
         Text            =   "EC003"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   13
         Text            =   "TOTAL"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.PictureBox PicTopBar 
      BackColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   1
      Top             =   4920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox Picentry 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4935
      Left            =   0
      ScaleHeight     =   4935
      ScaleWidth      =   12615
      TabIndex        =   2
      Top             =   0
      Width           =   12615
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Print Option"
         ForeColor       =   &H000000C0&
         Height          =   4695
         Left            =   7200
         TabIndex        =   29
         Top             =   120
         Width           =   4935
         Begin VB.CommandButton cmdPrint 
            BackColor       =   &H00E0E0E0&
            Height          =   495
            Left            =   4080
            MouseIcon       =   "FrmSS.frx":473A
            MousePointer    =   99  'Custom
            Picture         =   "FrmSS.frx":5004
            Style           =   1  'Graphical
            TabIndex        =   55
            ToolTipText     =   "Print"
            Top             =   3840
            Width           =   735
         End
         Begin VB.CommandButton CmdGo 
            Caption         =   "Update"
            Height          =   285
            Left            =   3960
            TabIndex        =   54
            Top             =   240
            Width           =   855
         End
         Begin VB.PictureBox hline2 
            Height          =   30
            Left            =   240
            ScaleHeight     =   30
            ScaleWidth      =   6135
            TabIndex        =   53
            Top             =   2880
            Width           =   6135
         End
         Begin VB.PictureBox hline1 
            Height          =   30
            Left            =   240
            ScaleHeight     =   30
            ScaleWidth      =   6015
            TabIndex        =   52
            Top             =   960
            Width           =   6015
         End
         Begin VB.TextBox txtERIDN 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
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
            Height          =   285
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   42
            Top             =   600
            Width           =   2655
         End
         Begin VB.TextBox txtAPQTR 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
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
            Height          =   285
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   41
            Top             =   240
            Width           =   1695
         End
         Begin VB.ComboBox CboMonth 
            Height          =   315
            Left            =   2280
            TabIndex        =   40
            Text            =   "Month"
            Top             =   1560
            Width           =   1695
         End
         Begin VB.ComboBox CboYear 
            Height          =   315
            Left            =   3960
            TabIndex        =   39
            Text            =   "Year"
            Top             =   1560
            Width           =   855
         End
         Begin VB.TextBox TxtSBR 
            Height          =   285
            Left            =   2280
            TabIndex        =   38
            Top             =   1920
            Width           =   1695
         End
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   325
            Left            =   2280
            ScaleHeight     =   300
            ScaleWidth      =   1665
            TabIndex        =   35
            Top             =   1200
            Width           =   1695
            Begin VB.TextBox TxtDATETRANS 
               Height          =   315
               Left            =   0
               TabIndex        =   36
               Top             =   0
               Width           =   1455
            End
            Begin MSComCtl2.DTPicker DTPicker1 
               Height          =   255
               Left            =   1440
               TabIndex        =   37
               Top             =   0
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   450
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   20709377
               CurrentDate     =   38811
            End
         End
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   325
            Left            =   2280
            ScaleHeight     =   300
            ScaleWidth      =   1665
            TabIndex        =   32
            Top             =   2280
            Width           =   1695
            Begin VB.TextBox txtDATEPAID 
               Height          =   315
               Left            =   0
               TabIndex        =   33
               Top             =   0
               Width           =   1455
            End
            Begin MSComCtl2.DTPicker DTPicker2 
               Height          =   255
               Left            =   1440
               TabIndex        =   34
               Top             =   0
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   450
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               CustomFormat    =   "mm/dd/yyyy"
               Format          =   20709377
               CurrentDate     =   38811
            End
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "TRANSMITTAL CERTIFCATION"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   31
            Top             =   3480
            Value           =   -1  'True
            Width           =   255
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "EMPLOYEE LIST"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   30
            Top             =   3840
            Width           =   255
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ER ID Number:"
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
            Height          =   240
            Left            =   600
            TabIndex        =   51
            Top             =   600
            Width           =   1425
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Applicable Quarter:"
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
            Height          =   240
            Left            =   120
            TabIndex        =   50
            Top             =   240
            Width           =   1905
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date Transmitted:"
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
            Left            =   480
            TabIndex        =   49
            Top             =   1200
            Width           =   1590
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Applicable Month:"
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
            Left            =   480
            TabIndex        =   48
            Top             =   1560
            Width           =   1605
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "SBR Number / OR No.:"
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
            Left            =   120
            TabIndex        =   47
            Top             =   1920
            Width           =   2040
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date Paid:"
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
            Left            =   1200
            TabIndex        =   46
            Top             =   2280
            Width           =   945
         End
         Begin VB.Label disklabel 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "NR3001DK"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   240
            TabIndex        =   45
            Top             =   3120
            Width           =   825
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PRINT EMPLOYEE FILE"
            Height          =   195
            Left            =   600
            TabIndex        =   44
            Top             =   3840
            Width           =   1785
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PRINT TRANSMITTAL CERTIFICATION"
            Height          =   195
            Left            =   600
            TabIndex        =   43
            Top             =   3480
            Width           =   2940
         End
      End
      Begin VB.Frame Frame2 
         Height          =   4695
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   6975
         Begin VB.PictureBox PicInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   4335
            Left            =   120
            ScaleHeight     =   4335
            ScaleWidth      =   6705
            TabIndex        =   4
            Top             =   240
            Width           =   6705
            Begin VB.TextBox txtEntry 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H001D1D1D&
               BorderStyle     =   0  'None
               ForeColor       =   &H00FFFFFF&
               Height          =   300
               Index           =   6
               Left            =   4440
               TabIndex        =   18
               Top             =   3720
               Width           =   1200
            End
            Begin VB.TextBox txtEntry 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H001D1D1D&
               BorderStyle     =   0  'None
               ForeColor       =   &H00FFFFFF&
               Height          =   300
               Index           =   5
               Left            =   3120
               TabIndex        =   17
               Top             =   3720
               Width           =   1200
            End
            Begin VB.TextBox txtEntry 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H001D1D1D&
               BorderStyle     =   0  'None
               ForeColor       =   &H00FFFFFF&
               Height          =   300
               Index           =   4
               Left            =   1800
               TabIndex        =   16
               Top             =   3720
               Width           =   1200
            End
            Begin VB.TextBox txtEntry 
               Appearance      =   0  'Flat
               BackColor       =   &H001D1D1D&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   300
               Index           =   3
               Left            =   2400
               TabIndex        =   15
               Top             =   2520
               Width           =   735
            End
            Begin VB.TextBox txtEntry 
               Appearance      =   0  'Flat
               BackColor       =   &H001D1D1D&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   300
               Index           =   2
               Left            =   2400
               TabIndex        =   14
               Top             =   2160
               Width           =   3495
            End
            Begin VB.TextBox txtEntry 
               Appearance      =   0  'Flat
               BackColor       =   &H001D1D1D&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   300
               Index           =   1
               Left            =   2400
               TabIndex        =   13
               Top             =   1800
               Width           =   3495
            End
            Begin VB.TextBox txtEntry 
               Appearance      =   0  'Flat
               BackColor       =   &H001D1D1D&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   300
               Index           =   0
               Left            =   2400
               TabIndex        =   12
               Top             =   1440
               Width           =   3495
            End
            Begin VB.PictureBox Picture5 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   615
               Left            =   120
               ScaleHeight     =   615
               ScaleWidth      =   6495
               TabIndex        =   5
               Top             =   120
               Width           =   6495
               Begin VB.TextBox txtERNME 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H00E0E0E0&
                  BorderStyle     =   0  'None
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
                  Height          =   285
                  Left            =   1200
                  Locked          =   -1  'True
                  TabIndex        =   6
                  Top             =   10
                  Width           =   4455
               End
               Begin VB.Label lblTIME 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "TIME"
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Left            =   5880
                  TabIndex        =   11
                  Top             =   360
                  Width           =   390
               End
               Begin VB.Label lblDATE 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "DATE"
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Left            =   5880
                  TabIndex        =   10
                  Top             =   0
                  Width           =   435
               End
               Begin VB.Label Label2 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Not For Sale"
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Left            =   120
                  TabIndex        =   9
                  Top             =   360
                  Width           =   885
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "SSS 1998 V3"
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Index           =   0
                  Left            =   120
                  TabIndex        =   8
                  Top             =   0
                  Width           =   960
               End
               Begin VB.Label Label3 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "R3 TAPE/DISKETTE PROJECT"
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Left            =   2040
                  TabIndex        =   7
                  Top             =   360
                  Width           =   2355
               End
            End
            Begin VB.Shape Shape1 
               BorderColor     =   &H00FFFFFF&
               Height          =   3480
               Index           =   1
               Left            =   40
               Top             =   800
               Width           =   6615
            End
            Begin VB.Shape Shape1 
               BorderColor     =   &H00FFFFFF&
               Height          =   800
               Index           =   0
               Left            =   40
               Top             =   30
               Width           =   6615
            End
            Begin VB.Label Label13 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "ECC"
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
               Height          =   300
               Left            =   4440
               TabIndex        =   28
               Top             =   3360
               Width           =   1200
            End
            Begin VB.Label Label12 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "MED"
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
               Height          =   300
               Left            =   3120
               TabIndex        =   27
               Top             =   3360
               Width           =   1200
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "AMOUNT:"
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
               Left            =   840
               TabIndex        =   26
               Top             =   3750
               Width           =   870
            End
            Begin VB.Label Label10 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "SSS"
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
               Height          =   300
               Left            =   1800
               TabIndex        =   25
               Top             =   3360
               Width           =   1200
            End
            Begin VB.Label Label8 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "GIVEN NAME :"
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
               Left            =   1005
               TabIndex        =   24
               Top             =   2160
               Width           =   1290
            End
            Begin VB.Label Label7 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "SURNAME :"
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
               Left            =   1245
               TabIndex        =   23
               Top             =   1800
               Width           =   1050
            End
            Begin VB.Label Label5 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               Caption         =   " PREMIUM CONTRIBUTIONS "
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
               Left            =   2055
               TabIndex        =   22
               Top             =   3000
               Width           =   2625
            End
            Begin VB.Label Label6 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "SS NUMBER :"
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
               Left            =   1065
               TabIndex        =   21
               Top             =   1440
               Width           =   1230
            End
            Begin VB.Label Label9 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "MIDDLE INITIAL :"
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
               Left            =   735
               TabIndex        =   20
               Top             =   2520
               Width           =   1560
            End
            Begin VB.Label Label4 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   " << UPDATE EMPLOYEE FILE >> "
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
               Left            =   1890
               TabIndex        =   19
               Top             =   960
               Width           =   2955
            End
         End
      End
   End
End
Attribute VB_Name = "FrmSS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim rsHD  As ADODB.Recordset
Attribute rsHD.VB_VarHelpID = -1
Dim rsDB   As ADODB.Recordset
Attribute rsDB.VB_VarHelpID = -1

Private dSSS As Double, dSSS02 As Double, dSSS03 As Double
Private dMED As Double, dMED02 As Double, dMED03 As Double
Private dEC As Double, dEC002 As Double, dEC003 As Double
Private total_SSS As Double, total_MED As Double, total_EC As Double
Private iTotal As Double
Private iCount As Double
Private store_apqtr As String
Private LastRec As Double


Private Sub SS_PRINT()
On Error GoTo printErr
Dim ans
If Not IsDate(TxtDATETRANS.text) Then
    MsgBox "Invalid Date!", vbOKOnly, "Date Transmitted!"
    Exit Sub
End If
If Not IsDate(txtDATEPAID.text) Then
   MsgBox "Invalid Date!", vbOKOnly, "Date Transmitted!"
   Exit Sub
End If
If Len(TxtSBR.text) = 0 Then
   MsgBox "Blank SBR/OR Number !", vbOKOnly, "Warning !!"
   Exit Sub
End If
If CboMonth.text = "Month" Then
   MsgBox "Invalid Month!", vbOKOnly, "Warning !!"
   Exit Sub
End If
Dim strFont As String, sngSize As Single

ans = MsgBox("Print Transmittal Certificate?", vbYesNo + vbQuestion, "Confirm")
 If ans = vbYes Then
'         FrmPrintStatus.Show , FrmSSS
'         FrmPrintStatus.lblPrnStat.Caption = "Printing .... please wait!"
         '//save current printer settings
         strFont = Printer.Font
         sngSize = Printer.FontSize
        
         Printer.Orientation = 1
         Printer.Font = "ms sans serif"
         Printer.FontSize = 8
         Printer.Print
         Printer.Print
         Printer.Print
         Printer.Print
         Printer.Print
         Printer.Print
         Printer.Print
         Printer.Print
         Printer.Print Tab(7); "DISKETTE/TAPE NUMBER"; Tab(35); "-----------"; Tab(50); "NR3001DK_" & txtAPQTR.text; Tab(100); "DATE TRANSMITTED: " & TxtDATETRANS
         Printer.Print
         Printer.Print Tab(7); "EMPLOYER NAME"; Tab(35); "-----------"; Tab(50); txtERNME.text; Tab(100); "APPLICABLE MONTH: " & CboMonth.text & " " & CboYear.text
         Printer.Print
         Printer.Print Tab(7); "ER ID NUMBER"; Tab(35); "-----------"; Tab(50); txtERIDN.text
         Printer.Print
         Printer.Print
         Printer.FontUnderline = True
         Printer.Print Tab(25); "SSS"; Tab(45); "MEDICARE"; Tab(70); "EC"; Tab(85); "TOTAL"; Tab(105); "SBR NUMBER/OR#"; Tab(130); "DATE PAID"
         Printer.FontUnderline = False
         Printer.Print
         Printer.Print Tab(7); "AMOUNT"; Tab(25); Format(Val(dSSS), "Standard"); Tab(45); Format(Val(dMED), "Standard"); Tab(70); Format(Val(dEC), "Standard"); Tab(85); Format(Val(iTotal), "standard"); Tab(105); TxtSBR.text; Tab(130); txtDATEPAID
         Printer.Print
         Printer.Print
         Printer.Print
         Printer.Print
         Printer.Print
         Printer.Print
         Printer.Print
         Printer.Print
         Printer.Print
         Printer.Print
         Printer.Print
         Printer.Print
         Printer.Print
         Printer.Print
         Printer.Print
         Printer.Print
         Printer.Print
         Printer.Print
         Printer.Print
         Printer.Print
         Printer.Print
         Printer.Print
         Printer.Print Tab(7); "TOTAL EMPLOYEES IN THIS DISKETTE/TAPE  ----------         " & iCount
         Printer.Print
         Printer.Print
         Printer.Print
         Printer.Print Tab(65); "CERTIFIED CORRECT AND PAID:"
         Printer.Print
         Printer.Print
         Printer.Print
         Printer.Print Tab(7); "RECEIVED BY: ________________________";
         Printer.FontUnderline = True
         Printer.Print Tab(105); "       EDWIN DELOS SANTOS        "
         Printer.FontUnderline = False
         Printer.Print
         Printer.Print Tab(7); "DATE RECEIVED: ______________________"; Tab(105); "_____________________________"
         Printer.Print
   
   MsgBox "Done !!!", vbInformation, "Printing - Certificate"
Else
   MsgBox "Cancelled!!!", vbInformation, "Printing - Certificate"
 End If
   Printer.EndDoc
   
printErr:
   errorMsg Err, Me.Name

End Sub





Private Function monthNUMBER(srcStr As String, srcYR As String) As String
Dim i_mo As String
Select Case Mid(srcStr, 1, 3)
Case Is = "JAN"
    i_mo = "01" & srcYR
Case Is = "FEB"
    i_mo = "02" & srcYR
Case Is = "MAR"
    i_mo = "03" & srcYR
Case Is = "APR"
    i_mo = "04" & srcYR
Case Is = "MAY"
    i_mo = "05" & srcYR
Case Is = "JUN"
    i_mo = "06" & srcYR
Case Is = "JUL"
    i_mo = "07" & srcYR
Case Is = "AUG"
    i_mo = "08" & srcYR
Case Is = "SEP"
    i_mo = "09" & srcYR
Case Is = "OCT"
    i_mo = "10" & srcYR
Case Is = "NOV"
    i_mo = "11" & srcYR
Case Is = "DEC"
    i_mo = "12" & srcYR
End Select
 monthNUMBER = i_mo
End Function

Private Sub CboMonth_Click()
Dim mo As String
Dim yr As String
mo = CboMonth.text
yr = CboYear.text
txtAPQTR.text = monthNUMBER(mo, yr)
disklabel.Caption = "NR3001DK_" & txtAPQTR.text
End Sub

Private Sub DTPicker1_CloseUp()
   TxtDATETRANS.text = Format(DTPicker1.Value, "mm/dd/yyyy")
   TxtDATETRANS.SetFocus
End Sub

Private Sub DTPicker2_CloseUp()
   txtDATEPAID.text = Format(DTPicker2.Value, "mm/dd/yyyy")
   txtDATEPAID.SetFocus
End Sub

Private Sub Form_Activate()
   If WindowState <> vbMinimized Then
           PicStatusBar.Top = Me.ScaleHeight - PicStatusBar.Height
           PicStatusBar.Width = Me.ScaleWidth
           ListView1.Width = Me.ScaleWidth
           ListView1.Top = PicTopBar.Top
           ListView1.Height = Me.ScaleHeight - (PicEntry.Height + 350)
    End If
End Sub

Private Sub Form_Load()
'StatusBar1.Panels(1).text = FrmMain.lblCurruser
lblDate = Format(Now(), "mm/dd/yyyy")
DTPicker1.Value = Format(Now(), "mm/dd/yyyy")
DTPicker2.Value = Format(Now(), "mm/dd/yyyy")
show
ListView1.SetFocus
'// List BackColour Formatting
Call SetListViewColor(ListView1, PicLv, vbWhite, &HFAD9AF, 0.1)

Dim yr As Long
Dim iyr As Long
Dim nxyr As Long
iyr = Mid(lblDate, 7, 4)
      CboYear.text = iyr   ' current year
iyr = Val(iyr) - 2
nxyr = iyr + 10
    For yr = iyr To nxyr Step 1
    CboYear.AddItem yr
    Next yr
'//

With CboMonth
    .AddItem "JANUARY"
    .AddItem "FEBRUARY"
    .AddItem "MARCH"
    .AddItem "APRIL"
    .AddItem "MAY"
    .AddItem "JUNE"
    .AddItem "JULY"
    .AddItem "AUGUST"
    .AddItem "SEPTEMBER"
    .AddItem "OCTOBER"
    .AddItem "NOVEMBER"
    .AddItem "DECEMBER"
End With
 '//

    Set rsHD = New ADODB.Recordset
    rsHD.Open "SELECT * From HEADER", CnPay, adOpenStatic, adLockOptimistic
    Call HDBindData
    Set rsDB = New ADODB.Recordset
    rsDB.Open "SELECT * From EMPLOYEE order by ESURN", CnPay, adOpenStatic, adLockOptimistic
    LoadSS
End Sub
Sub HDBindData()
  If rsHD.EOF = True Then Exit Sub
  If rsHD.BOF = True Then Exit Sub
On Error Resume Next
  txtERNME = rsHD!ERNME
  txtAPQTR.text = rsHD!apQTR
  txtERIDN.text = rsHD!ERIDN
  store_apqtr = rsHD!apQTR
End Sub

Private Sub LoadSS()
Dim lstmain1 As ListItem
Dim total_PREM As Double
'// init variable
     iCount = 0
     iTotal = 0
     dSSS = 0
     dSSS02 = 0
     dSSS03 = 0
     dMED = 0
     dMED02 = 0
     dMED03 = 0
     dEC = 0
     dEC002 = 0
     dEC003 = 0

ListView1.ListItems.Clear
If rsDB.RecordCount > 0 Then
rsDB.MoveFirst
Do While Not rsDB.EOF
  
    Set lstmain1 = ListView1.ListItems.Add(, , rsDB!esurn, 1, 1)
'/* papulate list */
     With lstmain1.ListSubItems.Add(, , rsDB!ename)
           .ForeColor = vbBlack
     End With
    
    If Not IsNull(rsDB.Fields("EENMI")) Then
       lstmain1.SubItems(2) = rsDB.Fields("EENMI")
    End If
    If Not IsNull(rsDB.Fields("SSNUM")) Then
     With lstmain1.ListSubItems.Add(, , rsDB!ssnum)
           .ForeColor = vbBlack
           .Bold = True
     End With
    End If
    If Not IsNull(rsDB.Fields("SSS01")) Then
       lstmain1.SubItems(4) = rsDB.Fields("SSS01")
    End If
    If Not IsNull(rsDB.Fields("MED01")) Then
       lstmain1.SubItems(5) = rsDB.Fields("MED01")
    End If
    If Not IsNull(rsDB.Fields("EC001")) Then
       lstmain1.SubItems(6) = rsDB.Fields("EC001")
    End If
    If Not IsNull(rsDB.Fields("SSS02")) Then
       lstmain1.SubItems(7) = rsDB.Fields("SSS02")
    End If
    If Not IsNull(rsDB.Fields("MED02")) Then
       lstmain1.SubItems(8) = rsDB.Fields("MED02")
    End If
    If Not IsNull(rsDB.Fields("EC002")) Then
       lstmain1.SubItems(9) = rsDB.Fields("EC002")
    End If
    If Not IsNull(rsDB.Fields("SSS03")) Then
       lstmain1.SubItems(10) = rsDB.Fields("SSS03")
    End If
    If Not IsNull(rsDB.Fields("MED03")) Then
       lstmain1.SubItems(11) = rsDB.Fields("MED03")
    End If
    If Not IsNull(rsDB.Fields("EC003")) Then
       lstmain1.SubItems(12) = rsDB.Fields("EC003")
    End If
    
        
    
'// total
    total_PREM = rsDB!SSS01 + rsDB!MED01 + rsDB!EC001
    total_PREM = total_PREM + rsDB!SSS02 + rsDB!MED02 + rsDB!EC002
    total_PREM = total_PREM + rsDB!SSS03 + rsDB!MED03 + rsDB!EC003
    With lstmain1.ListSubItems.Add(, , total_PREM)
           .ForeColor = vbBlack
    End With
     iCount = iCount + 1
     dSSS = dSSS + rsDB!SSS01
     dSSS02 = dSSS02 + rsDB!SSS02
     dSSS03 = dSSS03 + rsDB!SSS03
     dMED = dMED + rsDB!MED01
     dMED02 = dMED02 + rsDB!MED02
     dMED03 = dMED03 + rsDB!MED03
     dEC = dEC + rsDB!EC001
     dEC002 = dEC002 + rsDB!EC002
     dEC003 = dEC003 + rsDB!EC003
     iTotal = iTotal + total_PREM
    rsDB.MoveNext
Loop
End If
LvTotal
End Sub
Sub LvTotal()
Dim lstmain1 As ListItem
If ListView1.ListItems.Count = 0 Then Exit Sub
 Set lstmain1 = ListView1.ListItems.Add(, , iCount)
     lstmain1.ForeColor = vbRed
     lstmain1.SubItems(4) = (Format(dSSS, "sTANDARD"))
     lstmain1.ListSubItems(4).ForeColor = vbRed
     lstmain1.SubItems(5) = (Format(dMED, "sTANDARD"))
     lstmain1.ListSubItems(5).ForeColor = vbRed
     lstmain1.SubItems(6) = (Format(dEC, "sTANDARD"))
     lstmain1.ListSubItems(6).ForeColor = vbRed
     '//
     lstmain1.SubItems(7) = (Format(dSSS02, "sTANDARD"))
     lstmain1.ListSubItems(7).ForeColor = vbRed
     lstmain1.SubItems(8) = (Format(dMED02, "sTANDARD"))
     lstmain1.ListSubItems(8).ForeColor = vbRed
     lstmain1.SubItems(9) = (Format(dEC002, "sTANDARD"))
     lstmain1.ListSubItems(9).ForeColor = vbRed
     '//
     lstmain1.SubItems(10) = (Format(dSSS03, "sTANDARD"))
     lstmain1.ListSubItems(10).ForeColor = vbRed
     lstmain1.SubItems(11) = (Format(dMED03, "sTANDARD"))
     lstmain1.ListSubItems(11).ForeColor = vbRed
     lstmain1.SubItems(12) = (Format(dEC003, "sTANDARD"))
     lstmain1.ListSubItems(12).ForeColor = vbRed
'//
     lstmain1.SubItems(13) = (Format(iTotal, "sTANDARD"))
     lstmain1.ListSubItems(13).ForeColor = vbBlue
    
End Sub



Private Sub Form_Unload(Cancel As Integer)
   
    rsHD.Close
    Set rsHD = Nothing
    rsDB.Close
    Set rsDB = Nothing
    Set FrmSS = Nothing
End Sub

Private Sub CmdGo_Click()
Dim apMonth As Integer
total_SSS = 0
total_MED = 0
total_EC = 0

Dim ans
apMonth = Val(Mid(txtAPQTR.text, 1, 2))
ans = MsgBox("Applicable Month < " & apMonth & " >", vbYesNo + vbQuestion, "Update!")
If ans = vbYes Then

With rsDB
    .MoveFirst
     While Not .EOF = True
     total_SSS = !SSS01 + !SSS02 + !SSS03
     total_MED = !MED01 + !MED02 + !MED03
     total_EC = !EC001 + !EC002 + !EC003
      Select Case apMonth
       Case Is = 1, 4, 7, 10
        !SSS01 = total_SSS
        !MED01 = total_MED
        !EC001 = total_EC
        !SSS02 = 0
        !SSS03 = 0
        !MED02 = 0
        !MED03 = 0
        !EC002 = 0
        !EC003 = 0
       Case Is = 2, 5, 8, 11
        !SSS02 = total_SSS
        !MED02 = total_MED
        !EC002 = total_EC
        !SSS01 = 0
        !SSS03 = 0
        !MED01 = 0
        !MED03 = 0
        !EC001 = 0
        !EC003 = 0
       Case Is = 3, 6, 9, 12
        !SSS03 = total_SSS
        !MED03 = total_MED
        !EC003 = total_EC
        !SSS01 = 0
        !SSS02 = 0
        !MED01 = 0
        !MED02 = 0
        !EC001 = 0
        !EC002 = 0
        
       End Select
    .Update
    .MoveNext
    Wend
End With
End If
UpdateHeader
End Sub

Private Sub CmdPrint_Click()
If Option1(0).Value = True Then
   Call SS_PRINT
Else
   Call EMP_FILE
End If
End Sub
Private Sub EMP_FILE()
On Error GoTo printErr
Dim jrecc As Integer
Dim jrec As Integer
Dim juser As String
Dim co As String 'company
Dim apQTR As String
Dim erID As String
'//
Dim sss_01 As Double
Dim sss_02 As Double
Dim sss_03 As Double
Dim med_01 As Double
Dim med_02 As Double
Dim med_03 As Double
Dim ec_001 As Double
Dim ec_002 As Double
Dim ec_003 As Double
'//
Dim lCOUNT As Long
Dim ret
Dim strFont As String, sngSize As Single

'//init var
co = txtERNME.text
apQTR = txtAPQTR.text
erID = txtERIDN.text
'//------------------------
ret = MsgBox("EMPLOYEE FILE", vbYesNo + vbQuestion, "Print")
 If ret = vbYes Then
         '//save current printer settings
         strFont = Printer.Font
         sngSize = Printer.FontSize
         '// orientation: 2 for landscape
         Printer.Orientation = 2
         Printer.Font = "ms sans serif"
         Printer.FontSize = 9
         Printer.Print
         Printer.Print
         Printer.Print
         Printer.Print
         Printer.Print Tab(10); "00" & co; Tab(55); apQTR & erID
       With rsDB
            .MoveFirst
         While Not .EOF = True
             '// calling prnamt function you need some adjustment ... see details on your print-out ...
             Printer.Print Tab(10); !reccd & !esurn; Tab(35); !ename; Tab(55); !eenmi & !ssnum;
             Call PrnAmt(75, !SSS01)
             Call PrnAmt(86, !SSS02)
             Call PrnAmt(97, !SSS03)
             Call PrnAmt(108, !MED01)
             Call PrnAmt(119, !MED02)
             Call PrnAmt(130, !MED03)
             Call PrnAmt(141, !EC001)
             Call PrnAmt(152, !EC002)
             Call PrnAmt(163, !EC003)
             Printer.Print Tab(178); !remks
             '//*one space - line spacing *//
             ' Printer.Print
               sss_01 = sss_01 + !SSS01
               sss_02 = sss_02 + !SSS02
               sss_03 = sss_03 + !SSS03
               med_01 = med_01 + !MED01
               med_02 = med_02 + !MED02
               med_03 = med_03 + !MED03
               ec_001 = ec_001 + !EC001
               ec_002 = ec_002 + !EC002
               ec_003 = ec_003 + !EC003
             lCOUNT = lCOUNT + 1
            '//space after one record printed...
            'Printer.Print
             .MoveNext
             '// store record printed  to jrec ....
             jrec = jrec + 1
             jrecc = jrecc + 1
             '// if record = 40 procedd to next page ...
             If jrec = 40 Then
'//-- next page indicator ----------------
'             Printer.Print Tab(6); "----------------------- next page ------------------- "
'//-----------------------------
                jrec = 0
                If jrecc < .RecordCount Then
                   Printer.NewPage
                   Printer.Print
                   Printer.Print
 '//-- next page headings -------------------------------
         Printer.FontSize = 9
         Printer.Print
         Printer.Print
         Printer.Print
         Printer.Print
         Printer.Print Tab(6); "00" & co; Tab(38); apQTR & erID
         Printer.Print
'//------------------------------------------------
                End If
             End If
        Wend  'while not EOF
         ''=======================================
         '// footer
      
         Printer.Print Tab(10); "99";
             Call PrnAmt(75, sss_01)
             Call PrnAmt(86, sss_02)
             Call PrnAmt(97, sss_03)
             Call PrnAmt(108, med_01)
             Call PrnAmt(119, med_02)
             Call PrnAmt(130, med_03)
             Call PrnAmt(141, ec_001)
             Call PrnAmt(152, ec_002)
             Call PrnAmt(163, ec_003)
         Printer.Print
       End With
           
        '//send information to the printer
         Printer.EndDoc
         '//reset printer setting
         Printer.Font = strFont
         Printer.FontSize = sngSize
         If rsDB.EOF = True Then
           'CmdPrint.Caption = "Print"
           MsgBox "D O N E !!", vbInformation, "Printing..."
         End If
         rsDB.MoveFirst
     Else
     Exit Sub
   End If
printErr:
  errorMsg Err, Me.Name

End Sub


Private Sub UpdateHeader()
store_apqtr = txtAPQTR.text
With rsHD
  .Fields(1) = txtERNME.text
  .Fields(2) = txtAPQTR.text
  .Fields(3) = txtERIDN.text
   rsHD.Update
End With
MsgBox "D O N E !!!", vbOKOnly, "Update"
End Sub



Private Sub ListView1_BeforeLabelEdit(Cancel As Integer)
Cancel = True
End Sub

Private Function Format_SSNUM(srcStr As String) As String
Dim snum As String
Dim l As String
Dim c As String
Dim r As String
l = Mid(srcStr, 1, 2)
c = Mid(srcStr, 3, 7)
r = Mid(srcStr, 10, 1)
snum = l & "-" & c & "-" & r
Format_SSNUM = snum
End Function



Private Sub Listview1_Click()
  Call showDATA(rsDB, "SSNUM", ListView1, 3)
End Sub


Private Sub showDATA(srcRS As Recordset, srcFLD As String, srcLV As ListView, colNUM As Integer)
On Error Resume Next
Dim srcREF
  If srcRS.RecordCount = 0 Then Exit Sub
  If srcLV.ListItems.Count = 0 Then Exit Sub
  srcREF = srcLV.SelectedItem.ListSubItems(colNUM).text  'SSNUM
 With srcRS
    .MoveFirst
   Do Until .EOF
   If .Fields(srcFLD) = srcREF Then
      BindData
     Exit Sub
   Else
     .MoveNext
   End If
   Loop
 End With
End Sub

Private Sub BindData()
On Error Resume Next
total_SSS = 0
total_MED = 0
total_EC = 0
  
  If rsDB.EOF = True Then Exit Sub
  If rsDB.BOF = True Then Exit Sub
  If Not IsNull(rsDB!ssnum) Then
    txtEntry(0) = Format_SSNUM(rsDB!ssnum)
  End If
  If Not IsNull(rsDB!esurn) Then
    txtEntry(1) = rsDB!esurn
  End If
  If Not IsNull(rsDB!ename) Then
    txtEntry(2) = rsDB!ename
  End If
  If Not IsNull(rsDB!eenmi) Then
    txtEntry(3) = rsDB!eenmi
  End If
'//TOTAL
total_SSS = rsDB!SSS01 + rsDB!SSS02 + rsDB!SSS03
total_MED = rsDB!MED01 + rsDB!MED02 + rsDB!MED03
total_EC = rsDB!EC001 + rsDB!EC002 + rsDB!EC003
'//
 txtEntry(4) = Format(total_SSS, "standard")
 txtEntry(5) = Format(total_MED, "standard")
 txtEntry(6) = Format(total_EC, "standard")
  
End Sub



Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   SortListView ListView1, ColumnHeader
End Sub

Private Sub Listview1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 115 Then 'F4 -Refresh
  If rsDB.State = adStateOpen Then
    rsDB.Close
  End If
 '//
   Dim SQL
   SQL = "select * from EMPLOYEE order by ESURN"
   rsDB.Open SQL, CnPay
  '//
   If rsDB.RecordCount > 0 Then
    LoadSS
   End If
End If  'keycode
End Sub

Private Sub Listview1_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case Is = 37, 38, 39, 40
 Call showDATA(rsDB, "SSNUM", ListView1, 3)
End Select
End Sub

Private Sub Option1_Click(Index As Integer)
Dim mo As String
Dim yr As String
If Index = 1 Then
  txtAPQTR.text = store_apqtr
  disklabel.Caption = txtAPQTR.text & txtERIDN
Else
  mo = CboMonth.text
  yr = CboYear.text
  txtAPQTR.text = monthNUMBER(mo, yr)
  disklabel.Caption = "NR3001DK_" & txtAPQTR.text
End If
End Sub

Private Sub Timer1_Timer()
LblTime.Caption = Format(Now, "hh:mm:ss AMPM")
End Sub

'//procedure to align amount to the left: as in  31220.00
'//                                                200.00
'//original coding by myself
'---------------------------------------------------------
Private Function PrnAmt(iTab As Integer, amt As Double) As Double
 Dim intLEN As Integer, currtab As Integer, tf As Boolean
  intLEN = Len(Trim(Format(amt, "Standard")))
  Select Case intLEN
    Case Is = 0
      currtab = iTab + 9
    Case Is = 1
      currtab = iTab + 8
    Case Is = 2
      currtab = iTab + 7
    Case Is = 3
      currtab = iTab + 6
    Case Is = 4
      currtab = iTab + 5
    Case Is = 5
      currtab = iTab + 4
    Case Is = 6
      currtab = iTab + 3
    Case Is = 7
      currtab = iTab + 2
    Case Is = 8
      currtab = iTab + 1
  End Select
   Printer.Print Tab(currtab); Format(amt, "standard");
End Function

