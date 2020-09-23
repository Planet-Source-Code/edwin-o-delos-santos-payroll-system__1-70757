VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmStockReceive 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Stock Received"
   ClientHeight    =   8865
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11550
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "StockReceived.frx":0000
   ScaleHeight     =   8865
   ScaleWidth      =   11550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
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
      Height          =   1470
      Left            =   1440
      TabIndex        =   47
      Top             =   1440
      Visible         =   0   'False
      Width           =   1935
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3075
      Left            =   2760
      TabIndex        =   42
      Top             =   2760
      Visible         =   0   'False
      Width           =   4785
      _ExtentX        =   8440
      _ExtentY        =   5424
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "i16x16"
      SmallIcons      =   "i16x16"
      ForeColor       =   12582912
      BackColor       =   15727868
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.TextBox txtEntry 
      Alignment       =   1  'Right Justify
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
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   11
      Left            =   9720
      Locked          =   -1  'True
      TabIndex        =   46
      Text            =   "0.00"
      Top             =   1920
      Width           =   1530
   End
   Begin VB.TextBox txtEntry 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E8FBFB&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   10
      Left            =   9720
      Locked          =   -1  'True
      TabIndex        =   44
      Text            =   "0.00"
      Top             =   1560
      Width           =   1530
   End
   Begin VB.CommandButton CmdView 
      BackColor       =   &H00C0FFFF&
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   2280
      Width           =   255
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C5CFD3&
      BorderStyle     =   0  'None
      Height          =   2775
      Left            =   9480
      TabIndex        =   32
      Top             =   3000
      Width           =   1695
      Begin VB.PictureBox Hotkey 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         Picture         =   "StockReceived.frx":14EC02
         ScaleHeight     =   285
         ScaleWidth      =   90
         TabIndex        =   41
         Top             =   3000
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.PictureBox HotKey2 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   360
         Picture         =   "StockReceived.frx":14EE14
         ScaleHeight     =   285
         ScaleWidth      =   90
         TabIndex        =   40
         Top             =   3000
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.CommandButton CmdDelete 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Delete"
         Height          =   315
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   1920
         Width           =   1455
      End
      Begin VB.CommandButton cmdREFRESH 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Refresh"
         Height          =   315
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   2280
         Width           =   1455
      End
      Begin VB.CommandButton CmdEdit 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Edit"
         Height          =   315
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   840
         Width           =   1455
      End
      Begin VB.CommandButton CmdAdd 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Add"
         Height          =   315
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton CmdUpdate 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Update"
         Height          =   315
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   1200
         Width           =   1455
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Save"
         Height          =   315
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   480
         Width           =   1455
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Cancel"
         Height          =   315
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   1560
         Width           =   1455
      End
   End
   Begin VB.TextBox txtEntry 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   9
      Left            =   9720
      TabIndex        =   31
      Text            =   "0"
      Top             =   1200
      Width           =   1530
   End
   Begin VB.TextBox txtEntry 
      BackColor       =   &H00E8FBFB&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   8
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   30
      Text            =   "0"
      Top             =   2280
      Width           =   1530
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   7
      Left            =   5400
      TabIndex        =   29
      Text            =   "0"
      Top             =   1920
      Width           =   1530
   End
   Begin VB.TextBox txtEntry 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   0
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "SN"
      Top             =   840
      Width           =   2370
   End
   Begin VB.TextBox txtEntry 
      BackColor       =   &H00E8FBFB&
      Height          =   285
      Index           =   4
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   2280
      Width           =   2490
   End
   Begin VB.TextBox txtEntry 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   6
      Left            =   5400
      TabIndex        =   28
      Text            =   "0"
      Top             =   1560
      Width           =   1530
   End
   Begin VB.TextBox txtEntry 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   5
      Left            =   5400
      TabIndex        =   27
      Top             =   1200
      Width           =   2490
   End
   Begin VB.PictureBox PicClose 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   11235
      MouseIcon       =   "StockReceived.frx":14F026
      MousePointer    =   99  'Custom
      Picture         =   "StockReceived.frx":14F8F0
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   25
      Top             =   50
      Width           =   240
   End
   Begin VB.PictureBox PicMinimize 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   10680
      MouseIcon       =   "StockReceived.frx":14FE7A
      MousePointer    =   99  'Custom
      Picture         =   "StockReceived.frx":150744
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   24
      Top             =   50
      Width           =   240
   End
   Begin VB.PictureBox PicRestore 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   10965
      MouseIcon       =   "StockReceived.frx":150CCE
      MousePointer    =   99  'Custom
      Picture         =   "StockReceived.frx":151598
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   23
      Top             =   50
      Width           =   240
   End
   Begin VB.PictureBox ctrlLiner2 
      Height          =   30
      Left            =   240
      ScaleHeight     =   30
      ScaleWidth      =   11055
      TabIndex        =   22
      Top             =   2760
      Width           =   11055
   End
   Begin MSComctlLib.ListView lvList 
      Height          =   5475
      Left            =   240
      TabIndex        =   21
      Top             =   3000
      Width           =   9105
      _ExtentX        =   16060
      _ExtentY        =   9657
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
      Height          =   285
      Index           =   1
      Left            =   1440
      TabIndex        =   4
      Top             =   1155
      Width           =   2235
   End
   Begin VB.TextBox txtEntry 
      BackColor       =   &H00E8FBFB&
      Height          =   285
      Index           =   2
      Left            =   1470
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1560
      Width           =   2475
   End
   Begin VB.TextBox txtEntry 
      BackColor       =   &H00E8FBFB&
      Height          =   285
      Index           =   3
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1920
      Width           =   2475
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   285
      Left            =   1440
      TabIndex        =   5
      Top             =   1080
      Width           =   2505
      _ExtentX        =   4419
      _ExtentY        =   503
      _Version        =   393216
      CustomFormat    =   "MMM-dd-yyyy"
      Format          =   54132739
      CurrentDate     =   38207
   End
   Begin MSComctlLib.ImageList i16x16 
      Left            =   7800
      Top             =   2160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "StockReceived.frx":151B22
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "StockReceived.frx":151B97
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "StockReceived.frx":1525A9
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "StockReceived.frx":152943
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "StockReceived.frx":152CDD
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "StockReceived.frx":153077
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "StockReceived.frx":153411
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "StockReceived.frx":153E23
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "StockReceived.frx":154835
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "StockReceived.frx":155247
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "StockReceived.frx":155C59
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "StockReceived.frx":15666B
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "StockReceived.frx":15707D
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "StockReceived.frx":157A8F
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "StockReceived.frx":15802B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblPcsBox 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   7560
      TabIndex        =   49
      Top             =   1980
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label lblPcsCase 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   7560
      TabIndex        =   48
      Top             =   1605
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Category:"
      ForeColor       =   &H0000011D&
      Height          =   240
      Index           =   4
      Left            =   120
      TabIndex        =   45
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   0
      Picture         =   "StockReceived.frx":1585C7
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Image2 
      Height          =   240
      Left            =   360
      Picture         =   "StockReceived.frx":158B51
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Stock Received"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   360
      TabIndex        =   26
      Top             =   360
      Width           =   1470
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Total Amount"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000011D&
      Height          =   240
      Left            =   8250
      TabIndex        =   20
      Top             =   1920
      Width           =   1395
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Date Receive"
      Height          =   240
      Index           =   0
      Left            =   195
      TabIndex        =   19
      Top             =   1140
      Width           =   1215
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Product Code"
      Height          =   240
      Index           =   1
      Left            =   195
      TabIndex        =   18
      Top             =   1515
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Stock Receive"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   210
      Left            =   4320
      TabIndex        =   17
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cases"
      Height          =   195
      Index           =   12
      Left            =   4935
      TabIndex        =   16
      Top             =   1560
      Width           =   435
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Boxes"
      Height          =   240
      Index           =   13
      Left            =   4155
      TabIndex        =   15
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Pieces"
      Height          =   240
      Index           =   14
      Left            =   4155
      TabIndex        =   14
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      ForeColor       =   &H0000011D&
      Height          =   240
      Index           =   2
      Left            =   195
      TabIndex        =   13
      Top             =   1890
      Width           =   1215
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reference"
      Height          =   195
      Index           =   3
      Left            =   4620
      TabIndex        =   12
      Top             =   1245
      Width           =   750
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Total Qty"
      ForeColor       =   &H0000011D&
      Height          =   240
      Left            =   8250
      TabIndex        =   11
      Top             =   1200
      Width           =   1395
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Unit Cost(Each)"
      ForeColor       =   &H0000011D&
      Height          =   240
      Left            =   8250
      TabIndex        =   10
      Top             =   1560
      Width           =   1395
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pcs"
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   7125
      TabIndex        =   9
      Top             =   1605
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pcs"
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   7125
      TabIndex        =   8
      Top             =   1980
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Product"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   195
      Left            =   480
      TabIndex        =   7
      Top             =   840
      Width           =   660
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Cost"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   210
      Left            =   8280
      TabIndex        =   6
      Top             =   840
      Width           =   2535
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H80000010&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   240
      Left            =   4200
      Top             =   840
      Width           =   3765
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000010&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   240
      Left            =   315
      Top             =   840
      Width           =   3645
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H80000010&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   240
      Left            =   8160
      Top             =   840
      Width           =   3045
   End
End
Attribute VB_Name = "FrmStockReceive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim rsProdRcv        As Recordset  'product received
Dim rs               As Recordset  'temp recset for product
Private lf As Integer, tp As Integer, wd As Integer, ht As Integer
Private bRestore     As Boolean 'hanlde restore to maximize
Private bMin         As Boolean 'handle minimize state
Private Sub BtnCancel_Click()
'With FmePrintOption
'   If .Visible = True Then
'      .Visible = False
'   End If
'End With
End Sub


Private Sub CmdAdd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Call Btn_Focus(CmdAdd, Hotkey, HotKey2)
End Sub


Private Sub cmdRefresh_Click()
  If rsProdRcv.State = adStateOpen Then
     rsProdRcv.Close
  End If
rsProdRcv.Open "SELECT * From tbl_IC_StockReceive order by SN ", CN, adOpenStatic, adLockOptimistic
 Call FillListView(lvList, rsProdRcv, 2)
 Call Listview_Total(lvList, rsProdRcv)
 lvList.SetFocus
End Sub

Private Sub cmdSave_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Btn_Focus(cmdSave, Hotkey, HotKey2)
End Sub

Private Sub CmdEdit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Btn_Focus(CmdEdit, Hotkey, HotKey2)
End Sub

Private Sub CmdUpdate_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Btn_Focus(CmdUpdate, Hotkey, HotKey2)
End Sub

Private Sub CmdCancel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Btn_Focus(cmdCancel, Hotkey, HotKey2)
End Sub

Private Sub CmdDelete_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Btn_Focus(CmdDelete, Hotkey, HotKey2)
End Sub

Private Sub cmdREFRESH_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Btn_Focus(cmdRefresh, Hotkey, HotKey2)
End Sub

Private Sub CmdView_Click()
AlignObj txtEntry(4), ListView1, 1
ListView1.SetFocus
End Sub





Private Sub dtpDate_Click()
  dtpDate.Visible = False
End Sub

Private Sub dtpDate_CloseUp()
   txtEntry(1).text = Format(dtpDate.Value, "mmm-dd-yyyy")
   txtEntry(1).SetFocus
End Sub



Private Sub Form_Activate()
MainForm.PicClose.Enabled = False

End Sub

Private Sub Form_Load()
'initialize
bRestore = False
bMin = True
showButton "C", Me
dtpDate.Value = Format(Now(), "mmm-dd-yyyy")
'//
Dim xd As String
xd = Format(Now(), "mm-dddd-yyyy")
Call lstCalendar(List1, xd)

With MainForm
      'For listview
        Set lvList.SmallIcons = .i16x16
        Set lvList.Icons = .i16x16
        
'        Set ListView1.SmallIcons = .i16x16
'        Set ListView1.Icons = .i16x16
        
End With
Set rsProdRcv = New ADODB.Recordset
rsProdRcv.Open "SELECT * From tbl_IC_StockReceive order by SN ", CN, adOpenStatic, adLockOptimistic
Load_DATA
'//list to print

Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient
rs.Open "SELECT ProductCode,Description,Category From tbl_IC_Products order by SN ", CN, adOpenStatic, adLockOptimistic
Load_Product

End Sub
Private Sub Load_DATA()
On Error GoTo ERRORHANDLE
'// set columnheaders
 Call InsertColumn(lvList, rsProdRcv)
'//set details
 Call FillListView(lvList, rsProdRcv, 2)
 Call Listview_Total(lvList, rsProdRcv)
ERRORHANDLE:
    errorMsg Err, Me.Name
End Sub
Private Sub Load_Product()
'// set columnheaders
 Call InsertColumn(ListView1, rs)
'//set details
 Call FillListView(ListView1, rs, 1)
 rs.Close
 Set rs = Nothing
End Sub




Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 down = False
End Sub

Private Sub Form_Resize()
'can not be set if form has no border
 'SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set FrmStockReceive = Nothing
  MainForm.PicClose.Enabled = True
End Sub



Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case Is = 13
   txtEntry(1).text = Format(List1.text, "mmm-dd-yyyy")
   List1.Visible = False
   txtEntry(1).SetFocus
Case Is = 27
   List1.Visible = False
   txtEntry(1).SetFocus
End Select
End Sub



Private Sub Listview1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
  txtEntry(2).text = ListView1.SelectedItem.text
  txtEntry(3).text = ListView1.SelectedItem.ListSubItems(1).text
  txtEntry(4).text = ListView1.SelectedItem.ListSubItems(2).text
'  txtEntry(10).text = ListView1.SelectedItem.ListSubItems(3).text
'  lblPcsCase.Caption = ListView1.SelectedItem.ListSubItems(4).text
'  lblPcsBox.Caption = ListView1.SelectedItem.ListSubItems(5).text
  txtEntry(2).SetFocus
  ListView1.Visible = False
ElseIf KeyCode = 27 Then
  txtEntry(2).SetFocus
  ListView1.Visible = False
End If
End Sub

Private Sub lvList_Click()
On Error GoTo ERRORHANDLE
If addRec = True Or editRec = True Then Exit Sub
Call BindDatasource(Me, rsProdRcv, lvList, True, 11)
ERRORHANDLE:
 errorMsg Err, Me.Name
End Sub

Private Sub lvList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    ' If the ListView is already sorted by the clicked column, _
    ' just reverse the order. Otherwise, sort the clicked column ascending.
    If lvList.Sorted = True And ColumnHeader.SubItemIndex = lvList.SortKey Then
        If lvList.SortOrder = lvwAscending Then
            lvList.SortOrder = lvwDescending
        Else
            lvList.SortOrder = lvwAscending
        End If
    Else
        lvList.Sorted = True
        lvList.SortKey = ColumnHeader.SubItemIndex
        lvList.SortOrder = lvwAscending
    End If

End Sub

Private Sub lvList_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo ERRORHANDLE
 
If addRec = True Or editRec = True Then Exit Sub
Call BindDatasource(Me, rsProdRcv, lvList, True, 11)
ERRORHANDLE:
 errorMsg Err, Me.Name
End Sub

Private Sub PicClose_Click()
  Unload FrmStockReceive
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
Private Sub CmdAdd_Click()
 Dim NextNo As Long
On Error GoTo ERRORHANDLE
NextNo = Last_Recc(rsProdRcv)
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

Private Sub cmdSave_Click()
  showButton "S", Me, True, True
  Call WriteData(Me, rsProdRcv, True, 11)
End Sub

Private Sub CmdEdit_Click()
 On Error GoTo ERRORHANDLE
  showButton "E", Me, True, True
  txtEntry(1).SetFocus
ERRORHANDLE:
 errorMsg Err, Me.Name
End Sub

Private Sub CmdUpdate_Click()
On Error GoTo errorMsg
 showButton "U", Me, True, True
Call WriteData(Me, rsProdRcv, False, 11)
errorMsg:
  errorMsg Err, Me.Name, "save"
End Sub

Private Sub cmdCancel_Click()
 showButton "C", Me, True, True
 lvList.SetFocus
End Sub


Private Sub PicMinimize_Click()
 If bMin = False Then Exit Sub
 Set picRestore.Picture = Image2
 bRestore = True
 Save_FrmMax
 Me.Move 100, Me.Height + 1000, Me.Width, 600
 bMin = False
End Sub

Private Sub PicRestore_Click()
  If bRestore = False Then Exit Sub
  Set picRestore.Picture = Image1
  Me.Move lf, tp, wd, ht
 bMin = True
End Sub

Private Sub Save_FrmMax()
   lf = Me.Left
   tp = Me.Top
   wd = Me.Width
   ht = Me.Height
End Sub



Private Sub txtEntry_Change(Index As Integer)
If editRec = False Then Exit Sub
Select Case Index
Case Is = 9, 10
 txtEntry(11).text = Val(txtEntry(9)) * Val(txtEntry(10))
Case Is = 6
If Val(txtEntry(7).text) = 0 Then
  txtEntry(8).text = Val(toNumber(lblPcsCase.Caption))
  txtEntry(9).text = Val(txtEntry(8)) * Val(txtEntry(6).text)
End If
Case Is = 7
If Val(txtEntry(6).text) = 0 Then
   txtEntry(8).text = Val(toNumber(lblPcsBox.Caption))
   txtEntry(9).text = Val(txtEntry(8)) * Val(txtEntry(7).text)
End If
End Select
End Sub
Private Sub txtEntry_GotFocus(Index As Integer)
Dim idx As Integer
idx = Index
AlignObj txtEntry(1), dtpDate, 2
nxTab = idx
txtEntry(idx).SelStart = 0
txtEntry(idx).SelLength = Len(txtEntry(idx).text)
End Sub

Private Sub txtEntry_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim lastTab As Integer
lastTab = txtEntry.UBound
Select Case KeyCode
Case Is = 13 'eNTER KEY
   If nxTab = lastTab Then Exit Sub
     nxTab = nxTab + 1
Case Is = 38  'up arrow key
 If nxTab = 0 Or nxTab = 1 Then Exit Sub
     nxTab = nxTab - 1
End Select
txtEntry(nxTab).SetFocus
End Sub

Private Sub txtEntry_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case Is = 40   'down arrow key
 If Len(txtEntry(1).text) = 0 Then
   If List1.Visible = False Then
     List1.Visible = True
     List1.SetFocus
   End If  'Visible
  End If  'LEN = 0
End Select
End Sub


