VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmProdList 
   BorderStyle     =   0  'None
   Caption         =   "Product List"
   ClientHeight    =   8895
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11565
   FillStyle       =   0  'Solid
   Icon            =   "ProductLIst.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "ProductLIst.frx":08CA
   ScaleHeight     =   8895
   ScaleWidth      =   11565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   WhatsThisHelp   =   -1  'True
   Begin VB.Frame Frame1 
      BackColor       =   &H00C5CFD3&
      BorderStyle     =   0  'None
      Height          =   3135
      Left            =   9480
      TabIndex        =   47
      Top             =   960
      Width           =   1695
      Begin VB.PictureBox HotKey2 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   360
         Picture         =   "ProductLIst.frx":14F4CC
         ScaleHeight     =   285
         ScaleWidth      =   90
         TabIndex        =   61
         Top             =   2640
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.PictureBox Hotkey 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         Picture         =   "ProductLIst.frx":14F6DE
         ScaleHeight     =   285
         ScaleWidth      =   90
         TabIndex        =   49
         Top             =   2640
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   315
         Left            =   120
         TabIndex        =   55
         Top             =   1560
         Width           =   1455
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   315
         Left            =   120
         TabIndex        =   54
         Top             =   480
         Width           =   1455
      End
      Begin VB.CommandButton CmdUpdate 
         Caption         =   "&Update"
         Height          =   315
         Left            =   120
         TabIndex        =   53
         Top             =   1200
         Width           =   1455
      End
      Begin VB.CommandButton CmdAdd 
         Caption         =   "&Add"
         Height          =   315
         Left            =   120
         TabIndex        =   52
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton CmdEdit 
         Caption         =   "&Edit"
         Height          =   315
         Left            =   120
         TabIndex        =   51
         Top             =   840
         Width           =   1455
      End
      Begin VB.CommandButton cmdREFRESH 
         Caption         =   "&Refresh"
         Height          =   315
         Left            =   120
         TabIndex        =   50
         Top             =   2280
         Width           =   1455
      End
      Begin VB.CommandButton CmdDelete 
         Caption         =   "&Delete"
         Height          =   315
         Left            =   120
         TabIndex        =   48
         Top             =   1920
         Width           =   1455
      End
   End
   Begin VB.TextBox TxtSearch 
      Appearance      =   0  'Flat
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
      Height          =   285
      Left            =   3600
      TabIndex        =   60
      Text            =   "TYPE ITEM TO SEARCH ..."
      Top             =   4800
      Width           =   4095
   End
   Begin VB.PictureBox PicRestore 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   11010
      MouseIcon       =   "ProductLIst.frx":14F8F0
      MousePointer    =   99  'Custom
      Picture         =   "ProductLIst.frx":1501BA
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   59
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
      Left            =   10725
      MouseIcon       =   "ProductLIst.frx":150744
      MousePointer    =   99  'Custom
      Picture         =   "ProductLIst.frx":15100E
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   58
      Top             =   50
      Width           =   240
   End
   Begin VB.PictureBox PicClose 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   11280
      MouseIcon       =   "ProductLIst.frx":151598
      MousePointer    =   99  'Custom
      Picture         =   "ProductLIst.frx":151E62
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   56
      Top             =   50
      Width           =   240
   End
   Begin VB.TextBox txtEntry 
      BackColor       =   &H00EFFCFC&
      Height          =   285
      Index           =   4
      Left            =   1850
      Locked          =   -1  'True
      TabIndex        =   46
      Top             =   2280
      Width           =   2260
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   16
      Left            =   6840
      TabIndex        =   45
      Top             =   4320
      Width           =   1035
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   15
      Left            =   6840
      TabIndex        =   44
      Text            =   "0"
      Top             =   3600
      Width           =   1515
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   14
      Left            =   6840
      TabIndex        =   43
      Text            =   "0"
      Top             =   3240
      Width           =   1515
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   13
      Left            =   6840
      TabIndex        =   42
      Text            =   "0"
      Top             =   2880
      Width           =   1515
   End
   Begin VB.PictureBox ctrlLiner2 
      Height          =   30
      Left            =   3600
      ScaleHeight     =   30
      ScaleWidth      =   7695
      TabIndex        =   41
      Top             =   4680
      Width           =   7695
   End
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
      Left            =   10800
      MousePointer    =   99  'Custom
      Picture         =   "ProductLIst.frx":1523EC
      Style           =   1  'Graphical
      TabIndex        =   39
      ToolTipText     =   "Last"
      Top             =   4800
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
      Left            =   10440
      MousePointer    =   99  'Custom
      Picture         =   "ProductLIst.frx":1526A1
      Style           =   1  'Graphical
      TabIndex        =   40
      ToolTipText     =   "Next"
      Top             =   4800
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
      Left            =   10080
      MousePointer    =   99  'Custom
      Picture         =   "ProductLIst.frx":152956
      Style           =   1  'Graphical
      TabIndex        =   38
      ToolTipText     =   "Previous"
      Top             =   4800
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
      Left            =   9720
      MaskColor       =   &H00404040&
      MousePointer    =   99  'Custom
      Picture         =   "ProductLIst.frx":152C0B
      Style           =   1  'Graphical
      TabIndex        =   37
      ToolTipText     =   "First"
      Top             =   4800
      Width           =   375
   End
   Begin MSComctlLib.ListView lvList 
      Height          =   3195
      Left            =   240
      TabIndex        =   36
      Top             =   5280
      Width           =   11025
      _ExtentX        =   19447
      _ExtentY        =   5636
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
      BackColor       =   -2147483643
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
   Begin VB.PictureBox ctrlLiner1 
      Height          =   30
      Left            =   360
      ScaleHeight     =   30
      ScaleWidth      =   10935
      TabIndex        =   35
      Top             =   5160
      Width           =   10935
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   2
      Left            =   1800
      TabIndex        =   11
      Top             =   1560
      Width           =   2490
   End
   Begin VB.TextBox txtEntry 
      BackColor       =   &H00EFFCFC&
      Height          =   285
      Index           =   3
      Left            =   1850
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   1920
      Width           =   2260
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   5
      Left            =   1845
      TabIndex        =   9
      Text            =   "0.00"
      Top             =   3000
      Width           =   1515
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   6
      Left            =   1845
      TabIndex        =   8
      Text            =   "0.00"
      Top             =   3360
      Width           =   1515
   End
   Begin VB.TextBox txtEntry 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   7
      Left            =   1920
      TabIndex        =   7
      Text            =   "0.00"
      Top             =   4080
      Width           =   1515
   End
   Begin VB.TextBox txtEntry 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   8
      Left            =   1920
      TabIndex        =   6
      Text            =   "0.00"
      Top             =   4440
      Width           =   1515
   End
   Begin VB.TextBox txtEntry 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   9
      Left            =   1920
      TabIndex        =   5
      Text            =   "0.00"
      Top             =   4800
      Width           =   1515
   End
   Begin VB.TextBox txtEntry 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   10
      Left            =   6795
      TabIndex        =   4
      Text            =   "0.00"
      Top             =   1200
      Width           =   1515
   End
   Begin VB.TextBox txtEntry 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   11
      Left            =   6840
      TabIndex        =   3
      Text            =   "0.00"
      Top             =   1560
      Width           =   1515
   End
   Begin VB.TextBox txtEntry 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   12
      Left            =   6840
      TabIndex        =   2
      Text            =   "0.00"
      Top             =   2160
      Width           =   1515
   End
   Begin VB.TextBox txtEntry 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   0
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "SN"
      Top             =   840
      Width           =   1710
   End
   Begin VB.ComboBox cmbStat 
      BackColor       =   &H00C0FFC0&
      Height          =   315
      ItemData        =   "ProductLIst.frx":152EC0
      Left            =   6840
      List            =   "ProductLIst.frx":152ECA
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   4290
      Width           =   1300
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   1
      Left            =   1845
      TabIndex        =   12
      Top             =   1140
      Width           =   2490
   End
   Begin VB.ComboBox CboSupplier 
      Height          =   315
      Left            =   2160
      TabIndex        =   62
      Text            =   "Combo1"
      Top             =   2280
      Width           =   2775
   End
   Begin VB.ComboBox CboCateg 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2160
      TabIndex        =   63
      Text            =   "Combo1"
      Top             =   1920
      Width           =   2775
   End
   Begin VB.CheckBox ChkLvwSearch 
      BackColor       =   &H8000000A&
      Caption         =   "Searh in listview"
      Height          =   255
      Left            =   8160
      TabIndex        =   64
      Top             =   4800
      Width           =   1455
   End
   Begin VB.Image Image2 
      Height          =   240
      Left            =   7920
      Picture         =   "ProductLIst.frx":152ED4
      Top             =   120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   7560
      Picture         =   "ProductLIst.frx":15345E
      Top             =   120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Product List"
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
      Left            =   480
      TabIndex        =   57
      Top             =   360
      Width           =   1125
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Product Code"
      Height          =   240
      Index           =   0
      Left            =   570
      TabIndex        =   34
      Top             =   1140
      Width           =   1215
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   240
      Index           =   1
      Left            =   570
      TabIndex        =   33
      Top             =   1515
      Width           =   1215
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Supplier"
      Height          =   240
      Index           =   2
      Left            =   570
      TabIndex        =   32
      Top             =   2250
      Width           =   1215
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Category"
      Height          =   240
      Index           =   3
      Left            =   570
      TabIndex        =   31
      Top             =   1905
      Width           =   1215
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Size"
      Height          =   240
      Index           =   4
      Left            =   570
      TabIndex        =   30
      Top             =   2940
      Width           =   1215
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Pack"
      Height          =   240
      Index           =   5
      Left            =   570
      TabIndex        =   29
      Top             =   3315
      Width           =   1215
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Packing"
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
      Left            =   570
      TabIndex        =   28
      Top             =   2640
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Delivery Pricing"
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
      Left            =   570
      TabIndex        =   27
      Top             =   3690
      Width           =   3015
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Price/Pc"
      Height          =   240
      Index           =   6
      Left            =   570
      TabIndex        =   26
      Top             =   3990
      Width           =   1215
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Price/Case"
      Height          =   240
      Index           =   7
      Left            =   600
      TabIndex        =   25
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Price/Box"
      Height          =   240
      Index           =   8
      Left            =   600
      TabIndex        =   24
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Suggested Retail Price"
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
      Left            =   5595
      TabIndex        =   23
      Top             =   840
      Width           =   2265
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "SRP Price/Pc"
      Height          =   240
      Index           =   9
      Left            =   5520
      TabIndex        =   22
      Top             =   1140
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Supplier Price"
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
      Left            =   5595
      TabIndex        =   21
      Top             =   1890
      Width           =   3015
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "SRP Price/Pack"
      Height          =   240
      Index           =   10
      Left            =   5520
      TabIndex        =   20
      Top             =   1515
      Width           =   1215
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Unit Cost (Each)"
      Height          =   240
      Index           =   11
      Left            =   5520
      TabIndex        =   19
      Top             =   2190
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Qty Setup"
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
      Left            =   5595
      TabIndex        =   18
      Top             =   2565
      Width           =   3015
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Pieces/Box"
      Height          =   240
      Index           =   12
      Left            =   5520
      TabIndex        =   17
      Top             =   2865
      Width           =   1215
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Pieces/Case"
      Height          =   240
      Index           =   13
      Left            =   5520
      TabIndex        =   16
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Box/Case"
      Height          =   240
      Index           =   14
      Left            =   5520
      TabIndex        =   15
      Top             =   3615
      Width           =   1215
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Main Product"
      Height          =   240
      Index           =   15
      Left            =   5670
      TabIndex        =   14
      Top             =   4290
      Width           =   1065
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Product Information"
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
      Left            =   570
      TabIndex        =   13
      Top             =   840
      Width           =   1725
   End
   Begin VB.Shape shpBar 
      BackColor       =   &H80000010&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   240
      Left            =   495
      Top             =   2640
      Width           =   2880
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H80000010&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   240
      Left            =   495
      Top             =   840
      Width           =   3870
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000010&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   240
      Left            =   495
      Top             =   3690
      Width           =   2865
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H80000010&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   240
      Left            =   5520
      Top             =   840
      Width           =   2790
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H80000010&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   240
      Left            =   5520
      Top             =   1890
      Width           =   2790
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H80000010&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   240
      Left            =   5520
      Top             =   2565
      Width           =   2790
   End
End
Attribute VB_Name = "FrmProdList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private rsProd        As Recordset
Dim rs                As Recordset
Private lf As Integer, tp As Integer, wd As Integer, ht As Integer
Private bRestore     As Boolean 'hanlde restore to maximize
Private bMin         As Boolean 'handle minimize state


Private Sub CboCateg_Click()
 txtEntry(3).text = CboCateg.text
End Sub

Private Sub CboSupplier_Click()
 txtEntry(4).text = CboSupplier.text
End Sub

Private Sub ChkLvwSearch_Click()
  TxtSearch.Enabled = (ChkLvwSearch.Value = 0)
End Sub

Private Sub cmbStat_Click()
  txtEntry(16).text = cmbStat.text
End Sub

Private Sub CmdAdd_Click()
Dim NextNo As Long
On Error GoTo ERRORHANDLE
NextNo = Last_Recc(rsProd)
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

Private Sub CmdDelete_Click()
 Call Delete_Record(rsProd, lvList)
End Sub


Private Sub cmdSave_Click()
 Dim X As Boolean
  X = isExist(rsProd, "ProductCode", txtEntry(1).text)
  If X = True Then
    MsgBox "Product code:[" & txtEntry(1).text & "] is Already Exist!", vbCritical, "Warning!"
    Exit Sub
  End If
 showButton "S", Me, True, True
 Call WriteData(Me, rsProd, True, 16)
End Sub

Private Sub CmdEdit_Click()
On Error GoTo ERRORHANDLE
  showButton "E", Me, True, True
  txtEntry(1).SetFocus
ERRORHANDLE:
 errorMsg Err, Me.Name
End Sub

Private Sub CmdUpdate_Click()
On Error GoTo errMsg
showButton "U", Me, True, True
 Call WriteData(Me, rsProd, False, 16)
errMsg:
    errorMsg Err, Me.Name, "Update"
End Sub

Private Sub cmdCancel_Click()
 showButton "C", Me, True, True
 lvList.SetFocus
End Sub


Private Sub CmdAdd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Call Btn_Focus(CmdAdd, Hotkey, HotKey2)
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
Call Btn_Focus(cmdREFRESH, Hotkey, HotKey2)
End Sub

Private Sub CmdFirst_Click()
rsProd.MoveFirst
 Call BindDatasource(Me, rsProd, lvList, False, 16)
End Sub

Private Sub CmdLast_Click()
rsProd.MoveLast
 Call BindDatasource(Me, rsProd, lvList, False, 16)
End Sub

Private Sub CmdNext_Click()
If rsProd.EOF = True Then
 Exit Sub
Else
 rsProd.MoveNext
Call BindDatasource(Me, rsProd, lvList, False, 16)
End If

End Sub

Private Sub CmdPrev_Click()
If rsProd.BOF = True Then
 Exit Sub
Else
 rsProd.MovePrevious
Call BindDatasource(Me, rsProd, lvList, False, 16)
End If

End Sub

Private Sub cmdRefresh_Click()
'//set details
 Call FillListView(lvList, rsProd, 2)
 Call Listview_Total(lvList, rsProd)
 lvList.SetFocus
End Sub




Private Sub Form_Activate()
 MainForm.PicClose.Enabled = False
End Sub

Private Sub Form_Load()
'//initialize
bRestore = False
bMin = True
showButton "C", Me
With MainForm
      'For listview
        Set lvList.SmallIcons = .i16x16
        Set lvList.Icons = .i16x16
End With

'//Bind the data combo
Set rs = New ADODB.Recordset
rs.Open "SELECT * FROM tbl_AP_Supplier", CN
Call Add_Item(rs, "NAME", CboSupplier)
rs.Close
'//open another table
rs.Open "SELECT * FROM tbl_IC_Category", CN
Call Add_Item(rs, "CATEGORYNAME", CboCateg)
rs.Close
Set rs = Nothing

Set rsProd = New ADODB.Recordset
Dim fldOrder As String
rsProd.Open "SELECT * From tbl_IC_Products order by sn", CN, adOpenStatic, adLockOptimistic
Load_DATA
End Sub
Private Sub Load_DATA()
On Error GoTo ERRORHANDLE
'// set columnheaders
 Call InsertColumn(lvList, rsProd)
'//set details
 Call FillListView(lvList, rsProd, 2)
 Call Listview_Total(lvList, rsProd)
ERRORHANDLE:
    errorMsg Err, Me.Name
End Sub


Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 down = False
End Sub


Private Sub Form_Unload(Cancel As Integer)
 rsProd.Close
 Set rsProd = Nothing
 Set FrmProdList = Nothing
 MainForm.PicClose.Enabled = True
End Sub


Private Sub lvList_Click()
On Error GoTo ERRORHANDLE
If addRec = True Or editRec = True Then Exit Sub
Call BindDatasource(Me, rsProd, lvList, True, 16)
ERRORHANDLE:
 errorMsg Err, Me.Name
End Sub


Private Sub lvList_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 70 Then  'F
 TxtSearch.SelStart = 0
 TxtSearch.SelLength = Len(TxtSearch.text)
 TxtSearch.SetFocus
End If
End Sub

Private Sub lvList_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo ERRORHANDLE
If addRec = True Or editRec = True Then Exit Sub
If addRec = True Then Exit Sub
If editRec = True Then Exit Sub
Select Case KeyCode
  Case Is = 33, 34, 38, 40
Call BindDatasource(Me, rsProd, lvList, True, 16)
End Select
ERRORHANDLE:
 errorMsg Err, Me.Name

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

Private Sub PicClose_Click()
  Unload FrmProdList
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

Private Sub txtEntry_Change(Index As Integer)
If ChkLvwSearch.Value = 1 Then
 If Index = 0 Then
   Call ListView_Search(lvList, txtEntry(0).text, 0)
 End If
End If
End Sub

Private Sub txtEntry_GotFocus(Index As Integer)
Dim idx As Integer
idx = Index
nxTab = idx
AlignObj txtEntry(3), CboCateg, 2
AlignObj txtEntry(4), CboSupplier, 2
txtEntry(idx).SelStart = 0
SendKeys "{home}+{end}"
'txtEntry(idx).SelLength = Len(txtEntry(idx).text)
End Sub

Private Sub txtEntry_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim lastTab As Integer
lastTab = txtEntry.UBound
If KeyCode = 13 Then
 If nxTab = lastTab Then Exit Sub
     nxTab = nxTab + 1
ElseIf KeyCode = 38 Then  'up arrow key
 If nxTab = 0 Or nxTab = 1 Then Exit Sub
     nxTab = nxTab - 1
End If
txtEntry(nxTab).SetFocus
End Sub

Private Sub TxtSearch_Change()
   Call ListView_Search(lvList, TxtSearch)
End Sub
Private Sub TxtSearch_GotFocus()
  TxtSearch.Alignment = 0
  TxtSearch.SelLength = Len(TxtSearch.text)
End Sub

Private Sub TxtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 13 Then
   lvList.SetFocus
 ElseIf KeyCode = 27 Then
   lvList.SetFocus
 End If
End Sub

Private Sub TxtSearch_LostFocus()
 TxtSearch.text = "TYPE ITEM TO SEARCH ..."
End Sub


