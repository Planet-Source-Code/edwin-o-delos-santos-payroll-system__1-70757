VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "edwinSoft Treeview Menu (c) 2008"
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8325
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   8325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00DE9A72&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4095
      Left            =   3120
      MouseIcon       =   "Form1.frx":000C
      MousePointer    =   99  'Custom
      ScaleHeight     =   4095
      ScaleWidth      =   135
      TabIndex        =   12
      Top             =   0
      Width           =   135
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Cancel"
      Height          =   735
      Left            =   5400
      Picture         =   "Form1.frx":0776
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Update"
      Height          =   735
      Left            =   4320
      Picture         =   "Form1.frx":1440
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Save"
      Height          =   735
      Left            =   3240
      Picture         =   "Form1.frx":210A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Refresh"
      Height          =   735
      Left            =   7080
      Picture         =   "Form1.frx":2DD4
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4200
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00DE9A72&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4095
      Left            =   3240
      ScaleHeight     =   4095
      ScaleWidth      =   4935
      TabIndex        =   4
      Top             =   0
      Width           =   4935
      Begin VB.CommandButton Command3 
         Height          =   375
         Left            =   840
         Picture         =   "Form1.frx":3A9E
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton Command2 
         Height          =   375
         Left            =   480
         Picture         =   "Form1.frx":4208
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   120
         Width           =   375
      End
      Begin MSComctlLib.ListView LvList 
         Height          =   3495
         Left            =   0
         TabIndex        =   8
         Top             =   600
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   6165
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "PK"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "DEPARTMENT"
            Object.Width           =   5292
         EndProperty
      End
      Begin VB.CommandButton Command1 
         Height          =   375
         Left            =   120
         Picture         =   "Form1.frx":4972
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   120
         Width           =   375
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Title"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   1560
         TabIndex        =   9
         Top             =   240
         Width           =   3255
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Help"
      Height          =   735
      Left            =   120
      Picture         =   "Form1.frx":50DC
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4200
      Width           =   975
   End
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
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":5DA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":6E50
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":7EFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":8FA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":A04E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":B0F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":C1A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":D24C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   3495
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   6165
      _Version        =   393217
      Indentation     =   0
      LabelEdit       =   1
      Style           =   5
      FullRowSelect   =   -1  'True
      SingleSel       =   -1  'True
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub Form_Load()

Dim nd As Node

'setting imagelist1 which contain list of
'images which will be displayed on the tree
Set Me.TreeView1.ImageList = Me.ImageList1
Me.TreeView1.LabelEdit = tvwManual


' adding the main node 1 with key "chk"
Set nd = Me.TreeView1.Nodes.Add(, , "mnu", "Option", 1) '// "checkparent")
'expanding the inserted node
nd.Expanded = True
'inserting the sub nodes to the node which contains key as "chk"
Set nd = Me.TreeView1.Nodes.Add("mnu", tvwChild, "mnu1", "Department", 2) '// "chkoff")
Set nd = Me.TreeView1.Nodes.Add("mnu", tvwChild, "mnu2", "Interface Setting", 3) '//"chkoff")
Set nd = Me.TreeView1.Nodes.Add("mnu", tvwChild, "mnu3", "Employment Type", 4)  '//"chkoff")
Set nd = Me.TreeView1.Nodes.Add("mnu", tvwChild, "mnu4", "Display Setting", 5) '//"house")
Set nd = Me.TreeView1.Nodes.Add("mnu", tvwChild, "mnu5", "Tools", 6) '// "chkoff")
Set nd = Me.TreeView1.Nodes.Add("mnu", tvwChild, "mnu6", "Sick Leave", 7) '//"chkoff")
Set nd = Me.TreeView1.Nodes.Add("mnu", tvwChild, "mnu7", "Vacation Leave", 8)  '//"chkoff")

Call Gradient(Me)
End Sub



Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
 Dim menuKey As String
 menuKey = TreeView1.SelectedItem.Key
Select Case menuKey
Case "mnu1"
  Picture1.Visible = True
  lblTitle.Caption = "Department"
Case "mnu3"
  Picture1.Visible = True
  lblTitle.Caption = "Employment Type"
End Select
End Sub

Private Sub Gradient(ByRef frm As Form)
Dim y As Single
'calculation variables for r,g,b gradiency
Dim VR, VG, VB As Single
'colors of the picture boxes
Dim Color1, Color2 As Long
'r,g,b variables for each picture box
Dim R, G, B, R2, G2, B2 As Integer
'calculation variable for extracting the rgb values
Dim temp As Long

Color1 = &HFCEFF1
Color2 = &HDD686F

'extract the r,g,b values from the first picture box
temp = (Color1 And 255)
R = temp And 255
temp = Int(Color1 / 256)
G = temp And 255
temp = Int(Color1 / 65536)
B = temp And 255
temp = (Color2 And 255)
R2 = temp And 255
temp = Int(Color2 / 256)
G2 = temp And 255
temp = Int(Color2 / 65536)
B2 = temp And 255

'//direction = "Vertical" Then
'create a calculation variable for determining the step between
'each level of the gradient; this also allows the user to create
'a perfect gradient regardless of the form size
VR = Abs(R - R2) / frm.ScaleHeight
VG = Abs(G - G2) / frm.ScaleHeight
VB = Abs(B - B2) / frm.ScaleHeight
'if the second value is lower then the first value, make the step
'negative
If R2 < R Then VR = -VR
If G2 < G Then VG = -VG
If B2 < B Then VB = -VB
'run a loop through the form height, incrementing the gradient color
'according to the height of the line being drawn
For y = 0 To frm.ScaleHeight
R2 = R + VR * y
G2 = G + VG * y
B2 = B + VB * y
'draw the line and continue through the loop
frm.Line (0, y)-(frm.ScaleWidth, y), RGB(R2, G2, B2)
Next y

End Sub

