VERSION 5.00
Begin VB.Form frmMsgBox 
   BackColor       =   &H00DEA576&
   Caption         =   "Instant Report"
   ClientHeight    =   1665
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5415
   Icon            =   "frmMsgBox.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   5415
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   320
      Left            =   4440
      TabIndex        =   1
      Top             =   1080
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtMsg 
      Appearance      =   0  'Flat
      BackColor       =   &H00DEA576&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1335
      Left            =   720
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   120
      Width           =   3615
   End
   Begin VB.CommandButton cmdYes 
      Caption         =   "&Yes"
      Height          =   320
      Left            =   4440
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdNo 
      Caption         =   "&No"
      Height          =   320
      Left            =   4440
      TabIndex        =   2
      Top             =   600
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   320
      Left            =   4440
      MouseIcon       =   "frmMsgBox.frx":038A
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Image ImgInfo 
      Height          =   720
      Left            =   0
      Picture         =   "frmMsgBox.frx":0C54
      Top             =   240
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image imgHelp 
      Height          =   480
      Left            =   120
      Picture         =   "frmMsgBox.frx":291E
      Top             =   360
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgCri 
      Height          =   480
      Left            =   120
      Picture         =   "frmMsgBox.frx":35E8
      Top             =   360
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "frmMsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'// This msgbox can only be use on a continues action within a procedure.
    'or in information/help for the user
'// use built-in msgbox otherwise ...
Option Explicit

Private Sub cmdYes_Click()
    pb_vbYes = True
    Unload Me
End Sub
Private Sub cmdNo_Click()
    pb_vbNo = True
    Unload Me
End Sub
Private Sub cmdCancel_Click()
    pb_vbCancel = True
    Unload Me
End Sub

Private Sub cmdOK_Click()
    pb_vbOK = True
    Unload Me
End Sub

Private Sub Form_Load()
   FormRndCorner Me, 365, 140  'wd=370
   DisableX Me
End Sub

Private Sub Form_Resize()
With frmMsgBox
  If .WindowState = 0 Then
   .Height = 2265
   .Width = 5505
  End If
End With
 SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub

Private Sub Form_Unload(Cancel As Integer)
    imgHelp.Visible = False
    imgCri.Visible = False
    ImgInfo.Visible = False
    txtMsg.text = ""
    Me.Caption = ""
    cmdOk.Visible = False
    cmdYes.Visible = False
    cmdNo.Visible = False
    cmdCancel.Visible = False
End Sub

