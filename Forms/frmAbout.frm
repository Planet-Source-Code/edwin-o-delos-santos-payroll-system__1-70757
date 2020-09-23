VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H0063854B&
   Caption         =   "Instant Report"
   ClientHeight    =   2745
   ClientLeft      =   2355
   ClientTop       =   1950
   ClientWidth     =   5790
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   1894.647
   ScaleMode       =   0  'User
   ScaleWidth      =   5437.109
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Supported operations"
      Height          =   255
      Left            =   3720
      TabIndex        =   8
      Top             =   480
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H0063854B&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   720
      Left            =   120
      Picture         =   "frmAbout.frx":08CA
      ScaleHeight     =   720
      ScaleWidth      =   720
      TabIndex        =   6
      Top             =   120
      Width           =   720
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Supported queries"
      Height          =   255
      Left            =   3720
      TabIndex        =   5
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Exit"
      Height          =   330
      Left            =   4485
      TabIndex        =   1
      Top             =   2190
      Width           =   1125
   End
   Begin InstantReport.Hline ctrlLiner1 
      Height          =   30
      Left            =   120
      TabIndex        =   7
      Top             =   1680
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   53
   End
   Begin VB.Image CamaGniLeonte 
      Height          =   480
      Left            =   5040
      Picture         =   "frmAbout.frx":2594
      ToolTipText     =   "Raziel Soft."
      Top             =   1275
      Width           =   480
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Print instant report from any access database!"
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
      Height          =   240
      Left            =   240
      TabIndex        =   4
      Top             =   1200
      Width           =   4050
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contact us : cyber_edu2005@yahoo.com"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1290
      TabIndex        =   3
      Top             =   2235
      Width           =   2955
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "edwinSoftware"
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
      Left            =   210
      TabIndex        =   2
      Top             =   1965
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Instant Report v1.0i"
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
      Height          =   240
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   1680
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
    MsgBox "SELECT... queries", vbInformation, "Supported queries"
End Sub

Private Sub Command3_Click()
Dim msg As String
msg = "INSTANT REPORT"
msg = msg & vbCrLf & "<< Operations >>"
msg = msg & vbCrLf & "Add    - New Record"
msg = msg & vbCrLf & "Save   - New Record "
msg = msg & vbCrLf & "Edit   - Existing Record"
msg = msg & vbCrLf & "Update - Existing Record"
msg = msg & vbCrLf & "Delete - Current Record"
msg = msg & vbCrLf & "<< Convertion from DB to Excel >>"
msg = msg & vbCrLf & "<< Convertion from Excel to DB >>"
myMsg msg, "Suppordted Operations", 2, True

End Sub

Private Sub Form_Load()
 DisableX Me
End Sub



