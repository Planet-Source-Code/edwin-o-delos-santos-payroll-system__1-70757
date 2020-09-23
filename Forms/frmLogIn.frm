VERSION 5.00
Object = "{9ACEED45-5983-4474-BF17-55AA24019736}#1.0#0"; "esGuard.ocx"
Begin VB.Form frmLogIn 
   Caption         =   "Login User"
   ClientHeight    =   2040
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5115
   Icon            =   "frmLogIn.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   5115
   StartUpPosition =   2  'CenterScreen
   Begin eSRedAlert.esGuard RedAlert 
      Height          =   375
      Left            =   0
      TabIndex        =   13
      Top             =   1680
      Visible         =   0   'False
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
   End
   Begin VB.ListBox CboUser 
      Appearance      =   0  'Flat
      BackColor       =   &H00E8FBFB&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   750
      Left            =   1320
      TabIndex        =   10
      Top             =   1200
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   240
      Picture         =   "frmLogIn.frx":08CA
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   9
      Top             =   0
      Width           =   480
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   0
      Picture         =   "frmLogIn.frx":1194
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   8
      Top             =   0
      Width           =   480
   End
   Begin InstantReport.Hline Hline2 
      Height          =   30
      Left            =   240
      TabIndex        =   7
      Top             =   600
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   53
   End
   Begin InstantReport.Hline Hline1 
      Height          =   30
      Left            =   240
      TabIndex        =   6
      Top             =   1800
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   53
   End
   Begin VB.TextBox txtPass 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1320
      MaxLength       =   15
      MousePointer    =   99  'Custom
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1320
      Width           =   2175
   End
   Begin VB.TextBox txtname 
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
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1320
      MousePointer    =   99  'Custom
      TabIndex        =   2
      ToolTipText     =   "Down arrow key to view user list!"
      Top             =   840
      Width           =   2175
   End
   Begin VB.CommandButton CmdLogIn 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3720
      Picture         =   "frmLogIn.frx":1A5E
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CheckBox chkAdmin 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Admin"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3720
      TabIndex        =   0
      Top             =   840
      Width           =   855
   End
   Begin VB.Label lblUserID 
      AutoSize        =   -1  'True
      Caption         =   "User ID     : "
      Height          =   195
      Left            =   960
      TabIndex        =   12
      Top             =   360
      Width           =   855
   End
   Begin VB.Label lblUserName 
      AutoSize        =   -1  'True
      Caption         =   "User Name: Logged Off"
      Height          =   195
      Left            =   960
      TabIndex        =   11
      Top             =   120
      Width           =   1680
   End
   Begin VB.Image imgHelp 
      Height          =   360
      Left            =   4560
      MouseIcon       =   "frmLogIn.frx":2668
      MousePointer    =   99  'Custom
      Picture         =   "frmLogIn.frx":2F32
      Top             =   120
      Width           =   360
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H000000C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
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
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   1320
      Width           =   885
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User ID"
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
      Left            =   360
      TabIndex        =   4
      Top             =   960
      Width           =   660
   End
End
Attribute VB_Name = "frmLogIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rs      As Recordset
Private dbPass  As String    '//Database Password container
Private iTRY    As Integer

Private Sub CboUser_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
 txtname.text = CboUser.text
 CboUser.Visible = False
 txtPass.SetFocus
ElseIf KeyCode = 27 Then
 CboUser.Visible = False
 txtPass.SetFocus
End If
End Sub
Private Sub chkAdmin_Click()
  txtname.SetFocus
End Sub

Private Sub LockLogIN()
 CmdLogIn.Enabled = False
End Sub



Private Sub cmdLogin_Click()
On Error GoTo ERRORHANDLE
If iTRY = 3 Then
: Rem lblTry.Caption = "No more try! Exit now!"
   Exit Sub
End If
 iTRY = iTRY + 1

If txtname.text = "" Or txtPass.text = "" Then
    txtname.SetFocus
End If
If rs.State = adStateOpen Then
  rs.Close
End If
rs.Open "select * from users ", Con, 1, 1, 1
rs.MoveFirst
Do Until rs.EOF
   If rs.Fields("userID") = txtname.text And _
      rs.Fields("password") = RedAlert.xPASS Then
       CurrUser.user_id = rs.Fields("UserID")
      CurrUser.USER_PASS = rs.Fields("password")
      CurrUser.USER_NAME = rs.Fields("FIRSTNAME") & " " & rs.Fields("LASTNAME")
      CurrUser.USER_isADMIN = RedAlert.xADMIN
      iTRY = 0
      GoTo iFound
   Else
      rs.MoveNext
   End If
   Loop
 
   txtPass.text = Empty
   txtname.SetFocus
If iTRY = 1 Then
Rem lblTry.Caption = "Invalid Username or pasword! Try again"
ElseIf iTRY = 2 Then
Rem lblTry.Caption = "Invalid Username or pasword! Try again"
ElseIf iTRY = 3 Then
Rem lblTry.Caption = "Invalid Username or pasword!"
   End If
   
   Exit Sub

iFound:

With lblUserName
  .Caption = "User Name: "
  .Caption = .Caption & CurrUser.USER_NAME
End With
With lblUserID
  .Caption = "UserID:          "
  .Caption = .Caption & CurrUser.user_id
End With
'// determine access
Menu_List rs, CurrUser.USER_PASS

txtname.text = Empty
txtPass = Empty

LockLogIN

ERRORHANDLE:
 errorMsg Err, Me.Name, "CmdLogin"
End Sub

Private Sub Form_Load()
'[====================]
'<database password   >
'[====================]
RedAlert.xEncrypt RedAlert.xdbPassword
dbPass = RedAlert.xPASS
'[============================]
'< open connection user menu  >
'[============================]
Call OpenDB("USERS.MDB", Con, True, dbPass)
Set rs = New ADODB.Recordset
rs.Open "select * from users order by userID, menu ", Con, 1, 1, 1
Call Add_Item(rs, "userID", CboUser)
show
txtname.SetFocus

End Sub

Private Sub Form_Resize()
With frmLogIn
  If .WindowState = 0 Then
   .Height = 2595
   .Width = 5235
  End If
End With
SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub

Private Sub Form_Unload(Cancel As Integer)
  rs.Close
  Set rs = Nothing
  Set frmLogIn = Nothing
End Sub

Private Sub imgHelp_Click()
  myMsg "Check Admin if you want to log-in " _
  & "as admin." & vbCrLf & vbCrLf _
  & "Login button will be locked upon login correctly. " _
  & "To enable Login button type your User ID.", "Password Help", 1, True
End Sub

Private Sub txtname_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then
    txtPass.SetFocus
  End If
  If KeyCode = 40 Then  'down arrow
     AlignObj txtname, CboUser, 1
     CboUser.SetFocus
     CboUser.ListIndex = 0
  End If
  If KeyCode = 27 Then
    CboUser.Visible = False
  End If
End Sub

Private Sub txtname_Change()
 CboUser.ListIndex = SendMessage(CboUser.hWnd, LB_FINDSTRING, -1, ByVal txtname.text)
  If UCase$(TrimSpaces(txtname)) = UCase$(CboUser.text) Then
     CmdLogIn.Enabled = True
     If iTRY = 3 Then
        iTRY = 0
: Rem       lblTry.Caption = "Access granted!"
     End If
  End If
End Sub
Private Sub txtPass_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then
   If CmdLogIn.Enabled = True Then
     CmdLogIn.SetFocus
   End If
  ElseIf KeyCode = 38 Then
   txtname.SetFocus
  End If
  If KeyCode = 27 Then
    CboUser.Visible = False
  End If
End Sub

Private Sub txtPass_Change()
On Error GoTo errMsg
Dim isAdmin As Integer
If chkAdmin.Value = 0 Then
   isAdmin = 0
Else
   isAdmin = -1
End If
 RedAlert.xText = txtPass.text
 RedAlert.xEncrypt RedAlert.xText, isAdmin
errMsg:
 errorMsg Err, Me.Name, "TxtPass_change"
End Sub

