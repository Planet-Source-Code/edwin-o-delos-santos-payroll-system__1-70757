VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmlogin 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Log In User"
   ClientHeight    =   2850
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   4965
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "Password.frx":0000
   LinkTopic       =   "Form10"
   MaxButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   4965
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   720
      TabIndex        =   10
      Top             =   1920
      Width           =   3975
      Begin VB.CommandButton CmdLogIn 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&LogIn"
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
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Cancel"
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
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   120
         Width           =   1455
      End
      Begin VB.PictureBox HotKey2 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   240
         Picture         =   "Password.frx":0FA2
         ScaleHeight     =   285
         ScaleWidth      =   90
         TabIndex        =   12
         Top             =   120
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
         Left            =   0
         Picture         =   "Password.frx":11B4
         ScaleHeight     =   285
         ScaleWidth      =   90
         TabIndex        =   11
         Top             =   120
         Visible         =   0   'False
         Width           =   90
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   30
      Left            =   240
      Picture         =   "Password.frx":13C6
      ScaleHeight     =   30
      ScaleWidth      =   1500
      TabIndex        =   9
      Top             =   2880
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   240
      Picture         =   "Password.frx":1688
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   7
      Top             =   120
      Width           =   510
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   6
      Top             =   2565
      Width           =   4965
      _ExtentX        =   8758
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   5662
            Text            =   "Use down arrow key to view userlist! "
            TextSave        =   "Use down arrow key to view userlist! "
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            TextSave        =   "CAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ListBox CboUser 
      Appearance      =   0  'Flat
      BackColor       =   &H00D38545&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   990
      Left            =   2040
      TabIndex        =   5
      Top             =   840
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.TextBox TextPass 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   0
      Locked          =   -1  'True
      MousePointer    =   99  'Custom
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   22920
      Visible         =   0   'False
      Width           =   1170
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
      Left            =   2040
      MaxLength       =   15
      MousePointer    =   99  'Custom
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1200
      Width           =   2535
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
      Left            =   2040
      Locked          =   -1  'True
      MousePointer    =   99  'Custom
      TabIndex        =   0
      ToolTipText     =   "Arrow key to view user list!"
      Top             =   840
      Width           =   2535
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Select your user name from the list !"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1080
      TabIndex        =   8
      Top             =   240
      Width           =   2505
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User name:"
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
      Left            =   840
      TabIndex        =   2
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
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
      Left            =   960
      TabIndex        =   1
      Top             =   1200
      Width           =   885
   End
End
Attribute VB_Name = "frmlogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rs As ADODB.Recordset
Attribute rs.VB_VarHelpID = -1
Private iTRY As Integer


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


Private Sub CmdCancel_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
 Call Btn_Focus(cmdCancel, Hotkey, HotKey2)
End Sub

Private Sub cmdLogin_Click()

If iTRY = 3 Then
   StatusBar1.Panels(1).text = "No more try! Exit now!"
   cmdCancel.SetFocus
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
   If rs.Fields("username") = txtname.text And _
      rs.Fields("password") = TextPass.text Then
      CurrUser.USER_ID = rs.Fields("USERNAME")
      CurrUser.USER_PASS = rs.Fields("password")
      CurrUser.USER_NAME = rs.Fields("FIRSTNAME") & " " & rs.Fields("LASTNAME")
      iTRY = 0
      GoTo ifound
   Else
      rs.MoveNext
   End If
   Loop
 
   txtPass.text = Empty
   txtname.SetFocus
If iTRY = 1 Then
     StatusBar1.Panels(1).text = "Invalid Username or pasword! 2nd Try"
ElseIf iTRY = 2 Then
     StatusBar1.Panels(1).text = "Invalid Username or pasword! Last Try"

ElseIf iTRY = 3 Then
     StatusBar1.Panels(1).text = "Invalid Username or pasword!"
     cmdCancel.SetFocus
   End If
   
   Exit Sub

ifound:
With MainForm.lblUserName
  .Caption = "User Name: "
  .Caption = .Caption & CurrUser.USER_NAME
End With
With MainForm.lblUserID
  .Caption = "User ID: "
  .Caption = .Caption & CurrUser.USER_ID
End With


Menu_List rs, CurrUser.USER_PASS

txtname.text = Empty
txtPass = Empty

cmdCancel_Click

End Sub
Private Sub cmdCancel_Click()
  rs.Close
  Set rs = Nothing
  Unload Me
End Sub



Private Sub CmdLogIn_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
 Call Btn_Focus(CmdLogin, Hotkey, HotKey2)
End Sub

Private Sub Form_Load()
CboUser.ZOrder
Set Con = New ADODB.Connection
Set rs = New ADODB.Recordset

Con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=false;Data Source= " & App.Path & "\DB\Users.mdb;Jet OLEDB:Database Password=¶£®«²£°"
rs.Open "select * from users order by username, menu ", Con, 1, 1, 1

Call iList(rs, "username", CboUser)

txtname.text = Empty
txtPass.text = Empty


End Sub
Private Sub text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdLogin_Click
End Sub


Private Sub Form_Resize()
With frmlogin
  If .WindowState = 0 Then
   .Height = 3360
   .Width = 5085
  End If
 End With

SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub

Private Sub Form_Unload(Cancel As Integer)

Set frmlogin = Nothing
End Sub


Private Sub txtName_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then
    txtPass.SetFocus
  End If
  If KeyCode = 40 Then
     CboUser.Visible = True
     CboUser.SetFocus
     CboUser.ListIndex = 0
  End If
  If KeyCode = 27 Then
    CboUser.Visible = False
  End If
End Sub

Private Sub txtPass_Change()
 encrypt txtPass
 TextPass.text = xPASS
End Sub

Private Sub txtpass_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then
   CmdLogin.SetFocus
  ElseIf KeyCode = 38 Then
   txtname.SetFocus
  End If
    If KeyCode = 40 Then
     CboUser.Visible = True
     CboUser.SetFocus
     CboUser.ListIndex = 0
  End If
  If KeyCode = 27 Then
    CboUser.Visible = False
  End If
End Sub

Private Sub TxtPass_KeyPress(KeyAscii As Integer)
'    Dim strValid As String
'    strValid = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"
'    KeyAscii = Asc(Chr(KeyAscii))
'    If KeyAscii > 26 Then ' if it's not a control code
'        If InStr(strValid, Chr(KeyAscii)) = 0 Then
'            KeyAscii = 0
'        Else
'          TxtPASS.ToolTipText = " Numeric not allowed! "
'        End If
'    End If
End Sub
