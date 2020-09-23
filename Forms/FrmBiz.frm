VERSION 5.00
Begin VB.Form FrmBiz 
   BackColor       =   &H00D37527&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Biz Information"
   ClientHeight    =   2070
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   5265
   Icon            =   "FrmBiz.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmBiz.frx":0ECA
   ScaleHeight     =   2070
   ScaleWidth      =   5265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdCancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cancel"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton CmdUpdate 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Update"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1440
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   960
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1320
      TabIndex        =   7
      Top             =   240
      Width           =   3735
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1320
      TabIndex        =   6
      Top             =   600
      Width           =   3735
   End
   Begin VB.CommandButton CmdEdit 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Edit"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "CONTACT:"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   810
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ADDRESS:"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   825
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "COMPANY:"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   840
   End
End
Attribute VB_Name = "FrmBiz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsBiz As ADODB.Recordset
Attribute rsBiz.VB_VarHelpID = -1



Private Sub CmdCancel_Click()
  CmdEdit.Enabled = True
  CmdUpdate.Enabled = False
  CmdCancel.Enabled = False
End Sub

Private Sub CmdEdit_Click()
CmdEdit.Enabled = False
CmdUpdate.Enabled = True
CmdCancel.Enabled = True
End Sub

Private Sub CmdUpdate_Click()
With rsBiz
  .Fields(0) = Text1
  .Fields(1) = Text2
  .Fields(2) = Text3
  rsBiz.update
End With
  MAIN.lblCOMPANY = rsBiz!company
  MAIN.lblADDRESS = rsBiz!address
  MAIN.lblCONTACT = rsBiz!contact
  CmdEdit.Enabled = True
  CmdUpdate.Enabled = False
  CmdCancel.Enabled = False

End Sub




Private Sub Form_Load()
OpenDB
Set rsBiz = New ADODB.Recordset
rsBiz.Open "SELECT * From BIZ", CN, adOpenStatic, adLockOptimistic
RefBindData
End Sub
Sub RefBindData()
  If rsBiz.EOF = True Then Exit Sub
  If rsBiz.BOF = True Then Exit Sub
On Error Resume Next
  Text1.text = rsBiz!company
  Text2.text = rsBiz!address
  Text3.text = rsBiz!contact
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'On Error Resume Next
'If MsgBox("Close this system?", vbYesNo + vbQuestion, "Biz Information") = vbNo Then
'  Cancel = 1 'True
'Else
'  rsBiz.Close
'  Set rsBiz = Nothing
'  Unload Me
'End If

End Sub

Private Sub Form_Resize()

SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub

Private Sub Form_Unload(Cancel As Integer)
  rsBiz.Close
  Set rsBiz = Nothing
  Set FrmBiz = Nothing
  Unload Me
End Sub

