VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmCalcu 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   Caption         =   "Edwin's Calculator"
   ClientHeight    =   5025
   ClientLeft      =   165
   ClientTop       =   255
   ClientWidth     =   4110
   FillColor       =   &H00C0C0C0&
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00E0E0E0&
   Icon            =   "FrmCalcu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   4110
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton Picture2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   120
      MouseIcon       =   "FrmCalcu.frx":0FA2
      Picture         =   "FrmCalcu.frx":186C
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   240
      Width           =   615
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Authocopy"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3120
      TabIndex        =   24
      Top             =   0
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2295
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   4048
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "DETAILS"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Object.Width           =   4233
      EndProperty
   End
   Begin VB.CommandButton CmdPercent 
      BackColor       =   &H00E0E0E0&
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   3960
      Width           =   615
   End
   Begin VB.CommandButton cmdBackSpace 
      Caption         =   "ç"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   3240
      TabIndex        =   21
      ToolTipText     =   "Delete last digit"
      Top             =   3000
      Width           =   615
   End
   Begin VB.CommandButton digits1 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   240
      TabIndex        =   7
      Top             =   3960
      Width           =   615
   End
   Begin VB.CommandButton equals 
      BackColor       =   &H00E0E0E0&
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   4440
      Width           =   615
   End
   Begin VB.CommandButton clearbttn 
      BackColor       =   &H00E0E0E0&
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   4440
      Width           =   615
   End
   Begin VB.CommandButton dotbttn 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1680
      TabIndex        =   11
      Top             =   4440
      Width           =   615
   End
   Begin VB.CommandButton div 
      BackColor       =   &H00E0E0E0&
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   3000
      Width           =   615
   End
   Begin VB.CommandButton times 
      BackColor       =   &H00E0E0E0&
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3480
      Width           =   615
   End
   Begin VB.CommandButton minus 
      BackColor       =   &H00E0E0E0&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3960
      Width           =   615
   End
   Begin VB.CommandButton plus 
      BackColor       =   &H00E0E0E0&
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4440
      Width           =   615
   End
   Begin VB.CommandButton over 
      BackColor       =   &H00E0E0E0&
      Caption         =   "1/x"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4680
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton plusminus 
      BackColor       =   &H00E0E0E0&
      Caption         =   "+/-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3480
      Width           =   615
   End
   Begin VB.CommandButton digits1 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   960
      TabIndex        =   10
      Top             =   4440
      Width           =   615
   End
   Begin VB.CommandButton digits1 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   1680
      TabIndex        =   6
      Top             =   3480
      Width           =   615
   End
   Begin VB.CommandButton digits1 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   1680
      TabIndex        =   3
      Top             =   3000
      Width           =   615
   End
   Begin VB.CommandButton digits1 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   1680
      TabIndex        =   9
      Top             =   3960
      Width           =   615
   End
   Begin VB.CommandButton digits1 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   960
      TabIndex        =   8
      Top             =   3960
      Width           =   615
   End
   Begin VB.CommandButton digits1 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   960
      TabIndex        =   5
      Top             =   3480
      Width           =   615
   End
   Begin VB.CommandButton digits1 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   960
      TabIndex        =   2
      Top             =   3000
      Width           =   615
   End
   Begin VB.CommandButton digits1 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   240
      TabIndex        =   4
      Top             =   3480
      Width           =   615
   End
   Begin VB.CommandButton digits1 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   3000
      Width           =   615
   End
   Begin InstantReport.Hline Hline1 
      Height          =   30
      Left            =   0
      TabIndex        =   27
      Top             =   1080
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   53
   End
   Begin InstantReport.Hline Hline2 
      Height          =   30
      Left            =   0
      TabIndex        =   28
      Top             =   2040
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   53
   End
   Begin VB.Label display 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   840
      TabIndex        =   20
      Top             =   2520
      Width           =   3015
   End
   Begin VB.Label lblOperator 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   330
      Left            =   360
      TabIndex        =   23
      Top             =   1680
      Width           =   90
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Current Operation"
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      TabIndex        =   26
      Top             =   1200
      Width           =   840
   End
End
Attribute VB_Name = "FrmCalcu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'//original coding by :  Edwin delos santos
Option Explicit
Private operand1 As Double, operand2 As Double
Private operator As String
Private cleardisplay As Boolean
Private dispnum As Integer        'COUNTER FOR DISPLAY NUMBER
Private addnum As Boolean         'HANDLE CONTINOUS ADDITION
Private newnum As Boolean         'NEW NUMBER IS ENTERED
Private memori As Double          'STORE NUMBER ENTERED
Private fmat As Boolean           'HANDLE FORMAT

Private Sub capi_Click()
   Clipboard.Clear
    Clipboard.SetText Val(display.Caption)
End Sub

Private Sub clearbttn_Click()
  display = 0
  memori = 0
  operand1 = 0
  lblOperator.Caption = ""
End Sub

Private Sub CmdPercent_Click()
display = Val(display) / 100
ListView1.SetFocus
lblOperator.Caption = "%"
equals_Click
End Sub


Private Sub DISPLAY_Change()
If display = "." Then
   display = "0."
   Exit Sub
End If
'//autocopy
If Check1.Value Then
   Clipboard.Clear
   Clipboard.SetText display.Caption, vbCFText
End If
End Sub

Private Sub DISPLAY_KeyPress(KeyAscii As Integer)
Dim strValid As String
    '
    strValid = "0123456789+-/*."
    '
    If KeyAscii > 26 Then ' if it's not a control code
        If InStr(strValid, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If

If KeyAscii = 43 Then ' +
  plus_click
ElseIf KeyAscii = 45 Then ' -
 minus_click
ElseIf KeyAscii = 42 Then ' *
 times_click
ElseIf KeyAscii = 47 Then ' /
 div_click
ElseIf KeyAscii = 13 Then ' = or ENTER
 equals_Click
End If
End Sub

Private Sub display_Click()
 display = Format(memori, "fixed")
 fmat = False
End Sub


Private Sub dotbttn_Click()
If InStr(display, ".") Then
 Exit Sub
Else
 display = display + "."
End If
End Sub

Private Sub equals_Click()
On Error GoTo errhdle:
Dim result As Double
addnum = False
newnum = False
operand2 = Val(display)

If operator = "+" Then result = operand1 + operand2
If operator = "-" Then result = operand1 - operand2
If operator = "*" Then result = operand1 * operand2
If operator = "/" And operand1 <> "0" Then result = operand1 / operand2
If Val(operand2) <> 0 Then
  Lvdisplay
End If
memori = result
display = result
'/RESET THE VALUE
operand1 = 0
operand2 = 0
'/============

operator = "="
Lvdisplay

display = Format(result, "standard")
fmat = True
'/==============
dispnum = 0
operator = ""
ListView1.SetFocus
errhdle:
 Exit Sub
End Sub


Private Sub digits1_Click(Index As Integer)
On Error Resume Next
If newnum = False Then
  display = ""
  newnum = True
End If
If Len(display) >= 12 Then
   Exit Sub
Else
  display = display + digits1(Index).Caption
End If
memori = toMoney(display)
ListView1.SetFocus
End Sub



Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode >= vbKeyNumpad0 And KeyCode <= vbKeyNumpad9 Then
      digits1_Click (Chr(KeyCode - 48))
    ElseIf KeyCode = 13 Then
      If addnum = False Then
        If Val(display) = 0 Then Exit Sub
      End If
      equals_Click
    ElseIf KeyCode = 107 Then
        plus_click
    ElseIf KeyCode = vbKeyBack Then
        cmdBackSpace_Click
    ElseIf KeyCode = 109 Then
        minus_click
    ElseIf KeyCode = 106 Then
        times_click
    ElseIf KeyCode = 111 Then
        div_click
    ElseIf KeyCode = 110 Then
       dotbttn_Click
    ElseIf Shift = 1 And KeyCode = 53 Then
       CmdPercent_Click
    End If
End Sub

Private Sub Form_Load()
    newnum = True
    addnum = True
    lblOperator.Caption = ""
    lblOperator.ZOrder
End Sub



Private Sub Form_LostFocus()
If Me.WindowState = 0 Then
   Me.WindowState = 1
End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'If MsgBox("Exit?", vbYesNo + vbQuestion, "My calculator") = vbNo Then
'  Cancel = 1 'true
'Else
'  Unload Me
'End If

End Sub

Private Sub Form_Resize()
With FrmCalcu
  If .WindowState = 0 Then
   .Height = 5535
   .Width = 4290
  End If
End With
SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub


Private Sub Form_Unload(Cancel As Integer)
Set FrmCalcu = Nothing
Unload Me
End Sub

Private Sub ListView1_DblClick()
   Dim x
   If ListView1.ListItems.Count <> 0 Then
    x = ListView1.SelectedItem.ListSubItems(1).text
     If Val(x) > 0 Then
       display = toMoney(ListView1.SelectedItem.ListSubItems(1).text)
       memori = toMoney(ListView1.SelectedItem.ListSubItems(1).text)
    End If
   End If
End Sub



Private Sub minus_click()
If fmat = True Then
 display = memori
 fmat = False
End If

If dispnum = 0 And Val(display) = 0 Then Exit Sub
dispnum = dispnum + 1
If dispnum >= 1 Then
   newnum = True
End If

If dispnum = 2 Then
 equals_Click
Else
 operand1 = Val(display)
 Lvdisplay
 operator = "-"
 display = ""
 lblOperator.Caption = operator
End If
End Sub



Private Sub Picture2_Click()
 ListView1.ListItems.Clear
 clearbttn_Click
 lblOperator.Caption = ""
End Sub

Private Sub plus_click()
If fmat = True Then
  display = memori
  fmat = False
End If
If dispnum >= 1 Then
   newnum = True
End If
If Val(display) = 0 Then
   Exit Sub
Else
   addnum = True
End If
If addnum = True Then
  Lvdisplay
  operand1 = operand1 + Val(display)
Else
  operand1 = Val(display)
End If
operator = "+"
 lblOperator.Caption = operator
'/================

If addnum = False Then
  Lvdisplay
End If
 display = ""
End Sub

Private Sub div_click()

If fmat = True Then
 display = memori
 fmat = False
End If

If dispnum = 0 And Val(display) = 0 Then Exit Sub
If dispnum >= 1 Then
   newnum = True
End If

'dispnum = dispnum + 1

If dispnum = 2 Then
 equals_Click
Else
 operand1 = Val(display)
 Lvdisplay
 operator = "/"
 display = ""
 lblOperator.Caption = operator
End If
End Sub

Private Sub plusminus_click()
display = -Val(display)
End Sub

Private Sub times_click()
If fmat = True Then
 display = memori
 fmat = False
End If

If dispnum = 0 And Val(display) = 0 Then Exit Sub
dispnum = dispnum + 1
If dispnum >= 1 Then
   newnum = True
End If

If dispnum = 2 Then
 equals_Click
Else
 operand1 = Val(display)
 Lvdisplay
 operator = "*"
 display = ""
 lblOperator.Caption = operator
End If
End Sub
Private Sub over_click()
If Val(display) <> 0 Then display = 1 / Val(display)
End Sub
Private Sub cmdBackSpace_Click()
    If Len(display.Caption) = 0 Then
        Exit Sub
    Else
        With display
            .Caption = Left(.Caption, Len(.Caption) - 1)
        End With
    End If

End Sub


Sub Lvdisplay()
Dim lstmain1 As ListItem
 Set lstmain1 = ListView1.ListItems.Add(, , operator)
     lstmain1.ForeColor = vbBlue
     lstmain1.SubItems(1) = Format(display.Caption, "Standard")
     If operator = "=" Then
      LvLine
      lstmain1.ListSubItems(1).ForeColor = vbRed
     Else
       lstmain1.ListSubItems(1).ForeColor = vbBlack
     End If
     
End Sub
Sub LvLine()
Dim lstmain1 As ListItem
 Set lstmain1 = ListView1.ListItems.Add(, , " ")
     lstmain1.Bold = False
     lstmain1.SubItems(1) = Format("---------------------------")
     lstmain1.Bold = False

End Sub
