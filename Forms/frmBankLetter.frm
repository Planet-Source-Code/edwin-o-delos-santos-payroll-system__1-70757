VERSION 5.00
Begin VB.Form frmBankLetter 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Bank Letter"
   ClientHeight    =   7500
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10770
   LinkTopic       =   "Form1"
   ScaleHeight     =   7500
   ScaleWidth      =   10770
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPreview 
      Caption         =   "&Preview"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   9480
      Picture         =   "frmBankLetter.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5520
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   9480
      Picture         =   "frmBankLetter.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6480
      Width           =   975
   End
   Begin VB.PictureBox PicDocument 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7335
      Left            =   0
      ScaleHeight     =   7335
      ScaleWidth      =   7815
      TabIndex        =   9
      Top             =   0
      Width           =   7815
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Text            =   "Attention : <B> LITO OBLEFIAS MDC </B> - HUGAS PUWET"
         Top             =   3240
         Width           =   7215
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         TabIndex        =   17
         Text            =   "FEBRUARY 2008"
         Top             =   3840
         Width           =   2655
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         TabIndex        =   16
         Text            =   "Combo1"
         Top             =   3840
         Width           =   2535
      End
      Begin VB.TextBox Text2 
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
         Height          =   1215
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   14
         Top             =   4560
         Width           =   7455
      End
      Begin VB.TextBox Text1 
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
         Height          =   1215
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Top             =   1320
         Width           =   7455
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Very Truly yours,"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   15
         Top             =   6840
         Width           =   1500
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         Caption         =   "date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   12
         Top             =   120
         Width           =   405
      End
   End
   Begin VB.Frame fraBorder 
      Caption         =   "Borders"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   7920
      TabIndex        =   6
      Top             =   1680
      Width           =   2655
      Begin VB.CheckBox chkBorder 
         Appearance      =   0  'Flat
         Caption         =   "Show Input borders"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.Label lblBorder 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Displays borders around Textbox/List/Combo items."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   2415
      End
   End
   Begin VB.Frame fraAlign 
      Caption         =   "Form Alignment"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   7920
      TabIndex        =   0
      Top             =   120
      Width           =   2655
      Begin VB.OptionButton optAlign 
         Appearance      =   0  'Flat
         Caption         =   "Left"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton optAlign 
         Appearance      =   0  'Flat
         Caption         =   "Centre"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   855
      End
      Begin VB.OptionButton optAlign 
         Appearance      =   0  'Flat
         Caption         =   "Right"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   1320
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton optAlign 
         Appearance      =   0  'Flat
         Caption         =   "Full"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   1320
         TabIndex        =   1
         Top             =   480
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.Label lblAlign 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "How the form will be shown horizontally on the page."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   2415
      End
   End
End
Attribute VB_Name = "frmBankLetter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim qPrint As qcPrinter
Dim eAlignment As qePrinterAlign

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdPreview_Click()
'Update text from the Form contents
qPrint.FormPrint_Update Me
qPrint.Preview

End Sub

Private Sub Form_Load()
'//Set up the information on the form
lblDate.Caption = Now()
Combo1.AddItem "MONTHLY PAYROLL"
Combo1.AddItem "SEMI-MONTHLY PAYROLL"
Combo1.AddItem "WEEKLY PAYROLL"
Combo1.AddItem "DAILY PAYROLL"
Combo1.ListIndex = 1
Text1.text = "<B>BANK OF THE PHILIPPINE ISLAND </B>" & vbCrLf _
 & "Madrigal Business Park Branch" & vbCrLf _
 & "777 Madrigal Business Park I" & vbCrLf _
 & "Acacia Ave., corner Commerce Avenue," & vbCrLf

Text2.text = "Please debit our BPI CA #17777-77131-77 the amount of PESOS the amount of " _
 & FrmPayroll.lblAmtInWord & Space(1) & FrmPayroll.lblnetpay _
 & " and credit the to individual payroll account (please see attached sheet)"

eAlignment = eJustify
'//Initialise the qPrint object in Form_Load - updates are much quicker
Set qPrint = New qcPrinter
qPrint.MarginTop = 567
qPrint.MarginLeft = 567
qPrint.MarginRight = 567
' FormPrint parameters:
' frmPrint As Object: 'Me'
'     The form holding the controls
' Optional ParentContain As Object: 'picDocument'
'     The container holding the controls to be printed
'     The scalewidth of the container is used for positioning
'     if no container is specified, the Form (Me) is used
' Optional FormAlign As qePrinterAlign = eLeft: 'eAlignment'
'     How to display the form. Left/Right/Centre/Justify
'     'Jusitfy' will stretch the controls to fit the width of the page
' Optional InputBorder As Boolean = False: 'True'
'     Will display a border equivalent to the border of TextBox,
'     ListBox or ComboBox.  Labels, CheckBoxes and OptionButtons are
'     Not given a border.
' Optional TopOffset As Single: 800
'     Distance from Top of page to Offset the Control contents
'     In this instance to allow for the 'docTitle' added afterwards
' Optional AutoHeight As Boolean: True
'     Where the contents of an item are taller than the bounding
'     control, qPrinter will adjust the height and position of
'     subsequent items.
' Optional ExcludeList As String: "*chkExpand"
'     List of controls prefixed by * to identify them
'     to be excluded from the document.  If a container control
'     eg. PictureBox/Frame is included, all the controls in the
'     container will be excluded.  In this instance 'chkExpand'
'     is excluded because it is held in 'picDocument' - the container
'     we want to print.

qPrint.FormPrint Me, PicDocument, eAlignment, True, 1000, True, "*chkExpand"
' Add the title
'qPrint.AddText "qbd software ltd" & vbCrLf & "qPrinter:FormPrint example", "Verdana", 18, , , , , eCentre, , , "docTitle"


End Sub
Private Sub Form_Unload(Cancel As Integer)
'Destroy qPrint object
Set qPrint = Nothing

End Sub

Private Sub optAlign_Click(Index As Integer)

If Index <> eAlignment Then
eAlignment = Index
' Change the Alignment of the document
qPrint.FormPrint Me, PicDocument, eAlignment, CBool(chkBorder.Value = vbChecked), 1000, True, "*chkExpand"
' Add the title
qPrint.AddText "qbd software ltd" & vbCrLf & "qPrinter:FormPrint example", "Verdana", 18, , , , , eCentre, , , "docTitle"
'If eAlignment = eLeft Then
'qPrint.TextItem("docTitle").IndentRight = qPrint.TextItem("lstInfo").IndentRight
'ElseIf eAlignment = eRight Then
'qPrint.TextItem("docTitle").IndentLeft = qPrint.TextItem("lblInfo/0").IndentLeft
'End If
End If

End Sub

Private Sub chkBorder_Click()
' Change the borders of the document
qPrint.FormPrint Me, PicDocument, eAlignment, CBool(chkBorder.Value = vbChecked), 1000, True, "*chkExpand"
' Add the title
'qPrint.AddText "qbd software ltd" & vbCrLf & "qPrinter:FormPrint example", "Verdana", 18, , , , , eCentre, , , "docTitle"
If eAlignment = eLeft Then
qPrint.TextItem("docTitle").IndentRight = qPrint.TextItem("lstInfo").IndentRight
ElseIf eAlignment = eRight Then
qPrint.TextItem("docTitle").IndentLeft = qPrint.TextItem("lblInfo/0").IndentLeft
End If

End Sub


