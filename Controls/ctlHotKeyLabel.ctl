VERSION 5.00
Begin VB.UserControl ctlHotKeyLabel 
   ClientHeight    =   270
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1950
   ScaleHeight     =   270
   ScaleWidth      =   1950
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Action"
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
      Height          =   240
      Left            =   600
      MouseIcon       =   "ctlHotKeyLabel.ctx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   0
      Width           =   555
   End
   Begin VB.Label lblHotKey 
      AutoSize        =   -1  'True
      Caption         =   "KeyIn"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   495
   End
End
Attribute VB_Name = "ctlHotKeyLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//UNFINISHED CONTROL
'// DESIGNED BY EDWIN DELOS SANTOS
Private PixelX, PixelY
Private ky As Long  'key in width
Private ac As Long  'action key width

'Event Declarations:
Event Click() 'MappingInfo=Label1,Label1,-1,Click
Event Change()
Event Resize()


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1,Label1,-1,Caption
Public Property Get Caption() As String
    Caption = Label1.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Label1.Caption() = New_Caption
    PropertyChanged "Caption"
End Property

Public Property Get CaptionHotkey() As String
    CaptionHotkey = lblHotKey.Caption
End Property

Public Property Let CaptionHotkey(ByVal New_Caption As String)
    lblHotKey.Caption() = New_Caption
    PropertyChanged "CaptionHotkey"
End Property

Private Sub Label1_Click()
   RaiseEvent Click
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1,Label1,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = Label1.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    Label1.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

Public Property Get ForeColorHotKey() As OLE_COLOR
    ForeColorHotKey = lblHotKey.ForeColor
End Property

Public Property Let ForeColorHotKey(ByVal New_ForeColor As OLE_COLOR)
    lblHotKey.ForeColor() = New_ForeColor
    PropertyChanged "ForeColorHotkey"
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = Label1.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    Label1.BackColor() = New_BackColor
    UserControl.BackColor = Label1.BackColor
    lblHotKey.BackColor = Label1.BackColor
    PropertyChanged "BackColor"
    
End Property


'

Public Sub UserControl_Initialize()
  PixelX = Screen.TwipsPerPixelX
 UserControl_Resize
End Sub

Private Sub UserControl_Resize()

lblHotKey.Left = 0
lblHotKey.Height = UserControl.ScaleHeight
Label1.Width = UserControl.ScaleWidth
Label1.Height = UserControl.ScaleHeight
Label1.Left = lblHotKey.Width + PixelX

End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Label1.Caption = PropBag.ReadProperty("Caption", Label1.Caption)
    lblHotKey.Caption = PropBag.ReadProperty("CaptionHotkey", lblHotKey.Caption)
    Label1.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    Label1.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    lblHotKey.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    lblHotKey.ForeColor = PropBag.ReadProperty("ForeColorHotKey", &H0&)
    Set Font = PropBag.ReadProperty("Font", Ambient.Font)
    Set Label1.Font = PropBag.ReadProperty("Font", Ambient.Font)
    Set lblHotKey.Font = PropBag.ReadProperty("Font", Ambient.Font)
    lblHotKey.Width = PropBag.ReadProperty("WidthHotkey", 495)
    Label1.Left = PropBag.ReadProperty("Left", 495)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Caption", Label1.Caption, &H80000012)
    Call PropBag.WriteProperty("CaptionHotkey", lblHotKey.Caption, &H80000012)
    Call PropBag.WriteProperty("ForeColor", Label1.ForeColor, &H80000012)
    Call PropBag.WriteProperty("BackColor", Label1.BackColor, &H8000000F)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("BackColor", lblHotKey.BackColor, &H8000000F)
    Call PropBag.WriteProperty("ForeColorHotKey", lblHotKey.ForeColor, &H0&)
'    Call PropBag.WriteProperty("BackColorHotkey", lblHotKey.BackColor, &H8000000F)
    Call PropBag.WriteProperty("Font", Font, Ambient.Font)
    Call PropBag.WriteProperty("Font", Label1.Font, Ambient.Font)
    Call PropBag.WriteProperty("Font", lblHotKey.Font, Ambient.Font)
    Call PropBag.WriteProperty("WidthHotkey", lblHotKey.Width, 495)
    Call PropBag.WriteProperty("Left", Label1.Left, 495)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1,Label1,-1,BackColor
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1,Label1,-1,Font

Public Property Get Font() As Font
    Set Font = Label1.Font
    Set Font = lblHotKey.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set Label1.Font = New_Font
    Set lblHotKey.Font = New_Font
    PropertyChanged "Font"
End Property

Public Property Get WidthHotkey() As Long
    WidthHotkey = lblHotKey.Width
End Property

Public Property Let WidthHotkey(ByVal New_Width As Long)
    lblHotKey.Width() = New_Width
    PropertyChanged "WidthHotkey"
End Property

Public Property Get Left() As Long
    Left = Label1.Left
End Property

Public Property Let Left(ByVal New_Left As Long)
    Label1.Left() = New_Left
    PropertyChanged "Left"
End Property






