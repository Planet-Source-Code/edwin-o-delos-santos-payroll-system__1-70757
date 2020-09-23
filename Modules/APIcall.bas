Attribute VB_Name = "APIcall"
Option Explicit

'//use for searhING item from listbox
Public Declare Function SendMessage Lib "user32" _
Alias "SendMessageA" (ByVal hwnd As Long, ByVal _
wMsg As Integer, ByVal wParam As Integer, lParam _
As Any) As Long
Public Const LB_FINDSTRING = &H18F

'// FLAT HEADERS

Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" _
   (ByVal hwnd As Long, _
    ByVal Msg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long) As Long

Public Declare Function GetWindowLong Lib "user32" _
   Alias "GetWindowLongA" _
   (ByVal hwnd As Long, _
   ByVal nIndex As Long) As Long
   
Public Declare Function SetWindowLong Lib "user32" _
   Alias "SetWindowLongA" _
   (ByVal hwnd As Long, _
   ByVal nIndex As Long, _
   ByVal dwNewLong As Long) As Long
   
Public Declare Function SetWindowPos Lib "user32" _
   (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
    ByVal x As Long, ByVal Y As Long, _
    ByVal cx As Long, ByVal cy As Long, _
    ByVal wFlags As Long) As Long

Public Const GWL_STYLE = (-16)
Public Const LVM_FIRST = &H1000
Public Const LVM_GETHEADER = (LVM_FIRST + 31)
Public Const HDS_BUTTONS = &H2
Const SWP_DRAWFRAME = &H20
Const SWP_NOMOVE = &H2
Const SWP_NOSIZE = &H1
Const SWP_NOZORDER = &H4
Public Const SWP_FLAGS = SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_DRAWFRAME

