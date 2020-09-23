Attribute VB_Name = "modAPIelite"


Option Explicit

' Declaration for Stay on Top sub
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, y, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1

' Win32 Declarations for DisableX
Private Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Private Const MF_BYPOSITION = &H400&

' Win32 Declarations for INI Access
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

' Win32 Declarations for Cut, Copy, Paste and Delete
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, LParam As Any) As Long
Public Const WM_CUT = &H300
Public Const WM_COPY = &H301
Public Const WM_PASTE = &H302
Public Const WM_CLEAR = &H303
Public Const WM_USER = &H400
Public Const EM_CANUNDO = &HC6
Public Const EM_UNDO = &HC7

Public Const EM_LINEINDEX = &HBB
Private Const EM_SETTARGETDEVICE = (WM_USER + 72)
Public Const EM_GETLINECOUNT = &HBA
Public Const EM_LINEFROMCHAR = &HC9

' Win32 Declarations for FolderView
Public Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long

' Win 32 Declarations for View Mode
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal LParam As Long) As Long
Public Enum ERECViewModes
    ercDefault = 0
    ercWordWrap = 1
    ercWYSIWYG = 2
End Enum

' Declarations for FormatSize
Public Declare Function StrFormatByteSize Lib "shlwapi" Alias "StrFormatByteSizeA" (ByVal dw As Long, ByVal pszBuf As String, ByRef cchBuf As Long) As String
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const WM_GETTEXTLENGTH = &HE

' Win32 Declarations for enum fonts
Private Const LF_FACESIZE = 32
Private Type LOGFONT
        lfHeight As Long
        lfWidth As Long
        lfEscapement As Long
        lfOrientation As Long
        lfWeight As Long
        lfItalic As Byte
        lfUnderline As Byte
        lfStrikeOut As Byte
        lfCharSet As Byte
        lfOutPrecision As Byte
        lfClipPrecision As Byte
        lfQuality As Byte
        lfPitchAndFamily As Byte
        lfFaceName(LF_FACESIZE) As Byte
End Type
Private Type NEWTEXTMETRIC
        tmHeight As Long
        tmAscent As Long
        tmDescent As Long
        tmInternalLeading As Long
        tmExternalLeading As Long
        tmAveCharWidth As Long
        tmMaxCharWidth As Long
        tmWeight As Long
        tmOverhang As Long
        tmDigitizedAspectX As Long
        tmDigitizedAspectY As Long
        tmFirstChar As Byte
        tmLastChar As Byte
        tmDefaultChar As Byte
        tmBreakChar As Byte
        tmItalic As Byte
        tmUnderlined As Byte
        tmStruckOut As Byte
        tmPitchAndFamily As Byte
        tmCharSet As Byte
        ntmFlags As Long
        ntmSizeEM As Long
        ntmCellHeight As Long
        ntmAveWidth As Long
End Type

Private Declare Function EnumFontFamilies Lib "gdi32" Alias "EnumFontFamiliesA" (ByVal hdc As Long, ByVal lpszFamily As String, ByVal lpEnumFontFamProc As Long, LParam As Any) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long

' Win32 Declarations for Print sub
Private Type Rect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type CharRange
    cpMin As Long     ' First character of range (0 for start of doc)
    cpMax As Long     ' Last character of range (-1 for end of doc)
End Type

Private Type FormatRange
    hdc As Long       ' Actual DC to draw on
    hdcTarget As Long ' Target DC for determining text formatting
    rc As Rect        ' Region of the DC to draw to (in twips)
    rcPage As Rect    ' Region of the entire DC (page size) (in twips)
    chrg As CharRange ' Range of text to draw (see above declaration)
End Type

Const EM_FORMATRANGE As Long = WM_USER + 57
Const PHYSICALOFFSETX As Long = 112
Const PHYSICALOFFSETY As Long = 113

Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long

Public Sub SetViewMode(ByVal eViewMode As ERECViewModes)
    '** Description:
    '** Set View Mode
    Select Case eViewMode 'Set View Mode
        Case 0 'to No Wrap
            SendMessageLong frmMDI.ActiveForm.rtfText.hWnd, EM_SETTARGETDEVICE, 0, 1
        Case 1 'to Word Wrap
            SendMessageLong frmMDI.ActiveForm.rtfText.hWnd, EM_SETTARGETDEVICE, 0, 0
        Case 2 'to WYSIWYG
            On Error Resume Next
            SendMessageLong frmMDI.ActiveForm.rtfText.hWnd, EM_SETTARGETDEVICE, Printer.hdc, Printer.Width
   End Select
End Sub

Public Sub DisableX(TheForm As Form)
    '** Description:
    '** Disable X in upper right corner of the form
    Dim lngMenu As Long
    lngMenu = GetSystemMenu(TheForm.hWnd, False)
    DeleteMenu lngMenu, 6, MF_BYPOSITION
End Sub

Public Sub NotOnTop(TheForm As Form)
    '** Description:
    '** Remove window from top
    SetWindowPos TheForm.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub

Public Sub OnTop(TheForm As Form)
    '** Description:
    '** Put window on top
    SetWindowPos TheForm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub

Public Function ReadINI(Section As String, Key As String, Optional sDefault As String)
    '** Description:
    '** Get settings from ini file
    Dim sRet As String
    ' Fill sRet with Null Chars
    sRet = String(255, Chr(0))
    ' Get data from INI file
    ReadINI = Left(sRet, GetPrivateProfileString(Section, Key, sDefault, sRet, Len(sRet), App.Path & "\ElitePad.ini"))
End Function

Public Sub WriteINI(Section As String, Key As String, Value As String)
    '** Description:
    '** Write settings to ini file
    ' Write to INI file
    WritePrivateProfileString Section, Key, Value, App.Path & "\ElitePad.ini"
End Sub


Private Function EnumFontFamProc(lpNLF As LOGFONT, lpNTM As NEWTEXTMETRIC, ByVal FontType As Long, LParam As ComboBox) As Long
    Dim FaceName As String
    ' Convert font name
    FaceName = StrConv(lpNLF.lfFaceName, vbUnicode)
    ' Add font
    LParam.AddItem Left$(FaceName, InStr(FaceName, vbNullChar) - 1)
    EnumFontFamProc = 1
End Function

Public Function LoadFonts()
    Dim hdc As Long
    Dim I As Integer
    frmMDI.cboFontName.Clear 'Clear combobox
    hdc = GetDC(frmMDI.cboFontName.hWnd) 'Get combobox DC
    ' Enum fonts
    EnumFontFamilies hdc, vbNullString, AddressOf EnumFontFamProc, frmMDI.cboFontName
    ReleaseDC frmMDI.cboFontName.hWnd, hdc 'Release combobox DC
    
    For I = 5 To 72
        ' Fills combobox with font size from 5 to 75
        frmMDI.cboFontSize.AddItem I
    Next I
    
    frmMDI.cboFontName.text = "Tahoma" 'Set default font name
    frmMDI.cboFontSize.text = "9" 'Set default font size
End Function


