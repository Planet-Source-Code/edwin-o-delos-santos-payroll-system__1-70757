Attribute VB_Name = "modAlignLVCol"
'Public Declare Function SendMessage Lib "user32" _
'    Alias "SendMessageA" (ByVal hwnd As Long, _
'    ByVal wMsg As Long, _
'    ByVal wParam As Long, _
'    lParam As Any) As Long
Const LVM_SETCOLUMNWIDTH = &H1000 + 30
Const LVSCW_AUTOSIZE = -1
Const LVSCW_AUTOSIZE_USEHEADER = -2

'*LV is the ListView control
Public Sub autoAlignCol(ByVal lv As ListView)
Dim col As Long
For col = 0 To lv.ColumnHeaders.Count - 1
    SendMessage lv.hWnd, LVM_SETCOLUMNWIDTH, col, LVSCW_AUTOSIZE_USEHEADER
Next col
End Sub
