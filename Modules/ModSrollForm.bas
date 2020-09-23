Attribute VB_Name = "ModScrollForm"
Option Explicit
'Public ctl As Control
'Public fullhtemp As Integer
'Public fullvtemp As Integer
Public iFullFormHeigth As Integer
Public iFullFormWidth As Integer
Public oldvPos As Integer
Public oldhPos As Integer


Public Sub GetFullSize(ByRef frm As Form)
Dim ctl As Control
Dim fullhtemp As Integer
Dim fullvtemp As Integer

fullhtemp = 0
fullvtemp = 0
For Each ctl In frm.Controls
        If ctl.Top + ctl.Height > fullvtemp Then fullvtemp = ctl.Top + ctl.Height
        If ctl.Left + ctl.Width > fullhtemp Then fullhtemp = ctl.Left + ctl.Width
Next
iFullFormHeigth = fullvtemp + frm.HScroll1.Height
iFullFormWidth = fullhtemp + frm.VScroll1.Width
End Sub

Public Sub ME_Resize(ByRef mefrm As Form)

mefrm.VScroll1.Left = mefrm.Width - (1.45 * mefrm.VScroll1.Width)

mefrm.HScroll1.Top = mefrm.Height - (2.75 * mefrm.HScroll1.Height)

mefrm.Shape1.Left = mefrm.VScroll1.Left
mefrm.Shape1.Top = mefrm.HScroll1.Top

'If the full screen is already showing,
'then disable the scrollbar
mefrm.VScroll1.Enabled = (iFullFormHeigth - mefrm.Height) >= 0

'First, make sure we aren't minimized
If mefrm.ScaleHeight > mefrm.HScroll1.Height And mefrm.Width > mefrm.VScroll1.Width Then
    
    'If there is any more screen to see,
    'modify the scrollbar
    If mefrm.VScroll1.Enabled Then
        With mefrm.VScroll1
            .Height = mefrm.ScaleHeight - mefrm.HScroll1.Height
            .Min = 0
            .Max = iFullFormHeigth - mefrm.Height
            .SmallChange = Screen.TwipsPerPixelY * 10
            .LargeChange = mefrm.ScaleHeight - mefrm.HScroll1.Height
        End With

    'Otherwise, just resize the scrollbar for neatness
    Else: mefrm.VScroll1.Height = mefrm.ScaleHeight - mefrm.HScroll1.Height
    End If

    mefrm.HScroll1.Enabled = (iFullFormWidth - mefrm.Width) >= 0
    If mefrm.HScroll1.Enabled Then
        With mefrm.HScroll1
            .Width = mefrm.ScaleWidth - mefrm.VScroll1.Width
            .Min = 0
            .Max = iFullFormWidth - mefrm.Width
            .SmallChange = Screen.TwipsPerPixelX * 10
            .LargeChange = mefrm.ScaleWidth - mefrm.VScroll1.Width
        End With

    Else: mefrm.HScroll1.Width = mefrm.ScaleWidth - mefrm.VScroll1.Width
    End If
End If
End Sub

Public Sub pScrollForm(meform As Form)
Dim ctl As Control

'Moves each textbox and Control if the scrollbar is clicked
For Each ctl In meform.Controls
    If Not (TypeOf ctl Is VScrollBar) And _
        Not (TypeOf ctl Is Frame) And _
        Not (TypeOf ctl Is CommandButton) And _
        Not (TypeOf ctl Is ComboBox) And _
        Not (TypeOf ctl Is ListView) And _
        Not (TypeOf ctl Is Label) And _
        Not (TypeOf ctl Is HScrollBar) Then
        ctl.Top = ctl.Top + oldvPos - meform.VScroll1.Value
        ctl.Left = ctl.Left + oldhPos - meform.HScroll1.Value
    End If
Next

oldvPos = meform.VScroll1.Value
oldhPos = meform.HScroll1.Value
End Sub

