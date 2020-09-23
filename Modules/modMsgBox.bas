Attribute VB_Name = "modMsgBox"
Option Explicit
Public pb_vbOK As Boolean
Public pb_vbYes As Boolean
Public pb_vbNo As Boolean
Public pb_vbCancel As Boolean

Public Sub myMsg(ByVal sMessage As String, _
                 ByVal sCaption As String, _
                 Optional ByVal iIcon As Integer = 2, _
                 Optional ByVal bOk As Boolean = True)
'myMsg "Your msg here", "Caption", 1, True 'Show message
    '** Description:
    '** Show message with custom MsgBox
    initMsgVar
    If bOk = True Then
        frmMsgBox.cmdOk.Visible = True
    Else
        frmMsgBox.cmdCancel.Visible = True
        frmMsgBox.cmdNo.Visible = True
        frmMsgBox.cmdYes.Visible = True
    End If
    initImg
    ' See which icon is
    Select Case iIcon
     Case 0  'Critical
        frmMsgBox.imgCri.Visible = True
    Case 1   'help
        frmMsgBox.imgHelp.Visible = True
    Case 2  'info
        frmMsgBox.ImgInfo.Visible = True
    Case Else
         'nothing to do
    End Select
    frmMsgBox.Caption = sCaption 'Set msgbox caption
    frmMsgBox.txtMsg.text = sMessage 'Set message
    frmMsgBox.show 'vbModal 'Show form
End Sub
Private Sub initImg()
    frmMsgBox.imgHelp.Visible = False
    frmMsgBox.imgCri.Visible = False
    frmMsgBox.ImgInfo.Visible = False
End Sub

Public Sub initMsgVar()
 pb_vbOK = False
 pb_vbYes = False
 pb_vbNo = False
 pb_vbCancel = False
End Sub
