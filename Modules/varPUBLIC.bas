Attribute VB_Name = "varPUBLIC"
Option Explicit

'// handle add & edit//
Public addrec As Boolean
Public editrec As Boolean
'// handle access menu//
Public accss As Boolean
Public mymenu As String
'// connection//
Public CN As New Connection
'// handle autocomplete in listview//
Public Execute As Boolean
Public Mytextbox As TextBox
'//
Public userADMIN As String
Public userNAME As String

'// hanlde formatdate funtion
Public idate
Public iyear
Public imonth
Public iday
'//  handle form resize
Public frm_h As Long
Public frm_w As Long
'// handle show updn popup
Public showPopup As Boolean
'// handle to show/hide entry/option form
Public showform As Boolean
'// handle convert number to words
Public convert As numTOword
'// handle preview
Public PV As PreviewTxt
