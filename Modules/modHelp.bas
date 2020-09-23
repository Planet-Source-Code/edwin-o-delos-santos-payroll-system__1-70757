Attribute VB_Name = "modHelp"
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
(ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Function HyperJump(ByVal URL As String) As Long
On Error Resume Next
   HyperJump = ShellExecute(0&, vbNullString, URL, vbNullString, vbNullString, vbNormalFocus)
End Function

'<< syntax>
'Dim xHlp
'private sub cmdHelp_click()
' xHlp = HyperJump(App.HelpFile)
'End sub
