Attribute VB_Name = "modOpenProcess"
Option Explicit

Public Const PROCESS_QUERY_INFORMATION = &H400

'// launching a New Application
Declare Function OpenProcess Lib "kernel32" _
  (ByVal dwDesiredAccess As Long, ByVal bInheritHandle _
  As Long, ByVal dwProcessID As Long) As Long

'-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-


