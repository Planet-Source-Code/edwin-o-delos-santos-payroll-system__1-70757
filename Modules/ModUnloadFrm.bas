Attribute VB_Name = "ModUnloadFrm"
Option Explicit
Public Sub UnloadAllForms()
Dim Form As Form
   For Each Form In Forms
      Unload Form
      Set Form = Nothing
   Next Form
End Sub

