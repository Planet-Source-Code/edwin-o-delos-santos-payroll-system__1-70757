Attribute VB_Name = "modDetectFormLoaded"

'//Detect If Form Is Already Loaded
Option Explicit

'Add a module to your project (In the menu choose Project -> Add Module, Then click Open)
'Add 2 Command Buttons to your form. Add Another Form (Form 2).
'Click on the first button, and the program will detect that Form2 is not loaded.
'Now click on the second button to load Form2 and click again on the first button.
'The program will dedtect that Form2 is now loaded.
'Insert the following code to your module:

Public Function FormLoadedByName(FormName As String) As Boolean
Dim i As Integer, fnamelc As String
fnamelc = LCase$(FormName)
FormLoadedByName = False
For i = 0 To Forms.Count - 1
If LCase$(Forms(i).Name) = fnamelc Then
  FormLoadedByName = True
  Exit Function
End If
Next
End Function

'Insert the following code to your form (Form1):

'//Private Sub Command1_Click()
'Replace 'Form2' with the name of the form you want to detect his state.
'//If FormLoadedByName("Form2") = True Then
'//MsgBox "The Form is loaded"
'//Else
'//MsgBox "The Form is not loaded"
'//End If
'//End Sub

'//Private Sub Command2_Click()
'//Load Form2
'//End Sub



