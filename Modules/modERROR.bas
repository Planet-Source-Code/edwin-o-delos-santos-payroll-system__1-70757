Attribute VB_Name = "modERROR"
Option Explicit

Public Sub errorMsg(ByVal errNUM As ErrObject, _
                    Optional ByVal ModuleName As String, _
                    Optional ByVal OccurIn As String)
 Select Case errNUM
 Case Is = 0
   Exit Sub
 Case Is = 5
   MsgBox "Invalid procedure call or argument", vbCritical, "Warning!"
   Exit Sub
 Case Is = 13
   MsgBox "Data type mismatch!", vbCritical, "Warning!"
   Exit Sub
' Case Is = 3021  'requested operations require a curren record. Current Record has been deleted
' Case Is = 340   'Array doesnot exist
 Case Is = 32755 'Cancelled open
   Exit Sub
' Case Is = 3704  'Operation is not allowed whent the object is close
' Case Is = 9     'Subscrip out of range
 Case Is = 7005
   MsgBox "RowSet not available!", vbInformation, "Warning!"
   Exit Sub
 Case Is = -2147217843
   MsgBox "Not a valid password!", vbInformation, "Enter valid password"
   Exit Sub
 Case Is = -2147217887
   MsgBox "Cannot update (expression)!", vbInformation, "Field not updatable."
   Exit Sub
 Case Is = 3709
   Dim errMsg As String
   errMsg = "The connection cannot be used"
   errMsg = errMsg & Chr(10) & "to perform this operation"
   errMsg = errMsg & Chr(10) & "It is either closed or invalid"
   errMsg = errMsg & Chr(10) & "in this context.!"
   MsgBox errMsg, vbCritical, "Disconnected Recordset"
   Exit Sub
  Case Else
   MsgBox "Error From: " & ModuleName & vbNewLine & _
           "Occur In: " & OccurIn & vbNewLine & _
           "Error Number: " & errNUM.Number & vbNewLine & _
           "Description: " & errNUM.Description, vbCritical, "Application Error"
    'Save the error log (The save error log will be display later on in the program)
    Open App.Path & "\ErrorLog.log" For Append As #1
                Print #1, Format(Date, "MMM-dd-yyyy") & "]~~~~[" & Time & "]~~~~[" & Err.Number & "]~~~~[" & Err.Description & "]~~~~[" & ModuleName & "]~~~~[" & OccurIn
    Close #1
 End Select
End Sub


