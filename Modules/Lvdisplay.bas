Attribute VB_Name = "Lvdisplay"
Option Explicit
'Public Execute As Boolean
'Public Mytextbox As TextBox

Public Function PopulateList(pList As ListView, pRst As ADODB.Recordset)
    On Error Resume Next
    Dim i As Integer, iColCount As Integer
    Dim sColName As String
    Dim sColValue As String
    Dim oCH As ColumnHeader
    Dim oLI As ListItem
    Dim oSI As ListSubItem
    Dim oFld As ADODB.Field

    With pList
        .View = lvwReport
        pRst.MoveFirst
        For Each oFld In pRst.Fields
            sColName = CkNuL(oFld.name)
            Set oCH = .ColumnHeaders.Add()
            oCH.text = sColName
            iColCount = iColCount + 1
        
        Next oFld

        While Not pRst.EOF
            i = 0
           
            sColValue = CkNuL(pRst.Fields(i).Value)
            Set oLI = .ListItems.Add()
            oLI.text = sColValue
                        
            For i = 1 To iColCount
                Set oSI = oLI.ListSubItems.Add()
                oSI.text = CkNuL(pRst(i))
            Next
            pRst.MoveNext
        Wend ' next record
    pRst.Close
    Set pRst = Nothing
    End With
End Function

Private Function CkNuL(pVal As String) As String
    If IsMissing(pVal) Then
        CkNuL = ""
    ElseIf IsNull(pVal) Then
        CkNuL = ""
    Else
        CkNuL = Format(pVal)
    End If
End Function


