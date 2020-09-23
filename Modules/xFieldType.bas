Attribute VB_Name = "xFieldType"
Option Explicit
Public Sub TypeX(ByRef rs As Recordset, ByRef lst As ListBox)
   Dim fldLoop As ADODB.Field
   ' Enumerate Fields collection of table.
   lst.Clear
   For Each fldLoop In rs.Fields
      lst.AddItem fldLoop.Name & "  > " & FieldType(fldLoop.Type)
'      Debug.Print "  Name: " & fldLoop.Name & vbCr & _
         "  Type: " & FieldType(fldLoop.Type) & vbCr
   Next fldLoop
End Sub

Public Function FieldType(intType As Integer) As String
   Select Case intType
      Case 16
         FieldType = "adTinyInt-16"
      Case 2
         FieldType = "adSmallInt-2"
      Case 3
         FieldType = "adInteger-3"
      Case 20
         FieldType = "adBigInt-20"
      Case 17
          FieldType = "adUnsignedTinyInt-17"
      Case 18
         FieldType = "adUnsignedSmallInt-18"
      Case 19
         FieldType = "adUnsignedInt-19"
      Case 21
         FieldType = "adUnsignedBigInt-21"
      Case 4
         FieldType = "adSingle-4"
      Case 5
         FieldType = "adDouble-5"
      Case 6
         FieldType = "adCurrency-6"
      Case 14
         FieldType = "adDecimal-14"
      Case 131
         FieldType = "adNumeric-131"
      Case 11
         FieldType = "adBoolean-11"
      Case 10
         FieldType = "adError-10"
      Case 72
         FieldType = "adGuid-72"
      Case 7
         FieldType = "adDate-7"
      Case 133
         FieldType = "adDBDate-133"
      Case 134
         FieldType = "adDBTime-134"
      Case 135
         FieldType = "adDBTimeStamp-135"
      Case 8
         FieldType = "adBSTR-8"
      Case 129
         FieldType = "adChar-129"
      Case 200
         FieldType = "adVarChar-200"
      Case 201
         FieldType = "adLongVarChar-201"
      Case 130
         FieldType = "adWChar-130"
      Case 202
         FieldType = "adVarWChar-202"
      Case 203
         FieldType = "adLongVarWChar-203"
      Case 128
         FieldType = "adBinary-128"
      Case 204
         FieldType = "adVarBinary-204"
      Case 205
         FieldType = "adLongVarBinary-205"
   End Select
End Function




