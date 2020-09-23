Attribute VB_Name = "modLVfind"
Option Explicit

' =====================================================================================================================
'
' Function:     EnhListView_Find
'
' Imputs:
'               Variable Name       Type        Optional    Description
'               lstListName         ListView    No          Name of the ListView to find in
'               strStringToFind     String      No          What to find in the list
'               bolWholeWordOnly    Boolean     Yes         Only 'find' if the 'found' is exactly like the 'search'
'               bolCaseSensitive    Boolean     Yes         Only 'find' if the 'found' is the same case as the 'search'
'
' Returns:      Integer of the 'found' item
'               Also selects the item and makes sure the item is visible
'
' =====================================================================================================================
Public Function EnhListView_Find(lstListName As ListView, _
                                strStringToFind As String, _
                                Optional bolWholeWordOnly As Boolean, _
                                Optional bolCaseSensitive As Boolean) _
                                As Integer
    ' setup variables
    Dim lngIndex As Long        ' used for the current index of the parent items
    Dim lngIndexSub As Long     ' used for the current index of the subitems
    Dim strCurrItem As String   ' used to store the text of the currently selected item for compare
    
    ' if we want to be sensitive about the case then make the 'search' all upper case
    If bolCaseSensitive = True Then strStringToFind = UCase(strStringToFind)
    
    ' set the return to the default zero
    EnhListView_Find = 0
    
    ' if there is nothing to search then exit
    If lstListName.ListItems.Count < 1 Then Exit Function
    
    ' if no item is currently selected then select the first item
    If lstListName.SelectedItem.Index = -1 Then lstListName.SelectedItem.Index = 1
    
    ' move through the rows
    For lngIndex = lstListName.SelectedItem.Index - -1 To lstListName.ListItems.Count
        
        ' if we want to be sensitive about the case then...
        If bolCaseSensitive = True Then
            ' fill our variable with the uppercase version of the current text
            strCurrItem = UCase(lstListName.ListItems.item(lngIndex).text)
        Else
            ' otherwise, fill our variable with the current text
            strCurrItem = lstListName.ListItems.item(lngIndex).text
        End If
        
        If bolWholeWordOnly = True Then
            ' if the current item and the 'search' is an exact match then finalize
            If strCurrItem = strStringToFind Then GoTo Finalize
        Else
            ' if the current item contains the 'search' then finalize
            If InStr(strCurrItem, strStringToFind) > 0 Then GoTo Finalize
        End If
        
        ' if we have subitems...
        If lstListName.ColumnHeaders.Count > 1 Then
            
            ' move through the subitems of the current row
            For lngIndexSub = 1 To lstListName.ColumnHeaders.Count - 1
                ' if we want to be sensitive about the case then...
                If bolCaseSensitive = True Then
                    ' fill our variable with the uppercase version of the current text
                    strCurrItem = UCase(lstListName.ListItems.item(lngIndex).SubItems(lngIndexSub))
                Else
                    ' otherwise, fill our variable with the current text
                    strCurrItem = lstListName.ListItems.item(lngIndex).SubItems(lngIndexSub)
                End If
                
                If bolWholeWordOnly = True Then
                    ' if the current item and the 'search' is an exact match then finalize
                    If strCurrItem = strStringToFind Then GoTo Finalize
                Else
                    ' if the current item contains the 'search' then finalize
                    If InStr(strCurrItem, strStringToFind) > 0 Then GoTo Finalize
                End If
            ' move to next subitem
            Next lngIndexSub
        
        End If
        
    ' move to next row
    Next lngIndex
    
    Exit Function
    
Finalize:
    EnhListView_Find = lngIndex                             ' send back the index of the found item
    lstListName.ListItems.item(lngIndex).EnsureVisible      ' make sure the item is visible
    lstListName.ListItems.item(lngIndex).Selected = True    ' make sure the item is selected
End Function
' =====================================================================================================================




