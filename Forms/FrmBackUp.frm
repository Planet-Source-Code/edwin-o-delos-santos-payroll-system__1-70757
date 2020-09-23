VERSION 5.00
Begin VB.Form FrmBackUp 
   Caption         =   "Back-Up Database Files..."
   ClientHeight    =   4965
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6705
   Icon            =   "FrmBackUp.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   6705
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00C0C0C0&
      Caption         =   "E&xit"
      Height          =   315
      Left            =   1440
      TabIndex        =   12
      Top             =   4560
      Width           =   1200
   End
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Search"
      Default         =   -1  'True
      Height          =   315
      Left            =   120
      TabIndex        =   11
      Top             =   4560
      Width           =   1200
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   2175
      Left            =   0
      ScaleHeight     =   2175
      ScaleWidth      =   6615
      TabIndex        =   7
      Top             =   2280
      Width           =   6615
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   4320
         TabIndex        =   18
         Top             =   490
         Width           =   200
      End
      Begin VB.CommandButton cmdCheckDir 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Verify Folder"
         Height          =   315
         Left            =   2880
         TabIndex        =   17
         Top             =   480
         Width           =   1320
      End
      Begin VB.CommandButton cmdMkDir 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Create New Folder"
         Enabled         =   0   'False
         Height          =   315
         Left            =   4560
         TabIndex        =   16
         Top             =   480
         Width           =   1800
      End
      Begin VB.CheckBox chkOverWrite 
         Caption         =   "Overwrite?"
         Height          =   255
         Left            =   1680
         TabIndex        =   15
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox TextPath 
         Height          =   285
         Left            =   840
         TabIndex        =   13
         Text            =   "C:\BACKUP\"
         Top             =   120
         Width           =   5535
      End
      Begin VB.ListBox lstFoundFiles 
         Appearance      =   0  'Flat
         Height          =   1155
         Left            =   240
         Style           =   1  'Checkbox
         TabIndex        =   8
         Top             =   840
         Width           =   6135
      End
      Begin VB.Label Label1 
         Caption         =   "Copy to:"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   120
         Width           =   615
      End
      Begin VB.Label lblfound 
         AutoSize        =   -1  'True
         Caption         =   "&Files Found:"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   480
         Width           =   855
      End
      Begin VB.Label lblCount 
         AutoSize        =   -1  'True
         Caption         =   "0"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   1200
         TabIndex        =   9
         Top             =   480
         Width           =   90
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   2175
      Left            =   0
      ScaleHeight     =   2175
      ScaleWidth      =   6615
      TabIndex        =   1
      Top             =   0
      Width           =   6615
      Begin VB.TextBox txtSearchSpec 
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   5
         Text            =   "*.MDB"
         Top             =   120
         Width           =   1935
      End
      Begin VB.FileListBox filList 
         Height          =   1455
         Left            =   3480
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   4
         Top             =   600
         Width           =   2895
      End
      Begin VB.DirListBox dirList 
         Height          =   1215
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   3135
      End
      Begin VB.DriveListBox drvList 
         Height          =   315
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   3135
      End
      Begin VB.Image imgHelp 
         Height          =   360
         Left            =   6000
         MouseIcon       =   "FrmBackUp.frx":1CCA
         MousePointer    =   99  'Custom
         Picture         =   "FrmBackUp.frx":2594
         Top             =   120
         Width           =   360
      End
      Begin VB.Label lblCriteria 
         Caption         =   "Search &Criteria:"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   120
         Width           =   1335
      End
   End
   Begin VB.CommandButton CmdCopy 
      Caption         =   "Copy Selected Files"
      Height          =   315
      Left            =   4680
      TabIndex        =   0
      Top             =   4560
      Width           =   1695
   End
End
Attribute VB_Name = "FrmBackUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_path As String
Dim SearchFlag As Integer   ' Used as flag for cancel and other operations.
Private mb_vbYes As Boolean

Private Sub Check1_Click()
 cmdMkDir.Enabled = (Check1.Value = 1)
End Sub

Private Sub cmdCheckDir_Click()
Dim FSys As New FileSystemObject
Dim msg As String
Dim b_folderExist As Boolean     'determine if folder distination exists
b_folderExist = FSys.FolderExists(TextPath.text)
  If b_folderExist = False Then
     msg = "Folder does not exists." & vbCrLf
     msg = msg + "Create new folder?"
     myMsg msg, TextPath, 2, False
  Else  'if already exists do nothing
     myMsg "Already exists", TextPath, 2, True
  End If
  msg = ""

End Sub



Private Sub CmdCopy_Click()
On Error Resume Next
Dim FSys As New FileSystemObject
Dim thisFile As File
Dim b_folderExist As Boolean     'determine if folder distination exists
Dim b_fileExist As Boolean       'determine if files to be copied exists
Dim b_nofile As Boolean          'flag for selected record
Dim i_FileCount As Integer       'flag for number of file copied
Dim msg As String
Dim mfileCopy As String
Dim mfileExist As String
Dim i As Integer
Dim success As Boolean
'//initialize
success = False
b_nofile = True
i_FileCount = 0
b_folderExist = FSys.FolderExists(TextPath.text)
If b_folderExist = False Then
   myMsg "Distination folder does not exist!", TextPath, 2, True
   Exit Sub
End If
   For i = 0 To lstFoundFiles.ListCount - 1
     Set thisFile = FSys.GetFile(lstFoundFiles.List(i))
     If Not (thisFile Is Nothing) Then
        If lstFoundFiles.Selected(i) = True Then
           b_nofile = False
           If chkOverWrite.Value = 1 Then
               '//overwrite existing files
               thisFile.Copy m_path & thisFile.Name, True
               mfileCopy = mfileCopy & thisFile.Name & vbCrLf
               i_FileCount = i_FileCount + 1
            Else
                b_fileExist = FSys.FileExists(thisFile.Name)
                If b_fileExist = True Then
                    mfileExist = mfileExist & thisFile.Name & vbCrLf
                Else
                    thisFile.Copy m_path & thisFile.Name
                    mfileCopy = mfileCopy & thisFile.Name & vbCrLf
                    i_FileCount = i_FileCount + 1
                End If
            End If  '//check value = 1
        Else
           If b_nofile = False Then    'there are selected record
              b_nofile = False
           Else
              b_nofile = True   'of no selected record flag = true
           End If
        End If    '//selected = true
          If i_FileCount > 0 Then
              success = True
          Else
              success = False
          End If
     End If
   Next i
If b_nofile = True Then
    myMsg "No selected Fiel/s to copy!", "Copy selected files", 2, True
    Exit Sub
End If
If success = True Then
   msg = "Copying Successfull!" & vbCrLf
   msg = msg & vbCrLf & "[ File/s copied: ]"
   msg = msg & vbCrLf & mfileCopy
   msg = msg & vbCrLf & "[ Already Exist: ]"
   msg = msg & vbCrLf & mfileExist
   myMsg msg, "Copy selected files", 2, True
Else
    msg = "Copying Not Successfull!" & vbCrLf
    msg = msg & vbCrLf & "[ Already Exist: ]"
    msg = msg & vbCrLf & mfileExist
    myMsg msg, "Copy selected files", 2, True
End If
   success = False
   mfileCopy = ""
   mfileExist = ""
   b_nofile = True
   i_FileCount = 0
   msg = ""

End Sub

Private Sub cmdExit_Click()
    If cmdExit.Caption = "E&xit" Then
      Unload FrmBackUp
      Set FrmBackUp = Nothing
    Else                    ' If user chose Cancel, just end Search.
        SearchFlag = False
    End If
End Sub




Private Sub cmdMkDir_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim FSys As New FileSystemObject
Dim b_folderExist As Boolean     'determine if folder distination exists
b_folderExist = FSys.FolderExists(TextPath.text)
 If b_folderExist = True Then
     myMsg TextPath.text & " already exists", "Folder", 2, True
     Check1.Value = 0
    Exit Sub
 End If
  mb_vbYes = pb_vbYes
 If mb_vbYes = True Then
      Dim ms_Path As String
      ms_Path = TextPath.text
      MkDir ms_Path
      myMsg ms_Path & " is created", "New folder", 2, True
Else
      myMsg "Verify First", "Create New folder", 2, True
End If
      mb_vbYes = False
      pb_vbYes = False
      ms_Path = ""
End Sub

Private Sub cmdSearch_Click()
' Initialize for search, then perform recursive search.
Dim FirstPath As String, DirCount As Integer, NumFiles As Integer
Dim result As Integer
  ' Check what the user did last.
    If cmdSearch.Caption = "&Reset" Then  ' If just a reset, initialize and exit.
        ResetSearch
        txtSearchSpec.SetFocus
        Exit Sub
    End If

    ' Update dirList.Path if it is different from the currently
    ' selected directory, otherwise perform the search.
    If dirList.Path <> dirList.List(dirList.ListIndex) Then
        dirList.Path = dirList.List(dirList.ListIndex)
        Exit Sub         ' Exit so user can take a look before searching.
    End If

    ' Continue with the search.
'    Picture2.Move 0, 0
'    Picture1.Visible = False
    Picture2.Visible = True

    cmdExit.Caption = "Cancel"

    filList.Pattern = txtSearchSpec.text
    FirstPath = dirList.Path
    DirCount = dirList.ListCount

    ' Start recursive direcory search.
    NumFiles = 0                       ' Reset found files indicator.
    result = DirDiver(FirstPath, DirCount, "")
    filList.Path = dirList.Path
    cmdSearch.Caption = "&Reset"
    cmdSearch.SetFocus
    cmdExit.Caption = "E&xit"
End Sub

Private Function DirDiver(NewPath As String, DirCount As Integer, BackUp As String) As Integer
'  Recursively search directories from NewPath down...
'  NewPath is searched on this recursion.
'  BackUp is origin of this recursion.
'  DirCount is number of subdirectories in this directory.
Static FirstErr As Integer
Dim DirsToPeek As Integer, AbandonSearch As Integer, ind As Integer
Dim OldPath As String, ThePath As String, entry As String
Dim retval As Integer
    SearchFlag = True           ' Set flag so the user can interrupt.
    DirDiver = False            ' Set to True if there is an error.
    retval = DoEvents()         ' Check for events (for instance, if the user chooses Cancel).
    If SearchFlag = False Then
        DirDiver = True
        Exit Function
    End If
    On Local Error GoTo DirDriverHandler
    DirsToPeek = dirList.ListCount                  ' How many directories below this?
    Do While DirsToPeek > 0 And SearchFlag = True
        OldPath = dirList.Path                      ' Save old path for next recursion.
        dirList.Path = NewPath
        If dirList.ListCount > 0 Then
            ' Get to the node bottom.
            dirList.Path = dirList.List(DirsToPeek - 1)
            AbandonSearch = DirDiver((dirList.Path), DirCount%, OldPath)
        End If
        ' Go up one level in directories.
        DirsToPeek = DirsToPeek - 1
        If AbandonSearch = True Then Exit Function
    Loop
    ' Call function to enumerate files.
    If filList.ListCount Then
        If Len(dirList.Path) <= 3 Then             ' Check for 2 bytes/character
            ThePath = dirList.Path                  ' If at root level, leave as is...
        Else
            ThePath = dirList.Path + "\"            ' Otherwise put "\" before the filename.
        End If
        For ind = 0 To filList.ListCount - 1        ' Add conforming files in this directory to the list box.
            entry = ThePath + filList.List(ind)
            lstFoundFiles.AddItem entry
            lblCount.Caption = str(Val(lblCount.Caption) + 1)
        Next ind
    End If
    If BackUp <> "" Then        ' If there is a superior directory, move it.
        dirList.Path = BackUp
    End If
    Exit Function
DirDriverHandler:
    If Err = 7 Then             ' If Out of Memory error occurs, assume the list box just got full.
        DirDiver = True         ' Create Msg and set return value AbandonSearch.
        MsgBox "You've filled the list box. Abandoning search..."
        Exit Function           ' Note that the exit procedure resets Err to 0.
    Else                        ' Otherwise display error message and quit.
        MsgBox Error
        End
    End If
End Function




Private Sub DirList_Change()
    ' Update the file list box to synchronize with the directory list box.
    filList.Path = dirList.Path
End Sub


Private Sub DirList_LostFocus()
    dirList.Path = dirList.List(dirList.ListIndex)
End Sub

Private Sub DrvList_Change()
    On Error GoTo DriveHandler
    dirList.Path = drvList.Drive
    Exit Sub

DriveHandler:
    drvList.Drive = dirList.Path
    Exit Sub
End Sub


Private Sub filList_Click()
'lstFoundFiles.AddItem filList.FileName
End Sub

Private Sub Form_Load()
    mb_vbYes = False
    m_path = App.Path & "\"
    TextPath.text = m_path
    DisableX Me
End Sub

Private Sub Form_Resize()
With FrmBackUp
  If .WindowState = 0 Then
   .Height = 5475
   .Width = 6825
  End If
End With
 SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload FrmBackUp
    Set FrmBackUp = Nothing
End Sub

Private Sub ResetSearch()
    ' Reinitialize before starting a new search.
    lstFoundFiles.Clear
    lblCount.Caption = 0
    SearchFlag = False                  ' Flag indicating search in progress.
'    Picture2.Visible = False
    cmdSearch.Caption = "&Search"
    cmdExit.Caption = "E&xit"
    Picture1.Visible = True
    dirList.Path = CurDir: drvList.Drive = dirList.Path ' Reset the path.
End Sub



Private Sub imgHelp_Click()
Dim msg As String
msg = "Notice: The Folder " & m_path & " must exist" & vbCrLf
msg = msg & vbCrLf & "<< Search >> This will search files specified in Search Criteria"
msg = msg & vbCrLf & "<< Verify folder >> This will check if folder already exists"

 myMsg msg, "Back-Up", 1, True
End Sub

Private Sub txtSearchSpec_Change()
    ' Update file list box if user changes pattern.
    filList.Pattern = txtSearchSpec.text
End Sub

Private Sub txtSearchSpec_GotFocus()
    txtSearchSpec.SelStart = 0          ' Highlight the current entry.
    txtSearchSpec.SelLength = Len(txtSearchSpec.text)
End Sub


