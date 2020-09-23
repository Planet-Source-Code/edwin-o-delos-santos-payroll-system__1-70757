Attribute VB_Name = "modPUBLIC"
Option Explicit
Public nxTab     As Integer     'TO HANDLE TAB ORDER/keyEvents [Enter] & Up arrow key
Public addRec As Boolean        'to handle new record
Public editRec As Boolean       'to handle existing record
Public currLen(15) As Integer   'array to store lenght of string/value - used by PrintValue procedure
Public printIndex(50) As String 'store field to print
Public initPrint As Boolean     'initialize list to print
'//handle to move form
Public down As Boolean
Public t As Integer
Public w As Integer
'==========================
Public CurrUser                    As USER_INFO  '(see modEncrypt)
Public textpass                    As String 'temporary storge for password
Public isFilter As Boolean        'flag for search /filter

