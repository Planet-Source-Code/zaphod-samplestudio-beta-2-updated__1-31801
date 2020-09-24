Attribute VB_Name = "Module1"
'SampleStudio
'by: Paul Bryan in 2002
'Allows for graphical isolation of sample ranges
'Load, Record, Zoom, Play, Loop, Sound Bank Saving,
' Hotkey combinations, Triggers Mutiple Formats
'and will Paste Selected Data Into New Files
' feel free to re-use this code. but give me some credit :)

' Main Module

Public Declare Function FindFirstFile& Lib "kernel32" _
       Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData _
       As WIN32_FIND_DATA) 'FileExsists Stuff

Public Declare Function FindClose Lib "kernel32" _
       (ByVal hFindFile As Long) As Long 'FileExsists Stuff


'Class holders for WavForm.frm (Sample) Instances
Public SCount(255) As Boolean   ' Multiple Open Wave Files
Public Scope(255) As Form ' Session Filecount
Public SFilePath(255) As String
Public FileCount As Integer ' New File Number Counter
Public HotKey(255) As Integer
' Stat holders for MDI form
Public CurFormFocus As Integer
Public CurFilePath As String
Public CurBitRate As Integer
Public CurSampRate As Long
Public CurChannels As String

'More File Exsists Stuff
Public Const MAX_PATH = 260
Type FILETIME ' 8 Bytes
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type
Type WIN32_FIND_DATA ' 318 Bytes
        dwFileAttributes As Long
        ftCreationTime As FILETIME
        ftLastAccessTime As FILETIME
        ftLastWriteTime As FILETIME
        nFileSizeHigh As Long
        nFileSizeLow As Long
        dwReserved¯ As Long
        dwReserved1 As Long
        cFileName As String * MAX_PATH
        cAlternate As String * 14
End Type
    
Public Sub main() ' Program flow starts here!
    
    MDIMain.Show
    
End Sub
Public Sub LoadNewFile(Fpath As String, FName As String, Optional Spawn As Integer) ' Open another file
        'creates a new Instance of Wavform
        'and point it to Sample data (Fname)
        Dim n As Integer
            For n = 1 To 255
                If SCount(n) = False Then Exit For
            Next n
                SCount(n) = True
                Set Scope(n) = New WavForm
                Scope(n).SetFocus
    Call Scope(n).LoadFileData(n, Spawn, Fpath, FName)
 
End Sub
Function FExsists(strFileName As String) As Boolean ' Does File Already Exsist?
Dim lpFindFileData As WIN32_FIND_DATA
Dim hFindFirst As Long
       hFindFirst = FindFirstFile(strFileName, lpFindFileData)
              If hFindFirst > 0 Then
                      FindClose hFindFirst
                      FExsists = True
              Else
                      FExsists = False
              End If
End Function
