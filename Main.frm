VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm MDIMain 
   BackColor       =   &H00818787&
   Caption         =   "Sample Sequencer (by: Pbryan^ 2k And 2)"
   ClientHeight    =   4680
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   10725
   Icon            =   "Main.frx":0000
   LinkTopic       =   "MainForm"
   Picture         =   "Main.frx":08CA
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   2
      Left            =   360
      Top             =   2760
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   615
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   10665
      TabIndex        =   0
      Top             =   0
      Width           =   10725
      Begin VB.Label lblSampRate 
         Height          =   255
         Left            =   8400
         TabIndex        =   7
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblSelectedSamples 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   7560
         TabIndex        =   6
         Top             =   330
         Width           =   75
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Samples Selected:"
         Height          =   195
         Left            =   6120
         TabIndex        =   5
         Top             =   360
         Width           =   1320
      End
      Begin VB.Image Image8 
         BorderStyle     =   1  'Fixed Single
         Height          =   540
         Left            =   3560
         Picture         =   "Main.frx":2AC0E
         Stretch         =   -1  'True
         ToolTipText     =   "Clear Selection"
         Top             =   0
         Width           =   540
      End
      Begin VB.Image Image7 
         BorderStyle     =   1  'Fixed Single
         Height          =   540
         Left            =   3000
         Picture         =   "Main.frx":2AF9E
         Stretch         =   -1  'True
         ToolTipText     =   "Paste as a New File"
         Top             =   0
         Width           =   540
      End
      Begin VB.Image Image6 
         BorderStyle     =   1  'Fixed Single
         Height          =   540
         Left            =   2482
         Picture         =   "Main.frx":2B4D2
         Stretch         =   -1  'True
         ToolTipText     =   "Restore (Zoom Out)"
         Top             =   0
         Width           =   540
      End
      Begin VB.Image Image5 
         BorderStyle     =   1  'Fixed Single
         Height          =   540
         Left            =   1920
         Picture         =   "Main.frx":2B894
         Stretch         =   -1  'True
         ToolTipText     =   "Zoom In"
         Top             =   0
         Width           =   540
      End
      Begin VB.Label lblChannels 
         Height          =   255
         Left            =   10200
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblBitRate 
         Height          =   255
         Left            =   9480
         TabIndex        =   3
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblSecLen 
         Height          =   255
         Left            =   4200
         TabIndex        =   2
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label lblFilePath 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4200
         TabIndex        =   1
         Top             =   0
         Width           =   5295
      End
      Begin VB.Image Image2 
         BorderStyle     =   1  'Fixed Single
         Height          =   540
         Left            =   960
         Picture         =   "Main.frx":2C15E
         ToolTipText     =   "Pause"
         Top             =   0
         Width           =   540
      End
      Begin VB.Image Image3 
         BorderStyle     =   1  'Fixed Single
         Height          =   540
         Left            =   1440
         Picture         =   "Main.frx":2C468
         ToolTipText     =   "Stop"
         Top             =   0
         Width           =   540
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   540
         Left            =   0
         Picture         =   "Main.frx":2C772
         ToolTipText     =   "Play"
         Top             =   0
         Width           =   540
      End
      Begin VB.Image Image4 
         BorderStyle     =   1  'Fixed Single
         Height          =   540
         Left            =   480
         Picture         =   "Main.frx":2CA7C
         ToolTipText     =   "Play Loop"
         Top             =   0
         Width           =   540
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   360
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuOpenItem 
         Caption         =   "Open Sample (*.wav file)"
      End
      Begin VB.Menu mnuRecord 
         Caption         =   "Record New Sample (*.wav file)"
      End
      Begin VB.Menu mnuSepr 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLoadBank 
         Caption         =   "Load Sample Bank (*.bnk file)"
      End
      Begin VB.Menu mnuSaveBank 
         Caption         =   "Save Sample Bank (*.bnk file)"
      End
      Begin VB.Menu mnuExitItem 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "MDIMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'SampleStudio
'by: Paul Bryan in 2002
'Allows for graphical isolation of sample ranges
'Load, Record, Zoom, Play, Loop, Sound Bank Saving,
' Hotkey combinations, Triggers Mutiple Formats
'and will Paste Selected Data Into New Files
' feel free to re-use this code. but give me some credit :)

'Main Interface

Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Dim retVal As Long
Dim LastScope As Integer

    'Key Codes for Trigger types (I.E.> Loop, Play, and Stop types)
    '-----------------------------------------------------------
    '9=    tab                      40=    DnArrow
    '13=   vbCrLf                   41=    Select
    '17=    ctrl                    44=    PrintScr
    '18=    alt                     45=    Insert
    '19=    pause                   46=    Del
    '20=    capslock toggled        47=    Hlp
    '27=    Esc                     91, 92=    WinKey
    '32=   " "                      112-123=    F1 - F12
    '33=    PgUp
    '34=    PgDn               (These Are Used for KeyState
    '35=    End                         Detection)
    '36=    Home
    '37=    LeftArrow
    '38=    UpArrow
    '39=    RightArrow
    '-----------------------------------------------------------
           
' Trigger detect functions:
Public Function Shift() As Boolean
Shift = CBool(GetAsyncKeyState(vbKeyShift))
End Function
Public Function Ctrl() As Boolean
    Ctrl = CBool(GetAsyncKeyState(17))
End Function
Public Function Alt() As Boolean
    Alt = CBool(GetAsyncKeyState(18))
End Function
Public Function Caps() As Boolean
Caps = CBool(GetKeyState(vbKeyCapital) And 1)
End Function

Public Function TheKey(ByVal lKey As KeyCodeConstants) As Boolean
  TheKey = GetAsyncKeyState(lKey)
End Function

Private Sub Image1_Click() ' Play
    If SCount(CurFormFocus) = True Then
    Scope(CurFormFocus).PlayIt
    End If
End Sub

Private Sub Image2_Click() ' Pause
If SCount(CurFormFocus) = True Then
Scope(CurFormFocus).PausePlay
End If
End Sub

Private Sub Image3_Click() ' Stop Playing
If SCount(CurFormFocus) = True Then
Scope(CurFormFocus).StopPlay
End If
End Sub

Private Sub Image4_Click() ' Play Loop
If SCount(CurFormFocus) = True Then
Scope(CurFormFocus).PlayLoop
End If
End Sub

Private Sub Image5_Click() ' Zoom In
    If SCount(CurFormFocus) = True Then
    MousePointer = vbHourglass
    Scope(CurFormFocus).ZoomIn
    MousePointer = 0
    End If
End Sub

Private Sub Image6_Click() ' Restore view (Zoom Out)
    If SCount(CurFormFocus) = True Then
    MousePointer = vbHourglass
    Scope(CurFormFocus).ZoomOut
    MousePointer = 0
    End If
End Sub

Private Sub Image7_Click() ' Paste as a New File
    If SCount(CurFormFocus) = True Then
    MousePointer = vbHourglass
    Scope(CurFormFocus).PasteAsNewFile
    MousePointer = 0
    End If
End Sub

Private Sub Image8_Click() ' Cancel Selection
    If SCount(CurFormFocus) = True Then
    Scope(CurFormFocus).Reset
    End If
End Sub


Private Sub MDIForm_Load() ' Load Needed Forms

 frmAbout.Show , Me
 frmAbout.StartFlash
 frmSelection.Show , Me
 
End Sub

Private Sub mnuAbout_Click() ' Show Credits
    frmAbout.Show , Me
End Sub

Private Sub mnuLoadBank_Click() ' Load new Sample Array
    Dim n As Integer
    n = 1
    CommonDialog1.Filter = "Sound Bank Files (*.bnk)|*.bnk"
    CommonDialog1.CancelError = True
    On Error GoTo ErrHandler
    CommonDialog1.ShowOpen
    If CommonDialog1.FileName <> "" Then
        MousePointer = 11 'hourglass
         Open CommonDialog1.FileName For Input As #12
GetLine:
         If EOF(12) = True Then
            Me.Caption = "Sampler (" & CommonDialog1.FileTitle & ") by: Pbryan"
            GoTo ErrHandler
         End If
         
         Input #12, a$
         Debug.Print a$ & " - FileName"
    If a$ = "$" Then n = n + 1: GoTo GetLine
         LoadNewFile a$, a$, n - 1
         
         Input #12, a$
         Debug.Print a$ & " - HotKey"
         Scope(n).txtHotKey = Chr$(Val(a$))
         
        GoTo GetLine
    End If
ErrHandler:
    MousePointer = 0
    Close #12
    Exit Sub
End Sub

Private Sub mnuOpenItem_Click() ' Open a New Sample instance
    CommonDialog1.Filter = "Wave Files (*.wav)|*.wav"
    CommonDialog1.CancelError = True
    On Error GoTo ErrHandler
    CommonDialog1.ShowOpen
    If CommonDialog1.FileName <> "" Then
        MousePointer = 11 'hourglass
        LoadNewFile CommonDialog1.FileName, CommonDialog1.FileTitle, CurFormFocus
        MousePointer = 0 ' Arrow
        Exit Sub
    End If
ErrHandler:
    Exit Sub

End Sub

Private Sub mnuExitItem_Click() ' Quit
    End
End Sub

Private Sub mnuRecord_Click() ' Record new Sound File
    frmRecWav.Show , Me
    frmRecWav.RecordWav App.Path & "\Untitled" & FileCount & ".wav"
    ' This passes file info to the Recording form
End Sub

Private Sub mnuSaveBank_Click() ' Save Current Sample Array
    CommonDialog1.Filter = "Sound Bank Files (*.bnk)|*.bnk"
    CommonDialog1.CancelError = True
    On Error GoTo ErrHandler
    CommonDialog1.ShowSave
    If CommonDialog1.FileName <> "" Then
        Open CommonDialog1.FileName For Output As #12
    For n = 0 To 255
        If SCount(n) = True Then
            Print #12, SFilePath(n)
            Print #12, HotKey(n)
            Print #12, "$"
        End If
    Next n
        Me.Caption = "Sampler (" & CommonDialog1.FileTitle & ") by: Pbryan"
        GoTo ErrHandler
    End If
ErrHandler:
    Close #12
    Exit Sub
               
End Sub

Private Sub Timer1_Timer() ' Check for new Sample Focus
Dim n As Integer           ' and scan for Hotkey triggers
Dim Tmp As Long, Tmp2, num, chrCode
    On Error Resume Next
' if form focus has not changed then check for a pressed Hotkeys
If LastScope = CurFormFocus Then GoTo chkKeys
    If SCount(LastScope) = True Then ' Make sure old window is still open
    Scope(LastScope).ImgActive.Visible = False ' Turn Focus Icon Off for Old location
    End If
    Scope(CurFormFocus).ImgActive.Visible = True ' Turn Focus Icon On for new location
    LastScope = CurFormFocus
    
chkKeys: ' Check for Ctrl, Shift & Hotkey Triggering
For n = 65 To 90 ' A-Z
Tmp = GetAsyncKeyState(n)
Tmp2 = GetAsyncKeyState(vbKeyShift)

If Tmp = -32767 Then

    DoEvents
    If n > 64 And n < 91 Then
            chrCode = n
        End If
     If n < 58 Then
            chrCode = n
       ElseIf (n > 96 & n < 138 & num <> 0) Then
            chrCode = n - 48
       Else
            chrCode = n
       End If
        
    End If
    
Next n
    If chrCode = 0 Then Exit Sub
    
    For n = 0 To 27 ' Scan the first 28 Sample Instances' Hotkey assignments
    
    If SCount(n) = False Then GoTo skipit ' if Sample instance don't exsist then skip to the next one
        
        If Ctrl = True And Shift = True And chrCode = HotKey(n) Then
                ' Ctrl + Shift + Hotkey Detected
                CtrlShiftAction n
                Exit For
        End If
        If Ctrl = True And chrCode = HotKey(n) Then
                ' Ctrl + Hotkey Detected
                CtrlAction n
                Exit For
        End If
        If Shift = True And chrCode = HotKey(n) Then
                ' Shift + Hotkey Detected
                ShiftAction n
                Exit For
        End If
skipit:
    Next n
DoEvents

End Sub
Private Sub CtrlAction(Which As Integer) 'Hotkey Play
        'which = Sample Instance
        Scope(Which).PlayIt
End Sub
Private Sub ShiftAction(Which As Integer) 'Hotkey Loop
        'which = Sample Instance
        Scope(Which).PlayLoop
End Sub
Private Sub CtrlShiftAction(Which As Integer) 'Hotkey Stop
        'which = Sample Instance
        Scope(Which).StopPlay
End Sub


