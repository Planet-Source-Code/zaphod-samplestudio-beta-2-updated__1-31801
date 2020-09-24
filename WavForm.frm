VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form WavForm 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Waveform Viewer"
   ClientHeight    =   1680
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9660
   ControlBox      =   0   'False
   LinkTopic       =   "WavForm"
   MDIChild        =   -1  'True
   MousePointer    =   99  'Custom
   ScaleHeight     =   1680
   ScaleMode       =   0  'User
   ScaleWidth      =   9660
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtHotKey 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   110
      MaxLength       =   1
      TabIndex        =   19
      Top             =   840
      Width           =   615
   End
   Begin VB.Frame Frame3 
      Caption         =   "Effects:"
      Height          =   615
      Left            =   120
      TabIndex        =   13
      Top             =   3960
      Width           =   4935
      Begin VB.CommandButton Command9 
         Caption         =   "Reverse"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   240
         Width           =   1215
      End
   End
   Begin MCI.MMControl MMControl1 
      Height          =   330
      Left            =   3480
      TabIndex        =   4
      Top             =   5040
      Visible         =   0   'False
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   582
      _Version        =   393216
      AutoEnable      =   0   'False
      Enabled         =   0   'False
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1305
      Left            =   840
      ScaleHeight     =   1305
      ScaleWidth      =   8655
      TabIndex        =   0
      Top             =   0
      Width           =   8655
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000001&
         Height          =   645
         Left            =   0
         ScaleHeight     =   615
         ScaleWidth      =   8565
         TabIndex        =   15
         Top             =   600
         Visible         =   0   'False
         Width           =   8595
         Begin VB.PictureBox Picture6 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1240
            Left            =   1320
            ScaleHeight     =   1245
            ScaleWidth      =   1815
            TabIndex        =   16
            Top             =   0
            Visible         =   0   'False
            Width           =   1815
            Begin VB.PictureBox Picture7 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000001&
               ForeColor       =   &H80000005&
               Height          =   1240
               Left            =   480
               ScaleHeight     =   1215
               ScaleWidth      =   585
               TabIndex        =   17
               Top             =   0
               Visible         =   0   'False
               Width           =   615
               Begin VB.Line Line4 
                  BorderColor     =   &H80000005&
                  Visible         =   0   'False
                  X1              =   240
                  X2              =   240
                  Y1              =   0
                  Y2              =   1200
               End
            End
         End
         Begin VB.Line Line2 
            Visible         =   0   'False
            X1              =   600
            X2              =   600
            Y1              =   0
            Y2              =   1200
         End
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H8000000D&
         Height          =   600
         Left            =   0
         ScaleHeight     =   570
         ScaleWidth      =   8565
         TabIndex        =   1
         Top             =   0
         Width           =   8595
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   600
            Left            =   1200
            ScaleHeight     =   600
            ScaleWidth      =   2175
            TabIndex        =   2
            Top             =   0
            Visible         =   0   'False
            Width           =   2175
            Begin VB.PictureBox Picture4 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H8000000D&
               ForeColor       =   &H80000009&
               Height          =   600
               Left            =   600
               ScaleHeight     =   570
               ScaleWidth      =   585
               TabIndex        =   3
               Top             =   0
               Width           =   615
               Begin VB.Line Line3 
                  BorderColor     =   &H80000005&
                  Visible         =   0   'False
                  X1              =   240
                  X2              =   240
                  Y1              =   0
                  Y2              =   2400
               End
            End
         End
         Begin VB.Line Line1 
            Visible         =   0   'False
            X1              =   600
            X2              =   600
            Y1              =   0
            Y2              =   2400
         End
      End
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      X1              =   0
      X2              =   9720
      Y1              =   1640
      Y2              =   1640
   End
   Begin VB.Image ImgActive 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   435
      Left            =   10
      Picture         =   "WavForm.frx":0000
      Stretch         =   -1  'True
      ToolTipText     =   "Current Focus"
      Top             =   10
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   480
      Picture         =   "WavForm.frx":0451
      Stretch         =   -1  'True
      ToolTipText     =   "Close File"
      Top             =   0
      Width           =   345
   End
   Begin VB.Label lblHotKey 
      Alignment       =   2  'Center
      Caption         =   "HotKey:"
      Height          =   255
      Left            =   0
      TabIndex        =   18
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label17 
      Caption         =   "17"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   3000
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label6 
      Caption         =   "6/1"
      Height          =   255
      Index           =   1
      Left            =   4200
      TabIndex        =   11
      Top             =   480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label15 
      Caption         =   "15/1"
      Height          =   255
      Index           =   1
      Left            =   4080
      TabIndex        =   10
      Top             =   840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label15 
      Caption         =   "15/0"
      Height          =   255
      Index           =   0
      Left            =   3960
      TabIndex        =   9
      Top             =   840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "13"
      Height          =   195
      Left            =   2640
      TabIndex        =   8
      Top             =   1080
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "12"
      Height          =   195
      Left            =   1200
      TabIndex        =   7
      Top             =   1080
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Label Label11 
      Caption         =   "11"
      Height          =   255
      Left            =   840
      TabIndex        =   6
      Top             =   1080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label4 
      Caption         =   "4"
      Height          =   255
      Left            =   4920
      TabIndex        =   5
      Top             =   360
      Visible         =   0   'False
      Width           =   255
   End
End
Attribute VB_Name = "WavForm"
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

'Sample data Form (very busy)

Dim IsPlaying As Boolean
Dim xLast As Integer
Dim RepeatIt As Integer, PlayControls As Boolean
Dim MMCPart As Boolean ' Play Switches
Dim MovEsq As Boolean, MovDir As Boolean
Dim IniPlay As Long, FimPlay As Long ' Begin And End Markers
Dim RePlay As Boolean ' for change play selection while playing (Auto re-resume)
Dim SpawnFrom As Integer
Dim Keeper As Double
Dim MvLiner As Integer ' Play Position Holder
Dim ThisFileName As String
Dim ThisFilePath As String, ThisBitRate As String, ThisSampRate As Long
Dim ThisChannels As String, ThisLength As String, ThisSampSelect As Double
Dim ThisNumSamps As Double
Dim ThisSampBegin As Double, ThisTimeBegin As Double
Dim ThisSampEnd As Double, ThisTimeEnd As Double
Dim ThisSampTotal As Double, ThisTimeTotal As Double
Dim LastWidth As Long, LastLeft As Long
Dim FormInstance As Integer ' Allows for Multiple Files to be Opened
Dim CurFile As String ' File Name open this Session
Dim HeightControl As Long ' Max Height of Picture Boxes
Dim HOLDER$
Public Sub LoadFileData(Instance As Integer, Spawned As Integer, _
                        Fpath As String, FName As String)
    'Populates Current Display, with Sample Data from FPath Strings' File.
    CurFile = Fpath
    FormInstance = Instance
    SpawnFrom = Spawned
    Call InitStart
    SFilePath(Instance) = Fpath
    'On Error GoTo Errhandler
    Dim yLec As Long, ydate As Date, ysg As Single
    Dim yint As Integer, ybt As Byte
    Dim LenData As Long, InData As Long
    Dim Nbits As Integer, StMo As String
    Me.Caption = "Waveform View - " & "(" & FName & ")"
    
    ThisFilePath = "File: " & Fpath
    ThisFileName = FName
    Open FName For Binary Access Read As #1
    Label4.Caption = 0 'IniPlay of present Zoom
    Label17.Caption = 0 '= selected Samples
    For n = 1 To 100
        X$ = Input(4, #1)
    If n = 2 Then HOLDER$ = X$ ' Hold This for Saving a New Wav
    If X$ = "fmt " Then Exit For 'Ignore everything else till this
    Next n
    'Get the Wave File Header Info
    Get #1, , yLec ' 16 - don't know what this is for, so ignore it.
    Get #1, , yint 'Compression Type (1=PCM)
    Get #1, , yint 'is Channels, 1 if mono and 2 if stereo

    If yint = 2 Then
        ThisChannels = "Stereo"
      ElseIf yint = 1 Then
        ThisChannels = "Mono"
      Else
        ThisChannels = "Error!"
        GoTo ErrHandler
    End If
    Get #1, , yLec 'is the Sampling frequency of the file

    ThisSampRate = yLec
    Get #1, , yLec 'is a multiple of the sample frequency

    Get #1, , yint 'is the divisor of the number of bytes of
          'data which gives the number of Samples in the .wav
    yDiv = yint
    Label12.Caption = yDiv
    Get #1, , yint 'is the number of bits (8 or 16)

    If yint = 8 Or yint = 16 Then
        ThisBitRate = yint & " bit"
      Else
        ThisBitRate = "Error"
        GoTo ErrHandler
    End If
GotTheData:
    For n = 1 To 100 ' Seek for start of Wav Data
        Y$ = Input(1, #1)

        If Y$ = "d" Then Exit For ' Ignore everything till this

    Next n
    Z$ = Input(3, #1)
  If Z$ <> "ata" Then 'Sample Data Layer Starts when this is found
        If n > 90 Then GoTo ErrHandler
        Temp = Seek(1)
        Seek #1, Temp - 3
        GoTo GotTheData
  End If
    Get #1, , yLec '= num of bytes of data, start reading data here.
                   ' Sample Data Follows after this...
    Label13.Caption = yLec
    LenData = yLec / yDiv
    ThisNumSamps = LenData
    Label6(1).Caption = LenData
    LenTemp = LenData / (ThisSampRate)
    Extemp = (Int(LenTemp * 1000)) / 1000
    If LenTemp - Extemp >= 0.0005 Then
        Extemp = Extemp + 0.001
    End If
    ThisLength = "Length: " & Extemp & " seconds"
    Label15(0).Caption = LenTemp
    Label15(1).Caption = LenTemp
    FimPlay = Int(LenTemp * 1000)
    InData = Seek(1) 'Loc(1) + 1 is the number of the first sound data byte of the file.
    Label11.Caption = InData
    StMo = ThisChannels
    Nbits = Val(ThisBitRate)
    MousePointer = vbHourglass
    
    Call GraphWave(InData, LenData, Nbits, StMo)
    
    Close #1
        MousePointer = 0
        PlayControls = True
    Exit Sub
ErrHandler:
    MsgBox "Error!!", vbOKOnly
    Close #1
    Call InitStart
  
    Exit Sub
End Sub

Sub InitStart() ' Initialize Sample Positons & Captions
    Caption = ""
    Cls
    Me.Left = 10
    Me.Top = (SpawnFrom * (Me.Height)) + 2
        
    IniPlay = 0
    RepeatIt = 0
    PlayControls = False
    MMCPart = True
    Picture3.Width = 0
    Picture3.Visible = False
    Picture6.Width = 0
    Picture6.Visible = False
    Picture2.Cls: Picture5.Cls: Picture4.Cls: Picture7.Cls
    Label6(1).Caption = "": ThisChannels = ""
    Label11.Caption = "": Label12.Caption = "": Label13.Caption = ""
    

End Sub

Public Sub PasteAsNewFile() ' Paste as New File
    Dim Resp As String ' Filename from InputBox
    Dim InData As Long, LenData As Long
    Dim InDataSel As Long, LenDataSel As Long
    Dim SampIni As Long, BytInic As Long
    Dim Nbits As Integer, StMo As String
    Dim yDiv As Integer, SampFreq As Long
    
    Resp = InputBox("File Name", "Select Filename", App.Path & "\Untitled" & FileCount & ".wav")
    
    If MMCPart = False Then
        Picture2.SetFocus
        Exit Sub
    End If
    If ThisSampBegin = 0 And ThisSampEnd = 0 Then
        msg = "No Selection Made!"
        MsgBox msg, vbOKOnly
        Exit Sub
    End If
    
    Picture2.SetFocus
    FileCount = FileCount + 1
    
    
    InData = Label11.Caption
    SampIni = ThisSampBegin ' Selection Begins
    yDiv = Label12.Caption ' FileSize in Bytes
    BytInic = SampIni * yDiv ' Location in the Wav of the Visible Selection
    InDataSel = InData + BytInic
    LenDataSel = ThisSampTotal 'Selection Length in Samples
    SampFreq = ThisSampRate ' Sampling Frequency
    StMo = ThisChannels ' Stereo Mono
    Nbits = Val(ThisBitRate) ' Sampling Bits
    
    Open CurFile For Binary Access Read As #1
    
    Call SaveWave(Resp, InDataSel, LenDataSel, SampFreq, Nbits, StMo)
    
    Close #1
        LoadNewFile Resp, Resp, FormInstance
End Sub

Public Sub ZoomIn() ' Zoom on Selected
    Dim InData As Long, LenData As Long
    Dim InDataSel As Long, LenDataSel As Long
    Dim SampIni As Long, BytInic As Long
    Dim Nbits As Integer, StMo As String
    Dim yDiv As Integer
    If MMCPart = False Then
        Picture2.SetFocus
        Exit Sub
    End If
    If ThisSampBegin = 0 Then
        msg = "No Zoom Selection Made"
        MsgBox msg, vbOKOnly
        Exit Sub
    End If
    
    Picture2.SetFocus
    Picture3.Width = 0
    Picture3.Visible = False
    Picture6.Width = 0
    Picture6.Visible = False
    Picture2.Cls: Picture5.Cls: Picture4.Cls: Picture7.Cls
    InData = Label11.Caption
    SampIni = ThisSampBegin
    yDiv = Label12.Caption
    BytInic = SampIni * yDiv
    InDataSel = InData + BytInic
    LenDataSel = ThisSampTotal + 1
    Open CurFile For Binary Access Read As #1
    StMo = ThisChannels
    Nbits = Val(ThisBitRate)
    Call GraphWave(InDataSel, LenDataSel, Nbits, StMo)
    Close #1
    Label4.Caption = IniPlay 'IniPlay of present zoom, without selection
    Label6(1).Caption = LenDataSel 'LenData of present zoom
    Label15(1).Caption = LenDataSel / ThisSampRate
    'Label8 contains file frequency
    'Label15(1) will be LenTemp of actual zoom
    Label17.Caption = SampIni

End Sub

Public Sub ZoomOut() ' Restore The "Whole" Wav View
    Dim InData As Long, LenData As Long
    Dim Nbits As Integer, StMo As String
    Dim LenTemp As Double
    If MMCPart = False Then
        Picture2.SetFocus
        Exit Sub
    End If
    If ThisSampBegin = 0 Then
        Picture2.SetFocus
        Exit Sub
    End If
    
    Picture2.SetFocus
    IniPlay = 0
    Label4.Caption = 0
    Label15(1).Caption = Label15(0).Caption
    Label6(1).Caption = ThisNumSamps
    Label17.Caption = 0
    RepeatIt = 0
    Picture3.Width = 0
    Picture3.Visible = False
    Picture6.Width = 0
    Picture6.Visible = False
    Picture2.Cls
    Picture5.Cls
    Picture4.Cls
    Picture7.Cls
    InData = Label11.Caption
    LenData = ThisNumSamps
    Open CurFile For Binary Access Read As #1
    StMo = ThisChannels
    Nbits = Val(ThisBitRate)
    Call GraphWave(InData, LenData, Nbits, StMo)
    Close #1
    LenTemp = Label15(1).Caption
    FimPlay = Int(LenTemp * 1000)
    PlayControls = True
    Call Reset

End Sub

Public Sub Reset() ' Reset Selection
    Picture2.SetFocus
    If PlayControls = False Then Exit Sub
    If MMCPart = False Then Exit Sub
    Picture3.Width = 0
    Picture3.Visible = False
    Picture6.Width = 0
    Picture6.Visible = False
    IniPlay = Label4.Caption
    FimPlay = IniPlay + Label15(1).Caption * 1000
    If ThisTimeBegin <> 0 Then
        ThisTimeBegin = IniPlay / 1000
        ThisTimeEnd = FimPlay / 1000
        ThisTimeTotal = (FimPlay - IniPlay) / 1000
        ThisSampBegin = Label17.Caption
        ThisSampEnd = Val(Label6(1).Caption) + Val(Label17.Caption)
        ThisSampTotal = Label6(1).Caption 'Data Length of Zoom
        
        Call PosChange
    End If
End Sub

Private Sub Form_Activate() ' Change Focus Info
    CurFormFocus = FormInstance
    MDIMain.lblSelectedSamples = ThisSampSelect
    MDIMain.lblFilePath.Caption = ThisFilePath
    MDIMain.lblSecLen.Caption = ThisLength
    MDIMain.lblBitRate.Caption = ThisBitRate
    MDIMain.lblChannels.Caption = ThisChannels
    MDIMain.lblSampRate.Caption = ThisSampRate
    frmSelection.lblHotKey.Caption = "HotKey: " & Chr$(HotKey(FormInstance))
    frmSelection.Caption = ThisFileName
    frmSelection.lblSampBegin.Caption = ThisSampBegin
    frmSelection.lblTimeBegin.Caption = ThisTimeBegin
    frmSelection.lblSampEnd.Caption = ThisSampEnd
    frmSelection.lblTimeEnd.Caption = ThisTimeEnd
    frmSelection.lblSampTotal.Caption = ThisSampTotal
    frmSelection.lblTimeTotal.Caption = ThisTimeTotal
    
End Sub

Private Sub Form_Load() ' Ini
    Me.Icon = MDIMain.Icon
    Me.Caption = "Waveform View"
    Me.Height = 1682
    Me.Width = MDIMain.Width - 300
    Line5.X2 = Me.Width
    HeightControl = 800
    Picture2.Height = HeightControl
    Picture5.Height = HeightControl
    Picture7.Height = HeightControl
    
    Picture4.Width = Picture2.Width
    Picture7.Width = Picture2.Width
    MMControl1.Notify = False
    MMControl1.Wait = True
    MMControl1.Shareable = False ' Set properties needed by MCI to open.
    MMControl1.DeviceType = "WaveAudio"
    MMControl1.TimeFormat = mciFormatMilliseconds
End Sub

Private Sub Form_Resize() ' Fix Width & Height
    
    Picture1.Height = (HeightControl * 2) + 2

    
    Picture1.Width = Me.Width - 20
    Picture2.Width = Picture1.Width - 60
    Picture5.Width = Picture2.Width
    Picture4.Width = Picture2.Width
    Picture3.Width = Picture2.Width
    Picture6.Width = Picture2.Width
    Picture7.Width = Picture2.Width
    
End Sub


Private Sub Form_Unload(Cancel As Integer) ' Clean Up
    MMControl1.Command = "Close"
    Close #1
    HotKey(FormInstance) = 0
    SCount(FormInstance) = False
    SFilePath(FormInstance) = ""
End Sub

Public Sub PlayIt() 'Play
        
    If IsPlaying = True Then
            IsPlaying = False
            RepeatIt = 0
            MMControl1.Command = "Stop"
            PlayControls = True
            DoEvents
    End If
    If PlayControls = False Then
        Picture2.SetFocus
        Exit Sub
    End If
    IsPlaying = True
    RepeatIt = 0
    Call MMControl1_PlayClick(False)
    MMControl1.Command = "Play"
    Picture2.SetFocus
End Sub

Public Sub PausePlay() ' Pause
    If PlayControls = False Or MMCPart = True Then
        Picture2.SetFocus
        Exit Sub
    End If
    MMControl1.Command = "Pause"
    Picture2.SetFocus
End Sub

Public Sub StopPlay() ' Stop
    If PlayControls = False Then
        Picture2.SetFocus
        Exit Sub
    End If
    IsPlaying = False
    RepeatIt = 0
    MMControl1.Command = "Stop"
    Picture2.SetFocus
End Sub

Public Sub PlayLoop() ' Play Looped
    If IsPlaying = True Then
            IsPlaying = False
            RepeatIt = 0
            MMControl1.Command = "Stop"
            PlayControls = True
            DoEvents
    End If
    If PlayControls = False Then
        Picture2.SetFocus
        Exit Sub
    End If
    IsPlaying = True
    Call MMControl1_PlayClick(False)
    RepeatIt = 1
    MMControl1.Command = "Play"
    Picture2.SetFocus
End Sub

Private Sub Image1_Click() ' Close
    Unload Me
End Sub

Private Sub MMControl1_Done(NotifyCode As Integer)
    MMControl1.UpdateInterval = 0
    If RepeatIt = 1 Then ' Play Selection Again
        Call MMControl1_PlayClick(False)
        MMControl1.Command = "Play"
        Exit Sub
      Else ' Stop
       IsPlaying = False
    End If
    MMControl1.Command = "Close"
    Line1.Visible = False
    Line2.Visible = False
    Line3.Visible = False
    Line4.Visible = False
    MMCPart = True
End Sub

Private Sub MMControl1_PlayClick(Cancel As Integer) ' Play
    If RepeatIt = 1 Then ' Loop Mode
        MMControl1.UpdateInterval = 50
        MMControl1.From = IniPlay&
        MMControl1.To = FimPlay&
        'Track Playing Position
        Line1.X1 = MvLiner
        Line1.X2 = MvLiner
        Line2.X1 = MvLiner
        Line2.X2 = MvLiner
        Line3.X1 = MvLiner
        Line3.X2 = MvLiner
        Line4.X1 = MvLiner
        Line4.X2 = MvLiner
        Exit Sub
    End If
    ' Single Mode
    MMControl1.FileName = CurFile
    MMControl1.Command = "Open"
    MMControl1.From = IniPlay&
    MMControl1.To = FimPlay&
    MMControl1.UpdateInterval = 50
    LenTemp = Label15(1).Caption
    Keeper = Picture2.ScaleWidth / (LenTemp * 1000)
    MvLiner = Int((IniPlay - Label4.Caption) * Keeper)
'Track Playing Position
    Line1.X1 = MvLiner
    Line1.X2 = MvLiner
    Line2.X1 = MvLiner
    Line2.X2 = MvLiner
    Line3.X1 = MvLiner
    Line3.X2 = MvLiner
    Line4.X1 = MvLiner
    Line4.X2 = MvLiner
    If FimPlay - IniPlay > 500 Then
        If Picture3.Width > 100 Then
            Line3.Visible = True
            Line4.Visible = True
          Else
            Line1.Visible = True
            Line2.Visible = True
        End If
    End If
    MMCPart = False
End Sub

Private Sub MMControl1_StatusUpdate() ' Mark Play Position
    Z = Int((MMControl1.Position - Label4.Caption) * Keeper)
    Line1.X1 = Z
    Line1.X2 = Z
    Line2.X1 = Z
    Line2.X2 = Z
    Line3.X1 = Z
    Line3.X2 = Z
    Line4.X1 = Z
    Line4.X2 = Z
End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Mono or Left Channel
    If IsPlaying = True Then
            IsPlaying = False
            RepeatIt = 0
            MMControl1.Command = "Stop"
            MMCPart = True
            PlayControls = True
            RePlay = True
            DoEvents
    End If
    If PlayControls = False Then Exit Sub
    If MMCPart = False Then Exit Sub
    If Button = vbRightButton Then
        If Picture3.Width < 40 Then Exit Sub
        If Picture3.Left - X < 50 And Picture3.Left - X > 0 Then
            Picture2.MousePointer = 9
            MovEsq = True
            LastWidth = Picture3.Width
            LastLeft = Picture3.Left
        End If
        If (X - Picture3.Left - Picture3.Width) < 50 And (X - Picture3.Left - Picture3.Width) > 0 Then
            Picture2.MousePointer = 9
            MovDir = True
        End If
    End If
    If Button = vbLeftButton Then
        Picture3.Width = 0
        Picture3.Left = X
        Picture3.Visible = True
        Picture4.Left = -X
        Picture6.Width = 0
        Picture6.Left = X
        Picture6.Visible = True
        Picture7.Left = -X
        IniPlay = Label4.Caption + X * Label15(1).Caption * 1000 / Picture2.ScaleWidth
        FimPlay = Label4.Caption + Label15(1).Caption * 1000
        ThisTimeBegin = IniPlay / 1000
        ThisTimeEnd = FimPlay / 1000
        ThisTimeTotal = (FimPlay - IniPlay) / 1000
        ThisSampBegin = Label17.Caption + Int(X * (Label6(1).Caption - 1) / Picture2.ScaleWidth)
        ThisSampEnd = Val(Label17.Caption) + Val(Label6(1).Caption) - 1
        ThisSampTotal = ThisSampEnd - ThisSampBegin + 1
        Call PosChange
    End If
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Mono or Left Channel
    If Picture3.Visible = False Then Exit Sub
    If MMCPart = False Then Exit Sub
    If Button = vbLeftButton Then
        If X > Picture3.Left And X <= Picture2.ScaleWidth Then
            Picture3.Width = X - Picture3.Left
            Picture6.Width = X - Picture6.Left
            FimPlay = Label4.Caption + X * Label15(1).Caption * 1000 / Picture2.ScaleWidth
            ThisTimeEnd = FimPlay / 1000
            ThisTimeTotal = (FimPlay - IniPlay) / 1000
            ThisSampEnd = Label17.Caption + Int(X * (Label6(1).Caption - 1) / Picture2.ScaleWidth)
            ThisSampTotal = ThisSampEnd - ThisSampBegin + 1
            Call PosChange
        End If
    End If
    If Button = vbRightButton Then
        If X = xLast Then
            Picture4.Visible = True
            Picture7.Visible = True
            Exit Sub
          ElseIf MovEsq = True Then
            Call ShadeArea(X)
            xLast = X
            If X >= 0 Then
                IniPlay = Label4.Caption + Picture3.Left * Label15(1).Caption * 1000 / Picture2.ScaleWidth
                ThisTimeBegin = IniPlay / 1000
                ThisTimeTotal = (FimPlay - IniPlay) / 1000
                ThisSampBegin = Label17.Caption + Int(Picture3.Left * (Label6(1).Caption - 1) / Picture2.ScaleWidth)
                ThisSampTotal = ThisSampEnd - ThisSampBegin + 1
                Call PosChange
            End If
          ElseIf MovDir = True And X - Picture3.Left > 50 And X <= Picture2.ScaleWidth Then
            Picture3.Width = X - Picture3.Left
            Picture6.Width = Picture3.Width
            FimPlay = Label4.Caption + X * Label15(1).Caption * 1000 / Picture2.ScaleWidth
            ThisTimeEnd = FimPlay / 1000
            ThisTimeTotal = (FimPlay - IniPlay) / 1000
            ThisSampEnd = Label17.Caption + Int(X * (Label6(1).Caption - 1) / Picture2.ScaleWidth)
            ThisSampTotal = ThisSampEnd - ThisSampBegin + 1
            Call PosChange
        End If
    End If
End Sub
Sub ShadeArea(X)
    If LastWidth + LastLeft - X < 50 Then Exit Sub
    Picture4.Visible = False
    Picture7.Visible = False
    Picture3.Left = X
    Picture6.Left = X
    Picture4.Left = -X
    Picture7.Left = -X
    Picture3.Width = LastWidth + LastLeft - X
    Picture6.Width = Picture3.Width
    Picture4.Visible = True
    Picture7.Visible = True

End Sub
Private Sub Picture2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Mono or Left Channel
    If RePlay = True Then
        RePlay = False
        PlayLoop
    End If
    Picture2.MousePointer = 0
    MovEsq = False
    MovDir = False
End Sub

Private Sub Picture4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Mono or Left Channel
    If IsPlaying = True Then
            IsPlaying = False
            RepeatIt = 0
            MMControl1.Command = "Stop"
            MMCPart = True
            PlayControls = True
            RePlay = True
            DoEvents
    End If

    If MMCPart = False Then Exit Sub
    If Button = vbRightButton Then
        If Picture3.Width < 100 Then
            Picture4.MousePointer = 9
            If X - Picture3.Left < Picture3.Width / 3 Then
                MovEsq = True
                LastWidth = Picture3.Width
                LastLeft = Picture3.Left
              Else
                MovDir = True
            End If
          ElseIf X - Picture3.Left < 50 Then
            Picture4.MousePointer = 9
            MovEsq = True
            LastWidth = Picture3.Width
            LastLeft = Picture3.Left
          ElseIf Picture3.Width + Picture3.Left - X < 100 Then
            Picture4.MousePointer = 9
            MovDir = True
        End If
    End If
    If Button = vbLeftButton Then
        Picture3.Width = 0
        Picture3.Left = X
        Picture3.Visible = True
        Picture4.Left = -X
        Picture6.Width = 0
        Picture6.Left = X
        Picture6.Visible = True
        Picture7.Left = -X
        IniPlay = Label4.Caption + X * Label15(1).Caption * 1000 / Picture2.ScaleWidth
        FimPlay = Label4.Caption + Label15(1).Caption * 1000
        ThisTimeBegin = IniPlay / 1000
        ThisTimeEnd = FimPlay / 1000
        ThisTimeTotal = (FimPlay - IniPlay) / 1000
        ThisSampBegin = Label17.Caption + Int(X * (Label6(1).Caption - 1) / Picture2.ScaleWidth)
        ThisSampEnd = Val(Label17.Caption) + Val(Label6(1).Caption) - 1
        ThisSampTotal = ThisSampEnd - ThisSampBegin + 1
        Call PosChange
    End If

End Sub

Private Sub Picture4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Mono or Left Channel
    If MMCPart = False Then Exit Sub
    If Button = vbRightButton Then
        If X = xLast Then
            Picture4.Visible = True
            Picture7.Visible = True
            Exit Sub
          ElseIf MovEsq = True Then
            Call ShadeArea(X)
            xLast = X
            If X >= 0 Then
                IniPlay = Label4.Caption + Picture3.Left * Label15(1).Caption * 1000 / Picture2.ScaleWidth
                ThisTimeBegin = IniPlay / 1000
                ThisTimeTotal = (FimPlay - IniPlay) / 1000
                ThisSampBegin = Label17.Caption + Int(Picture3.Left * (Label6(1).Caption - 1) / Picture2.ScaleWidth)
                ThisSampTotal = ThisSampEnd - ThisSampBegin + 1
                Call PosChange
            End If
          ElseIf MovDir = True And X - Picture3.Left > 50 And X <= Picture2.ScaleWidth Then
            Picture3.Width = X - Picture3.Left
            Picture6.Width = Picture3.Width
            FimPlay = Label4.Caption + X * Label15(1).Caption * 1000 / Picture2.ScaleWidth
            ThisTimeEnd = FimPlay / 1000
            ThisTimeTotal = (FimPlay - IniPlay) / 1000
            ThisSampEnd = Label17.Caption + Int(X * (Label6(1).Caption - 1) / Picture2.ScaleWidth)
            ThisSampTotal = ThisSampEnd - ThisSampBegin + 1
            Call PosChange
        End If
    End If

End Sub

Private Sub Picture4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Mono or Left Channel
    If RePlay = True Then
        RePlay = False
        PlayLoop
    End If
    Picture4.MousePointer = 0
    MovDir = False
    MovEsq = False
End Sub

Private Sub Picture5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Right Channel
    If IsPlaying = True Then
            IsPlaying = False
            RepeatIt = 0
            MMControl1.Command = "Stop"
            MMCPart = True
            PlayControls = True
            RePlay = True
            DoEvents
    End If
    If PlayControls = False Then Exit Sub
    If MMCPart = False Then Exit Sub
    If Button = vbRightButton Then
        If Picture3.Width < 40 Then Exit Sub
        If Picture3.Left - X < 50 And Picture3.Left - X > 0 Then
            Picture5.MousePointer = 9
            MovEsq = True
            LastWidth = Picture3.Width
            LastLeft = Picture3.Left
        End If
        If (X - Picture3.Left - Picture3.Width) < 50 And (X - Picture3.Left - Picture3.Width) > 0 Then
            Picture5.MousePointer = 9
            MovDir = True
        End If
    End If
    If Button = vbLeftButton Then
        Picture3.Width = 0
        Picture3.Left = X
        Picture3.Visible = True
        Picture4.Left = -X
        Picture6.Width = 0
        Picture6.Left = X
        Picture6.Visible = True
        Picture7.Left = -X
        IniPlay = Label4.Caption + X * Label15(1).Caption * 1000 / Picture2.ScaleWidth
        FimPlay = Label4.Caption + Label15(1).Caption * 1000
        ThisTimeBegin = IniPlay / 1000
        ThisTimeEnd = FimPlay / 1000
        ThisTimeTotal = (FimPlay - IniPlay) / 1000
        ThisSampBegin = Label17.Caption + Int(X * (Label6(1).Caption - 1) / Picture2.ScaleWidth)
        ThisSampEnd = Val(Label17.Caption) + Val(Label6(1).Caption) - 1
        ThisSampTotal = ThisSampEnd - ThisSampBegin + 1
        Call PosChange
    End If
End Sub

Private Sub Picture5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Right Channel
    If Picture3.Visible = False Then Exit Sub
    If MMCPart = False Then Exit Sub
    If Button = vbLeftButton Then
        If X > Picture3.Left And X <= Picture2.ScaleWidth Then
            Picture3.Width = X - Picture3.Left
            Picture6.Width = X - Picture6.Left
            FimPlay = Label4.Caption + X * Label15(1).Caption * 1000 / Picture2.ScaleWidth
            ThisTimeEnd = FimPlay / 1000
            ThisTimeTotal = (FimPlay - IniPlay) / 1000
            ThisSampEnd = Label17.Caption + Int(X * (Label6(1).Caption - 1) / Picture2.ScaleWidth)
            ThisSampTotal = ThisSampEnd - ThisSampBegin + 1
            Call PosChange
        End If
    End If
    If Button = vbRightButton Then
        If X = xLast Then
            Picture4.Visible = True
            Picture7.Visible = True
            Exit Sub
          ElseIf MovEsq = True Then
            Call ShadeArea(X)
            xLast = X
            If X >= 0 Then
                IniPlay = Label4.Caption + Picture3.Left * Label15(1).Caption * 1000 / Picture2.ScaleWidth
                ThisTimeBegin = IniPlay / 1000
                ThisTimeTotal = (FimPlay - IniPlay) / 1000
                ThisSampBegin = Label17.Caption + Int(Picture3.Left * (Label6(1).Caption - 1) / Picture2.ScaleWidth)
                ThisSampTotal = ThisSampEnd - ThisSampBegin + 1
                Call PosChange
            End If
          ElseIf MovDir = True And X - Picture3.Left > 50 And X <= Picture2.ScaleWidth Then
            Picture3.Width = X - Picture3.Left
            Picture6.Width = Picture3.Width
            FimPlay = Label4.Caption + X * Label15(1).Caption * 1000 / Picture2.ScaleWidth
            ThisTimeEnd = FimPlay / 1000
            ThisTimeTotal = (FimPlay - IniPlay) / 1000
            ThisSampEnd = Label17.Caption + Int(X * (Label6(1).Caption - 1) / Picture2.ScaleWidth)
            ThisSampTotal = ThisSampEnd - ThisSampBegin + 1
            Call PosChange
        End If
    End If
End Sub

Sub PosChange() 'Selection Change
    Dim Npont As Long
    If ThisSampTotal = 0 Then
        ThisSampSelect = 0
        Exit Sub
    End If
    frmSelection.lblSampBegin.Caption = ThisSampBegin
    frmSelection.lblTimeBegin.Caption = ThisTimeBegin
    frmSelection.lblSampEnd.Caption = ThisSampEnd
    frmSelection.lblTimeEnd.Caption = ThisTimeEnd
    frmSelection.lblSampTotal.Caption = ThisSampTotal
    frmSelection.lblTimeTotal.Caption = ThisTimeTotal
    Npont = ThisSampTotal
    ThisSampSelect = Npont
    MDIMain.lblSelectedSamples = ThisSampSelect
End Sub
Public Sub GraphWave(InData As Long, LenData As Long, Nbits As Integer, _
                StMo As String)
    Dim yByte As Byte
    Dim yzero As Double, xmax As Double, xmult As Double, ySelFat As Double
    Dim yint As Integer, yPos As Integer, yGraf As Integer
    Dim limsup As Integer
    Dim ySel As Long
    Dim nMult As Double, xPos As Integer
    
    If StMo = "Stereo" Then
        Picture2.Height = HeightControl
        Picture5.Visible = True
        Picture5.Top = (Picture2.Top + Picture2.Height) + 20
        Picture6.Visible = True
        Picture7.Visible = True
        Picture3.Height = HeightControl
        Picture4.Height = HeightControl
      Else
        Picture2.Height = HeightControl * 2
        Picture5.Visible = False
        Picture3.Height = HeightControl * 2
        Picture4.Height = HeightControl * 2
    End If
    ySelFat = LenData / Picture2.ScaleWidth
    xzero = 0
    yzero = Picture2.ScaleHeight / 2
    xmax = Picture2.ScaleWidth
    ymax = 128
    ymaxgraf = Picture2.ScaleHeight * 3 / 8
    ymult = ymaxgraf / ymax
    yPos = Int(yzero + 15 * 128)
    Picture2.Line (0, yzero)-(xmax, yzero)
    Picture2.Line (0, yzero - ymax * ymult)-(xmax, yzero - ymax * ymult), &H8000000F
    Picture2.Line (0, yzero + ymax * ymult)-(xmax, yzero + ymax * ymult), &H8000000F
    Picture4.Line (0, yzero)-(xmax, yzero)
    Picture4.Line (0, yzero - ymax * ymult)-(xmax, yzero - ymax * ymult), &H8000000F
    Picture4.Line (0, yzero + ymax * ymult)-(xmax, yzero + ymax * ymult), &H8000000F
    If StMo = "Stereo" Then GoTo Stereo8
    If Nbits = 16 Then GoTo Mono16
Mono8:
    Get #1, InData, yByte
    yGraf = (yPos - 15 * yByte)
    Picture2.PSet (xzero, yGraf)
    Picture4.PSet (xzero, yGraf)
    If LenData <= Picture2.ScaleWidth Then
        nMult = Picture2.ScaleWidth / LenData
        For n = 1 To LenData - 1
            xPos = Int(n * nMult)
            Get #1, , yByte
            yGraf = (yPos - 15 * yByte)
            Picture2.Line -(xPos, yGraf)
            Picture4.Line -(xPos, yGraf)
        Next n
      Else
        For n = 1 To Picture2.ScaleWidth
            ySel = InData + Int(n * ySelFat)
            Get #1, ySel, yByte
            yGraf = (yPos - 15 * yByte)
            Picture2.Line -(n, yGraf)
            Picture4.Line -(n, yGraf)
        Next n
    End If
    GoTo Done
Mono16:
    Get #1, InData, yint
    yGraf = yzero - yint / 17
    Picture2.PSet (xzero, yGraf)
    Picture4.PSet (xzero, yGraf)
    If LenData <= Picture2.ScaleWidth Then
        nMult = Picture2.ScaleWidth / LenData
        For n = 1 To LenData - 1
            xPos = Int(n * nMult)
            Get #1, , yint
            yGraf = yzero - yint / 17
            Picture2.Line -(xPos, yGraf)
            Picture4.Line -(xPos, yGraf)
        Next n
      Else
        For n = 1 To Picture2.ScaleWidth
            ySel = InData + 2 * Int(n * ySelFat)
            Get #1, ySel, yint
            yGraf = yzero - yint / 17
            Picture2.Line -(n, yGraf)
            Picture4.Line -(n, yGraf)
        Next n
    End If
    GoTo Done
Stereo8:
    Picture2.CurrentX = 0
    Picture2.CurrentY = 0
    Picture2.Print "Left"
    Picture4.CurrentX = 0
    Picture4.CurrentY = 0
    Picture4.Print "Left"
    Picture5.Line (0, yzero)-(xmax, yzero)
    Picture5.Line (0, yzero - ymax * ymult)-(xmax, yzero - ymax * ymult), &H8000000F
    Picture5.Line (0, yzero + ymax * ymult)-(xmax, yzero + ymax * ymult), &H8000000F
    Picture5.CurrentX = 0
    Picture5.CurrentY = 0
    Picture5.Print "Right"
    Picture7.Line (0, yzero)-(xmax, yzero)
    Picture7.Line (0, yzero - ymax * ymult)-(xmax, yzero - ymax * ymult), &H8000000F
    Picture7.Line (0, yzero + ymax * ymult)-(xmax, yzero + ymax * ymult), &H8000000F
    Picture7.CurrentX = 0
    Picture7.CurrentY = 0
    Picture7.Print "Right"
    If Nbits = 16 Then GoTo Stereo16
    yPos = Int(yzero + 7 * 128)
    Get #1, InData, yByte 'left Channel
    yGraf = yPos - 7 * yByte '15 * yByte
    Picture2.PSet (xzero, yGraf)
    Picture4.PSet (xzero, yGraf)
    Get #1, , yByte 'right Channel
    yGraf = yPos - 7 * yByte
    Picture5.PSet (xzero, yGraf)
    Picture7.PSet (xzero, yGraf)
    If LenData <= Picture2.ScaleWidth Then
        nMult = Picture2.ScaleWidth / LenData
        For n = 1 To LenData - 1
            xPos = Int(n * nMult)
            Get #1, , yByte 'left Channel
            yGraf = yPos - 7 * yByte '15 * yByte
            Picture2.Line -(xPos, yGraf)
            Picture4.Line -(xPos, yGraf)
            Get #1, , yByte 'right Channel
            yGraf = yPos - 7 * yByte
            Picture5.Line -(xPos, yGraf)
            Picture7.Line -(xPos, yGraf)
        Next n
      Else
        For n = 1 To Picture2.ScaleWidth
            ySel = InData + 2 * Int(n * ySelFat)
            Get #1, ySel, yByte 'left Channel
            yGraf = yPos - 7 * yByte '15 * yByte
            Picture2.Line -(n, yGraf)
            Picture4.Line -(n, yGraf)
            Get #1, , yByte 'right Channel
            yGraf = yPos - 7 * yByte
            Picture5.Line -(n, yGraf)
            Picture7.Line -(n, yGraf)
        Next n
    End If
    GoTo Done

Stereo16:
    Get #1, InData, yint 'left Channel
    yGraf = yzero - yint / 35 '17
    Picture2.PSet (xzero, yGraf)
    Picture4.PSet (xzero, yGraf)
    Get #1, , yint 'right Channel
    yGraf = yzero - yint / 35
    Picture5.PSet (xzero, yGraf)
    Picture7.PSet (xzero, yGraf)
    If LenData <= Picture2.ScaleWidth Then
        nMult = Picture2.ScaleWidth / LenData
        For n = 1 To LenData - 1
            xPos = Int(n * nMult)
            Get #1, , yint 'left Channel
            yGraf = yzero - yint / 35 '17
            Picture2.Line -(xPos, yGraf)
            Picture4.Line -(xPos, yGraf)
            Get #1, , yint 'right Channel
            yGraf = yzero - yint / 35
            Picture5.Line -(xPos, yGraf)
            Picture7.Line -(xPos, yGraf)
        Next n
      Else
        For n = 1 To Picture2.ScaleWidth
            ySel = InData + 4 * Int(n * ySelFat)
            Get #1, ySel, yint 'left Channel
            yGraf = yzero - yint / 35 '17
            Picture2.Line -(n, yGraf)
            Picture4.Line -(n, yGraf)
            Get #1, , yint 'right Channel
            yGraf = yzero - yint / 35
            Picture5.Line -(n, yGraf)
            Picture7.Line -(n, yGraf)
        Next n
    End If
    
Done:

End Sub
Private Sub WriteHeader(Chan As Integer, SampFreq As Long, Nbits As Integer, LenData As Long)
    
    Dim TmpR As Long
    Put #2, , "RIFF" ' RIFF Header Layer
    Put #2, 5, HOLDER$
    Put #2, 9, "WAVE" ' WAVE Header Layer
    Put #2, 13, "fmt "
    Put #2, 17, 16 '16
    Put #2, 21, 1 ' Compression (None=1(PCM))
    
    Put #2, 23, Chan ' Channels 1 or 2
    Put #2, 25, SampFreq ' Sampling Rate
    TmpR = SampFreq * (Chan * (Nbits / 8))
    Put #2, 29, TmpR '  Calculation
    TmpR = (Nbits / 8) * Chan
    Put #2, 33, TmpR 'Calculation
    Put #2, 35, Nbits ' Sampling bits
             ' End of WAVE Header Layer
    Put #2, 37, "data" ' Sound Data Layer
    Put #2, , LenData * TmpR ' Number of Samples in Wav
    
End Sub

Public Sub SaveWave(FName As String, InData As Long, LenData As Long, _
                    SampFreq As Long, Nbits As Integer, StMo As String)
    
    Dim ChanOut As Integer
    Dim yByte As Byte
    Dim yint As Integer
        
    Open FName For Binary Access Write As #2
    ' Create or Overwrite a File
    If StMo = "Stereo" Then
        ChanOut = 2
      Else
        ChanOut = 1
    End If
    
    WriteHeader ChanOut, SampFreq, Nbits, LenData ' Write Header Info
    
    'Start of Binary Copy from the Selected Area in the Wav File
    'to the Newly created Untitled Wav File.
    If ChanOut = 2 Then GoTo Stereo8
    If Nbits = 16 Then GoTo Mono16
Mono8:
    Get #1, InData, yByte ' Points to First Block of Selection in source wav
    Put #2, , yByte ' Writes to Next Block in New File
        For n = 1 To LenData - 1
            Get #1, , yByte ' Points to Next Block of Selection in source wav
            Put #2, , yByte ' Writes to Next Block in New File
        Next n
    GoTo Done
    
Mono16:
    Get #1, InData, yint ' Points to First Block of Selection in source wav
    Put #2, , yint ' Writes to Next Block in New File

        For n = 1 To LenData - 1

            Get #1, , yint ' Points to Next Block of Selection in source wav
            Put #2, , yint ' Writes to Next Block in New File
        Next n
    GoTo Done

Stereo8:
    
    If Nbits = 16 Then GoTo Stereo16

    Get #1, InData, yByte 'left Channel
    Put #2, , yByte ' Writes to Next Block in New File
   
    Get #1, , yByte 'right Channel
    Put #2, , yByte ' Writes to Next Block in New File
    
        For n = 1 To LenData - 1
            
            Get #1, , yByte 'left Channel
            Put #2, , yByte ' Writes to Next Block in New File
            
           
            Get #1, , yByte 'right Channel
            Put #2, , yByte ' Writes to Next Block in New File
        Next n
    GoTo Done

Stereo16:
    Get #1, InData, yint 'left Channel
    Put #2, , yint ' Writes to Next Block in New File
    
    Get #1, , yint 'right Channel
    Put #2, , yint ' Writes to Next Block in New File
    
        For n = 1 To LenData - 1
            
            Get #1, , yint 'left Channel
            Put #2, , yint ' Writes to Next Block in New File
         
            Get #1, , yint 'right Channel
            Put #2, , yint ' Writes to Next Block in New File
            
        Next n
         
Done:
    Close #2
End Sub

Private Sub Command9_Click() ' Paste New File Reversed
Dim FTemp As String
    Dim InData As Long, LenData As Long
    Dim InDataSel As Long, LenDataSel As Long
    Dim SampIni As Long, BytInic As Long
    Dim Nbits As Integer, StMo As String
    Dim yDiv As Integer, SampFreq As Long
    If MMCPart = False Then
        Picture2.SetFocus
        Exit Sub
    End If
    If ThisSampBegin = 0 Then
        msg = "No Selection Made!"
        MsgBox msg, vbOKOnly
        Exit Sub
    End If
    MousePointer = vbHourglass ' Busy
    Picture2.SetFocus
    FTemp = App.Path & "\Reversed-Untitled" & FormInstance & ".wav"
    
    InData = Label11.Caption
    SampIni = ThisSampBegin ' Selection Begins
    yDiv = Label12.Caption ' FileSize in Bytes
    BytInic = SampIni * yDiv ' Location in the Wav of the Visible Selection
    InDataSel = InData + BytInic
    LenDataSel = ThisSampTotal 'Selection Length in Samples
    SampFreq = ThisSampRate ' Sampling Frequency
    StMo = ThisChannels ' Stereo Mono
    Nbits = Val(ThisBitRate) ' Sampling Bits
    
    Open CurFile For Binary Access Read As #1
    
    Call SaveReversed(FTemp, InDataSel, LenDataSel, SampFreq, Nbits, StMo)
    
    Close #1
    
        LoadNewFile App.Path & "\Untitled" & FileCount & ".wav", FTemp, FormInstance
        MousePointer = 0 ' Arrow
End Sub
Private Sub SaveReversed(FName As String, InData As Long, LenData As Long, _
                    SampFreq As Long, Nbits As Integer, StMo As String)
    
    Dim ChanOut As Integer
    Dim yint As Integer
    Dim yByte As Byte
    Dim n As Long
    Dim EData As Long
    
    EData = InData + LenData ' Last Sample of Selection
    
    Open FName For Binary Access Write As #2
    ' Create or Overwrite a File
    If StMo = "Stereo" Then
        ChanOut = 2
      Else
        ChanOut = 1
    End If
    
    WriteHeader ChanOut, SampFreq, Nbits, LenData ' Write Header Info
    
    If ChanOut = 2 Then GoTo Stereo8
    If Nbits = 16 Then GoTo Mono16
Mono8:
    Get #1, EData, yByte ' Points to Last Block of Selection in source wav
    Put #2, , yByte ' Writes to Next Block in New File
        For n = EData To (InData + 1) Step -1
            Get #1, n, yByte ' Points to Next Block of Selection in source wav
            Put #2, , yByte ' Writes to Next Block in New File
        Next n
    GoTo Done
    
Mono16:
    Get #1, EData, yint ' Points to Last Block of Selection in source wav
    Put #2, , yint ' Writes to Next Block in New File

        For n = EData To (InData + 1) Step -1

            Get #1, n, yint ' Points to Next Block of Selection in source wav
            Put #2, , yint ' Writes to Next Block in New File
        Next n
    GoTo Done

Stereo8:
    
    If Nbits = 16 Then GoTo Stereo16

    Get #1, InData, yByte 'left Channel
    Put #2, , yByte ' Writes to Next Block in New File
   
    Get #1, , yByte 'right Channel
    Put #2, , yByte ' Writes to Next Block in New File
    
        For n = 1 To LenData - 1
            
            Get #1, , yByte 'left Channel
            Put #2, , yByte ' Writes to Next Block in New File
            
           
            Get #1, , yByte 'right Channel
            Put #2, , yByte ' Writes to Next Block in New File
        Next n
    GoTo Done

Stereo16:
    Get #1, EData - 1, yint 'left Channel
    Put #2, , yint ' Writes to Next Block in New File
    
    Get #1, EData, yint  'right Channel
    Put #2, , yint ' Writes to Next Block in New File
    
        For n = EData - 2 To (InData + 1) Step -2
            
            Get #1, n - 1, yint 'left Channel
            Put #2, , yint ' Writes to Next Block in New File
         
            Get #1, , yint  'right Channel
            Put #2, , yint ' Writes to Next Block in New File
            
        Next n
         
Done:
    Close #2
End Sub

Private Sub Picture5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Right Channel
    If RePlay = True Then
        RePlay = False
        PlayLoop
    End If
    Picture5.MousePointer = 0
    MovEsq = False
    MovDir = False

End Sub

Private Sub Picture7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Right Channel
    If IsPlaying = True Then
            IsPlaying = False
            RepeatIt = 0
            MMControl1.Command = "Stop"
            MMCPart = True
            PlayControls = True
            RePlay = True
            DoEvents
    End If

    If MMCPart = False Then Exit Sub
    If Button = vbRightButton Then
        If Picture3.Width < 100 Then
            Picture7.MousePointer = 9
            If X - Picture3.Left < Picture3.Width / 3 Then
                MovEsq = True
                LastWidth = Picture3.Width
                LastLeft = Picture3.Left
              Else
                MovDir = True
            End If
          ElseIf X - Picture3.Left < 50 Then
            Picture7.MousePointer = 9
            MovEsq = True
            LastWidth = Picture3.Width
            LastLeft = Picture3.Left
          ElseIf Picture3.Width + Picture3.Left - X < 100 Then
            Picture4.MousePointer = 9
            MovDir = True
        End If
    End If
    If Button = vbLeftButton Then
        Picture3.Width = 0
        Picture3.Left = X
        Picture3.Visible = True
        Picture4.Left = -X
        Picture6.Width = 0
        Picture6.Left = X
        Picture6.Visible = True
        Picture7.Left = -X
        IniPlay = Label4.Caption + X * Label15(1).Caption * 1000 / Picture2.ScaleWidth
        FimPlay = Label4.Caption + Label15(1).Caption * 1000
        ThisTimeBegin = IniPlay / 1000
        ThisTimeEnd = FimPlay / 1000
        ThisTimeTotal = (FimPlay - IniPlay) / 1000
        ThisSampBegin = Label17.Caption + Int(X * (Label6(1).Caption - 1) / Picture2.ScaleWidth)
        ThisSampEnd = Val(Label17.Caption) + Val(Label6(1).Caption) - 1
        ThisSampTotal = ThisSampEnd - ThisSampBegin + 1
        Call PosChange
    End If

End Sub

Private Sub Picture7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Right Channel
    If MMCPart = False Then Exit Sub
    If Button = vbRightButton Then
        If X = xLast Then
            Picture4.Visible = True
            Picture7.Visible = True
            Exit Sub
          ElseIf MovEsq = True Then
            Call ShadeArea(X)
            xLast = X
            If X >= 0 Then
                IniPlay = Label4.Caption + Picture3.Left * Label15(1).Caption * 1000 / Picture2.ScaleWidth
                ThisTimeBegin = IniPlay / 1000
                ThisTimeTotal = (FimPlay - IniPlay) / 1000
                ThisSampBegin = Label17.Caption + Int(Picture3.Left * (Label6(1).Caption - 1) / Picture2.ScaleWidth)
                ThisSampTotal = ThisSampEnd - ThisSampBegin + 1
                Call PosChange
            End If
          ElseIf MovDir = True And X - Picture3.Left > 50 And X <= Picture2.ScaleWidth Then
            Picture3.Width = X - Picture3.Left
            Picture6.Width = Picture3.Width
            FimPlay = Label4.Caption + X * Label15(1).Caption * 1000 / Picture2.ScaleWidth
            ThisTimeEnd = FimPlay / 1000
            ThisTimeTotal = (FimPlay - IniPlay) / 1000
            ThisSampEnd = Label17.Caption + Int(X * (Label6(1).Caption - 1) / Picture2.ScaleWidth)
            ThisSampTotal = ThisSampEnd - ThisSampBegin + 1
            Call PosChange
        End If
    End If


End Sub

Private Sub Picture7_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Right Channel
    If RePlay = True Then
        RePlay = False
        PlayLoop
    End If
    Picture7.MousePointer = 0
    MovDir = False
    MovEsq = False
End Sub

Private Sub txtHotKey_Change()
    If txtHotKey <> "" Then
    k = Asc(UCase(txtHotKey))
    For n = 0 To 255
    If SCount(n) = True And HotKey(n) = k Then
        MsgBox "HotKey already in use!", vbCritical, "Sampler - Error:"
        txtHotKey = "": Exit Sub
    End If
    Next n
    HotKey(FormInstance) = k
    frmSelection.lblHotKey.Caption = "HotKey: " & Chr$(HotKey(FormInstance))
    End If
End Sub
