VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmRecWav 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Record New Sample (*.wav file)"
   ClientHeight    =   1530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3870
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1530
   ScaleWidth      =   3870
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   3615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Finish"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   720
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   240
      Top             =   1440
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Play"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Record"
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Save As:"
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
      Left            =   240
      TabIndex        =   6
      Top             =   0
      Width           =   3495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   3615
   End
End
Attribute VB_Name = "frmRecWav"
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

'Record Sample Form
 Private Declare Function mciSendString Lib "winmm.dll" Alias _
    "mciSendStringA" (ByVal lpstrCommand As String, ByVal _
    lpstrReturnString As Any, ByVal uReturnLength As Long, ByVal _
    hwndCallback As Long) As Long
Dim Blinker As Long ' Record Indicator

Public Sub RecordWav(Fpath As String) ' Pass Filename here
    
    Text1.Text = Fpath
End Sub
Private Sub Command1_Click() 'Save & Finish
    ' This Don't work exactly right because the mci
    ' save command save the wav file with an error in it!
    
    If Text1.Text = "" Then GoTo ErrHandler
On Error GoTo ErrHandler
   

 If FExsists(Text1.Text) = True Then
        Kill Text1.Text 'If file already exists then delete
 End If

'MCI command to save the WAV file
     i = mciSendString("save TheSample " & Text1.Text, 0&, 0, 0)
        DoEvents: DoEvents: DoEvents
        DoEvents: DoEvents: DoEvents
     LoadNewFile Text1.Text, Text1.Text, CurFormFocus
ErrHandler:
    
    
    Unload Me
End Sub

Private Sub Command2_Click() ' Record

 
 i = mciSendString("seek TheSample to start", 0&, 0, 0) 'Always start at the beginning
 i = mciSendString("set TheSample samplespersec 44100", 0&, 0, 0) 'CD Quality
 i = mciSendString("set TheSample bitspersample 16", 0&, 0, 0)  '16 bits for better sound
 i = mciSendString("set TheSample channels 2", 0&, 0, 0) ' 2 channels for stereo
 i = mciSendString("record TheSample", 0&, 0, 0)  'Start the recording

Command3.Enabled = True  'Enable the STOP BUTTON
Command4.Enabled = False  'Disable the "PLAY" button
Command1.Enabled = False  'Disable the "SAVE AS" button
End Sub

Private Sub Command3_Click() ' Stop Recording
  i = mciSendString("stop TheSample", 0&, 0, 0)
Command2.BackColor = &HC0C0C0
Command1.Enabled = True 'Enable the "SAVE AS" button
Command4.Enabled = True 'Enable the "PLAY" button


End Sub

Private Sub Command4_Click() ' Playback Recorded
  
  i = mciSendString("play TheSample from 0", 0&, 0, 0)

End Sub

Private Sub Form_Load() 'ini

 'Close any MCI operations from previous VB programs
 i = mciSendString("close all", 0&, 0, 0)
 
 'MCI Command to open a new wav file
 i = mciSendString("open new type waveaudio alias TheSample", 0&, 0, 0)

End Sub

Private Sub Form_Unload(Cancel As Integer) ' clean up
 i = mciSendString("close TheSample", 0&, 0, 0)
End Sub


Private Sub Timer1_Timer() 'update Status
Dim mssg As String * 255
i = mciSendString("status TheSample mode", mssg, 255, 0)
Label1.Caption = " " & mssg
If Left(mssg, 9) <> "recording" Then Exit Sub
Blinker = Blinker + 1
If Blinker >= 3 Then ' Flash Button if Recording
    If Command2.BackColor = &HC0C0C0 Then
        Command2.BackColor = vbRed
    Else
        Command2.BackColor = &HC0C0C0
    End If
    Blinker = 0
End If
End Sub
