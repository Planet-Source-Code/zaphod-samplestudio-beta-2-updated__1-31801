VERSION 5.00
Begin VB.Form frmSelection 
   Caption         =   "Selection"
   ClientHeight    =   810
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7590
   ControlBox      =   0   'False
   Icon            =   "frmSelection.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   810
   ScaleWidth      =   7590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   5640
      TabIndex        =   16
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label lbltriggers 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Triggers:"
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
      Left            =   4080
      TabIndex        =   15
      Top             =   0
      Width           =   3495
   End
   Begin VB.Label lblHotKey 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "HotKey:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ctrl + Shift + Hotkey= Stop"
      Height          =   255
      Left            =   5640
      TabIndex        =   13
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Shift + Hotkey= Loop"
      Height          =   255
      Left            =   4080
      TabIndex        =   12
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ctrl + Hotkey= Play"
      Height          =   255
      Left            =   4080
      TabIndex        =   11
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label lblTimeTotal 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3000
      TabIndex        =   10
      Top             =   480
      Width           =   975
   End
   Begin VB.Label lblSampTotal 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3000
      TabIndex        =   9
      Top             =   240
      Width           =   975
   End
   Begin VB.Label lblTimeEnd 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2040
      TabIndex        =   8
      Top             =   480
      Width           =   975
   End
   Begin VB.Label lblSampEnd 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2040
      TabIndex        =   7
      Top             =   240
      Width           =   975
   End
   Begin VB.Label lblTimeBegin 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1080
      TabIndex        =   6
      Top             =   480
      Width           =   975
   End
   Begin VB.Label lblSampBegin 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1080
      TabIndex        =   5
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Time (sec)"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   4
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Samples"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   10
      Left            =   0
      TabIndex        =   3
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Beginning"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   11
      Left            =   1080
      TabIndex        =   2
      Top             =   0
      Width           =   975
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "End"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   7
      Left            =   2040
      TabIndex        =   1
      Top             =   0
      Width           =   975
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   9
      Left            =   3000
      TabIndex        =   0
      Top             =   0
      Width           =   975
   End
End
Attribute VB_Name = "frmSelection"
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

'Current Selection Focus
'
' Fields Update from the WavForm.Form_Activate Sub
