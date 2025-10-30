VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Begin VB.Form SoundPick 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Sounds"
   ClientHeight    =   5415
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5655
   Icon            =   "SoundPick.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "&Stop Test"
      Height          =   375
      Left            =   1440
      TabIndex        =   49
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&Defaults"
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   4920
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4695
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   8281
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Sounds"
      TabPicture(0)   =   "SoundPick.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "DoPlay(3)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "DoPlay(2)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "DoPlay(1)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "DoPlay(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "FileName(0)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "BrowseButton(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "FileName(1)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "BrowseButton(1)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "FileName(2)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "BrowseButton(2)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "FileName(3)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "BrowseButton(3)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "TestButton(0)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "TestButton(1)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "TestButton(2)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "TestButton(3)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).ControlCount=   17
      TabCaption(1)   =   "Music"
      TabPicture(1)   =   "SoundPick.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DoPlay(10)"
      Tab(1).Control(1)=   "DoPlay(9)"
      Tab(1).Control(2)=   "DoPlay(8)"
      Tab(1).Control(3)=   "DoPlay(7)"
      Tab(1).Control(4)=   "DoPlay(6)"
      Tab(1).Control(5)=   "DoPlay(5)"
      Tab(1).Control(6)=   "DoPlay(4)"
      Tab(1).Control(7)=   "FileName(4)"
      Tab(1).Control(8)=   "BrowseButton(4)"
      Tab(1).Control(9)=   "TestButton(4)"
      Tab(1).Control(10)=   "FileName(5)"
      Tab(1).Control(11)=   "BrowseButton(5)"
      Tab(1).Control(12)=   "TestButton(5)"
      Tab(1).Control(13)=   "FileName(6)"
      Tab(1).Control(14)=   "BrowseButton(6)"
      Tab(1).Control(15)=   "TestButton(6)"
      Tab(1).Control(16)=   "FileName(7)"
      Tab(1).Control(17)=   "BrowseButton(7)"
      Tab(1).Control(18)=   "TestButton(7)"
      Tab(1).Control(19)=   "FileName(8)"
      Tab(1).Control(20)=   "BrowseButton(8)"
      Tab(1).Control(21)=   "TestButton(8)"
      Tab(1).Control(22)=   "FileName(9)"
      Tab(1).Control(23)=   "BrowseButton(9)"
      Tab(1).Control(24)=   "TestButton(9)"
      Tab(1).Control(25)=   "FileName(10)"
      Tab(1).Control(26)=   "BrowseButton(10)"
      Tab(1).Control(27)=   "TestButton(10)"
      Tab(1).ControlCount=   28
      Begin VB.CommandButton TestButton 
         Caption         =   "Test"
         Height          =   285
         Index           =   10
         Left            =   -70200
         TabIndex        =   48
         Top             =   3120
         Width           =   495
      End
      Begin VB.CommandButton BrowseButton 
         Caption         =   "..."
         Height          =   285
         Index           =   10
         Left            =   -70560
         TabIndex        =   47
         Top             =   3120
         Width           =   255
      End
      Begin VB.TextBox FileName 
         Height          =   285
         Index           =   10
         Left            =   -74880
         TabIndex        =   45
         Top             =   3120
         Width           =   4215
      End
      Begin VB.CommandButton TestButton 
         Caption         =   "Test"
         Height          =   285
         Index           =   9
         Left            =   -70200
         TabIndex        =   44
         Top             =   1320
         Width           =   495
      End
      Begin VB.CommandButton BrowseButton 
         Caption         =   "..."
         Height          =   285
         Index           =   9
         Left            =   -70560
         TabIndex        =   43
         Top             =   1320
         Width           =   255
      End
      Begin VB.TextBox FileName 
         Height          =   285
         Index           =   9
         Left            =   -74880
         TabIndex        =   41
         Top             =   1320
         Width           =   4215
      End
      Begin VB.CommandButton TestButton 
         Caption         =   "Test"
         Height          =   285
         Index           =   8
         Left            =   -70200
         TabIndex        =   40
         Top             =   4320
         Width           =   495
      End
      Begin VB.CommandButton BrowseButton 
         Caption         =   "..."
         Height          =   285
         Index           =   8
         Left            =   -70560
         TabIndex        =   39
         Top             =   4320
         Width           =   255
      End
      Begin VB.TextBox FileName 
         Height          =   285
         Index           =   8
         Left            =   -74880
         TabIndex        =   37
         Top             =   4320
         Width           =   4215
      End
      Begin VB.CommandButton TestButton 
         Caption         =   "Test"
         Height          =   285
         Index           =   7
         Left            =   -70200
         TabIndex        =   36
         Top             =   3720
         Width           =   495
      End
      Begin VB.CommandButton BrowseButton 
         Caption         =   "..."
         Height          =   285
         Index           =   7
         Left            =   -70560
         TabIndex        =   35
         Top             =   3720
         Width           =   255
      End
      Begin VB.TextBox FileName 
         Height          =   285
         Index           =   7
         Left            =   -74880
         TabIndex        =   33
         Top             =   3720
         Width           =   4215
      End
      Begin VB.CommandButton TestButton 
         Caption         =   "Test"
         Height          =   285
         Index           =   6
         Left            =   -70200
         TabIndex        =   32
         Top             =   1920
         Width           =   495
      End
      Begin VB.CommandButton BrowseButton 
         Caption         =   "..."
         Height          =   285
         Index           =   6
         Left            =   -70560
         TabIndex        =   31
         Top             =   1920
         Width           =   255
      End
      Begin VB.TextBox FileName 
         Height          =   285
         Index           =   6
         Left            =   -74880
         TabIndex        =   29
         Top             =   1920
         Width           =   4215
      End
      Begin VB.CommandButton TestButton 
         Caption         =   "Test"
         Height          =   285
         Index           =   5
         Left            =   -70200
         TabIndex        =   28
         Top             =   2520
         Width           =   495
      End
      Begin VB.CommandButton BrowseButton 
         Caption         =   "..."
         Height          =   285
         Index           =   5
         Left            =   -70560
         TabIndex        =   27
         Top             =   2520
         Width           =   255
      End
      Begin VB.TextBox FileName 
         Height          =   285
         Index           =   5
         Left            =   -74880
         TabIndex        =   25
         Top             =   2520
         Width           =   4215
      End
      Begin VB.CommandButton TestButton 
         Caption         =   "Test"
         Height          =   285
         Index           =   4
         Left            =   -70200
         TabIndex        =   24
         Top             =   720
         Width           =   495
      End
      Begin VB.CommandButton BrowseButton 
         Caption         =   "..."
         Height          =   285
         Index           =   4
         Left            =   -70560
         TabIndex        =   23
         Top             =   720
         Width           =   255
      End
      Begin VB.CommandButton TestButton 
         Caption         =   "Test"
         Height          =   285
         Index           =   3
         Left            =   4800
         TabIndex        =   21
         Top             =   2520
         Width           =   495
      End
      Begin VB.CommandButton TestButton 
         Caption         =   "Test"
         Height          =   285
         Index           =   2
         Left            =   4800
         TabIndex        =   20
         Top             =   1920
         Width           =   495
      End
      Begin VB.CommandButton TestButton 
         Caption         =   "Test"
         Height          =   285
         Index           =   1
         Left            =   4800
         TabIndex        =   19
         Top             =   1320
         Width           =   495
      End
      Begin VB.CommandButton TestButton 
         Caption         =   "Test"
         Height          =   285
         Index           =   0
         Left            =   4800
         TabIndex        =   18
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox FileName 
         Height          =   285
         Index           =   4
         Left            =   -74880
         TabIndex        =   11
         Top             =   720
         Width           =   4215
      End
      Begin VB.CommandButton BrowseButton 
         Caption         =   "..."
         Height          =   285
         Index           =   3
         Left            =   4440
         TabIndex        =   10
         Top             =   2520
         Width           =   255
      End
      Begin VB.TextBox FileName 
         Height          =   285
         Index           =   3
         Left            =   120
         TabIndex        =   9
         Top             =   2520
         Width           =   4215
      End
      Begin VB.CommandButton BrowseButton 
         Caption         =   "..."
         Height          =   285
         Index           =   2
         Left            =   4440
         TabIndex        =   8
         Top             =   1920
         Width           =   255
      End
      Begin VB.TextBox FileName 
         Height          =   285
         Index           =   2
         Left            =   120
         TabIndex        =   7
         Top             =   1920
         Width           =   4215
      End
      Begin VB.CommandButton BrowseButton 
         Caption         =   "..."
         Height          =   285
         Index           =   1
         Left            =   4440
         TabIndex        =   6
         Top             =   1320
         Width           =   255
      End
      Begin VB.TextBox FileName 
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   1320
         Width           =   4215
      End
      Begin VB.CommandButton BrowseButton 
         Caption         =   "..."
         Height          =   285
         Index           =   0
         Left            =   4440
         TabIndex        =   4
         Top             =   720
         Width           =   255
      End
      Begin VB.TextBox FileName 
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   4215
      End
      Begin VB.CheckBox DoPlay 
         Caption         =   "Program Startup"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   480
         Value           =   1  'Checked
         Width           =   4215
      End
      Begin VB.CheckBox DoPlay 
         Caption         =   "Signon"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   15
         Top             =   1080
         Value           =   1  'Checked
         Width           =   4215
      End
      Begin VB.CheckBox DoPlay 
         Caption         =   "Chat Message"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   16
         Top             =   1680
         Value           =   1  'Checked
         Width           =   4215
      End
      Begin VB.CheckBox DoPlay 
         Caption         =   "Challenge"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   17
         Top             =   2280
         Value           =   1  'Checked
         Width           =   4215
      End
      Begin VB.CheckBox DoPlay 
         Caption         =   "Program Startup"
         Height          =   255
         Index           =   4
         Left            =   -74880
         TabIndex        =   22
         Top             =   480
         Value           =   1  'Checked
         Width           =   4215
      End
      Begin VB.CheckBox DoPlay 
         Caption         =   "GSC Battle"
         Height          =   255
         Index           =   5
         Left            =   -74880
         TabIndex        =   26
         Top             =   2280
         Value           =   1  'Checked
         Width           =   4215
      End
      Begin VB.CheckBox DoPlay 
         Caption         =   "RBY Battle"
         Height          =   255
         Index           =   6
         Left            =   -74880
         TabIndex        =   30
         Top             =   1680
         Value           =   1  'Checked
         Width           =   4215
      End
      Begin VB.CheckBox DoPlay 
         Caption         =   "Victory"
         Height          =   255
         Index           =   7
         Left            =   -74880
         TabIndex        =   34
         Top             =   3480
         Value           =   1  'Checked
         Width           =   4215
      End
      Begin VB.CheckBox DoPlay 
         Caption         =   "Loss"
         Height          =   255
         Index           =   8
         Left            =   -74880
         TabIndex        =   38
         Top             =   4080
         Value           =   1  'Checked
         Width           =   4215
      End
      Begin VB.CheckBox DoPlay 
         Caption         =   "Challenge"
         Height          =   255
         Index           =   9
         Left            =   -74880
         TabIndex        =   42
         Top             =   1080
         Value           =   1  'Checked
         Width           =   4215
      End
      Begin VB.CheckBox DoPlay 
         Caption         =   "Advance Battle"
         Height          =   255
         Index           =   10
         Left            =   -74880
         TabIndex        =   46
         Top             =   2880
         Value           =   1  'Checked
         Width           =   4215
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Please note that Program Startup and Challenge sounds will not play if Music is enabled."
         Height          =   495
         Left            =   120
         TabIndex        =   12
         Top             =   3000
         Width           =   4695
      End
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4320
      TabIndex        =   1
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3000
      TabIndex        =   0
      Top             =   4920
      Width           =   1215
   End
End
Attribute VB_Name = "SoundPick"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub BrowseButton_Click(Index As Integer)
    MainContainer.FileBox.DialogTitle = "Select Sound File"
    MainContainer.FileBox.Flags = cdlOFNHideReadOnly
    MainContainer.FileBox.CancelError = True
    MainContainer.FileBox.Filter = "Wave Files (*.wav)|*.wav|MIDI Files (*.mid)|*.mid|MP3 Files (*.mp3)|*.mp3|Windows Media audio (*.wma)|*.wma|Module audio (*.mod;*.s3m;*.xm;*.it)|*.mod;*.s3m;*.xm;*.it|All Files (*.*)|*.*"
    MainContainer.FileBox.FileName = SoundFile(Index)
    On Error GoTo Cancelled
    MainContainer.FileBox.ShowOpen
    FileName(Index).Text = MainContainer.FileBox.FileName
Cancelled:
End Sub

Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
    Dim X As Byte
    Dim Answer As Integer
    
    Answer = MsgBox("This will reset all sound files to the default.  Any custom sound settings will be lost.", vbQuestion + vbYesNo, "Are you sure?")
    If Answer = vbNo Then Exit Sub
    For X = 0 To 10
        Call ResetDefaultSound(X)
        FileName(X).Text = SoundFile(X)
        DoPlay(X).Value = 1
    Next
End Sub

Private Sub Command2_Click()
    Call StopSound
    Call StopMusic
End Sub

Private Sub Form_Load()
    Dim X As Integer
    
    On Error Resume Next
    SSTab1.Tab = 0
    For X = 0 To 10
        FileName(X).Text = SoundFile(X)
        If SoundEnable(X) Then DoPlay(X).Value = 1 Else DoPlay(X).Value = 0
    Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call StopSound
    Call StopMusic
End Sub

Private Sub OKButton_Click()
    Dim X As Integer
    
    For X = 0 To 10
        SoundFile(X) = FileName(X).Text
        If Not FileExists(SoundFile(X)) Then Call ResetDefaultSound(X)
        SaveSetting "NetBattle", "Sound", Trim(Str(X)), SoundFile(X)
        If DoPlay(X).Value = 1 Then SoundEnable(X) = True Else SoundEnable(X) = False
        SaveSetting "NetBattle", "Enable Sound", Trim(Str(X)), SoundEnable(X)
    Next
    Unload Me
End Sub

Private Sub TestButton_Click(Index As Integer)
    StopSound
    StopMusic
    If Not FileExists(SoundFile(Index)) Then
        MsgBox "File does not exist!", vbCritical, "Bad Filename"
        Call ResetDefaultSound(Index)
        FileName(Index).Text = SoundFile(Index)
        Exit Sub
    End If
    Select Case Index
        Case nbSoundOpening, nbSoundSignon, nbSoundChat, nbSoundChallenge
            Call PlaySound(Index)
        Case Else
            Call PlayMusic(Index, True)
    End Select
End Sub

