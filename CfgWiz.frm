VERSION 5.00
Begin VB.Form CfgWiz 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuration Wizard"
   ClientHeight    =   4995
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7095
   Icon            =   "CfgWiz.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   7095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox WF 
      BorderStyle     =   0  'None
      Height          =   4455
      Index           =   5
      Left            =   2280
      ScaleHeight     =   4455
      ScaleWidth      =   4695
      TabIndex        =   38
      Top             =   0
      Visible         =   0   'False
      Width           =   4695
      Begin VB.Label Label18 
         Caption         =   "Congratulations!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   53
         Top             =   0
         Width           =   4575
      End
      Begin VB.Label Label19 
         Caption         =   "The initial NetBattle configuration is finished.  You can change options at any time from the Options menu."
         Height          =   495
         Left            =   120
         TabIndex        =   52
         Top             =   1080
         Width           =   4575
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label20 
         Caption         =   "Press Finish to continue."
         Height          =   255
         Left            =   120
         TabIndex        =   51
         Top             =   4080
         Width           =   4575
         WordWrap        =   -1  'True
      End
   End
   Begin VB.PictureBox WF 
      BorderStyle     =   0  'None
      Height          =   4455
      Index           =   4
      Left            =   2280
      ScaleHeight     =   4455
      ScaleWidth      =   4695
      TabIndex        =   21
      Top             =   0
      Visible         =   0   'False
      Width           =   4695
      Begin VB.Frame Frame10 
         Caption         =   "Replays"
         Height          =   855
         Left            =   2520
         TabIndex        =   54
         Top             =   3000
         Width           =   2055
         Begin VB.CheckBox PromptR 
            Caption         =   "Prompt to Save"
            Height          =   255
            Left            =   120
            TabIndex        =   57
            Top             =   240
            Width           =   1815
         End
         Begin VB.CheckBox SaveIt 
            Caption         =   "Autosave"
            Height          =   255
            Left            =   120
            TabIndex        =   55
            Top             =   480
            Width           =   1815
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Automatic Updates"
         Height          =   855
         Left            =   2520
         TabIndex        =   48
         Top             =   120
         Width           =   2055
         Begin VB.CheckBox AutoUpdt 
            Caption         =   "Ask Before Updating"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   50
            ToolTipText     =   "Make updating optional"
            Top             =   480
            Value           =   1  'Checked
            Width           =   1815
         End
         Begin VB.CheckBox AutoUpdt 
            Caption         =   "Enabled"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   49
            ToolTipText     =   "Enable automatic updating"
            Top             =   240
            Value           =   1  'Checked
            Width           =   1815
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Team"
         Height          =   855
         Left            =   2520
         TabIndex        =   44
         Top             =   1080
         Width           =   2055
         Begin VB.PictureBox Picture3 
            BorderStyle     =   0  'None
            Height          =   495
            Left            =   120
            ScaleHeight     =   495
            ScaleWidth      =   1815
            TabIndex        =   45
            Top             =   240
            Width           =   1815
            Begin VB.OptionButton TeamHide 
               Caption         =   "Hide Until Battle"
               Height          =   255
               Index           =   1
               Left            =   0
               TabIndex        =   47
               ToolTipText     =   "Hide your team until battle starts"
               Top             =   240
               Width           =   1695
            End
            Begin VB.OptionButton TeamHide 
               Caption         =   "Show Other Players"
               Height          =   255
               Index           =   0
               Left            =   0
               TabIndex        =   46
               ToolTipText     =   "Make the team public on the server"
               Top             =   0
               Value           =   -1  'True
               Width           =   1695
            End
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Logs"
         Height          =   855
         Left            =   2520
         TabIndex        =   42
         Top             =   2040
         Width           =   2055
         Begin VB.CheckBox AutosaveL 
            Caption         =   "Autosave"
            Height          =   255
            Left            =   120
            TabIndex        =   56
            Top             =   480
            Width           =   1695
         End
         Begin VB.CheckBox PromptIt 
            Caption         =   "Prompt to Save"
            Height          =   255
            Left            =   120
            TabIndex        =   43
            ToolTipText     =   "If the other player disconnects, you will be asked if you want to save the log."
            Top             =   240
            Width           =   1815
         End
      End
      Begin VB.Label Label21 
         Caption         =   "Automatic Updates: NetBattle can check online for new versions.  It is recommended you keep the defaults."
         Height          =   855
         Left            =   120
         TabIndex        =   41
         Top             =   120
         Width           =   2295
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label22 
         Caption         =   "Team: You can choose to have your team visible on the Challenge Window, or to hide your Pokemon."
         Height          =   855
         Left            =   120
         TabIndex        =   40
         Top             =   1200
         Width           =   2295
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label23 
         Caption         =   $"CfgWiz.frx":1272
         Height          =   1215
         Left            =   120
         TabIndex        =   39
         Top             =   2160
         Width           =   2295
         WordWrap        =   -1  'True
      End
   End
   Begin VB.PictureBox WF 
      BorderStyle     =   0  'None
      Height          =   4455
      Index           =   3
      Left            =   2280
      ScaleHeight     =   4455
      ScaleWidth      =   4695
      TabIndex        =   20
      Top             =   0
      Visible         =   0   'False
      Width           =   4695
      Begin VB.Frame Frame11 
         Caption         =   "Background"
         Height          =   615
         Left            =   2520
         TabIndex        =   58
         Top             =   3720
         Width           =   2055
         Begin VB.CheckBox EnableBG 
            Caption         =   "Enable"
            Height          =   255
            Left            =   120
            TabIndex        =   59
            Top             =   240
            Width           =   1815
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Audio"
         Height          =   855
         Left            =   2520
         TabIndex        =   35
         Top             =   120
         Width           =   2055
         Begin VB.CheckBox Audio 
            Caption         =   "Music"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   37
            ToolTipText     =   "Enable music"
            Top             =   480
            Value           =   1  'Checked
            Width           =   1335
         End
         Begin VB.CheckBox Audio 
            Caption         =   "Sound"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   36
            ToolTipText     =   "Enable chat sounds & short audio clips"
            Top             =   240
            Value           =   2  'Grayed
            Width           =   1215
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Battle Messages"
         Height          =   855
         Left            =   2520
         TabIndex        =   31
         Top             =   1080
         Width           =   2055
         Begin VB.PictureBox Picture2 
            BorderStyle     =   0  'None
            Height          =   495
            Left            =   120
            ScaleHeight     =   495
            ScaleWidth      =   1815
            TabIndex        =   32
            Top             =   240
            Width           =   1815
            Begin VB.OptionButton MessStyle 
               Caption         =   "Extended (Log-Style)"
               Height          =   255
               Index           =   1
               Left            =   0
               TabIndex        =   34
               ToolTipText     =   "Extended text for creating logs."
               Top             =   240
               Width           =   1815
            End
            Begin VB.OptionButton MessStyle 
               Caption         =   "Normal"
               Height          =   255
               Index           =   0
               Left            =   0
               TabIndex        =   33
               ToolTipText     =   "Normal G/S/C messages"
               Top             =   0
               Value           =   -1  'True
               Width           =   1815
            End
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Animation"
         Height          =   855
         Left            =   2520
         TabIndex        =   28
         Top             =   2040
         Width           =   2055
         Begin VB.CheckBox AnimHP 
            Caption         =   "HP Bars"
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   240
            Width           =   1575
         End
         Begin VB.CheckBox PokeAnim 
            Caption         =   "Pokémon"
            Enabled         =   0   'False
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   480
            Width           =   1575
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Text Display"
         Height          =   615
         Left            =   2520
         TabIndex        =   26
         Top             =   3000
         Width           =   2055
         Begin VB.CheckBox FText 
            Caption         =   "Use Fancy Text"
            Height          =   255
            Left            =   120
            TabIndex        =   27
            ToolTipText     =   "If enabled, use colored text."
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.Label Label24 
         Caption         =   "Background: This will show a background, but force Ru/Sa images during battle."
         Height          =   615
         Left            =   120
         TabIndex        =   60
         Top             =   3840
         Width           =   2295
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label12 
         Caption         =   "Audio Options: Select Sound, Music, Both, or Neither."
         Height          =   735
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   2295
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label13 
         Caption         =   "Battle Messages: Normal will display Game Boy-style messages during battle.  Extended includes additional detail."
         Height          =   975
         Left            =   120
         TabIndex        =   24
         Top             =   960
         Width           =   2295
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label14 
         Caption         =   "Animation: HP Bars can be animated to drain/fill GameBoy style.  Pokémon animation is NOT supported yet."
         Height          =   855
         Left            =   120
         TabIndex        =   23
         Top             =   2040
         Width           =   2295
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label15 
         Caption         =   "Text Display: If Fancy Text is enabled, colored/bold/italic text will be used for messages."
         Height          =   615
         Left            =   120
         TabIndex        =   22
         Top             =   3000
         Width           =   2295
         WordWrap        =   -1  'True
      End
   End
   Begin VB.PictureBox WF 
      BorderStyle     =   0  'None
      Height          =   4455
      Index           =   2
      Left            =   2280
      ScaleHeight     =   4455
      ScaleWidth      =   4695
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   4695
      Begin VB.Frame Frame1 
         Caption         =   "Set Stored Password"
         Height          =   1695
         Left            =   120
         TabIndex        =   15
         Top             =   960
         Width           =   4455
         Begin VB.TextBox PW1 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   120
            MaxLength       =   16
            PasswordChar    =   "*"
            TabIndex        =   17
            Top             =   600
            Width           =   4215
         End
         Begin VB.TextBox PW2 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   120
            MaxLength       =   16
            PasswordChar    =   "*"
            TabIndex        =   16
            Top             =   1200
            Width           =   4215
         End
         Begin VB.Label Label5 
            Caption         =   "Enter Password"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   360
            Width           =   4215
         End
         Begin VB.Label Label6 
            Caption         =   "Re-enter Password for Verification"
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   960
            Width           =   4215
         End
      End
      Begin VB.Label Label4 
         Caption         =   $"CfgWiz.frx":130B
         Height          =   855
         Left            =   120
         TabIndex        =   14
         Top             =   120
         Width           =   4455
      End
      Begin VB.Label Label7 
         Caption         =   $"CfgWiz.frx":13D3
         Height          =   735
         Left            =   120
         TabIndex        =   13
         Top             =   3360
         Width           =   4455
      End
      Begin VB.Label Label8 
         Caption         =   "Please remember this password in case you ever have to reinstall NetBattle, or install it on a different machine."
         Height          =   495
         Left            =   120
         TabIndex        =   12
         Top             =   2880
         Width           =   4455
      End
   End
   Begin VB.PictureBox WF 
      BorderStyle     =   0  'None
      Height          =   4455
      Index           =   1
      Left            =   2280
      ScaleHeight     =   4455
      ScaleWidth      =   4695
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   4695
      Begin VB.ListBox LangBox 
         Height          =   2595
         Left            =   480
         TabIndex        =   9
         Top             =   600
         Width           =   3735
      End
      Begin VB.Label Label16 
         Caption         =   "Select a language from the installed plugins."
         Height          =   255
         Left            =   480
         TabIndex        =   11
         Top             =   240
         Width           =   3735
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label17 
         Caption         =   $"CfgWiz.frx":1474
         Height          =   855
         Left            =   480
         TabIndex        =   10
         Top             =   3360
         Width           =   3735
      End
   End
   Begin VB.PictureBox WF 
      BorderStyle     =   0  'None
      Height          =   4455
      Index           =   0
      Left            =   2280
      ScaleHeight     =   4455
      ScaleWidth      =   4695
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   4695
      Begin VB.Label Label1 
         Caption         =   "Welcome to Pokémon NetBattle!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   6
         Top             =   120
         Width           =   4575
      End
      Begin VB.Label Label2 
         Caption         =   "This Wizard will walk you through the initial NetBattle configuration."
         Height          =   495
         Left            =   0
         TabIndex        =   5
         Top             =   1200
         Width           =   4575
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label3 
         Caption         =   "Press Next to continue."
         Height          =   255
         Left            =   0
         TabIndex        =   4
         Top             =   4200
         Width           =   4575
         WordWrap        =   -1  'True
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<< &Back"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5760
      TabIndex        =   1
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "&Next >>"
      Height          =   375
      Left            =   4440
      TabIndex        =   0
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   4500
      Left            =   0
      Picture         =   "CfgWiz.frx":152F
      Top             =   0
      Width           =   2250
   End
End
Attribute VB_Name = "CfgWiz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim CurrentStep As Byte

Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
    Dim X As Byte
    
    CurrentStep = CurrentStep - 1
    If CurrentStep = WF.count - 1 Then
        OKButton.Caption = "&Finish"
        CancelButton.Enabled = False
    Else
        OKButton.Caption = "&Next >>"
        CancelButton.Enabled = True
    End If
    If CurrentStep = 0 Then
        Command1.Enabled = False
    Else
        Command1.Enabled = True
    End If
    For X = 0 To WF.count - 1
        If X = CurrentStep Then WF(X).Visible = True Else WF(X).Visible = False
    Next
End Sub

Private Sub Form_Load()
    Dim X As Integer
    '>>> Call WriteDebugLog("CfgWiz loaded")
    CurrentStep = 0
    Call AddMessage("Test: Line 1", , ":", vbRed, True)
    Call AddMessage("Test: Line 2", , ":", vbBlue, True)
    PW1.Text = SavedPassword
    PW2.Text = SavedPassword
    '>>> Call WriteDebugLog("Test Lines added, Passwords loaded")
    Audio(0).Value = SoundOption
    Audio(1).Value = MusicOption
    AnimHP.Value = AnimOption
    If SoundOption = 1 And MusicOption = 1 Then Audio(0).Value = 2
    '>>> Call WriteDebugLog("Music Options set")
    AutoUpdt(0).Value = Abs(AutoScan)
    AutoUpdt(1).Value = Abs(AskOnUpdate)
    If AutoUpdt(0).Value = 0 Then AutoUpdt(1).Enabled = False
    If AllowViewing = 1 Then TeamHide(0).Value = True Else TeamHide(1).Value = True
    If LogPrompt = 1 Then PromptIt.Value = 1
    If LogSave = 1 Then AutosaveL.Value = 1: PromptIt.Enabled = False
    If ReplayPrompt = 1 Then PromptR.Value = 1
    If Autosave = 1 Then SaveIt.Value = 1: PromptR.Value = False
    If FancyText Then FText.Value = 1 Else FText.Value = 0
    If UseBG Then EnableBG.Value = 1 Else EnableBG.Value = 0
    '>>> Call WriteDebugLog("Checkboxes set")
    Select Case BMessStyle
        Case 0
            MessStyle(0).Value = True
            MessStyle(1).Value = False
        Case 1
            MessStyle(1).Value = True
            MessStyle(0).Value = False
    End Select
    '>>> Call WriteDebugLog("BMessStyle set")
    For X = 0 To UBound(LFile)
        If LFile(X).Text <> "" Then
            LangBox.AddItem LFile(X).Text, X
            If LFile(X).FileName = CurrLang Then LangBox.ListIndex = X
        End If
    Next
    If CurrLang = "" Then LangBox.ListIndex = 0
    WF(0).Visible = True
    '>>> Call WriteDebugLog("Load Complete.")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Audio(0).Value > 0 Then
        SoundOption = 1
    Else
        SoundOption = 0
    End If
    SaveSetting "NetBattle", "Options", "ServerSound", SoundOption
    MusicOption = Audio(1).Value
    SaveSetting "NetBattle", "Options", "Music", MusicOption
    
    AnimOption = AnimHP.Value
    SaveSetting "NetBattle", "Options", "Animation", AnimOption
    
    If TeamHide(0).Value Then
        AllowViewing = 1
    Else
        AllowViewing = 0
    End If
    SaveSetting "NetBattle", "Options", "Allow Viewing", AllowViewing
    
    LogPrompt = PromptIt.Value
    SaveSetting "NetBattle", "Options", "Log Prompt", LogPrompt
    Autosave = SaveIt.Value
    SaveSetting "NetBattle", "Options", "Save Replays", Autosave
    
    AutoScan = AutoUpdt(0).Value
    AskOnUpdate = AutoUpdt(1).Value
    SaveSetting "NetBattle", "Options", "Auto Scan", AutoScan
    SaveSetting "NetBattle", "Options", "Ask On Update", AskOnUpdate
    
    If MessStyle(0).Value Then
        BMessStyle = 0
    Else
        BMessStyle = 1
    End If
    SaveSetting "NetBattle", "Options", "Message Style", BMessStyle
    
    If ServerPassword <> "" And SavedPassword <> ServerPassword Then
        SavedPassword = MD5(ServerPassword)
        SaveSetting "NetBattle", "Options", "Saved Password", SavedPassword
    End If
    
    If FText.Value = 1 Then FancyText = True Else FancyText = False
    SaveSetting "NetBattle", "Options", "Fancy Text", FancyText
    
'    If Option1(0).Value Then AddLineBreaks = True Else AddLineBreaks = False
'    SaveSetting "NetBattle", "Options", "Line Breaks", AddLineBreaks
    
    LogSave = AutosaveL.Value
    SaveSetting "NetBattle", "Options", "Log Save", LogSave
    ReplayPrompt = PromptR.Value
    SaveSetting "NetBattle", "Options", "Replay Prompt", ReplayPrompt
End Sub

Private Sub OKButton_Click()
    Dim X As Byte
    
    Select Case CurrentStep
        'Apply Language now
        Case 1
            If LFile(LangBox.ListIndex).FileName <> CurrLang Then
                CurrLang = LFile(LangBox.ListIndex).FileName
                SaveSetting "NetBattle", "Options", "Language", CurrLang
                Call Loader.DoLanguage
            End If
        Case 2
            If PW1.Text <> PW2.Text Then
                MsgBox "The passwords do not match.  Please retype them and try again.", vbCritical, "Error"
                Exit Sub
            Else
                SavedPassword = MD5(PW1.Text)
                SaveSetting "NetBattle", "Options", "Saved Password", SavedPassword
            End If
    End Select
    If CurrentStep = WF.count - 1 Then Unload Me: Exit Sub
    CurrentStep = CurrentStep + 1
    If CurrentStep = WF.count - 1 Then
        OKButton.Caption = "&Finish"
        CancelButton.Enabled = False
    Else
        OKButton.Caption = "&Next >>"
        CancelButton.Enabled = True
    End If
    If CurrentStep = 0 Then
        Command1.Enabled = False
    Else
        Command1.Enabled = True
    End If
    For X = 0 To WF.count - 1
        If X = CurrentStep Then WF(X).Visible = True Else WF(X).Visible = False
    Next
End Sub

Sub AddMessage(ByVal Message As String, Optional ByVal DebugMessage As Boolean = False, Optional ByVal BreakChar As String = "", Optional ByVal Color As Long = vbBlack, Optional ByVal Bold As Boolean = False, Optional ByVal Italic As Boolean = False)
    If DebugMessage And Not DebugMode Then Exit Sub
    'Call AddMessageMain(Messages, Message, BreakChar, Color, Bold, Italic)
End Sub

Private Sub AutoUpdt_Click(Index As Integer)
    If AutoUpdt(0).Value = 0 Then AutoUpdt(1).Enabled = False Else AutoUpdt(1).Enabled = True
End Sub

Private Sub AutoUpdt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    If AutoUpdt(0).Value = 0 Then AutoUpdt(1).Enabled = False Else AutoUpdt(1).Enabled = True
End Sub

Private Sub AutoUpdt_LostFocus(Index As Integer)
    If AutoUpdt(0).Value = 0 Then AutoUpdt(1).Enabled = False Else AutoUpdt(1).Enabled = True
End Sub

Private Sub Audio_Click(Index As Integer)
    If Audio(1).Value = 1 And Audio(0).Value = 1 Then Audio(0).Value = 2
    If Audio(1).Value = 0 And Audio(0).Value = 2 Then Audio(0).Value = 1
End Sub

Private Sub Audio_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Audio(1).Value = 1 And Audio(0).Value = 1 Then Audio(0).Value = 2
    If Audio(1).Value = 0 And Audio(0).Value = 2 Then Audio(0).Value = 1
End Sub

Private Sub Audio_LostFocus(Index As Integer)
    If Audio(1).Value = 1 And Audio(0).Value = 1 Then Audio(0).Value = 2
    If Audio(1).Value = 0 And Audio(0).Value = 2 Then Audio(0).Value = 1
End Sub

