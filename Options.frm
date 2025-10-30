VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Begin VB.Form Options 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options"
   ClientHeight    =   4170
   ClientLeft      =   135
   ClientTop       =   585
   ClientWidth     =   4710
   HelpContextID   =   20000
   Icon            =   "Options.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   4710
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton cmdHelp 
      Caption         =   "&Help"
      Height          =   375
      Left            =   3480
      TabIndex        =   52
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   3720
      Width           =   1095
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3495
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   6165
      _Version        =   393216
      Style           =   1
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "Display && Sound"
      TabPicture(0)   =   "Options.frx":212A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame4"
      Tab(0).Control(1)=   "Frame3"
      Tab(0).Control(2)=   "Frame2"
      Tab(0).Control(3)=   "Frame1"
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "System"
      TabPicture(1)   =   "Options.frx":2146
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame9"
      Tab(1).Control(1)=   "Frame7"
      Tab(1).Control(2)=   "Frame6"
      Tab(1).Control(3)=   "Frame5"
      Tab(1).Control(4)=   "LangBox"
      Tab(1).Control(5)=   "ChangePW"
      Tab(1).Control(6)=   "Command1"
      Tab(1).Control(7)=   "LangStat"
      Tab(1).ControlCount=   8
      TabCaption(2)   =   "Chat Windows"
      TabPicture(2)   =   "Options.frx":2162
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label1"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "lblSample"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label3"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label5"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "chkTimestamp"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Frame10"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "txtTSFormat"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "chkURLs"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "Timer1"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "txtLines"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "cmdFilter"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).ControlCount=   11
      Begin VB.CommandButton cmdFilter 
         Caption         =   "Word Filter..."
         Height          =   315
         Left            =   2640
         TabIndex        =   65
         Top             =   2940
         Width           =   1695
      End
      Begin VB.TextBox txtLines 
         Height          =   285
         Left            =   2640
         TabIndex        =   62
         Text            =   "1000"
         Top             =   2520
         Width           =   855
      End
      Begin VB.Timer Timer1 
         Interval        =   1
         Left            =   4200
         Top             =   3240
      End
      Begin VB.CheckBox chkURLs 
         Caption         =   "Parse URLs"
         Height          =   255
         Left            =   2640
         TabIndex        =   50
         Top             =   480
         Value           =   1  'Checked
         WhatsThisHelpID =   10030
         Width           =   1335
      End
      Begin VB.TextBox txtTSFormat 
         Height          =   285
         Left            =   2640
         TabIndex        =   48
         Text            =   "[hh:mm:ss]"
         Top             =   1260
         WhatsThisHelpID =   10029
         Width           =   1695
      End
      Begin VB.Frame Frame10 
         Caption         =   "Message Toggles"
         Height          =   2895
         Left            =   120
         TabIndex        =   37
         Top             =   360
         Width           =   2415
         Begin VB.CheckBox MessageToggle 
            Caption         =   "Show Battle Results"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   47
            Top             =   1080
            Value           =   1  'Checked
            WhatsThisHelpID =   10027
            Width           =   2200
         End
         Begin VB.CheckBox MessageToggle 
            Caption         =   "Show Nickname Change"
            Height          =   255
            Index           =   10
            Left            =   120
            TabIndex        =   46
            Top             =   2520
            Value           =   1  'Checked
            WhatsThisHelpID =   10027
            Width           =   2200
         End
         Begin VB.CheckBox MessageToggle 
            Caption         =   "Show Team Change"
            Height          =   255
            Index           =   9
            Left            =   120
            TabIndex        =   45
            Top             =   2280
            Value           =   1  'Checked
            WhatsThisHelpID =   10027
            Width           =   2200
         End
         Begin VB.CheckBox MessageToggle 
            Caption         =   "Show Ignore/Unignore"
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   44
            Top             =   2040
            Value           =   1  'Checked
            WhatsThisHelpID =   10027
            Width           =   2200
         End
         Begin VB.CheckBox MessageToggle 
            Caption         =   "Show Return From Away"
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   43
            Top             =   1800
            Value           =   1  'Checked
            WhatsThisHelpID =   10027
            Width           =   2200
         End
         Begin VB.CheckBox MessageToggle 
            Caption         =   "Show Away"
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   42
            Top             =   1560
            Value           =   1  'Checked
            WhatsThisHelpID =   10027
            Width           =   2200
         End
         Begin VB.CheckBox MessageToggle 
            Caption         =   "Show Return From Battle"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   41
            Top             =   1320
            Value           =   1  'Checked
            WhatsThisHelpID =   10027
            Width           =   2200
         End
         Begin VB.CheckBox MessageToggle 
            Caption         =   "Show Battle Start"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   40
            Top             =   840
            Value           =   1  'Checked
            WhatsThisHelpID =   10027
            Width           =   2200
         End
         Begin VB.CheckBox MessageToggle 
            Caption         =   "Show Player Sign Off"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   39
            Top             =   600
            Value           =   1  'Checked
            WhatsThisHelpID =   10027
            Width           =   2200
         End
         Begin VB.CheckBox MessageToggle 
            Caption         =   "Show Player Sign On"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   38
            Top             =   360
            Value           =   1  'Checked
            WhatsThisHelpID =   10027
            Width           =   2200
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Replays"
         Height          =   855
         Left            =   -74880
         TabIndex        =   29
         Top             =   2520
         Width           =   2055
         Begin VB.CheckBox PromptR 
            Caption         =   "Prompt to Save"
            Height          =   255
            Left            =   120
            TabIndex        =   35
            Top             =   240
            WhatsThisHelpID =   10019
            Width           =   1815
         End
         Begin VB.CheckBox SaveIt 
            Caption         =   "Autosave"
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Top             =   480
            WhatsThisHelpID =   10020
            Width           =   1815
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Logs"
         Height          =   855
         Left            =   -74880
         TabIndex        =   27
         Top             =   1560
         Width           =   2055
         Begin VB.CheckBox AutosaveL 
            Caption         =   "Autosave"
            Height          =   255
            Left            =   120
            TabIndex        =   36
            Top             =   480
            WhatsThisHelpID =   10018
            Width           =   1695
         End
         Begin VB.CheckBox PromptIt 
            Caption         =   "Prompt to Save"
            Height          =   255
            Left            =   120
            TabIndex        =   28
            ToolTipText     =   "If the other player disconnects, you will be asked if you want to save the log."
            Top             =   240
            WhatsThisHelpID =   10017
            Width           =   1815
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Team"
         Height          =   1095
         Left            =   -74880
         TabIndex        =   23
         Top             =   360
         Width           =   2055
         Begin VB.CheckBox chkAutoLoad 
            Caption         =   "Autoload at Startup"
            Height          =   255
            Left            =   120
            TabIndex        =   64
            ToolTipText     =   "Enable automatic updating"
            Top             =   720
            Value           =   1  'Checked
            WhatsThisHelpID =   10021
            Width           =   1815
         End
         Begin VB.PictureBox Picture2 
            BorderStyle     =   0  'None
            Height          =   495
            Left            =   120
            ScaleHeight     =   495
            ScaleWidth      =   1815
            TabIndex        =   24
            Top             =   240
            Width           =   1815
            Begin VB.OptionButton TeamHide 
               Caption         =   "Show Other Players"
               Height          =   255
               Index           =   0
               Left            =   0
               TabIndex        =   26
               ToolTipText     =   "Make the team public on the server"
               Top             =   0
               Value           =   -1  'True
               WhatsThisHelpID =   10015
               Width           =   1695
            End
            Begin VB.OptionButton TeamHide 
               Caption         =   "Hide Until Battle"
               Height          =   255
               Index           =   1
               Left            =   0
               TabIndex        =   25
               ToolTipText     =   "Hide your team until battle starts"
               Top             =   240
               WhatsThisHelpID =   10016
               Width           =   1695
            End
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Automatic Updates"
         Height          =   1335
         Left            =   -72720
         TabIndex        =   20
         Top             =   360
         Width           =   2055
         Begin VB.PictureBox Picture4 
            BorderStyle     =   0  'None
            Height          =   495
            Left            =   120
            ScaleHeight     =   495
            ScaleWidth      =   1815
            TabIndex        =   32
            Top             =   720
            Width           =   1815
            Begin VB.CommandButton UpdateScan 
               Caption         =   "&Check Version"
               Height          =   375
               Left            =   0
               TabIndex        =   33
               Top             =   120
               WhatsThisHelpID =   10023
               Width           =   1815
            End
         End
         Begin VB.CheckBox AutoUpdt 
            Caption         =   "Enabled"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   22
            ToolTipText     =   "Enable automatic updating"
            Top             =   240
            Value           =   1  'Checked
            WhatsThisHelpID =   10021
            Width           =   1815
         End
         Begin VB.CheckBox AutoUpdt 
            Caption         =   "Ask Before Updating"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   21
            ToolTipText     =   "Make updating optional"
            Top             =   480
            Value           =   1  'Checked
            WhatsThisHelpID =   10022
            Width           =   1815
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Text Display"
         Height          =   795
         Left            =   -74880
         TabIndex        =   17
         Top             =   480
         Width           =   2055
         Begin VB.CheckBox FixLines 
            Caption         =   "Fix Line Breaks"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   480
            WhatsThisHelpID =   10001
            Width           =   1575
         End
         Begin VB.CheckBox FText 
            Caption         =   "Use Fancy Text"
            Enabled         =   0   'False
            Height          =   255
            Left            =   120
            TabIndex        =   18
            ToolTipText     =   "If enabled, use colored text."
            Top             =   240
            Value           =   1  'Checked
            WhatsThisHelpID =   10000
            Width           =   1575
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Animation && Display"
         Height          =   1275
         Left            =   -72720
         TabIndex        =   14
         Top             =   480
         Width           =   2055
         Begin VB.CheckBox EnableBG 
            Caption         =   "Enable Background"
            Height          =   255
            Left            =   120
            TabIndex        =   56
            Top             =   960
            WhatsThisHelpID =   10008
            Width           =   1815
         End
         Begin VB.CheckBox UseOldInt 
            Caption         =   "Use Old Interface"
            Enabled         =   0   'False
            Height          =   255
            Left            =   120
            TabIndex        =   51
            ToolTipText     =   "This uses an 0.8.xx style Battle interface instead of the 0.9.x style."
            Top             =   720
            WhatsThisHelpID =   10011
            Width           =   1815
         End
         Begin VB.CheckBox PokeAnim 
            Caption         =   "Animate Pokémon"
            Enabled         =   0   'False
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   480
            WhatsThisHelpID =   10010
            Width           =   1815
         End
         Begin VB.CheckBox AnimHP 
            Caption         =   "Animate HP Bars"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   240
            WhatsThisHelpID =   10009
            Width           =   1815
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Battle Messages"
         Height          =   1935
         Left            =   -74880
         TabIndex        =   10
         Top             =   1320
         Width           =   2055
         Begin VB.CheckBox chkTrainername 
            Caption         =   "Trainer Name Prefix"
            Height          =   255
            Left            =   120
            TabIndex        =   61
            Top             =   960
            Width           =   1815
         End
         Begin VB.CheckBox chkNickname 
            Caption         =   "Use Nicknames"
            Height          =   315
            Left            =   120
            TabIndex        =   60
            Top             =   720
            Width           =   1865
         End
         Begin VB.TextBox txtDelay 
            Height          =   285
            Left            =   120
            MaxLength       =   4
            TabIndex        =   57
            Text            =   "2000"
            Top             =   1560
            Width           =   495
         End
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            Height          =   495
            Left            =   120
            ScaleHeight     =   495
            ScaleWidth      =   1815
            TabIndex        =   11
            Top             =   240
            Width           =   1815
            Begin VB.OptionButton MessStyle 
               Caption         =   "Normal"
               Height          =   255
               Index           =   0
               Left            =   0
               TabIndex        =   13
               ToolTipText     =   "Normal G/S/C messages"
               Top             =   0
               Value           =   -1  'True
               WhatsThisHelpID =   10002
               Width           =   1815
            End
            Begin VB.OptionButton MessStyle 
               Caption         =   "Extended (Log-Style)"
               Height          =   255
               Index           =   1
               Left            =   0
               TabIndex        =   12
               ToolTipText     =   "Extended text for creating logs."
               Top             =   240
               WhatsThisHelpID =   10003
               Width           =   1815
            End
         End
         Begin VB.Label Label4 
            Caption         =   "milliseconds"
            Height          =   255
            Left            =   660
            TabIndex        =   59
            Top             =   1620
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "Between-Move Delay "
            Height          =   255
            Left            =   120
            TabIndex        =   58
            Top             =   1320
            Width           =   1695
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Audio"
         Height          =   1455
         Left            =   -72720
         TabIndex        =   7
         Top             =   1800
         Width           =   2055
         Begin VB.PictureBox Picture3 
            BorderStyle     =   0  'None
            Height          =   375
            Left            =   120
            ScaleHeight     =   375
            ScaleWidth      =   1815
            TabIndex        =   30
            Top             =   840
            Width           =   1815
            Begin VB.CommandButton SoundButton 
               Caption         =   "Configure &Sounds..."
               Height          =   375
               Left            =   0
               TabIndex        =   31
               Top             =   0
               WhatsThisHelpID =   10014
               Width           =   1815
            End
         End
         Begin VB.CheckBox Audio 
            Caption         =   "Sound"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   9
            ToolTipText     =   "Enable chat sounds & short audio clips"
            Top             =   240
            Value           =   2  'Grayed
            WhatsThisHelpID =   10012
            Width           =   1215
         End
         Begin VB.CheckBox Audio 
            Caption         =   "Music"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   8
            ToolTipText     =   "Enable music"
            Top             =   480
            Value           =   1  'Checked
            WhatsThisHelpID =   10013
            Width           =   1335
         End
      End
      Begin VB.ComboBox LangBox 
         Height          =   315
         Left            =   -72720
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1920
         WhatsThisHelpID =   10024
         Width           =   2055
      End
      Begin VB.CommandButton ChangePW 
         Caption         =   "Stored &Password..."
         Height          =   375
         Left            =   -72720
         TabIndex        =   4
         Top             =   2520
         WhatsThisHelpID =   10025
         Width           =   2055
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Reregister .PNB Files"
         Height          =   375
         Left            =   -72720
         TabIndex        =   3
         Top             =   3000
         WhatsThisHelpID =   10026
         Width           =   2055
      End
      Begin VB.CheckBox chkTimestamp 
         Caption         =   "Show Timestamp"
         Height          =   255
         Left            =   2640
         TabIndex        =   49
         Top             =   720
         Value           =   1  'Checked
         WhatsThisHelpID =   10028
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "Number of Lines:"
         Height          =   255
         Left            =   2640
         TabIndex        =   63
         Top             =   2280
         Width           =   1275
      End
      Begin VB.Label Label3 
         Caption         =   "Format:"
         Height          =   255
         Left            =   2640
         TabIndex        =   55
         Top             =   1020
         Width           =   1575
      End
      Begin VB.Label lblSample 
         Caption         =   "Test"
         Height          =   255
         Left            =   2640
         TabIndex        =   54
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Sample:"
         Height          =   255
         Left            =   2640
         TabIndex        =   53
         Top             =   1680
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label LangStat 
         Height          =   735
         Left            =   -72720
         TabIndex        =   6
         Top             =   1080
         Width           =   1815
         WordWrap        =   -1  'True
      End
   End
End
Attribute VB_Name = "Options"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim NewPassword As String
'Dim WithEvents DX As clsDX
Dim Guid() As String

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

Private Sub AutosaveL_Click()
    If AutosaveL.Value = 1 Then PromptIt.Enabled = False Else PromptIt.Enabled = True
End Sub

Private Sub AutosaveL_KeyUp(KeyCode As Integer, Shift As Integer)
    If AutosaveL.Value = 1 Then PromptIt.Enabled = False Else PromptIt.Enabled = True
End Sub

Private Sub AutosaveL_LostFocus()
    If AutosaveL.Value = 1 Then PromptIt.Enabled = False Else PromptIt.Enabled = True
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

Private Sub CancelButton_Click()
    'UseDX = GetSetting("NetBattle", "DirectX", "Use", True)
    UseHiResTimer = GetSetting("NetBattle", "DirectX", "Timer", True)
    RefreshRate = GetSetting("NetBattle", "DirectX", "Interval", 2)
    DeviceGUID = GetSetting("NetBattle", "DirectX", "Device", "")
    Unload Me
End Sub

Private Sub ChangePW_Click()
    PasswordBoxTitle = "Edit Saved Password"
    PasswordBoxCaption = "Enter the user password to be saved."
    PWWindow.Show 1
    NewPassword = ServerPassword
    If NewPassword <> "" Then
        PasswordBoxTitle = "Verify Password"
        PasswordBoxCaption = "Enter the password again for verification."
        PWWindow.Show 1
        If NewPassword <> ServerPassword Then
            MsgBox "The two passwords do not match!", vbCritical, "Error"
        End If
    Else
        ServerPassword = ""
    End If
    ServerPassword = MD5(ServerPassword)
End Sub

'Private Sub chkMoveAnims_Click()
'    UseDX = (chkMoveAnims.Value = 1)
'    chkTimer.Enabled = UseDX
'    cmdTest.Enabled = UseDX
'    cmbDev.Enabled = UseDX
'    txtRefresh.Enabled = (chkTimer.Value = 1 And UseDX)
'End Sub
'
'Private Sub chkTimer_Click()
'    txtRefresh.Enabled = (chkTimer.Value = 1)
'    UseHiResTimer = txtRefresh.Enabled
'End Sub
'
Private Sub chkTimestamp_Click()
    txtTSFormat.Enabled = (chkTimestamp.Value <> 0)
    Label1.Visible = txtTSFormat.Enabled
    lblSample.Visible = txtTSFormat.Enabled
    Call Timer1_Timer
End Sub
'
'Private Sub cmbDev_Click()
'    DeviceGUID = Guid(cmbDev.ListIndex + 1)
'End Sub

Private Sub cmdFilter_Click()
    Dim Temp As String
    Dim X As Long
    Temp = InputBox("Please enter the words you would like filtered, delimited by semicolons (;).", "Word Filter", Join(CSFilter, ";"))
    CSFilter = Split(Temp, ";")
    For X = 0 To UBound(CSFilter)
        CSFilter(X) = LCase$(Trim$(CSFilter(X)))
    Next X
    SaveSetting "NetBattle", "Options", "CSFilter", Join(CSFilter, ";")
End Sub

Private Sub cmdHelp_Click()
    Call ShowHelpContext(20000)
End Sub

'Private Sub cmdTest_Click()
'    If cmdTest.Caption = "Start Test" Then
'        Set DX = New clsDX
'        If Not DX.InitDX(picTest) Then
'            Set DX = Nothing
'            Exit Sub
'        End If
'        MainContainer.DoPicture "bg0.gif"
'        DX.CreateSurfaceFromPBox MainContainer.SwapSpace
'        DX.AddAnim 2, 1, 0, 0
'        cmdTest.Caption = "Stop Test"
'        chkMoveAnims.Enabled = False
'        chkTimer.Enabled = False
'        txtRefresh.Enabled = False
'        cmbDev.Enabled = False
'        cmdHelp.Enabled = False
'        OKButton.Enabled = False
'        CancelButton.Enabled = False
'
'        Do While DX.Animating
'            Sleep 1
'            DoEvents
'        Loop
'
'        Set DX = Nothing
'        cmdHelp.Enabled = True
'        OKButton.Enabled = True
'        CancelButton.Enabled = True
'        chkMoveAnims.Enabled = True
'        Set picTest.Picture = Nothing
'        Call chkMoveAnims_Click
'        cmdTest.Caption = "Start Test"
'    Else
'        DX.AnimFinished 1
'    End If
'End Sub

Private Sub Command1_Click()
    Dim DidFileReg As Boolean
    Dim DidReplReg As Boolean
    
    DidFileReg = CreateFileAss(".pnb", "NetBattle.Team", "NetBattle Team", "Open", SlashPath & "PokeBattle.exe", , True, SlashPath & "PokeBattle.exe,1", True)
    DidReplReg = CreateFileAss(".btl", "NetBattle.Replay", "NetBattle Replay", "Open", SlashPath & "PokeBattle.exe", , True, SlashPath & "PokeBattle.exe,7", True)
    If DidFileReg And DidReplReg Then
        MsgBox "File Types successfully registered", vbInformation, "Done"
    Else
        MsgBox "Error registering file type", vbExclamation, "Error"
    End If
End Sub

Private Sub DX_BltError()
    MsgBox "DirectX was initialized sucessfully, but encountered an error when trying to render the display.  Try different settings to fix this.", vbCritical, "Rendering Failed"
End Sub

Private Sub Command2_Click()

End Sub

Private Sub Form_Initialize()
    Dim X As Long
    'X = 'InitCommonControls
End Sub

Private Sub Form_Load()
    Dim X As Integer
    
    Audio(0).Value = SoundOption
    Audio(1).Value = MusicOption
    AnimHP.Value = AnimOption
    If UseBG Then EnableBG.Value = 1 Else EnableBG.Value = 0
    If SoundOption = 1 And MusicOption = 1 Then Audio(0).Value = 2
    AutoUpdt(0).Value = AutoScan
    AutoUpdt(1).Value = AskOnUpdate
    If AutoUpdt(0).Value = 0 Then AutoUpdt(1).Enabled = False
    If AllowViewing = 1 Then TeamHide(0).Value = True Else TeamHide(1).Value = True
    If LogPrompt = 1 Then PromptIt.Value = 1
    If LogSave = 1 Then AutosaveL.Value = 1: PromptIt.Enabled = False
    If ReplayPrompt = 1 Then PromptR.Value = 1
    If Autosave = 1 Then SaveIt.Value = 1: PromptR.Value = False
    If FancyText Then FText.Value = 1 Else FText.Value = 0
    If AddLineBreaks Then FixLines.Value = 1
    If OldInterface Then UseOldInt.Value = 1
    If Autoload Then chkAutoLoad.Value = 1 Else chkAutoLoad.Value = 0
    txtDelay.Text = MoveDelay
    chkNickname.Value = IIf(UseNicks, 1, 0)
    chkTrainername.Value = IIf(UsePrefix, 1, 0)
    Select Case BMessStyle
        Case 0
            MessStyle(0).Value = True
            MessStyle(1).Value = False
        Case 1
            MessStyle(1).Value = True
            MessStyle(0).Value = False
    End Select
    chkTimestamp.Value = IIf(UseTS, 1, 0)
    Label1.Visible = UseTS
    lblSample.Visible = UseTS
    If UseTS Then Call Timer1_Timer
    txtTSFormat.Text = TSFormat
    chkURLs.Value = IIf(ParseURLs, 1, 0)
    txtLines.Text = CStr(DisplayLines)
    For X = 1 To 10
        MessageToggle(X).Value = IIf(MsgToggle(X), 1, 0)
    Next X
    
    On Error GoTo TheEnd
    For X = 0 To UBound(LFile)
        If LFile(X).Text <> "" Then
            LangBox.AddItem LFile(X).Text, X
            If LFile(X).FileName = CurrLang Then LangBox.ListIndex = X
        End If
    Next
    
    'FillInDeviceList
    'chkMoveAnims.Value = Abs(UseDX)
    'chkTimer.Value = Abs(UseHiResTimer)
    'txtRefresh = CStr(RefreshRate)
'    For X = 1 To UBound(Guid)
'        If Guid(X) = DeviceGUID Then
'            cmbDev.ListIndex = X - 1
'        End If
'    Next X
'    If X = UBound(Guid) + 1 Then cmbDev.ListIndex = X - 2
'    Call chkTimer_Click
'    Call chkMoveAnims_Click
    
TheEnd:
    If CurrLang = "" Then LangBox.ListIndex = 0
    SSTab1.Tab = 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'If Not (DX Is Nothing) Then Cancel = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Loader.Visible = True
End Sub

Private Sub Image1_Click()

End Sub

Private Sub LangBox_Click()
    Dim Temp As String
    
    Temp = ""
    If LFile(LangBox.ListIndex).HasMoves Then Temp = "Moves"
    If LFile(LangBox.ListIndex).HasPKMN Then
        If Temp = "" Then
            Temp = "Pokémon"
        Else
            Temp = Temp & ", Pokémon"
        End If
    End If
    If LFile(LangBox.ListIndex).HasBattle Then
        If Temp = "" Then
            Temp = "Battle"
        Else
            Temp = Temp & ", Battle"
        End If
    End If
    If LFile(LangBox.ListIndex).HasPDEX Then
        If Temp = "" Then
            Temp = "PokéDex"
        Else
            Temp = Temp & ", PokéDex"
        End If
    End If
    If LFile(LangBox.ListIndex).HasProgram Then
        If Temp = "" Then
            Temp = "Program"
        Else
            Temp = Temp & ", Program"
        End If
    End If
    If LFile(LangBox.ListIndex).HasMisc Then
        If Temp = "" Then
            Temp = "Misc."
        Else
            Temp = Temp & ", Misc."
        End If
    End If
    LangStat.Caption = Temp
End Sub

Private Sub LangBox_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim Temp As String
    
    Temp = ""
    If LFile(LangBox.ListIndex).HasMoves Then Temp = "Moves"
    If LFile(LangBox.ListIndex).HasPKMN Then
        If Temp = "" Then Temp = "Pokémon"
    Else
        Temp = Temp & ", Pokémon"
    End If
    If LFile(LangBox.ListIndex).HasBattle Then
        If Temp = "" Then Temp = "Battle"
    Else
        Temp = Temp & ", Battle"
    End If
    If LFile(LangBox.ListIndex).HasPDEX Then
        If Temp = "" Then Temp = "PokéDex"
    Else
        Temp = Temp & ", PokéDex"
    End If
    If LFile(LangBox.ListIndex).HasProgram Then
        If Temp = "" Then Temp = "Program"
    Else
        Temp = Temp & ", Program"
    End If
    If LFile(LangBox.ListIndex).HasMisc Then
        If Temp = "" Then Temp = "Misc."
    Else
        Temp = Temp & ", Misc."
    End If
    LangStat.Caption = Temp
End Sub

Private Sub OKButton_Click()
    Dim X As Long
    Dim Temp As String
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
    
    If EnableBG.Value = 1 Then UseBG = True Else UseBG = False
    SaveSetting "NetBattle", "Options", "Use Background", UseBG
    
    If TeamHide(0).Value Then
        AllowViewing = 1
    Else
        AllowViewing = 0
    End If
    SaveSetting "NetBattle", "Options", "Allow Viewing", AllowViewing
    
    Autoload = (chkAutoLoad.Value = 1)
    SaveSetting "NetBattle", "Options", "AutoLoad", Autoload
    
    LogPrompt = PromptIt.Value
    SaveSetting "NetBattle", "Options", "Log Prompt", LogPrompt
    LogSave = AutosaveL.Value
    SaveSetting "NetBattle", "Options", "Log Save", LogSave
    ReplayPrompt = PromptR.Value
    SaveSetting "NetBattle", "Options", "Replay Prompt", ReplayPrompt
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
        SavedPassword = ServerPassword
        SaveSetting "NetBattle", "Options", "Saved Password", SavedPassword
    End If
    
    If FText.Value = 1 Then FancyText = True Else FancyText = False
    SaveSetting "NetBattle", "Options", "Fancy Text", FancyText
    
    If FixLines.Value = 1 Then AddLineBreaks = True Else AddLineBreaks = False
    SaveSetting "NetBattle", "Options", "Line Breaks", AddLineBreaks
    
    If UseOldInt.Value = 1 Then OldInterface = True Else OldInterface = False
    SaveSetting "NetBattle", "Options", "Old Interface", OldInterface
    
    If LFile(LangBox.ListIndex).FileName <> CurrLang Then
        CurrLang = LFile(LangBox.ListIndex).FileName
        SaveSetting "NetBattle", "Options", "Language", CurrLang
        Call Loader.DoLanguage
    End If
    
    MoveDelay = Val(txtDelay.Text)
    SaveSetting "NetBattle", "Options", "MoveDelay", MoveDelay
    UseNicks = CBool(chkNickname.Value)
    UsePrefix = CBool(chkTrainername.Value)
    SaveSetting "NetBattle", "Options", "UseNicks", UseNicks
    SaveSetting "NetBattle", "Options", "UsePrefix", UsePrefix
    
    UseTS = txtTSFormat.Enabled
    SaveSetting "NetBattle", "Options", "UseTS", UseTS
    TSFormat = txtTSFormat.Text
    SaveSetting "NetBattle", "Options", "TSFormat", TSFormat
    ParseURLs = (chkURLs.Value <> 0)
    SaveSetting "NetBattle", "Options", "ParseURLs", ParseURLs
    DisplayLines = Val(txtLines.Text)
    SaveSetting "NetBattle", "Options", "Lines", DisplayLines
    Temp = ""
    For X = 1 To 10
        MsgToggle(X) = CBool(MessageToggle(X).Value)
        Temp = Temp & CStr(MessageToggle(X).Value)
    Next X
    X = Bin2Dec(Temp)
    SaveSetting "NetBattle", "Options", "MsgToggles", X
    
    SaveSetting "NetBattle", "DirectX", "Use", UseDX
    SaveSetting "NetBattle", "DirectX", "Timer", UseHiResTimer
    SaveSetting "NetBattle", "DirectX", "Interval", RefreshRate
    SaveSetting "NetBattle", "DirectX", "Device", DeviceGUID
    
    
    
       
    Unload Me
End Sub

Private Sub SaveIt_Click()
    If SaveIt.Value = 1 Then PromptR.Enabled = False Else PromptR.Enabled = True
End Sub

Private Sub SaveIt_KeyUp(KeyCode As Integer, Shift As Integer)
    If SaveIt.Value = 1 Then PromptR.Enabled = False Else PromptR.Enabled = True
End Sub

Private Sub SaveIt_LostFocus()
    If SaveIt.Value = 1 Then PromptR.Enabled = False Else PromptR.Enabled = True
End Sub

Private Sub SoundButton_Click()
    On Error Resume Next
    Unload SoundPick
    SoundPick.Show 1, Options
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    'If Not (DX Is Nothing) Then SSTab1.Tab = 3
End Sub

Private Sub tmrDX_Timer()
    'If Not (UseHiResTimer Or DX Is Nothing) Then DX.Blt
End Sub

Private Sub txtlines_LostFocus()
    If Not IsNumeric(txtLines.Text) Then txtLines.Text = "1000"
    If Val(txtLines.Text) > 99999 Or Val(txtLines.Text) < 100 Then txtLines.Text = "1000"
End Sub

Private Sub txtDelay_LostFocus()
    Dim X As Integer
    On Error GoTo Overflow
    X = 2000
    X = Val(txtDelay.Text)
    If X > 5000 Then X = 5000
    If X < 0 Then X = 0
Overflow:
    txtDelay.Text = CStr(X)
End Sub

Private Sub Timer1_Timer()
    Dim Temp As String
    If chkTimestamp.Value = 1 Then
        Temp = Replace(Format(Now, Replace(txtTSFormat.Text, "[", Chr(1))), Chr(1), "[")
        If lblSample.Caption <> Temp Then lblSample.Caption = Temp
    End If
End Sub

'Private Sub txtRefresh_LostFocus()
'    Dim X As Long
'    X = Val(txtRefresh.Text)
'    If X < 1 Then X = 1
'    If X > 16 Then X = 16
'    txtRefresh.Text = CStr(X)
'    RefreshRate = X
'End Sub
'
Private Sub UpdateScan_Click()
    Call Loader.DoVersionScan
End Sub








'Private Sub FillInDeviceList()
'    Dim DX1 As DirectX7
'    Dim DX2 As DirectDraw7
'    Dim DX3 As Direct3D7
'    Dim DX4 As Direct3DEnumDevices
'    Dim Z As Long
'    On Error GoTo Failed
'    cmbDev.Clear
'    Set DX1 = New DirectX7
'    Set DX2 = DX1.DirectDrawCreate("")
'    Set DX3 = DX2.GetDirect3D
'    Set DX4 = DX3.GetDevicesEnum
'    ReDim Guid(1 To DX4.GetCount)
'    For Z = 1 To DX4.GetCount
'        cmbDev.AddItem DX4.GetName(Z)
'        Guid(Z) = DX4.GetGuid(Z)
'    Next Z
'    Set DX4 = Nothing
'    Set DX3 = Nothing
'    Set DX2 = Nothing
'    Set DX1 = Nothing
'Failed:
'End Sub
