VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Begin VB.Form SetUsers 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Server Options"
   ClientHeight    =   4695
   ClientLeft      =   2805
   ClientTop       =   3825
   ClientWidth     =   4335
   Icon            =   "SetUsers.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   4335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton ApplyButton 
      Cancel          =   -1  'True
      Caption         =   "Apply"
      Height          =   375
      Left            =   1560
      TabIndex        =   45
      Top             =   4200
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3975
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   7011
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "SetUsers.frx":1272
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label8"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label10"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label9"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label22"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Slider3"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Slider1"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "ServerMessageBox"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "AutoLog"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtExpire"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Slider2"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).ControlCount=   14
      TabCaption(1)   =   "Security"
      TabPicture(1)   =   "SetUsers.frx":128E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label7"
      Tab(1).Control(1)=   "FloodCaption"
      Tab(1).Control(2)=   "Label6"
      Tab(1).Control(3)=   "Label5"
      Tab(1).Control(4)=   "Label20"
      Tab(1).Control(5)=   "Label21"
      Tab(1).Control(6)=   "FloodSlider"
      Tab(1).Control(7)=   "PWBox2"
      Tab(1).Control(8)=   "PWBox"
      Tab(1).Control(9)=   "OldVer"
      Tab(1).Control(10)=   "NewUsr"
      Tab(1).Control(11)=   "txtBanMsg"
      Tab(1).Control(12)=   "chkEncrypt"
      Tab(1).Control(13)=   "txtMaxIPs"
      Tab(1).ControlCount=   14
      TabCaption(2)   =   "Registry"
      TabPicture(2)   =   "SetUsers.frx":12AA
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "optServer(1)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "optServer(0)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Frame1"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Timer1"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "Random Numbers"
      TabPicture(3)   =   "SetUsers.frx":12C6
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame2"
      Tab(3).Control(1)=   "Frame3"
      Tab(3).ControlCount=   2
      Begin VB.Timer Timer1 
         Interval        =   1
         Left            =   -74700
         Top             =   2640
      End
      Begin VB.Frame Frame1 
         Caption         =   "Registry Settings"
         Height          =   2775
         Left            =   -74880
         TabIndex        =   46
         Top             =   960
         Width           =   3855
         Begin VB.CommandButton cmdReg 
            Caption         =   "Register"
            Height          =   255
            Left            =   2700
            TabIndex        =   55
            Top             =   2280
            Width           =   915
         End
         Begin VB.TextBox txtDesc 
            Height          =   1095
            Left            =   1200
            MaxLength       =   190
            MultiLine       =   -1  'True
            TabIndex        =   49
            Text            =   "SetUsers.frx":12E2
            Top             =   1080
            Width           =   2415
         End
         Begin VB.TextBox txtAdmin 
            Height          =   285
            Left            =   1200
            MaxLength       =   20
            TabIndex        =   48
            Text            =   "John Doe"
            Top             =   720
            Width           =   2415
         End
         Begin VB.TextBox txtName 
            Height          =   285
            Left            =   1200
            MaxLength       =   20
            TabIndex        =   47
            Text            =   "NetBattle Server"
            Top             =   360
            Width           =   2415
         End
         Begin VB.Label Label23 
            Alignment       =   1  'Right Justify
            Caption         =   "Registration:"
            Height          =   255
            Left            =   120
            TabIndex        =   54
            Top             =   2280
            Width           =   975
         End
         Begin VB.Label lblRegged 
            BackStyle       =   0  'Transparent
            Caption         =   "Unregistered"
            Height          =   255
            Left            =   1200
            TabIndex        =   53
            Top             =   2280
            Width           =   1695
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            Caption         =   "Description:"
            Height          =   255
            Left            =   120
            TabIndex        =   52
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            Caption         =   "Main Admin:"
            Height          =   255
            Left            =   120
            TabIndex        =   51
            Top             =   720
            Width           =   975
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            Caption         =   "Server Name:"
            Height          =   255
            Left            =   120
            TabIndex        =   50
            Top             =   360
            Width           =   975
         End
      End
      Begin MSComctlLib.Slider Slider2 
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   1920
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   873
         _Version        =   393216
         Max             =   25
         SelStart        =   10
         Value           =   10
      End
      Begin VB.TextBox txtExpire 
         Height          =   285
         Left            =   240
         TabIndex        =   40
         Text            =   "90"
         Top             =   3600
         Width           =   615
      End
      Begin VB.TextBox txtMaxIPs 
         Height          =   285
         Left            =   -74880
         TabIndex        =   38
         Text            =   "0"
         Top             =   3600
         Width           =   615
      End
      Begin VB.CheckBox chkEncrypt 
         Caption         =   "Use encryption  (Requires restart)"
         Height          =   255
         Left            =   -74880
         TabIndex        =   37
         Top             =   2760
         Value           =   1  'Checked
         Width           =   3735
      End
      Begin VB.TextBox txtBanMsg 
         Height          =   285
         Left            =   -74880
         TabIndex        =   35
         Top             =   1800
         Width           =   3855
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   -74880
         TabIndex        =   33
         Top             =   360
         Width           =   3735
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            Height          =   495
            Left            =   0
            ScaleHeight     =   495
            ScaleWidth      =   3735
            TabIndex        =   42
            Top             =   240
            Width           =   3735
            Begin VB.OptionButton OptRnd 
               Caption         =   "True Random: Uses numbers from Random.org."
               Height          =   255
               Index           =   1
               Left            =   0
               TabIndex        =   44
               Top             =   240
               Width           =   3735
            End
            Begin VB.OptionButton OptRnd 
               Caption         =   "Pseudo Random: Uses no extra bandwidth."
               Height          =   255
               Index           =   0
               Left            =   0
               TabIndex        =   43
               Top             =   0
               Width           =   3615
            End
         End
         Begin VB.Label Label15 
            Caption         =   "Choose random number system:"
            Height          =   255
            Left            =   0
            TabIndex        =   34
            Top             =   0
            Width           =   2415
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "True Random Settings"
         Height          =   2535
         Left            =   -74880
         TabIndex        =   26
         Top             =   1320
         Width           =   3855
         Begin VB.Timer RndTimer 
            Interval        =   1
            Left            =   3360
            Top             =   2040
         End
         Begin VB.TextBox txtRnd 
            Height          =   285
            Index           =   1
            Left            =   240
            TabIndex        =   30
            Text            =   "128"
            Top             =   1320
            Width           =   615
         End
         Begin VB.TextBox txtRnd 
            Height          =   285
            Index           =   0
            Left            =   240
            TabIndex        =   27
            Text            =   "1024"
            Top             =   600
            Width           =   615
         End
         Begin VB.Label Label19 
            Caption         =   "Curren State: Empty"
            Height          =   255
            Left            =   240
            TabIndex        =   32
            Top             =   2160
            Width           =   3015
         End
         Begin VB.Label Label18 
            Caption         =   "Current Cache: 0 Numbers"
            Height          =   255
            Left            =   240
            TabIndex        =   31
            Top             =   1920
            Width           =   3375
         End
         Begin VB.Label Label17 
            Caption         =   "Query server when cache gets how low?"
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   1080
            Width           =   3495
         End
         Begin VB.Label Label16 
            Caption         =   "Download how many random numbers at a time?"
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   360
            Width           =   3495
         End
      End
      Begin VB.OptionButton optServer 
         Caption         =   "Private Server - Do not list in Server Registry "
         Height          =   255
         Index           =   0
         Left            =   -74880
         TabIndex        =   25
         Top             =   360
         Width           =   3855
      End
      Begin VB.OptionButton optServer 
         Caption         =   "Public Server - List in Server Registry"
         Height          =   255
         Index           =   1
         Left            =   -74880
         TabIndex        =   24
         Top             =   600
         Value           =   -1  'True
         Width           =   3495
      End
      Begin VB.CheckBox AutoLog 
         Caption         =   "Auto-Logging"
         Height          =   255
         Left            =   2520
         TabIndex        =   20
         Top             =   3600
         Width           =   1455
      End
      Begin VB.CheckBox NewUsr 
         Caption         =   "Allow new users"
         Height          =   255
         Left            =   -74880
         TabIndex        =   19
         Top             =   2520
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.CheckBox OldVer 
         Caption         =   "Allow old versions"
         Height          =   255
         Left            =   -74880
         TabIndex        =   18
         Top             =   2280
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.TextBox PWBox 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   -74880
         MaxLength       =   16
         PasswordChar    =   "*"
         TabIndex        =   12
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox PWBox2 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   -74880
         MaxLength       =   16
         PasswordChar    =   "*"
         TabIndex        =   11
         Top             =   1200
         Width           =   1935
      End
      Begin VB.TextBox ServerMessageBox 
         Height          =   285
         Left            =   120
         MaxLength       =   128
         TabIndex        =   4
         Top             =   600
         Width           =   3855
      End
      Begin MSComctlLib.Slider Slider1 
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   873
         _Version        =   393216
         LargeChange     =   10
         Min             =   10
         Max             =   255
         SelStart        =   100
         TickFrequency   =   24
         Value           =   100
      End
      Begin MSComctlLib.Slider FloodSlider 
         Height          =   495
         Left            =   -72840
         TabIndex        =   13
         Top             =   3360
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   873
         _Version        =   393216
         Min             =   2
         Max             =   25
         SelStart        =   10
         Value           =   10
      End
      Begin MSComctlLib.Slider Slider3 
         Height          =   495
         Left            =   120
         TabIndex        =   21
         Top             =   2640
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   873
         _Version        =   393216
         LargeChange     =   20
         SmallChange     =   10
         Max             =   100
         SelStart        =   100
         TickFrequency   =   10
         Value           =   100
      End
      Begin VB.Label Label22 
         Caption         =   "User account inactivty expiration: (In Days)"
         Height          =   495
         Left            =   120
         TabIndex        =   41
         Top             =   3120
         Width           =   1815
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Number of lines to display:"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   2400
         Width           =   2415
      End
      Begin VB.Label Label21 
         Caption         =   "Max Connections per IP  (0 for no limit)"
         Height          =   495
         Left            =   -74880
         TabIndex        =   39
         Top             =   3120
         Width           =   1695
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Default Ban Message:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   36
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "10000"
         Height          =   255
         Left            =   2640
         TabIndex        =   23
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Connection Password"
         Height          =   255
         Left            =   -74880
         TabIndex        =   17
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Verify"
         Height          =   255
         Left            =   -74880
         TabIndex        =   16
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label FloodCaption 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "10"
         Height          =   255
         Left            =   -72480
         TabIndex        =   15
         Top             =   3120
         Width           =   1335
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Flood Tolerance"
         Height          =   255
         Left            =   -72720
         TabIndex        =   14
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Number of users:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "100"
         Height          =   255
         Left            =   2640
         TabIndex        =   9
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Welcome Message"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Send Rate (Lower = Faster)"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1680
         Width           =   2055
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "10"
         Height          =   255
         Left            =   2640
         TabIndex        =   6
         Top             =   1680
         Width           =   1335
      End
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3000
      TabIndex        =   0
      Top             =   4200
      Width           =   1215
   End
End
Attribute VB_Name = "SetUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim T As Single

Private Sub ApplyButton_Click()
    Call ApplySettings(False)
End Sub

Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdReg_Click()
    Dim Temp As String
    Temp = InputBox("By setting a password here, you can protect your Server Name from being used by others.  The name will be keyed to your Station ID.  If a hardware change causes your Station ID to change, you will need to enter this password again.  You can only register one Server Name at a time.", "Set Password")
    If Temp = InputBox("Please type the password again to confirm.", "Confirm") Then
        cmdReg.Enabled = False
        ServerWindow.SendRegData ("PASS:" & MD5(Temp))
    Else
        MsgBox "The passwords did not match.  Please try again."
    End If
    
    
End Sub

Private Sub FloodSlider_Click()
    FloodCaption.Caption = FloodSlider.Value
End Sub

Private Sub Form_Load()
    Dim B As Boolean
    Slider1.Value = MaxUsers
    Label2.Caption = MaxUsers
    FloodSlider.Value = FloodTolerance
    FloodCaption.Caption = FloodTolerance
    Slider2.Value = SendTimer
    Label8.Caption = SendTimer
    If Slider2.Value = 0 Then Label8.Caption = "No Queueing"
    ServerMessageBox.Text = ServerMessage
    PWBox.Text = ServerPassword
    PWBox2.Text = ServerPassword
    txtBanMsg.Text = DefaultBanMsg
    B = GetSetting("NetBattle", "Master Server", "Encrypt", True)
    chkEncrypt.Value = Abs(B)
    OldVer.Value = AllowOldVersions
    NewUsr.Value = AllowNewUsers
    AutoLog.Value = AutoLogging
    Slider3.Value = NumLines / 100
    Label10.Caption = NumLines
    txtName.Text = ServerName
    txtAdmin.Text = Admin
    txtDesc.Text = ServerDesc
    optServer(0).Value = Not PublicServer
    optServer(1).Value = PublicServer
    'txtIP.Text = RealIP
    txtExpire.Text = CStr(PurgeDays)
    OptRnd(0).Value = Not UseTrueRnd
    OptRnd(1).Value = UseTrueRnd
    txtRnd(0).Text = CStr(RndGroup)
    txtRnd(1).Text = CStr(RndThresh)
    If RndState = rReady Then Label19.Caption = "Current State: Ready"
    txtMaxIPs.Text = CStr(MaxIPs)
'    If RealIP = "" Then
'        Label14.Visible = False
'        txtIP.Visible = False
'        txtDesc.Height = 1335
'    Else
'        Label14.Visible = True
'        txtIP.Visible = True
'        txtDesc.Height = 855
'    End If
    Frame1.Enabled = PublicServer
End Sub

Private Sub OKButton_Click()
    Call ApplySettings(True)
End Sub
Sub ApplySettings(Final As Boolean)
    On Error Resume Next
    Dim X As Integer
    
    If PWBox.Text <> PWBox2.Text Then
        MsgBox "The two passwords must match!", vbCritical, "Verification Error"
        Exit Sub
    End If
    If optServer(1) And (txtName = "" Or txtAdmin = "" Or txtDesc = "") Then
        MsgBox "For a server to be public, all information must be supplied.", vbCritical, "Verification Error"
        Exit Sub
    End If
    If FloodSlider.Value <> FloodTolerance Then
        Call ServerWindow.SendAll("FTCG:" & FloodSlider.Value)
        FloodTolerance = FloodSlider.Value
        Call ServerWindow.AddMessage("Floodcount Tolerance changed to " & FloodTolerance)
        SaveSetting "NetBattle", "Master Server", "Floodcount Tolerance", FloodTolerance
    End If
    If Slider1.Value <> MaxUsers Then
        Call ServerWindow.SendAll("MUCG:" & Slider1.Value)
        MaxUsers = Slider1.Value
        Call ServerWindow.AddMessage("Max Users changed to " & MaxUsers)
        SaveSetting "NetBattle", "Master Server", "Max Users", MaxUsers
        Call ServerWindow.SendRegData("MAXC:" & MaxUsers)
    End If
    If Slider2.Value <> SendTimer Then
        SendTimer = Slider2.Value
        If SendTimer > 0 Then ServerWindow.SendAllQueue.Interval = SendTimer
        Call ServerWindow.AddMessage("Send Rate changed to " & SendTimer)
        SaveSetting "NetBattle", "Master Server", "Timer Interval", SendTimer
    End If
    If ServerMessageBox.Text <> ServerMessage Then
        ServerMessage = ServerMessageBox.Text
        Call ServerWindow.AddMessage("Welcome message changed to " & ServerMessage)
        SaveSetting "NetBattle", "Master Server", "Server Message", ServerMessage
    End If
    If PWBox.Text <> ServerPassword Then
        ServerPassword = PWBox.Text
        Call ServerWindow.AddMessage("Server Password Changed")
        SaveSetting "NetBattle", "Master Server", "Server Password", ServerPassword
    End If
    If AutoLog.Value = 1 And AutoLogging = 0 Then
        On Error Resume Next
        Close #LogFileNum
        Call ServerWindow.StartAutoLog
        SaveSetting "NetBattle", "Master Server", "AutoLogging", AutoLogging
    ElseIf AutoLog.Value = False And AutoLogging = 1 Then
        Call ServerWindow.StopAutoLog
        SaveSetting "NetBattle", "Master Server", "AutoLogging", AutoLogging
    End If
    If Slider3.Value * 100 <> NumLines Then
        NumLines = Slider3.Value * 100
        SaveSetting "NetBattle", "Master Server", "Display Lines", NumLines
    End If
    If ServerName <> txtName Then
        SaveSetting "NetBattle", "Master Server", "Name", txtName
        ServerName = txtName
        If ConnectedToReg And optServer(1).Value Then Call ServerWindow.SendRegData("NAMC:" & ServerName)
        Call ServerWindow.AddMessage("Server Name changed to " & ServerName)
    End If
    If Admin <> txtAdmin Then
        SaveSetting "NetBattle", "Master Server", "Admin", txtAdmin
        Admin = txtAdmin
        If ConnectedToReg And optServer(1).Value Then Call ServerWindow.SendRegData("ADMC:" & Admin)
        Call ServerWindow.AddMessage("Server Admin changed to " & Admin)
    End If
    If ServerDesc <> txtDesc Then
        SaveSetting "NetBattle", "Master Server", "Desc", txtDesc
        ServerDesc = txtDesc
        If ConnectedToReg And optServer(1).Value Then Call ServerWindow.SendRegData("DESC:" & ServerDesc)
        Call ServerWindow.AddMessage("Server Description changed to " & ServerDesc)
    End If
    If PublicServer <> optServer(1).Value Then
        PublicServer = optServer(1).Value
        SaveSetting "NetBattle", "Master Server", "Public", PublicServer
        With ServerWindow.RegSocket
            If PublicServer Then
                .Close
                .RemoteHost = RegAddress
                .RemotePort = RegPort
                .Connect
                ServerWindow.RegTimer.Enabled = True
            Else
                If .State = sckConnected Then Call ServerWindow.SendRegData("EXIT:")
            End If
        End With
    End If
    AllowNewUsers = NewUsr.Value
    SaveSetting "NetBattle", "Master Server", "Allow New Users", AllowNewUsers
    AllowOldVersions = OldVer.Value
    SaveSetting "NetBattle", "Master Server", "Allow Old Versions", AllowOldVersions
    UseTrueRnd = OptRnd(1).Value
    RndGroup = CLng(txtRnd(0).Text)
    RndThresh = CLng(txtRnd(1).Text)
    ServerWindow.RndTimer.Enabled = UseTrueRnd
    SaveSetting "NetBattle", "Master Server", "UseTrueRnd", UseTrueRnd
    SaveSetting "NetBattle", "Master Server", "RndGroup", RndGroup
    SaveSetting "NetBattle", "Master Server", "RndThresh", RndThresh
    MaxIPs = Val(txtMaxIPs.Text)
    SaveSetting "NetBattle", "Master Server", "MaxIPs", MaxIPs
    DefaultBanMsg = txtBanMsg.Text
    SaveSetting "NetBattle", "Master Server", "Ban Message", DefaultBanMsg
    PurgeDays = Val(txtExpire.Text)
    SaveSetting "NetBattle", "Master Server", "Purge Days", PurgeDays
    'UseXOR = CBool(chkEncrypt.Value)
    'with lag, this causes just too many problems.  Better to require a restart.
    SaveSetting "NetBattle", "Master Server", "Encrypt", CBool(chkEncrypt.Value)
    If Final Then Unload Me
End Sub
Private Sub OptRnd_Click(Index As Integer)
    txtRnd(0).Enabled = CBool(Index)
    txtRnd(1).Enabled = CBool(Index)
    Frame2.Enabled = CBool(Index)
    Label16.Enabled = CBool(Index)
    Label17.Enabled = CBool(Index)
    Label18.Enabled = CBool(Index)
    Label19.Enabled = CBool(Index)
End Sub

Private Sub optServer_Click(Index As Integer)
    Frame1.Enabled = optServer(1).Value
    txtName.Enabled = optServer(1).Value
    txtDesc.Enabled = optServer(1).Value
    txtAdmin.Enabled = optServer(1).Value
    'txtIP.Enabled = optServer(1).value
    Label11.Enabled = optServer(1).Value
    Label12.Enabled = optServer(1).Value
    Label13.Enabled = optServer(1).Value
    'Label14.Enabled = optServer(1).value
End Sub

Private Sub Slider1_Click()
    Label2.Caption = Slider1.Value
End Sub

Private Sub Slider2_Click()
    Label8.Caption = Slider2.Value
    If Slider2.Value = 0 Then Label8.Caption = "No Queueing"
End Sub

Private Sub Slider3_Change()
    Dim X As Long
    X = Slider3.Value * 100
    If X = 0 Then Label10.Caption = "No Limit" Else Label10.Caption = CStr(X)

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    If PreviousTab = 3 Then
        RndTimer.Enabled = False
    ElseIf SSTab1.Tab = 3 Then
        RndTimer.Enabled = True
        RndTimer_Timer
    End If
End Sub


Private Sub RndTimer_Timer()
    If Val(Mid(Label18.Caption, 16, Len(Label18.Caption) - 23)) <> RndCache Then
        Label18.Caption = "Current Cache: " & RndCache & " Numbers"
    End If
    If RndState = rReady Then
        If Label19.Caption <> "Current State: Ready" And T = 0 Then
            T = Timer
            Label19.Caption = "Current State: Download Successful"
        End If
    ElseIf RndState = rQuerying And Label19.Caption <> "Current State: Querying..." Then
        Label19.Caption = "Current State: Querying..."
    ElseIf RndState = rEmpty And Label19.Caption <> "Current State: Empty" Then
        Label19.Caption = "Current State: Empty"
    End If
    If T <> 0 Then
        If Timer - T > 1 Then
            Label19.Caption = "Current State: Ready"
            T = 0
        End If
    End If
    
    '
End Sub

Private Sub Text1_Change()

End Sub


Private Sub Timer1_Timer()
    If ServerWindow.SNRegged Or Not ConnectedToReg Then
        If cmdReg.Enabled Then cmdReg.Enabled = False
    Else
        If Not cmdReg.Enabled Then cmdReg.Enabled = True
    End If
    If ServerWindow.SNRegged Then
        If lblRegged.Caption <> "Registered" Then lblRegged.Caption = "Registered"
    Else
        If lblRegged.Caption <> "Unregistered" Then lblRegged.Caption = "Unregistered"
    End If
    
End Sub

Private Sub txtExpire_LinkClose()
    If Not IsNumeric(txtExpire.Text) Then
        txtExpire.Text = "90"
    ElseIf Val(txtExpire.Text) < 30 Then
        txtExpire.Text = "30"
    ElseIf Val(txtExpire.Text) > 360 Then
        txtExpire.Text = "360"
    End If
End Sub

Private Sub txtMaxIPs_LostFocus()
    If Not IsNumeric(txtMaxIPs.Text) Or Val(txtMaxIPs.Text) < 0 Or Val(txtMaxIPs.Text) > 16 Then txtMaxIPs.Text = "0"
End Sub

Private Sub txtRnd_LostFocus(Index As Integer)
    Dim X As Long
    Dim Y As Long
    On Error GoTo ETrap
    X = CLng(txtRnd(0).Text)
    Y = CLng(txtRnd(1).Text)
    If X > 16384 Then X = 16384
    If X < 256 Then X = 256
    If Y > Int(X / 8) Then Y = Int(X / 8)
    If Y < 8 Then Y = 8
    txtRnd(0).Text = CStr(X)
    txtRnd(1).Text = CStr(Y)
    Exit Sub
ETrap:
    txtRnd(0) = "1024"
    txtRnd(1) = "128"
End Sub
