VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{2C1EC115-F1BA-11D3-BF43-00A0CC32BE58}#9.1#0"; "DMC2.ocx"
Begin VB.MDIForm MainContainer 
   BackColor       =   &H8000000C&
   Caption         =   "Pokémon NetBattle"
   ClientHeight    =   7515
   ClientLeft      =   2595
   ClientTop       =   2040
   ClientWidth     =   10080
   Icon            =   "MainContainer.frx":0000
   LinkTopic       =   "MDIForm1"
   Begin VB.PictureBox picCont 
      Align           =   2  'Align Bottom
      Height          =   15
      Left            =   0
      ScaleHeight     =   1
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   672
      TabIndex        =   1
      Top             =   7500
      Width           =   10080
      Begin VB.PictureBox SwapSpace 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1935
         Left            =   180
         ScaleHeight     =   129
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   140
         TabIndex        =   2
         Top             =   120
         Visible         =   0   'False
         Width           =   2100
      End
   End
   Begin NetBattle.CompressZIt Compressor 
      Left            =   600
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.Timer LoaderKiller 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   600
      Top             =   1320
   End
   Begin VB.Timer ServerKiller 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   120
      Top             =   1320
   End
   Begin MSComctlLib.Toolbar IMWindowList 
      Align           =   2  'Align Bottom
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   7080
      Visible         =   0   'False
      Width           =   10080
      _ExtentX        =   17780
      _ExtentY        =   741
      ButtonWidth     =   1138
      ButtonHeight    =   582
      ToolTips        =   0   'False
      AllowCustomize  =   0   'False
      Appearance      =   1
      TextAlignment   =   1
      ImageList       =   "MiniTrainers"
      _Version        =   393216
   End
   Begin DMC2.DMC ModSFX 
      Left            =   1560
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin DMC2.DMC ModPlay 
      Left            =   1080
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.Timer MessageFlasher 
      Interval        =   1000
      Left            =   2040
      Top             =   840
   End
   Begin MSComDlg.CommonDialog FileBox 
      Left            =   120
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList MiniTrainers 
      Left            =   1920
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   22
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainContainer.frx":1272
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainContainer.frx":180C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainContainer.frx":1DA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainContainer.frx":2340
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainContainer.frx":28DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainContainer.frx":2E74
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainContainer.frx":340E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainContainer.frx":39A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainContainer.frx":3F42
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainContainer.frx":44DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainContainer.frx":4A76
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainContainer.frx":5010
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainContainer.frx":55AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainContainer.frx":5B44
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainContainer.frx":60DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainContainer.frx":6678
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainContainer.frx":6C12
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainContainer.frx":71AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainContainer.frx":7746
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainContainer.frx":7CE0
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainContainer.frx":8B32
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainContainer.frx":90CC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList Conditions 
      Left            =   720
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainContainer.frx":9666
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainContainer.frx":9C00
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainContainer.frx":A19A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainContainer.frx":A734
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainContainer.frx":ACCE
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainContainer.frx":B268
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainContainer.frx":B802
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainContainer.frx":BD9C
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainContainer.frx":C336
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainContainer.frx":C8D0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList Trainers 
      Left            =   1320
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainContainer.frx":CE6A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainContainer.frx":D744
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainContainer.frx":E01E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainContainer.frx":E8F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainContainer.frx":F1D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainContainer.frx":FAAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainContainer.frx":10386
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainContainer.frx":10C60
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainContainer.frx":1153A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainContainer.frx":11E14
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainContainer.frx":126EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainContainer.frx":12FC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainContainer.frx":138A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainContainer.frx":1417C
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainContainer.frx":14A56
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList Types 
      Left            =   120
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainContainer.frx":15330
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainContainer.frx":158CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainContainer.frx":15E64
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainContainer.frx":163FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainContainer.frx":16998
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainContainer.frx":16F32
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainContainer.frx":174CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainContainer.frx":17A66
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainContainer.frx":18000
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainContainer.frx":1859A
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainContainer.frx":18B34
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainContainer.frx":190CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainContainer.frx":19668
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainContainer.frx":19C02
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainContainer.frx":1A19C
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainContainer.frx":1A736
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainContainer.frx":1ACD0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFakeFile 
      Caption         =   "&File"
      Enabled         =   0   'False
   End
   Begin VB.Menu mnuFakeHelp 
      Caption         =   "&Help"
      Enabled         =   0   'False
   End
End
Attribute VB_Name = "MainContainer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cWeb As SHDocVw.InternetExplorer
Private Declare Function GetActiveWindow Lib "user32" () As Long
Private Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, _
        ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, _
        ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
        
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_SHOWWINDOW = &H40
Private Const SWP_NOREDRAW = &H8

Private Topped As Boolean
Private IMShift As Boolean

Public Sub DoPicture(ByVal FileName As String, Optional ByVal KillFile As Boolean = True)
    Dim X As Integer
    Dim GFXBytes() As Byte
    Dim FileNum As Integer
    Dim FToUse As String
    
    If FileName = "" Then
        If InVBMode Then Stop
        Exit Sub
    End If
    FileNum = FreeFile
ChooseImage:
    '>>> Call WriteDebugLog("Loading Image: " & FileName)
    For X = 0 To UBound(GFile.FileName)
        If GFile.FileName(X) & ".gif" = FileName Then Exit For
    Next
    If X > UBound(GFile.FileName) Then
        FileName = "000rs.gif"
        GoTo ChooseImage
    End If
    ReDim GFXBytes(GFile.ByteCount(X) - 1) As Byte
    Open GFXTempFile(GFile.InFile(X)) For Binary Access Read As #FileNum
    Get #FileNum, GFile.ByteStart(X), GFXBytes()
    Close #FileNum
    'Compressor.DecompressData GFXBytes(), GFile(X).ByteCount
    If KillFile Then
        SwapSpace.Picture = PictureFromArray(GFXBytes())
    Else
        FToUse = GFile.FileName(X) & ".gif"
        If FileExists(SlashPath & FToUse) Then Kill SlashPath & FToUse
        Open SlashPath & FToUse For Binary Access Write As #FileNum
        Put #FileNum, , GFXBytes()
        Close #FileNum
        SwapSpace.Picture = LoadPicture(SlashPath & FToUse)
    End If
End Sub

Sub HandleError(strLoc As String, strError$, lError As Long, varModule As Variant)

    Dim nCursorType As Integer

    nCursorType = Screen.MousePointer

    Screen.MousePointer = vbNormal
    MsgBox strLoc & ": " & strError & " (" & lError & ")", vbExclamation, varModule
    Screen.MousePointer = nCursorType

End Sub


Private Sub IMWindowList_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim X As Integer
    
    For X = 1 To MaxUsers
        If Player(X).Name = Button.Caption Then Exit For
    Next
    If IMShift Then
        Call KillIMWindow(X)
    ElseIf Not IMWindowArray(IMWindowID(X)).Visible Then
        IMWindowArray(IMWindowID(X)).WindowState = vbNormal
        IMWindowArray(IMWindowID(X)).Visible = True
        'IMWindowFlash(IMWindowID(X)) = False
        IMWindowArray(IMWindowID(X)).SetFocus
        IMWindowList.Buttons(Button.Index).Value = tbrPressed
    ElseIf IMWindowFlash(IMWindowID(X)) Then
        IMWindowArray(IMWindowID(X)).SetFocus
    Else
        IMWindowArray(IMWindowID(X)).Visible = False
        IMWindowList.Buttons(Button.Index).Value = tbrUnpressed
    End If
    Call MessageFlasher_Timer
End Sub

Sub PopButton(ByVal PNum As Integer)
    Dim X As Integer
    
    For X = 1 To IMWindowList.Buttons.count
        If IMWindowList.Buttons(X).Caption = Player(PNum).Name Then Exit For
    Next
    IMWindowList.Buttons(X).Value = tbrUnpressed
End Sub

Private Sub IMWindowList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    IMShift = (Shift Mod 2 = 1)
End Sub

Private Sub LoaderKiller_Timer()
    On Error Resume Next
    Unload Loader
    Unload Me
End Sub

Private Sub MDIForm_Load()
    Dim DidFileReg As Boolean
    Dim DidReplReg As Boolean
    Dim X As Integer
    Dim Temp As String
    Temp = "Nandeyanen!"
    Compressor.CompressString Temp
    Compressor.DecompressString Temp, 11
    If Temp <> "Nandeyanen!" Then
        If Left(Temp, 11) = "Nandeyanen!" And Len(Temp) > 11 Then
            MsgBox "NetBattle is not compatible with your system language.  In order to play, you'll need to either set your machine to use English mode (using the Windows Control Panel), or download AppLocale from Microsoft's web site and use that to launch NetBattle.", vbCritical, "Error"
        Else
            MsgBox "A required compression file is missing or corrupted.  Please reinstall the program." & vbCrLf & "This error may also mean NetBattle is not compatible with your system language.  In order to play, you'll need to either set your machine to use English mode (using the Windows Control Panel), or download AppLocale from Microsoft's web site and use that to launch NetBattle.", vbCritical, "Error"
        End If
        End
    End If
        
    Call SetEnglish
    Call DoInitialResize
    'If Me.Width > Screen.Width Then Me.Width = Screen.Width
    'If Me.Height > Screen.Height Then Me.Height = Screen.Height
    If Me.WindowState <> vbMaximized And Me.WindowState <> vbMinimized Then
        Me.Left = (Screen.Width - Me.Width) / 2
        Me.Top = (Screen.Height - Me.Height) / 2 - 500
    End If
    If InVBMode Then
        X = MsgBox("Click Yes for Server, No for Client", vbYesNo, "Debug Mode")
    End If
    'Check the command line, decide whether to run a server or a normal instance
    If Command$ = "SERVER" Or X = vbYes Then
        Call LoadGFXKeys
        If HasColGFX Then Call LoadGFXKeys("gfxcol.bin")
        ServerWindow.Show
    Else
        frmSplash.Show
        DoEvents
        Call LoadGFXKeys
        If HasColGFX Then Call LoadGFXKeys("gfxcol.bin")
        WriteDebugLog "Showing Loader"
        Loader.Show
        Call CenterWindow(Loader)
        If CmdReplay Then
            Loader.Hide
            'Playback.SetFocus
        End If
        DoneLoading = True
    End If

    DidFileReg = GetSetting("NetBattle", "Notification", "DidFileReg", False)
    DidReplReg = GetSetting("NetBattle", "Notification", "DidReplReg", False)
    'Comment this out before compiling!
    'DidFileReg = False
    If Not DidFileReg Then
        DidFileReg = CreateFileAss(".pnb", "NetBattle.Team", "NetBattle Team", "Open", SlashPath & "PokeBattle.exe", , True, SlashPath & "PokeBattle.exe,1", True)
        SaveSetting "NetBattle", "Notification", "DidFileReg", DidFileReg
    End If
    If Not DidReplReg Then
        DidReplReg = CreateFileAss(".btl", "NetBattle.Replay", "NetBattle Replay", "Open", SlashPath & "PokeBattle.exe", , True, SlashPath & "PokeBattle.exe,7", True)
        SaveSetting "NetBattle", "Notification", "DidReplReg", DidReplReg
    End If
End Sub

Private Sub MDIForm_Resize()
    If MainContainer.WindowState = vbMinimized And RunningServer Then MainContainer.Visible = False
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    Dim X As Byte
    
    On Error Resume Next
    Unload TaskBarIcon
    Call StopMusic
    Call StopSound
    CloseAll
    ModSFX.TerminateBASS
    ModPlay.TerminateBASS
    Unload Search 'Don't ask why...
    Unload MasterDex 'Ditto for these - End seems to leave certain windows up under some circumstances.  No idea.
    Unload Options
    Unload UserEdit
    Unload ScriptForm
    Unload SetUsers
    Call CloseDebugLog
    Call SetAttr(FTextFile, vbNormal)
    Kill FTextFile
    For X = 0 To UBound(GFXTempFile)
        Call SetAttr(GFXTempFile(X), vbNormal)
        Kill GFXTempFile(X)
    Next
    If Not RunningServer Then SaveSetting "NetBattle", "Options", "LastTBMode", TBMode
    If Me.WindowState <> vbMinimized Then
        If Me.WindowState = vbMaximized Then
            SaveSetting "NetBattle", "Main Window", "Maximized", True
        Else
            SaveSetting "NetBattle", "Main Window", "Maximized", False
        End If
        SaveSetting "NetBattle", "Main Window", "Width", Me.Width
        SaveSetting "NetBattle", "Main Window", "Height", Me.Height
    End If
    End
End Sub

Private Sub mnuOptions_Click()
    Options.Show 1
End Sub

Private Sub mnuHelpItem_Click(Index As Integer)
    Select Case Index
        Case 0
            ShellExecute 0, vbNullString, "http://www.netbattle.net", vbNullString, vbNullString, 0
        Case 2
            frmAbout.Show 1
    End Select
End Sub

Public Sub AddTBItem(ByVal TName As String, ByVal TIcon As Byte)
    Dim X As Integer
    
    IMWindowList.Buttons.Add , "U:" & TName, TName, , TIcon
    For X = 1 To IMWindowList.Buttons.count
        If IMWindowList.Buttons(X).Caption = TName Then IMWindowList.Buttons(X).Value = tbrPressed
    Next
    IMWindowList.Visible = True
End Sub

Public Sub DelTBItem(ByVal TName As String)
    IMWindowList.Buttons.Remove "U:" & TName
    If IMWindowList.Buttons.count = 0 Then IMWindowList.Visible = False
End Sub

Private Sub MessageFlasher_Timer()
    Static B As Boolean
    Dim X As Integer
    Dim Y As Integer
    Dim FoundPlayer As Boolean
    If IMWindowList.Buttons.count = 0 Then Exit Sub
    B = Not B
    MessageFlasher.Enabled = False
    For X = 1 To IMWindowList.Buttons.count
        FoundPlayer = False
        For Y = 1 To MaxUsers
            If Player(Y).Name = IMWindowList.Buttons(X).Caption Then
                FoundPlayer = True
                Exit For
            End If
        Next
        If FoundPlayer Then
            If Me.ActiveForm.hWnd = IMWindowArray(IMWindowID(Y)).hWnd Then IMWindowFlash(IMWindowID(Y)) = False
            If Not IMWindowFlash(IMWindowID(Y)) Then
                IMWindowList.Buttons(X).Value = Abs(IMWindowArray(IMWindowID(Y)).Visible)
            Else
                IMWindowList.Buttons(X).Value = IIf(B, 1, 0)
            End If
        Else
            'Player signed off - button didn't go away
            Call DelTBItem(IMWindowList.Buttons(X).Caption)
        End If
    Next
    MessageFlasher.Enabled = True
End Sub

Private Sub ServerKiller_Timer()
    If ServerWindow.ListView1.ListItems.count = 0 Then
        ServerKiller.Enabled = False
        On Error Resume Next
        Unload ServerWindow
        Unload TaskBarIcon
        Unload Me
    End If
End Sub

Private Sub DoInitialResize()
    Dim Maximized As Boolean
    Dim Width As Long
    Dim Height As Long
    
    If WindowState = vbMinimized Then Exit Sub
    Maximized = GetSetting("NetBattle", "Main Window", "Maximized", True)
    If Maximized Then Me.WindowState = vbMaximized: Exit Sub
    Width = GetSetting("NetBattle", "Main Window", "Width", 10200)
    Height = GetSetting("NetBattle", "Main Window", "Height", 8325)
    'If Width < MinWidth Then Width = MinWidth
    'If Height < MinHeight Then Height = MinHeight
    Me.Width = Width
    Me.Height = Height
    'Call CenterWindow(Me)
End Sub


Private Sub Timer1_Timer()
    Debug.Print Rnd
End Sub
