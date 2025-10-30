VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C8530F8A-C19C-11D2-99D6-9419F37DBB29}#1.1#0"; "ccrpprg6.ocx"
Begin VB.Form TeamLoader 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1695
   ClientLeft      =   5445
   ClientTop       =   2520
   ClientWidth     =   4575
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   Begin NetBattle.CompressZIt CompressZIt1 
      Left            =   3000
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin CCRProgressBar6.ccrpProgressBar Progress 
      Height          =   375
      Left            =   120
      Top             =   600
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Max             =   6
   End
   Begin VB.Timer AnimTimer 
      Interval        =   250
      Left            =   4080
      Top             =   1200
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3480
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TeamLoader.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TeamLoader.frx":08DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TeamLoader.frx":11B4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   2040
      Top             =   1080
      Width           =   480
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Loading Team..."
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4095
   End
End
Attribute VB_Name = "TeamLoader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim AnimFrame As Byte
Dim TeamChangeOK As Boolean
Public Canceled As Boolean
Private Sub AnimTimer_Timer()
    AnimFrame = AnimFrame + 1
    If AnimFrame = 5 Then AnimFrame = 1
    Select Case AnimFrame
        Case 1
            Image1.Picture = ImageList1.ListImages(1).Picture
        Case 2, 4
            Image1.Picture = ImageList1.ListImages(2).Picture
        Case 3
            Image1.Picture = ImageList1.ListImages(3).Picture
    End Select
End Sub

Private Sub Form_Load()
    Me.Visible = False 'True
    Progress.Value = 0
    AnimFrame = 1
    Image1.Picture = ImageList1.ListImages(1).Picture
End Sub

Public Sub ReadFile(ByVal FileToUse As String, Optional ByVal LoadName = True)
    Dim X As Integer
    Dim Y As Integer
    Dim Z As Byte
    Dim DummyData As Variant
    Dim Hax0red As Boolean
    Dim PKMNHasMoves As Boolean
    Dim DupePKMN As Boolean
    Dim Version As Byte
    Dim Temp As String
    Dim Temp2 As String
    Dim TempPlayer As Trainer
    Dim SwapPKMN(6) As Pokemon
    Dim MoveTotal As Integer
    Dim BlankPKMN As Pokemon
    Dim ByteArray() As Byte
    Dim FileSize As Long
    Dim Worked As Boolean
    Dim FileNum As Integer
    Dim TMove() As Integer
    MainContainer.MousePointer = vbHourglass
    DoEvents
    TeamChangeOK = False
    FileNum = FreeFile
    Open FileToUse For Binary Access Read As #FileNum
    Hax0red = False
    Temp2 = "      "
    Get #FileNum, 2, Temp2
    Close #FileNum
    Select Case Temp2
    Case "PNB4.0", "PNB4.1"
        Version = 4
        Open FileToUse For Binary Access Read As #FileNum
        Temp = String$(LOF(FileNum), vbNullChar)
        Get #FileNum, , Temp
        Close #FileNum
        DummyData = ChopString(Temp, 7)
        X = Asc(ChopString(Temp, 1))
        TempPlayer.Name = CorrectText(ChopString(Temp, X), True)
        X = Asc(ChopString(Temp, 1))
        TempPlayer.Extra = Replace(ChopString(Temp, X), Chr(1), " ")
        X = Asc(ChopString(Temp, 1))
        TempPlayer.WinMess = CorrectText(ChopString(Temp, X))
        X = Asc(ChopString(Temp, 1))
        TempPlayer.LoseMess = CorrectText(ChopString(Temp, X))
        If Len(TempPlayer.Name) > 20 Then TempPlayer.Name = Left(TempPlayer.Name, 20)
        Label1.Caption = "Loading Team: Player " & TempPlayer.Name
        'DoEvents
        TBMode = Asc(ChopString(Temp, 1))
        TempPlayer.Picture = Asc(ChopString(Temp, 1))
        TempPlayer.Version = Asc(ChopString(Temp, 1))
        For X = 1 To 6
            Label1.Caption = "Loading Team: " & Trim(Left$(Temp, 10))
            Progress.Value = X
            'DoEvents
            If Temp2 = "PNB4.0" Then
                Temp = Left$(Temp, 10) & "     " & Right$(Temp, Len(Temp) - 10)
            End If
            SwapPKMN(X) = Str2PKMN(ChopString(Temp, POKELEN), True)
            SwapPKMN(X).Image = ChooseImage(SwapPKMN(X), TempPlayer.Version)
        Next X
        If SwapPKMN(1).GameVersion = nbModAdv Then
            DBModName = Trim$(ChopString(Temp, 20))
            DBModStr = Temp
            ApplyDBMod
        End If
        For X = 1 To 6
            FillInPokeData SwapPKMN(X), SwapPKMN(X).GameVersion
        Next X
        
    Case Else
        Open FileToUse For Input As #FileNum
        Input #FileNum, DummyData
        TBMode = 1
        Select Case Temp2
        Case "PNB3.0"
            Version = 3
            Input #FileNum, FileSize
            Close #FileNum
            ReDim ByteArray(FileLen(FileToUse) - (Len(Str(FileSize)) + 12)) As Byte
            Open FileToUse For Binary Access Read As #FileNum
            Get #FileNum, Len(Str(FileSize)) + 12, ByteArray()
            Close #FileNum
            Worked = CompressZIt1.DecompressData(ByteArray(), FileSize)
            Open SlashPath & "team.tmp" For Binary Access Write As #FileNum
            Put #FileNum, , ByteArray()
            Close #FileNum
            Open SlashPath & "team.tmp" For Input As #FileNum
            Input #FileNum, TempPlayer.Name
        Case "PNB2.0"
            Version = 2
            Input #FileNum, TempPlayer.Name
        Case Else
            Version = 1
            TempPlayer.Name = Temp2
        End Select
        Input #FileNum, TempPlayer.Picture
        Input #FileNum, TempPlayer.Version
        TempPlayer.Version = nbGFXRS
        Input #FileNum, TempPlayer.Extra
        Input #FileNum, TempPlayer.WinMess
        Input #FileNum, TempPlayer.LoseMess
        If Len(TempPlayer.Name) > 20 Then TempPlayer.Name = Left(TempPlayer.Name, 20)
        TempPlayer.Name = CorrectText(TempPlayer.Name, True)
        Label1.Caption = "Loading Team: Player " & TempPlayer.Name
        Progress.Value = 1
        'DoEvents
        For X = 1 To 6
            SwapPKMN(X) = BlankPKMN
            With SwapPKMN(X)
                Input #FileNum, .No
                SwapPKMN(X) = BasePKMN(.No)
                Input #FileNum, DummyData
                Input #FileNum, .Item
                Input #FileNum, .Nickname
                If Version >= 2 Then
                    Input #FileNum, .Level
                Else
                    .Level = 100
                End If
                .Nickname = CorrectText(.Nickname)
                Label1.Caption = "Loading Team: " & .Nickname
                Progress.Value = X
                'DoEvents
                MoveTotal = 0
                For Y = 1 To 4
                    Input #FileNum, .Move(Y)
                Next
                Input #FileNum, DummyData
                Input #FileNum, DummyData
                Input #FileNum, DummyData
                Input #FileNum, .DV_Atk
                Input #FileNum, DummyData
                Input #FileNum, .DV_Def
                Input #FileNum, DummyData
                Input #FileNum, .DV_Spd
                Input #FileNum, DummyData
                Input #FileNum, DummyData
                Input #FileNum, .DV_SAtk
                Input #FileNum, .Gender
                .Image = ChooseImage(SwapPKMN(X), TempPlayer.Version)
            End With
            Call FillInPokeData(SwapPKMN(X), TBMode)
        Next
        Close #FileNum
    End Select
    
    
    'Legallity Checks
    TempPlayer.Extra = Left$(CorrectText(TempPlayer.Extra), 200)
    TempPlayer.WinMess = Left$(CorrectText(TempPlayer.WinMess), 200)
    TempPlayer.LoseMess = Left$(CorrectText(TempPlayer.LoseMess), 200)
    For X = 1 To 6
        Z = 1
        With SwapPKMN(X)
            If .No > 0 Then
                TMove = .Move
                For Y = 1 To 4
                    .Move(Y) = 0
                Next Y
                For Y = 1 To 4
                    .Move(Z) = TMove(Y)
                    If LegalMove(SwapPKMN(X)) <> "" Then
                        MsgBox "Illegal move detected on " & .Nickname & "." & vbNewLine & Moves(.Move(Y)).Name & " not allowed in this moveset.  Check this Pokemon in the Team Builder.", vbExclamation, "Illegal Move"
                        If TeamChangeFromMS Then
                            MsgBox "The team was not loaded.  Please fix the illegal moves in the Team Builder.", vbCritical, "Warning"
                            Close #FileNum
                            If Version = 3 Then Kill SlashPath & "team.tmp"
                            Unload Me
                            Exit Sub
                        End If
                        .Move(Z) = 0
                    Else
                        Z = Z + 1
                    End If
                Next Y
            End If
        End With
    Next X
    If Version = 3 Then Kill SlashPath & "team.tmp"
    DupePKMN = False
    For X = 1 To 6
        Select Case SwapPKMN(X).No
            Case 0
                'nothing
            Case 386, 387, 388, 389
                For Y = 1 To 6
                    If X <> Y And SwapPKMN(Y).No <> SwapPKMN(X).No And (SwapPKMN(Y).No = 386 Or SwapPKMN(Y).No = 387 Or SwapPKMN(Y).No = 388 Or SwapPKMN(Y).No = 389) Then
                        DupePKMN = True
                        MsgBox SwapPKMN(Y).Nickname & " is a duplicate, you have a " & BasePKMN(386).Name & " (" & SwapPKMN(X).Nickname & ") already." & SwapPKMN(Y).Nickname & " has been removed from this team.", vbCritical, "Error"
                        SwapPKMN(Y) = BlankPKMN
                    End If
                Next
            Case Else
'                For Y = 1 To 6
'                    If X <> Y And SwapPKMN(X).No = SwapPKMN(Y).No Then
'                        DupePKMN = True
'                        MsgBox SwapPKMN(Y).Nickname & " is a duplicate, you have a " & SwapPKMN(Y).Name & " (" & SwapPKMN(X).Nickname & ") already." & SwapPKMN(Y).Nickname & " has been removed from this team.", vbCritical, "Error"
'                        SwapPKMN(Y) = BlankPKMN
'                    End If
'                Next
        End Select
    Next
    If LoadName Then
       You = TempPlayer
       If BetaRel <> "" Then
           You.ProgVersion = App.Major & "." & App.Minor & "." & BetaRel
       Else
           You.ProgVersion = App.Major & "." & App.Minor & "." & App.Revision
       End If
    End If
    For X = 1 To 6
        PKMN(X) = SwapPKMN(X)
        StoredPKMN(X) = PKMN(X)
    Next
    Ranking = TeamRank
    Call ReadBinArray(CompatCheck(PKMN), Compatibility)
    Call UpdateListings(FileToUse)
    StoredFileName = FileToUse
    TeamChangeOK = True
    MainContainer.MousePointer = vbDefault
    Unload Me
End Sub

Public Function OpenTheFile() As Boolean
    Dim FileToUse As String
    Dim Temp As String
    With MainContainer.FileBox
        .DialogTitle = "Load Saved Trainer/Team"
        .Flags = cdlOFNHideReadOnly
        .CancelError = True
        .Filter = "Pokémon NetBattle Team (*.pnb)|*.pnb|All Files (*.*)|*.*"
        .DefaultExt = ".pnb"
        .FileName = ""
        Temp = GetSetting("NetBattle", "Options", "InitDir", "")
        If Temp <> "" Then .InitDir = Temp
        On Error GoTo Cancelled
        .ShowOpen
        FileToUse = .FileName
        SaveSetting "NetBattle", "Options", "InitDir", Left$(FileToUse, InStrRev(FileToUse, "\"))
        Call ReadFile(FileToUse)
        If TeamChangeFromMS And TeamChangeOK Then MasterServer.TeamChanged = True
        If TeamChangeFromTB Then TeamBuilder.TeamLoadOK = True
    End With
    OpenTheFile = True
    Exit Function
Cancelled:
    OpenTheFile = False
    Unload Me
End Function
