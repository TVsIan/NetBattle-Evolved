VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form DBManMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Database Manager"
   ClientHeight    =   2340
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   3135
   Icon            =   "DBManMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   3135
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "&Compile Database"
      Height          =   495
      Left            =   60
      TabIndex        =   3
      Top             =   1800
      Width           =   3015
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Edit &Battle Chart"
      Height          =   495
      Left            =   60
      TabIndex        =   2
      Top             =   1140
      Width           =   3015
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Edit &Moves"
      Height          =   495
      Left            =   60
      TabIndex        =   1
      Top             =   600
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Edit &Pokémon"
      Default         =   -1  'True
      Height          =   495
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   3015
   End
   Begin DBMan.CompressZIt CompressZIt1 
      Left            =   1920
      Top             =   2160
      _extentx        =   847
      _extenty        =   847
   End
   Begin MSComctlLib.ImageList Types 
      Left            =   2520
      Top             =   2160
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
            Picture         =   "DBManMain.frx":212A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DBManMain.frx":21C3
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DBManMain.frx":2414
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DBManMain.frx":2576
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DBManMain.frx":26CA
            Key             =   ""
            Object.Tag             =   "Grass"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DBManMain.frx":2C64
            Key             =   ""
            Object.Tag             =   "Ice"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DBManMain.frx":31FE
            Key             =   ""
            Object.Tag             =   "Fighting"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DBManMain.frx":3798
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DBManMain.frx":39E3
            Key             =   ""
            Object.Tag             =   "Ground"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DBManMain.frx":3F7D
            Key             =   ""
            Object.Tag             =   "Flying"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DBManMain.frx":4517
            Key             =   ""
            Object.Tag             =   "Psychic"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DBManMain.frx":4AB1
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DBManMain.frx":4C03
            Key             =   ""
            Object.Tag             =   "Rock"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DBManMain.frx":519D
            Key             =   ""
            Object.Tag             =   "Ghost"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DBManMain.frx":5737
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DBManMain.frx":598B
            Key             =   ""
            Object.Tag             =   "Dark"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DBManMain.frx":5F25
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileItem 
         Caption         =   "&CompileDatabase"
         Index           =   0
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "E&xit"
         Index           =   1
      End
   End
End
Attribute VB_Name = "DBManMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    EditHelp.Show
    Command1.Enabled = False
End Sub

Private Sub Command2_Click()
    MoveEdit.Show
    Command2.Enabled = False
End Sub

Private Sub Command3_Click()
    BatEdit.Show
    Command3.Enabled = False
End Sub

Sub ExportDB()
    'For Pokémon
    Dim No As Integer
    Dim GSNo As Integer
    Dim AdvNo As Integer
    Dim Name As String
    Dim Legendary As Boolean
    Dim Uber As Boolean
    Dim Type1 As Integer
    Dim Type2 As Integer
    Dim PAtt1 As Integer
    Dim PAtt2 As Integer
    Dim Color1 As Integer
    Dim Color2 As Integer
    Dim BaseHP As Integer
    Dim BaseAttack As Integer
    Dim BaseDefense As Integer
    Dim BaseSpeed As Integer
    Dim BaseSAttack As Integer
    Dim BaseSDefense As Integer
    Dim SpecialRBY As Integer
    Dim StartsWith As Integer
    Dim RawMoves As String
    Dim RawMachine As String
    Dim RawBreeding As String
    Dim RawRBY As String
    Dim RawRBYTM As String
    Dim RawSpecial As String
    Dim RawTutor As String
    Dim RawAdv As String
    Dim RawAdvTM As String
    Dim RawAdvBreed As String
    Dim RawAdvSpecial As String
    Dim RawAdvTutor As String
    Dim RawLFOnly As String
    Dim ExistRBY As Boolean
    Dim ExistGSC As Boolean
    Dim ExistAdv As Boolean
    Dim PercentFemale As Integer
    Dim RedBlue As String
    Dim Yellow As String
    Dim Gold As String
    Dim Silver As String
    Dim Crystal As String
    Dim Ruby As String
    Dim Sapphire As String
    Dim MyStage As Integer
    Dim MyMethod As Integer
    Dim Evo(1 To 5) As Integer
    Dim EvoM(1 To 5) As Integer
    Dim Offset As Byte
    Dim Weight As Integer
    Dim Height As Integer
    Dim LevelBal As Byte
    
    'For moves
    Dim ID As Integer
    Dim MoveName As String
    Dim MoveType As Integer
    Dim Power As Integer
    Dim Accuracy As Integer
    Dim PP As Integer
    Dim Percent As Integer
    Dim Special As Integer
    Dim Description As String
    Dim WorksProperly As Boolean
    Dim BrightPowder As Boolean
    Dim KingsRock As Boolean
    Dim RBYCompatible As Boolean
    Dim GSCCompatible As Boolean
    Dim ADVCompatible As Boolean
    Dim HitsBoth As Boolean
    Dim HitsAll As Boolean
    Dim SoundMove As Boolean
    Dim PhysMove As Boolean
    Dim AffectsSelf As Boolean
    Dim RBYTM As String
    Dim GSTM As String
    Dim ADVTM As String
    Dim BlockSubstitute As Boolean
    Dim Target As Integer
    Dim MagicCoat As Boolean
    
    'Battle Chart
    Dim RowID As Integer
    Dim RowData(17) As Single
    'Not for any specific data
    Dim QueryResults As ADODB.Recordset
    Dim X As Integer
    
    On Error Resume Next
    Set QueryResults = New ADODB.Recordset
    QueryResults.Open "SELECT * FROM Pokemon WHERE Number > 0 ORDER BY Number ASC", PokeData, adOpenStatic, adLockReadOnly, adCmdText
    QueryResults.MoveLast
    QueryResults.MoveFirst
    Open SlashPath & "PokeDB.csv" For Output As #1
    While Not QueryResults.EOF
        No = 0
        GSNo = 0
        AdvNo = 0
        Name = ""
        Legendary = False
        Uber = False
        Type1 = 0
        Type2 = 0
        PAtt1 = 0
        PAtt2 = 0
        Color1 = 0
        Color2 = 0
        BaseHP = 0
        BaseAttack = 0
        BaseDefense = 0
        BaseSpeed = 0
        BaseSAttack = 0
        BaseSDefense = 0
        SpecialRBY = 0
        StartsWith = 0
        RawMoves = ""
        RawMachine = ""
        RawBreeding = ""
        RawRBY = ""
        RawRBYTM = ""
        RawSpecial = ""
        RawTutor = ""
        RawAdv = ""
        RawAdvTM = ""
        RawAdvBreed = ""
        RawAdvSpecial = ""
        RawAdvTutor = ""
        RawLFOnly = ""
        ExistRBY = False
        ExistGSC = False
        ExistAdv = False
        PercentFemale = 0
        RedBlue = ""
        Yellow = ""
        Gold = ""
        Silver = ""
        Crystal = ""
        Ruby = ""
        Sapphire = ""
        MyStage = 0
        MyMethod = 0
        For X = 1 To 5
            Evo(X) = 0
            EvoM(X) = 0
        Next
        Offset = 0
        LevelBal = 0
        Weight = 0
        Height = 0
        No = QueryResults("Number")
        GSNo = QueryResults("GSCNumber")
        AdvNo = QueryResults("AdvNumber")
        Name = QueryResults("Name")
        Legendary = QueryResults("Legendary")
        Uber = QueryResults("Uber")
        Type1 = QueryResults("Type1")
        Type2 = QueryResults("Type2")
        PAtt1 = QueryResults("Attribute1")
        PAtt2 = QueryResults("Attribute2")
        Color1 = QueryResults("Color1")
        Color2 = QueryResults("Color2")
        BaseHP = QueryResults("HP")
        BaseAttack = QueryResults("Attack")
        BaseDefense = QueryResults("Defense")
        BaseSpeed = QueryResults("Speed")
        BaseSAttack = QueryResults("SpecialAttack")
        BaseSDefense = QueryResults("SpecialDefense")
        SpecialRBY = QueryResults("SpecialRBY")
        StartsWith = QueryResults("BornWith")
        RawMoves = QueryResults("Moves")
        RawMachine = QueryResults("Machine Moves")
        RawBreeding = QueryResults("Breeding Moves")
        RawRBY = QueryResults("R/B/Y Moves")
        RawRBYTM = QueryResults("R/B/Y Machines")
        RawSpecial = QueryResults("Special Moves")
        RawTutor = QueryResults("Move Tutor")
        RawAdv = QueryResults("Advance Moves")
        RawAdvTM = QueryResults("Advance TMs")
        RawAdvBreed = QueryResults("Advance Breeding")
        RawAdvSpecial = QueryResults("Advance Special")
        RawAdvTutor = QueryResults("Advance Tutor")
        RawLFOnly = QueryResults("LF Only")
        ExistRBY = QueryResults("RBY")
        ExistGSC = QueryResults("GSC")
        ExistAdv = QueryResults("Adv")
        PercentFemale = QueryResults("Percent Female")
        RedBlue = QueryResults("PokedexRB")
        Yellow = QueryResults("PokedexYellow")
        Gold = QueryResults("PokedexGold")
        Silver = QueryResults("PokedexSilver")
        Crystal = QueryResults("PokedexCrystal")
        Ruby = QueryResults("PokedexRuby")
        Sapphire = QueryResults("PokedexSapphire")
        MyStage = QueryResults("MyStage")
        MyMethod = QueryResults("MyMethod")
        Evo(1) = QueryResults("Evo1")
        EvoM(1) = QueryResults("EvoM1")
        Evo(2) = QueryResults("Evo2")
        EvoM(2) = QueryResults("EvoM2")
        Evo(3) = QueryResults("Evo3")
        EvoM(3) = QueryResults("EvoM3")
        Evo(4) = QueryResults("Evo4")
        EvoM(4) = QueryResults("EvoM4")
        Evo(5) = QueryResults("Evo5")
        EvoM(5) = QueryResults("EvoM5")
        Weight = QueryResults("Weight")
        Height = QueryResults("Height")
        Offset = QueryResults("Offset")
        LevelBal = QueryResults("LevelBal")
        RedBlue = Replace(RedBlue, Chr(34), "''")
        Yellow = Replace(Yellow, Chr(34), "''")
        Gold = Replace(Gold, Chr(34), "''")
        Silver = Replace(Silver, Chr(34), "''")
        Crystal = Replace(Crystal, Chr(34), "''")
        Ruby = Replace(Ruby, Chr(34), "''")
        Sapphire = Replace(Sapphire, Chr(34), "''")
        Write #1, No, GSNo, AdvNo, Name, Legendary, Uber, Type1, Type2, PAtt1, PAtt2, Color1, Color2, BaseHP, BaseAttack, BaseDefense, BaseSpeed, BaseSAttack, BaseSDefense, SpecialRBY, StartsWith, RawMoves, RawMachine, RawBreeding, RawRBY, RawRBYTM, RawSpecial, RawTutor, RawAdv, RawAdvTM, RawAdvBreed, RawAdvSpecial, RawAdvTutor, RawLFOnly, ExistRBY, ExistGSC, ExistAdv, PercentFemale, RedBlue, Yellow, Gold, Silver, Crystal, Ruby, Sapphire, MyStage, MyMethod, Evo(1), EvoM(1), Evo(2), EvoM(2), Evo(3), EvoM(3), Evo(4), EvoM(4), Evo(5), EvoM(5), Weight, Height, Offset, LevelBal
        QueryResults.MoveNext
    Wend
    Close
    Set QueryResults = New ADODB.Recordset
    QueryResults.Open "SELECT * FROM Moves WHERE ID > 0 ORDER BY ID ASC", PokeData, adOpenStatic, adLockReadOnly, adCmdText
    QueryResults.MoveLast
    QueryResults.MoveFirst
    Open SlashPath & "MoveDB.csv" For Output As #1
    While Not QueryResults.EOF
        ID = 0
        MoveName = ""
        MoveType = 0
        Power = 0
        Accuracy = 0
        PP = 0
        Percent = 0
        Special = 0
        Target = 0
        Description = ""
        WorksProperly = False
        BrightPowder = False
        KingsRock = False
        RBYCompatible = False
        GSCCompatible = False
        ADVCompatible = False
        HitsBoth = False
        AffectsSelf = False
        RBYTM = ""
        GSTM = ""
        ADVTM = ""
        BlockSubstitute = False
        ID = QueryResults("ID")
        MoveName = QueryResults("Name")
        MoveType = QueryResults("Type")
        Power = QueryResults("Power")
        Accuracy = QueryResults("Accuracy")
        PP = QueryResults("PP")
        Percent = QueryResults("Percent")
        Special = QueryResults("Special")
        Target = QueryResults("Target")
        Description = QueryResults("Description")
        WorksProperly = QueryResults("Works Properly")
        BrightPowder = QueryResults("BrightPowder")
        KingsRock = QueryResults("KingsRock")
        RBYCompatible = QueryResults("RBYCompatible")
        GSCCompatible = QueryResults("GSCCompatible")
        ADVCompatible = QueryResults("ADVCompatible")
        HitsBoth = QueryResults("HitsBoth")
        AffectsSelf = QueryResults("AffectsSelf")
        RBYTM = QueryResults("RBYTM")
        GSTM = QueryResults("GSTM")
        ADVTM = QueryResults("AdvTM")
        HitsAll = QueryResults("HitsAll")
        SoundMove = QueryResults("SoundMove")
        PhysMove = QueryResults("PhysMove")
        BlockSubstitute = QueryResults("BlockSubstitute")
        Description = Replace(Description, Chr(34), "''")
        MagicCoat = QueryResults("MagicCoat")
        Write #1, ID, MoveName, MoveType, Power, Accuracy, PP, Percent, Special, Target, Description, WorksProperly, BrightPowder, KingsRock, RBYCompatible, GSCCompatible, ADVCompatible, HitsBoth, AffectsSelf, RBYTM, GSTM, ADVTM, BlockSubstitute, HitsAll, SoundMove, PhysMove, MagicCoat
        QueryResults.MoveNext
    Wend
    Close
    Set QueryResults = New ADODB.Recordset
    QueryResults.Open "SELECT * FROM BattleChart WHERE ID > 0 ORDER BY ID ASC", PokeData, adOpenStatic, adLockReadOnly, adCmdText
    QueryResults.MoveLast
    QueryResults.MoveFirst
    Open SlashPath & "TypeDB.csv" For Output As #1
    While Not QueryResults.EOF
        RowID = 0
        For X = 1 To 17
            RowData(X) = 0
        Next
        RowID = QueryResults("ID")
        For X = 1 To 17
            RowData(X) = QueryResults(X)
        Next
        Write #1, RowID, RowData(1), RowData(2), RowData(3), RowData(4), RowData(5), RowData(6), RowData(7), RowData(8), RowData(9), RowData(10), RowData(11), RowData(12), RowData(13), RowData(14), RowData(15), RowData(16), RowData(17)
        QueryResults.MoveNext
    Wend
    Close
    MsgBox "Conversion Complete!", vbInformation, "Done"
End Sub

Private Sub Command4_Click()
    Call mnuFileItem_Click(0)
End Sub

Private Sub mnuFileItem_Click(Index As Integer)
    Select Case Index
        Case 0
            DBCompile.Show
            DBCompile.DoCompile
            'Unload DBCompile
'            Call ExportDB
'        Case 1
'            Call CompressDB
        Case 1
            Unload Me
    End Select
End Sub

Sub CompressDB()
    Dim CBytes() As Byte
    Dim HBytes() As Byte
    
    'PokeDB
    ReDim CBytes(FileLen(SlashPath & "PokeDB.csv") - 1) As Byte
    Open SlashPath & "PokeDB.csv" For Binary Access Read As #1
    Get #1, , CBytes()
    Close
    CompressZIt1.CompressData CBytes()
    If FileExists(SlashPath & "header.tmp") Then Kill SlashPath & "header.tmp"
    Open SlashPath & "header.tmp" For Output As #1
    Write #1, CompressZIt1.OriginalSize
    Close
    ReDim HBytes(FileLen(SlashPath & "header.tmp") - 1) As Byte
    Open SlashPath & "header.tmp" For Binary Access Read As #1
    Get #1, , HBytes()
    Close
    Kill SlashPath & "header.tmp"
    If FileExists(SlashPath & "PokeDB.cdb") Then Kill SlashPath & "PokeDB.cdb"
    Open SlashPath & "PokeDB.cdb" For Binary Access Write As #1
    Put #1, , HBytes
    Put #1, , CBytes
    Close
    MsgBox "PokeDB.csv compressed to PokeDB.cdb." & vbCrLf & "Original size: " & CompressZIt1.OriginalSize & vbCrLf & "Compressed Size: " & FileLen(SlashPath & "PokeDB.cdb"), vbInformation, "Done"
    'MoveDB
    ReDim CBytes(FileLen(SlashPath & "MoveDB.csv") - 1) As Byte
    Open SlashPath & "MoveDB.csv" For Binary Access Read As #1
    Get #1, , CBytes()
    Close
    CompressZIt1.CompressData CBytes()
    If FileExists(SlashPath & "header.tmp") Then Kill SlashPath & "header.tmp"
    Open SlashPath & "header.tmp" For Output As #1
    Write #1, CompressZIt1.OriginalSize
    Close
    ReDim HBytes(FileLen(SlashPath & "header.tmp") - 1) As Byte
    Open SlashPath & "header.tmp" For Binary Access Read As #1
    Get #1, , HBytes()
    Close
    Kill SlashPath & "header.tmp"
    If FileExists(SlashPath & "MoveDB.cdb") Then Kill SlashPath & "MoveDB.cdb"
    Open SlashPath & "MoveDB.cdb" For Binary Access Write As #1
    Put #1, , HBytes
    Put #1, , CBytes
    Close
    MsgBox "MoveDB.csv compressed to MoveDB.cdb." & vbCrLf & "Original size: " & CompressZIt1.OriginalSize & vbCrLf & "Compressed Size: " & FileLen(SlashPath & "MoveDB.cdb"), vbInformation, "Done"
    'TypeDB
    ReDim CBytes(FileLen(SlashPath & "TypeDB.csv") - 1) As Byte
    Open SlashPath & "TypeDB.csv" For Binary Access Read As #1
    Get #1, , CBytes()
    Close
    CompressZIt1.CompressData CBytes()
    If FileExists(SlashPath & "header.tmp") Then Kill SlashPath & "header.tmp"
    Open SlashPath & "header.tmp" For Output As #1
    Write #1, CompressZIt1.OriginalSize
    Close
    ReDim HBytes(FileLen(SlashPath & "header.tmp") - 1) As Byte
    Open SlashPath & "header.tmp" For Binary Access Read As #1
    Get #1, , HBytes()
    Close
    Kill SlashPath & "header.tmp"
    If FileExists(SlashPath & "TypeDB.cdb") Then Kill SlashPath & "TypeDB.cdb"
    Open SlashPath & "TypeDB.cdb" For Binary Access Write As #1
    Put #1, , HBytes
    Put #1, , CBytes
    Close
    MsgBox "TypeDB.csv compressed to TypeDB.cdb." & vbCrLf & "Original size: " & CompressZIt1.OriginalSize & vbCrLf & "Compressed Size: " & FileLen(SlashPath & "TypeDB.cdb"), vbInformation, "Done"
End Sub

Function FileExists(ByVal FileName As String) As Boolean
    'Determines if a file exists
    On Error GoTo Failed
    If Dir(FileName) = "" Then FileExists = False Else FileExists = True
    Exit Function
Failed:
    FileExists = False
End Function
