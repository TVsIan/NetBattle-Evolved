VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form BoxArrange 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Move Pokémon"
   ClientHeight    =   3975
   ClientLeft      =   2760
   ClientTop       =   3630
   ClientWidth     =   7935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   7935
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton BoxNav 
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   2520
      Picture         =   "BoxArrange.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Move to another box"
      Top             =   3000
      Width           =   375
   End
   Begin VB.CommandButton BoxNav 
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   2520
      Picture         =   "BoxArrange.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Delete"
      Top             =   3480
      Width           =   375
   End
   Begin VB.CommandButton BoxNav 
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   2520
      Picture         =   "BoxArrange.frx":0294
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Move down"
      Top             =   2520
      Width           =   375
   End
   Begin VB.CommandButton BoxNav 
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   2520
      Picture         =   "BoxArrange.frx":03DE
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Move up"
      Top             =   2040
      Width           =   375
   End
   Begin VB.CommandButton BoxNav 
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   2520
      Picture         =   "BoxArrange.frx":0528
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Copy Pokemon to box"
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton BoxNav 
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   2520
      Picture         =   "BoxArrange.frx":0672
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Insert at current position"
      Top             =   1080
      Width           =   375
   End
   Begin MSComctlLib.ImageList Balls 
      Left            =   7920
      Top             =   5280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483633
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711680
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BoxArrange.frx":07BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BoxArrange.frx":0D56
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TabStrip BoxPick 
      Height          =   735
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   1296
      TabWidthStyle   =   2
      MultiRow        =   -1  'True
      Style           =   1
      TabFixedWidth   =   2646
      TabFixedHeight  =   582
      HotTracking     =   -1  'True
      Separators      =   -1  'True
      TabMinWidth     =   0
      ImageList       =   "Balls"
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   10
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Box 1"
            ImageVarType    =   2
            ImageIndex      =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Box 2"
            ImageVarType    =   2
            ImageIndex      =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Box 3"
            ImageVarType    =   2
            ImageIndex      =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Box 4"
            ImageVarType    =   2
            ImageIndex      =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Box 5"
            ImageVarType    =   2
            ImageIndex      =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Box 6"
            ImageVarType    =   2
            ImageIndex      =   2
         EndProperty
         BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Box 7"
            ImageVarType    =   2
            ImageIndex      =   2
         EndProperty
         BeginProperty Tab8 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Box 8"
            ImageVarType    =   2
            ImageIndex      =   2
         EndProperty
         BeginProperty Tab9 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Box 9"
            ImageVarType    =   2
            ImageIndex      =   2
         EndProperty
         BeginProperty Tab10 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Box 10"
            ImageVarType    =   2
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView TeamBox 
      Height          =   2775
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   4895
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Active"
         Object.Width           =   3792
      EndProperty
   End
   Begin MSComctlLib.ListView PokeBox 
      Height          =   2775
      Left            =   3000
      TabIndex        =   4
      Top             =   1080
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   4895
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Active"
         Object.Width           =   3792
      EndProperty
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5400
      TabIndex        =   1
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   6720
      TabIndex        =   0
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Image lblMarker 
      Height          =   240
      Index           =   3
      Left            =   7200
      Picture         =   "BoxArrange.frx":12F0
      Top             =   3120
      Width           =   240
   End
   Begin VB.Image lblMarker 
      Height          =   240
      Index           =   2
      Left            =   6720
      Picture         =   "BoxArrange.frx":13C2
      Top             =   3120
      Width           =   240
   End
   Begin VB.Image lblMarker 
      Height          =   240
      Index           =   1
      Left            =   6240
      Picture         =   "BoxArrange.frx":1494
      Top             =   3120
      Width           =   240
   End
   Begin VB.Image lblMarker 
      Height          =   240
      Index           =   0
      Left            =   5760
      Picture         =   "BoxArrange.frx":1566
      Top             =   3120
      Width           =   240
   End
   Begin VB.Image imgInfo 
      Height          =   480
      Left            =   5475
      Top             =   1080
      Width           =   480
   End
   Begin VB.Label lblInfo 
      Caption         =   "Nickname"
      Height          =   255
      Index           =   0
      Left            =   6000
      TabIndex        =   31
      Top             =   1125
      Width           =   1455
   End
   Begin VB.Label lblInfo 
      Caption         =   "Lv.100 Pokemon (M)"
      Height          =   255
      Index           =   1
      Left            =   6000
      TabIndex        =   30
      Top             =   1365
      Width           =   1815
   End
   Begin VB.Label lblStat 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "999"
      Height          =   255
      Index           =   0
      Left            =   5415
      TabIndex        =   29
      Top             =   1965
      Width           =   375
   End
   Begin VB.Label lblStat 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "999"
      Height          =   255
      Index           =   5
      Left            =   7425
      TabIndex        =   28
      Top             =   1965
      Width           =   375
   End
   Begin VB.Label lblStat 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "999"
      Height          =   255
      Index           =   4
      Left            =   7020
      TabIndex        =   27
      Top             =   1965
      Width           =   375
   End
   Begin VB.Label lblStat 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "999"
      Height          =   255
      Index           =   3
      Left            =   6615
      TabIndex        =   26
      Top             =   1965
      Width           =   375
   End
   Begin VB.Label lblStat 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "999"
      Height          =   255
      Index           =   2
      Left            =   6225
      TabIndex        =   25
      Top             =   1965
      Width           =   375
   End
   Begin VB.Label lblStat 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "999"
      Height          =   255
      Index           =   1
      Left            =   5820
      TabIndex        =   24
      Top             =   1965
      Width           =   375
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "HP"
      Height          =   255
      Index           =   2
      Left            =   5415
      TabIndex        =   23
      Top             =   1725
      Width           =   375
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Atk"
      Height          =   255
      Index           =   3
      Left            =   5820
      TabIndex        =   22
      Top             =   1725
      Width           =   375
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Def"
      Height          =   255
      Index           =   4
      Left            =   6225
      TabIndex        =   21
      Top             =   1725
      Width           =   375
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Spd"
      Height          =   255
      Index           =   5
      Left            =   6615
      TabIndex        =   20
      Top             =   1725
      Width           =   375
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SAtk"
      Height          =   255
      Index           =   6
      Left            =   7020
      TabIndex        =   19
      Top             =   1725
      Width           =   375
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SDef"
      Height          =   255
      Index           =   7
      Left            =   7425
      TabIndex        =   18
      Top             =   1725
      Width           =   375
   End
   Begin VB.Line InfoLine 
      Index           =   0
      X1              =   5400
      X2              =   7800
      Y1              =   1665
      Y2              =   1665
   End
   Begin VB.Line InfoLine 
      Index           =   1
      X1              =   5400
      X2              =   7800
      Y1              =   2205
      Y2              =   2205
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Move 1"
      Height          =   255
      Index           =   8
      Left            =   5445
      TabIndex        =   17
      Top             =   2265
      Width           =   1095
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Move 2"
      Height          =   255
      Index           =   9
      Left            =   6660
      TabIndex        =   16
      Top             =   2265
      Width           =   1095
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Move 3"
      Height          =   255
      Index           =   10
      Left            =   5445
      TabIndex        =   15
      Top             =   2505
      Width           =   1095
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Move 4"
      Height          =   255
      Index           =   11
      Left            =   6660
      TabIndex        =   14
      Top             =   2505
      Width           =   1095
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Held Item: Held Item"
      Height          =   255
      Index           =   12
      Left            =   5400
      TabIndex        =   13
      Top             =   2805
      Width           =   2415
   End
   Begin VB.Line InfoLine 
      Index           =   2
      X1              =   5400
      X2              =   7800
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Image imgBoxVer 
      Height          =   240
      Left            =   7560
      Stretch         =   -1  'True
      ToolTipText     =   "True RBY"
      Top             =   1080
      Width           =   240
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Your Box"
      Height          =   255
      Left            =   3000
      TabIndex        =   3
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Your Team"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   2295
   End
End
Attribute VB_Name = "BoxArrange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CTemp(0 To 4) As Boolean
Dim CurrWindow As Byte
Dim NewPKMN(1 To 6) As Pokemon
Dim DontChange As Boolean
Dim CurrBox As Byte
Dim TempPKMN As Pokemon
Dim CurrPoke As Pokemon

Private Sub BoxPick_Click()
    Dim X As Integer
    Dim Y As Integer
    CurrBox = BoxPick.SelectedItem.Index
    On Error Resume Next
    X = 1
    Y = 1
    X = TeamBox.SelectedItem.Index
    Y = PokeBox.SelectedItem.Index
    Call RefreshLists(X, Y)
End Sub

Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim X As Byte
    
    Call ReadBoxPKMN
    DontChange = True
    'OnePKMN(0).Picture = LoadResPicture("RBYT", vbResIcon)
    'OnePKMN(1).Picture = LoadResPicture("GSCT", vbResIcon)
    'OnePKMN(2).Picture = LoadResPicture("ADV", vbResIcon)
    'OnePKMN(3).Picture = LoadResPicture("ADV+", vbResIcon)
    'OnePKMN(4).Picture = LoadResPicture("MOD", vbResIcon)
    'OnePKMN(5).Picture = LoadResPicture("RBY", vbResIcon)
    'OnePKMN(6).Picture = LoadResPicture("GSC", vbResIcon)
    For X = 1 To 6
        NewPKMN(X) = PKMN(X)
    Next
    CurrWindow = 1
    CurrBox = 1
    BoxPick.Tabs(1).Selected = True
    Call RefreshLists(1, 1)
End Sub

Private Sub RefreshLists(ByVal Index1 As Integer, ByVal Index2 As Integer)
    Dim X As Integer
    Dim TempPKMN As Pokemon
    Dim LeftPKMN As Integer
    Dim RightPKMN As Integer
    
    On Error Resume Next
    RightPKMN = -1
    LeftPKMN = TeamBox.SelectedItem.Index
    RightPKMN = Dec(Right(PokeBox.SelectedItem.Key, 4))
    
    TeamBox.ListItems.Clear
    For X = 1 To 6
        TeamBox.ListItems.Add , , NewPKMN(X).Nickname & " (" & NewPKMN(X).Name & ")"
    Next
    PokeBox.ListItems.Clear
    For X = 1 To 10
        BoxPick.Tabs(X).Image = 2
    Next X
    For X = 1 To UBound(BoxPKMN)
        BoxPick.Tabs(BoxPKMN(X).InBox).Image = 1
        If BoxPKMN(X).InBox = CurrBox Then
            PokeBox.ListItems.Add , "PKNUM" & FixedHex(X, 4), BoxPKMN(X).Nickname & " (" & BoxPKMN(X).Name & ")"
            If BoxPKMN(X).GameVersion <> Player(YourNumber).GameVersion Then
                PokeBox.ListItems("PKNUM" & FixedHex(X, 4)).ForeColor = vbRed
            Else
                PokeBox.ListItems("PKNUM" & FixedHex(X, 4)).ForeColor = vbBlack
            End If
        End If
    Next
    On Error Resume Next
    TeamBox.SelectedItem = TeamBox.ListItems(Index1)
    PokeBox.SelectedItem = PokeBox.ListItems(Index2)
    If CurrWindow = 1 Then
        Label1.ForeColor = vbRed
        Label2.ForeColor = vbBlack
        TempPKMN = NewPKMN(TeamBox.SelectedItem.Index)
    Else
        Label2.ForeColor = vbRed
        Label1.ForeColor = vbBlack
        TempPKMN = BoxPKMN(Val(PokeBox.SelectedItem.Tag))
    End If
    Call ReadBinArray(CompatCheck(PKMN), Compatibility)
    'For X = 0 To 6
    '    OnePKMN(X).Visible = Compatibility(X)
    'Next
    Select Case CurrWindow
        Case 1
            CurrPoke = NewPKMN(LeftPKMN)
        Case 2
            If RightPKMN = -1 Then CurrPoke = NewPKMN(LeftPKMN) Else CurrPoke = BoxPKMN(RightPKMN)
    End Select
    Call ShowData
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call WriteBoxPKMN
    Call ReadBinArray(CompatCheck(PKMN), Compatibility)
End Sub

Private Sub BoxNav_Click(Index As Integer)
    Dim Answer As Integer
    Dim X As Integer
    Dim Y As Integer
    Dim HasMoves As Boolean
    Dim TempPKMN As Pokemon
    Dim LeftPKMN As Integer
    Dim RightPKMN As Integer
    Dim BoxMove As Integer
    On Error Resume Next
    RightPKMN = -1
    LeftPKMN = TeamBox.SelectedItem.Index
    RightPKMN = Dec(Right(PokeBox.SelectedItem.Tag, 4))
    If RightPKMN = -1 And Not (((Index = 2 Or Index = 3) And CurrWindow = 1) Or Index = 1) Then Exit Sub
    Select Case Index
        'Left
        Case 0
            If BoxPKMN(RightPKMN).GameVersion <> Player(YourNumber).GameVersion Then MsgBox "This Pokémon is not compatible with your current team", vbInformation, "Cannot add to team": Exit Sub
'            For X = 1 To 6
'                If NewPKMN(X).No = BoxPKMN(RightPKMN).No _
'                And X <> LeftPKMN Then
'                    MsgBox "You already have a " & NewPKMN(X).Name & " on your team!", vbCritical, "Duplicate Pokémon"
'                    Exit Sub
'                End If
'            Next
            Y = 0
            For X = 1 To 4
                If BoxPKMN(RightPKMN).Move(X) <> 0 Then Y = Y + 1
            Next X
            If Y = 0 Then
                MsgBox "This Pokémon is imcomplete, therefore it cannot be added to your team.", vbCritical, "Incomplete Pokémon"
                Exit Sub
            End If
            DontChange = False
            NewPKMN(LeftPKMN) = BoxPKMN(RightPKMN)
            NewPKMN(LeftPKMN).Image = ChooseImage(NewPKMN(LeftPKMN), You.Version)
            Call RefreshLists(TeamBox.SelectedItem.Index, PokeBox.SelectedItem.Index)
        'Right
        Case 1
            ReDim Preserve BoxPKMN(UBound(BoxPKMN) + 1) As Pokemon
            BoxPKMN(UBound(BoxPKMN)) = NewPKMN(LeftPKMN)
            BoxPKMN(UBound(BoxPKMN)).InBox = CurrBox
            Call RefreshLists(TeamBox.SelectedItem.Index, UBound(BoxPKMN) - 1)
        'Up
        Case 2
            Select Case CurrWindow
                Case 1
                    If LeftPKMN = 1 Then Exit Sub
                    TempPKMN = NewPKMN(LeftPKMN)
                    NewPKMN(LeftPKMN) = NewPKMN(LeftPKMN - 1)
                    NewPKMN(LeftPKMN - 1) = TempPKMN
                    Call RefreshLists(TeamBox.SelectedItem.Index - 1, PokeBox.SelectedItem.Index)
                    DontChange = False
                Case 2
                    If PokeBox.SelectedItem.Index = 1 Then Exit Sub
                    BoxMove = PokeBox.SelectedItem.Tag
                    TempPKMN = BoxPKMN(RightPKMN)
                    BoxPKMN(RightPKMN) = BoxPKMN(BoxMove)
                    BoxPKMN(BoxMove) = TempPKMN
                    Call RefreshLists(TeamBox.SelectedItem.Index, PokeBox.SelectedItem.Index - 1)
            End Select
        'Down
        Case 3
            Select Case CurrWindow
                Case 1
                    If LeftPKMN = 6 Then Exit Sub
                    TempPKMN = NewPKMN(LeftPKMN)
                    NewPKMN(LeftPKMN) = NewPKMN(LeftPKMN + 1)
                    NewPKMN(LeftPKMN + 1) = TempPKMN
                    Call RefreshLists(TeamBox.SelectedItem.Index + 1, PokeBox.SelectedItem.Index)
                    DontChange = False
                Case 2
                    If PokeBox.SelectedItem.Index = PokeBox.ListItems.count Then Exit Sub
                    TempPKMN = BoxPKMN(RightPKMN)
                    BoxMove = Val(PokeBox.ListItems(PokeBox.SelectedItem.Index + 1).Tag)
                    BoxPKMN(PokeBox.SelectedItem.Index) = BoxPKMN(BoxMove)
                    BoxPKMN(BoxMove) = TempPKMN
                    Call RefreshLists(TeamBox.SelectedItem.Index, PokeBox.SelectedItem.Index + 1)
            End Select
        'Del
        Case 4
            With BoxPKMN(RightPKMN)
                Answer = MsgBox("Are you sure you want to delete " & .Nickname & " (" & .Name & ")?", vbYesNo + vbQuestion, "Confirm Delete")
            End With
            If Answer = vbNo Then Exit Sub
            For X = RightPKMN To UBound(BoxPKMN) - 1
                BoxPKMN(X) = BoxPKMN(X + 1)
            Next
            ReDim Preserve BoxPKMN(UBound(BoxPKMN) - 1) As Pokemon
            Call RefreshLists(TeamBox.SelectedItem.Index, 1)
        'Move
        Case 5
            MoveBoxNum = RightPKMN
            FromBox = CurrBox
            ToBox = -1
            MovePKMN.Show vbModal
            If ToBox = -1 Then Exit Sub
            BoxPKMN(RightPKMN).InBox = ToBox
            TempPKMN = BoxPKMN(RightPKMN)
            For X = RightPKMN To UBound(BoxPKMN) - 1
                BoxPKMN(X) = BoxPKMN(X + 1)
            Next X
            BoxPKMN(X) = TempPKMN
            Call RefreshLists(1, 1)
    End Select
End Sub

Private Sub OKButton_Click()
    Dim X As Integer
    Dim Y As Integer
    Dim Z As String
    For X = 1 To 6
        PKMN(X) = NewPKMN(X)
    Next X
    If TeamChangeFromMS Then MasterServer.TeamChanged = Not DontChange
    Unload Me
End Sub

Private Sub PokeBox_Click()
    Dim X As Integer
    Dim Y As Integer
    CurrWindow = 2
    On Error Resume Next
    X = 1
    Y = 1
    X = TeamBox.SelectedItem.Index
    Y = PokeBox.SelectedItem.Index
    Call RefreshLists(X, Y)
End Sub

Private Sub TeamBox_Click()
    Dim X As Integer
    Dim Y As Integer
    CurrWindow = 1
    On Error Resume Next
    X = 1
    Y = 1
    X = TeamBox.SelectedItem.Index
    Y = PokeBox.SelectedItem.Index
    Call RefreshLists(X, Y)
End Sub

Private Sub ShowData()
    Dim X As Byte
    Dim Z As Byte
    Dim ThisPKMN As Integer
    Dim Vis As Boolean
    'ThisPKMN = -1
    On Error Resume Next
    'ThisPKMN = Val(PokeItem.Tag)
    'If ThisPKMN = CurrentDisplay Then Exit Sub
    'Vis = (ThisPKMN <> -1)
    Vis = True
    If Vis Then
        With CurrPoke
            Call MainContainer.DoPicture(ChooseImage(CurrPoke, nbGFXSml))
            imgInfo.Picture = MainContainer.SwapSpace.Picture
'            GenderImage(1).Visible = (.Gender = 1)
'            GenderImage(2).Visible = (.Gender = 2)
            lblInfo(0).Caption = .Nickname
            lblInfo(1).Caption = "Lv." & .Level & " " & .Name
            If .Gender = 1 Then
                lblInfo(1).Caption = lblInfo(1).Caption & " (M)"
            ElseIf .Gender = 2 Then
                lblInfo(1).Caption = lblInfo(1).Caption & " (F)"
            End If
            lblStat(0).Caption = .MaxHP
            lblStat(1).Caption = .Attack
            lblStat(2).Caption = .Defense
            lblStat(3).Caption = .Speed
            lblStat(4).Caption = .SpecialAttack
            lblStat(5).Caption = .SpecialDefense
            For X = 8 To 11
                lblInfo(X).Caption = Moves(.Move(X - 7)).Name
            Next X
            lblInfo(12).Caption = "Held Item: " & Item(.Item)
            
            Select Case .GameVersion
            Case 0: imgBoxVer.Picture = LoadResPicture("RBYT", vbResIcon)
            Case 1: imgBoxVer.Picture = LoadResPicture("GSCT", vbResIcon)
            Case 2: imgBoxVer.Picture = LoadResPicture("ADV", vbResIcon)
            Case 3: imgBoxVer.Picture = LoadResPicture("ADV+", vbResIcon)
            Case 4: imgBoxVer.Picture = LoadResPicture("MOD", vbResIcon)
            Case 5: imgBoxVer.Picture = LoadResPicture("RBY", vbResIcon)
            Case 6: imgBoxVer.Picture = LoadResPicture("GSC", vbResIcon)
            End Select
        
            For Z = 0 To 3
                lblMarker(Z).Visible = ((.MarkerNum And 2 ^ Z) > 0)
                'mnuMarkerItem(Z).Checked = lblMarker(Z).Visible
            Next Z

        End With
        BoxNav(2).Enabled = (PokeBox.SelectedItem.Index <> 1)
        BoxNav(3).Enabled = (PokeBox.SelectedItem.Index <> PokeBox.ListItems.count)

    End If
    imgInfo.Visible = Vis
    For X = 0 To 12
        lblInfo(X).Visible = Vis
    Next X
    For X = 0 To 5
        lblStat(X).Visible = Vis
    Next X
    For X = 0 To 2
        InfoLine(X).Visible = Vis
    Next X
    'CurrentDisplay = ThisPKMN
End Sub
