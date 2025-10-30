VERSION 5.00
Begin VB.Form Expert 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Expert Build"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3630
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   3630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox HPSelect 
      Height          =   315
      Left            =   360
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   1920
      Width           =   1695
   End
   Begin VB.TextBox Lev 
      Height          =   285
      Left            =   3000
      MaxLength       =   3
      TabIndex        =   18
      Text            =   "100"
      Top             =   360
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   2400
      TabIndex        =   15
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CheckBox Shiny 
      Caption         =   "Shiny"
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   1200
      Width           =   735
   End
   Begin VB.ComboBox DV_SAtk 
      Height          =   315
      ItemData        =   "Expert.frx":0000
      Left            =   2280
      List            =   "Expert.frx":0034
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   360
      Width           =   615
   End
   Begin VB.ComboBox DV_Spd 
      Height          =   315
      ItemData        =   "Expert.frx":006E
      Left            =   1560
      List            =   "Expert.frx":00A2
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   360
      Width           =   615
   End
   Begin VB.ComboBox DV_Def 
      Height          =   315
      ItemData        =   "Expert.frx":00DC
      Left            =   840
      List            =   "Expert.frx":0110
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   360
      Width           =   615
   End
   Begin VB.ComboBox DV_Atk 
      Height          =   315
      ItemData        =   "Expert.frx":014A
      Left            =   120
      List            =   "Expert.frx":017E
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Level"
      Height          =   255
      Left            =   3000
      TabIndex        =   19
      Top             =   120
      Width           =   615
   End
   Begin VB.Label GenderDisp 
      BackStyle       =   0  'Transparent
      Caption         =   "Gender:"
      Height          =   255
      Left            =   1920
      TabIndex        =   17
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label HP 
      BackStyle       =   0  'Transparent
      Caption         =   "HP:"
      Height          =   255
      Left            =   360
      TabIndex        =   16
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label HiddenPower 
      BackStyle       =   0  'Transparent
      Caption         =   "Hidden Power:"
      Height          =   255
      Left            =   360
      TabIndex        =   14
      Top             =   1680
      Width           =   2535
   End
   Begin VB.Label SDefensePower 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   2280
      TabIndex        =   12
      Top             =   960
      Width           =   615
   End
   Begin VB.Label SAttackPower 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   2280
      TabIndex        =   11
      Top             =   720
      Width           =   615
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Special"
      Height          =   255
      Left            =   2280
      TabIndex        =   10
      Top             =   120
      Width           =   615
   End
   Begin VB.Label SpeedPower 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   1560
      TabIndex        =   8
      Top             =   720
      Width           =   615
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Speed"
      Height          =   255
      Left            =   1560
      TabIndex        =   7
      Top             =   120
      Width           =   615
   End
   Begin VB.Label DefensePower 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   840
      TabIndex        =   5
      Top             =   720
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Defense"
      Height          =   255
      Left            =   840
      TabIndex        =   4
      Top             =   120
      Width           =   615
   End
   Begin VB.Label AttackPower 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Attack"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "Expert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DoneLoading As Boolean

Private Sub DV_Atk_Click()
    If DoneLoading Then Call RefreshStats
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub DV_Def_Click()
    If DoneLoading Then Call RefreshStats
End Sub

Private Sub Form_Load()
    Dim X As Integer
    
    DV_Atk.Clear
    DV_Def.Clear
    DV_Spd.Clear
    DV_SAtk.Clear
    For X = 0 To 15
        DV_Atk.AddItem X, X
        DV_Def.AddItem X, X
        DV_Spd.AddItem X, X
        DV_SAtk.AddItem X, X
    Next
    DoneLoading = False
    DV_Atk.ListIndex = ExpertPKMN.DV_Atk
    DV_Def.ListIndex = ExpertPKMN.DV_Def
    DV_Spd.ListIndex = ExpertPKMN.DV_Spd
    DV_SAtk.ListIndex = ExpertPKMN.DV_SAtk
    DoneLoading = True
    Lev.Text = ExpertPKMN.Level
    HPSelect.AddItem Element(7)
    HPSelect.AddItem Element(10)
    HPSelect.AddItem Element(8)
    HPSelect.AddItem Element(9)
    HPSelect.AddItem Element(13)
    HPSelect.AddItem Element(12)
    HPSelect.AddItem Element(14)
    HPSelect.AddItem Element(17)
    HPSelect.AddItem Element(2)
    HPSelect.AddItem Element(3)
    HPSelect.AddItem Element(5)
    HPSelect.AddItem Element(4)
    HPSelect.AddItem Element(11)
    HPSelect.AddItem Element(6)
    HPSelect.AddItem Element(15)
    HPSelect.AddItem Element(16)
    If ExpertPKMN.GameVersion = nbRBYTrade Or ExpertPKMN.GameVersion = nbTrueRBY Then
        HPSelect.Visible = False
        HiddenPower.Visible = False
        GenderDisp.Visible = False
        Shiny.Visible = False
        Command1.Top = 1200
        Expert.Height = 2055
    Else
        HPSelect.Visible = True
        HiddenPower.Visible = True
        GenderDisp.Visible = True
        Shiny.Visible = True
        Command1.Top = 1920
        Expert.Height = 2775
    End If
    Call RefreshStats
End Sub

Sub RefreshStats()
    If DV_Atk.ListIndex = -1 Then DV_Atk.ListIndex = 15
    If DV_Def.ListIndex = -1 Then DV_Def.ListIndex = 15
    If DV_Spd.ListIndex = -1 Then DV_Spd.ListIndex = 15
    If DV_SAtk.ListIndex = -1 Then DV_SAtk.ListIndex = 15
    With ExpertPKMN
        .DV_Atk = DV_Atk.ListIndex
        .DV_Def = DV_Def.ListIndex
        .DV_Spd = DV_Spd.ListIndex
        .DV_SAtk = DV_SAtk.ListIndex
        .Level = Val(Lev.Text)
    
        'Adjust stats for DVs
        .Attack = GetStat(.Level, .BaseAttack, .DV_Atk)
        .Defense = GetStat(.Level, .BaseDefense, .DV_Def)
        .Speed = GetStat(.Level, .BaseSpeed, .DV_Spd)
        Select Case TBMode
        Case 0, 5
            .SpecialAttack = GetStat(.Level, .BaseSpecial, .DV_SAtk)
            .SpecialDefense = GetStat(.Level, .BaseSpecial, .DV_SAtk)
        Case Else
            .SpecialAttack = GetStat(.Level, .BaseSAttack, .DV_SAtk)
            .SpecialDefense = GetStat(.Level, .BaseSDefense, .DV_SAtk)
        End Select
        
        'Generate HP DV based on the others
        .DV_HP = 0
        If .DV_Atk Mod 2 = 1 Then .DV_HP = .DV_HP + 8
        If .DV_Def Mod 2 = 1 Then .DV_HP = .DV_HP + 4
        If .DV_Spd Mod 2 = 1 Then .DV_HP = .DV_HP + 2
        If .DV_SAtk Mod 2 = 1 Then .DV_HP = .DV_HP + 1
        .MaxHP = GetHP(.Level, .BaseHP, .DV_HP)
        .HP = .MaxHP
        AttackPower.Caption = .Attack
        DefensePower.Caption = .Defense
        SpeedPower.Caption = .Speed
        SAttackPower.Caption = .SpecialAttack
        SDefensePower.Caption = .SpecialDefense
        HP.Caption = "HP: " & .MaxHP
        
        'Set gender
        If .PercentFemale = -1 Then
            .Gender = 0
        Else
            If .DV_Atk <= .PercentFemale - 1 Then .Gender = 2 Else .Gender = 1
        End If
        If ShinyDV(.DV_Atk) And .DV_Def = 10 And .DV_Spd = 10 And .DV_SAtk = 10 And (.GameVersion <> nbTrueRBY And .GameVersion <> nbRBYTrade) Then Shiny.Value = 1 Else Shiny.Value = 0
        HiddenPower.Caption = "Hidden Power: " & Element(HiddenPowerType(DV_Atk.ListIndex, DV_Def.ListIndex)) & ", " & HiddenPowerStrength(DV_Atk.ListIndex, DV_Def.ListIndex, DV_Spd.ListIndex, DV_SAtk.ListIndex)
        GenderDisp.Caption = "Gender: " & Gender(.Gender)
        If Shiny.Value = 1 Then .Shiny = True Else .Shiny = False
    End With
End Sub

Private Sub HPSelect_Click()
    Select Case HPSelect.ListIndex
        Case 0
            DV_Atk.ListIndex = 12
            DV_Def.ListIndex = 12
        Case 1
            DV_Atk.ListIndex = 12
            DV_Def.ListIndex = 13
        Case 2
            DV_Atk.ListIndex = 12
            DV_Def.ListIndex = 14
        Case 3
            DV_Atk.ListIndex = 12
            DV_Def.ListIndex = 15
        Case 4
            DV_Atk.ListIndex = 13
            DV_Def.ListIndex = 12
        Case 5
            DV_Atk.ListIndex = 13
            DV_Def.ListIndex = 13
        Case 6
            DV_Atk.ListIndex = 13
            DV_Def.ListIndex = 14
        Case 7
            DV_Atk.ListIndex = 13
            DV_Def.ListIndex = 15
        Case 8
            DV_Atk.ListIndex = 14
            DV_Def.ListIndex = 12
        Case 9
            DV_Atk.ListIndex = 14
            DV_Def.ListIndex = 13
        Case 10
            DV_Atk.ListIndex = 14
            DV_Def.ListIndex = 14
        Case 11
            DV_Atk.ListIndex = 14
            DV_Def.ListIndex = 15
        Case 12
            DV_Atk.ListIndex = 15
            DV_Def.ListIndex = 12
        Case 13
            DV_Atk.ListIndex = 15
            DV_Def.ListIndex = 13
        Case 14
            DV_Atk.ListIndex = 15
            DV_Def.ListIndex = 14
        Case 15
            DV_Atk.ListIndex = 15
            DV_Def.ListIndex = 15
    End Select
    DV_Spd.ListIndex = 15
    DV_SAtk.ListIndex = 15
    Call RefreshStats
End Sub

Private Sub HPSelect_KeyUp(KeyCode As Integer, Shift As Integer)
    Call HPSelect_Click
End Sub

Private Sub Lev_Change()
    If Val(Lev.Text) < 1 Or Val(Lev.Text) > 100 Or Not IsNumeric(Lev.Text) Then Lev.Text = "1"
    Call RefreshStats
End Sub

Private Sub DV_SAtk_Click()
    If DoneLoading Then Call RefreshStats
End Sub

Private Sub DV_Spd_Click()
    If DoneLoading Then Call RefreshStats
End Sub
