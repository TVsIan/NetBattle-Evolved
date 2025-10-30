VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Stadium 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Three Pokémon"
   ClientHeight    =   3975
   ClientLeft      =   4935
   ClientTop       =   2925
   ClientWidth     =   5100
   ControlBox      =   0   'False
   Icon            =   "Stadium.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   5100
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   "Opponent's Pokémon"
      Height          =   855
      Left            =   120
      TabIndex        =   24
      Top             =   0
      Width           =   4935
      Begin VB.Image imgEnemyPoke 
         Height          =   480
         Index           =   6
         Left            =   4320
         Top             =   240
         Width           =   480
      End
      Begin VB.Image imgEnemyPoke 
         Height          =   480
         Index           =   5
         Left            =   3480
         Top             =   240
         Width           =   480
      End
      Begin VB.Image imgEnemyPoke 
         Height          =   480
         Index           =   4
         Left            =   2640
         Top             =   240
         Width           =   480
      End
      Begin VB.Image imgEnemyPoke 
         Height          =   480
         Index           =   3
         Left            =   1800
         Top             =   240
         Width           =   480
      End
      Begin VB.Image imgEnemyPoke 
         Height          =   480
         Index           =   2
         Left            =   960
         Top             =   240
         Width           =   480
      End
      Begin VB.Image imgEnemyPoke 
         Height          =   480
         Index           =   1
         Left            =   120
         Top             =   240
         Width           =   480
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Pokémon Info"
      Height          =   2295
      Left            =   2400
      TabIndex        =   4
      Top             =   1080
      Width           =   2655
      Begin VB.Image imgInfo 
         Height          =   480
         Left            =   195
         Top             =   240
         Width           =   480
      End
      Begin VB.Label lblInfo 
         Caption         =   "Nickname"
         Height          =   255
         Index           =   0
         Left            =   720
         TabIndex        =   23
         Top             =   285
         Width           =   1455
      End
      Begin VB.Label lblInfo 
         Caption         =   "Lv.100 Pokemon (M)"
         Height          =   255
         Index           =   1
         Left            =   720
         TabIndex        =   22
         Top             =   525
         Width           =   1815
      End
      Begin VB.Label lblStat 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "999"
         Height          =   255
         Index           =   0
         Left            =   135
         TabIndex        =   21
         Top             =   1125
         Width           =   375
      End
      Begin VB.Label lblStat 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "999"
         Height          =   255
         Index           =   5
         Left            =   2145
         TabIndex        =   20
         Top             =   1125
         Width           =   375
      End
      Begin VB.Label lblStat 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "999"
         Height          =   255
         Index           =   4
         Left            =   1740
         TabIndex        =   19
         Top             =   1125
         Width           =   375
      End
      Begin VB.Label lblStat 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "999"
         Height          =   255
         Index           =   3
         Left            =   1335
         TabIndex        =   18
         Top             =   1125
         Width           =   375
      End
      Begin VB.Label lblStat 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "999"
         Height          =   255
         Index           =   2
         Left            =   945
         TabIndex        =   17
         Top             =   1125
         Width           =   375
      End
      Begin VB.Label lblStat 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "999"
         Height          =   255
         Index           =   1
         Left            =   540
         TabIndex        =   16
         Top             =   1125
         Width           =   375
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "HP"
         Height          =   255
         Index           =   2
         Left            =   135
         TabIndex        =   15
         Top             =   885
         Width           =   375
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Atk"
         Height          =   255
         Index           =   3
         Left            =   540
         TabIndex        =   14
         Top             =   885
         Width           =   375
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Def"
         Height          =   255
         Index           =   4
         Left            =   945
         TabIndex        =   13
         Top             =   885
         Width           =   375
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Spd"
         Height          =   255
         Index           =   5
         Left            =   1335
         TabIndex        =   12
         Top             =   885
         Width           =   375
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "SAtk"
         Height          =   255
         Index           =   6
         Left            =   1740
         TabIndex        =   11
         Top             =   885
         Width           =   375
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "SDef"
         Height          =   255
         Index           =   7
         Left            =   2145
         TabIndex        =   10
         Top             =   885
         Width           =   375
      End
      Begin VB.Line InfoLine 
         Index           =   0
         X1              =   120
         X2              =   2520
         Y1              =   825
         Y2              =   825
      End
      Begin VB.Line InfoLine 
         Index           =   1
         X1              =   120
         X2              =   2520
         Y1              =   1365
         Y2              =   1365
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Move 1"
         Height          =   255
         Index           =   8
         Left            =   165
         TabIndex        =   9
         Top             =   1425
         Width           =   1095
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Move 2"
         Height          =   255
         Index           =   9
         Left            =   1380
         TabIndex        =   8
         Top             =   1425
         Width           =   1095
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Move 3"
         Height          =   255
         Index           =   10
         Left            =   165
         TabIndex        =   7
         Top             =   1665
         Width           =   1095
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Move 4"
         Height          =   255
         Index           =   11
         Left            =   1380
         TabIndex        =   6
         Top             =   1665
         Width           =   1095
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Held Item: Held Item"
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   5
         Top             =   1965
         Width           =   2415
      End
      Begin VB.Line InfoLine 
         Index           =   2
         X1              =   120
         X2              =   2520
         Y1              =   1920
         Y2              =   1920
      End
   End
   Begin VB.CommandButton RemoveButton 
      Caption         =   "&Remove"
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton AddButton 
      Caption         =   "&Add"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton ReadyButton 
      Caption         =   "&Ready!"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   2400
      TabIndex        =   0
      Top             =   3480
      Width           =   2655
   End
   Begin MSComctlLib.ListView YourPKMN 
      Height          =   1635
      Left            =   120
      TabIndex        =   3
      Top             =   1140
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   2884
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Your Pokémon"
         Object.Width           =   3016
      EndProperty
   End
   Begin VB.Image imgSel 
      Height          =   480
      Index           =   3
      Left            =   1800
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   480
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   5040
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Image imgSel 
      Height          =   480
      Index           =   2
      Left            =   960
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   480
   End
   Begin VB.Image imgSel 
      Height          =   480
      Index           =   1
      Left            =   120
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   480
   End
End
Attribute VB_Name = "Stadium"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SelectPKMN(1 To 2, 0 To 6) As Pokemon
Dim Selection(1 To 3) As Long
Dim CurrentDisplay As Integer
Private Sub AddButton_Click()
    Dim X As Integer
    Dim SelectedPKMNNo As Integer
    For X = 1 To 6
        If YourPKMN.ListItems(X).Selected = True Then SelectedPKMNNo = X
    Next
    If SelectedPKMNNo = 0 Then Exit Sub
    If Selection(3) > 0 Then Exit Sub
    For X = 1 To 3
        If Selection(X) = SelectedPKMNNo Then
            MsgBox "Duplicate Pokémon!", vbCritical, "Error"
            Exit Sub
        End If
        If Selection(X) = 0 Then
            Selection(X) = SelectedPKMNNo
            Exit For
        End If
    Next X
    Call RefreshSel
    ReadyButton.Enabled = (X > 2)
    AddButton.Enabled = (X < 3)
    RemoveButton.Enabled = True
End Sub

Private Sub Form_Load()
    Dim X As Byte
    Dim Y As Byte
    RemoveButton.Enabled = False
    Call CenterWindow(Stadium)
    For X = 1 To 2
        For Y = 1 To 6
            Call Battle.GetPKMN(X, Y)
            SelectPKMN(X, Y) = Code.SwapClassPKMN
        Next
    Next
    For X = 1 To 6
        Call MainContainer.DoPicture(ChooseImage(SelectPKMN(Battle.ONum, X), nbGFXSml))
        imgEnemyPoke(X).Picture = MainContainer.SwapSpace.Picture
        'OpponentPKMN.ListItems.Add X, "Key:" & X, SelectPKMN(Battle.ONum, X).Nickname
    Next
    For X = 1 To 6
        YourPKMN.ListItems.Add X, "Key:" & X, SelectPKMN(Battle.PNum, X).Nickname
    Next
    Selection(1) = 0
    Selection(2) = 0
    Selection(3) = 0
    YourPKMN.ListItems(1).Selected = True
    Call FillInData(1)
    Call RefreshSel
End Sub
Sub RefreshSel()
    Dim Temp As String
    Dim X As Long
    For X = 1 To 3
        Temp = ChooseImage(SelectPKMN(Battle.PNum, Selection(X)), nbGFXSml)
        If imgSel(X).Tag <> Temp Then
            Call MainContainer.DoPicture(Temp)
            imgSel(X).Picture = MainContainer.SwapSpace.Picture
        End If
    Next X
End Sub
Sub FillInData(ByVal Number As Integer)
    Dim X As Byte
    Dim Z As Byte
    Dim ThisPKMN As Integer
    Dim Vis As Boolean
    If Number = CurrentDisplay Then Exit Sub
    With SelectPKMN(Battle.PNum, Number)
        Call MainContainer.DoPicture(ChooseImage(SelectPKMN(Battle.PNum, Number), nbGFXSml))
        imgInfo.Picture = MainContainer.SwapSpace.Picture
        lblInfo(0).Caption = .Nickname
        lblInfo(1).Caption = "Lv." & .Level & " " & .Name
        If Not Battle.RBYMode Then
            If .Gender = 1 Then
                lblInfo(1).Caption = lblInfo(1).Caption & " (M)"
            ElseIf .Gender = 2 Then
                lblInfo(1).Caption = lblInfo(1).Caption & " (F)"
            End If
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
        lblInfo(12).Visible = Not Battle.RBYMode
        lblInfo(7).Visible = Not Battle.RBYMode
        lblStat(5).Visible = Not Battle.RBYMode
        If Battle.RBYMode Then lblInfo(6).Caption = "Spcl"
        lblInfo(12).Caption = "Held Item: " & Item(.Item)
        
    End With
    CurrentDisplay = Number

'    Dim X As Integer
'
'    InfoText(0).Caption = "Species: " & SelectPKMN(Battle.PNum, Number).Name
'    If SelectPKMN(Battle.PNum, Number).Type2 = 0 Then
'        InfoText(1).Caption = "Type: " & Element(SelectPKMN(Battle.PNum, Number).Type1)
'    Else
'        InfoText(1).Caption = "Type: " & Element(SelectPKMN(Battle.PNum, Number).Type1) & "/" & Element(SelectPKMN(Battle.PNum, Number).Type2)
'    End If
'    InfoText(2).Caption = "Gender: " & Gender(SelectPKMN(Battle.PNum, Number).Gender)
'    InfoText(3).Caption = "HP: " & SelectPKMN(Battle.PNum, Number).HP
'    InfoText(4).Caption = "Attack: " & SelectPKMN(Battle.PNum, Number).Attack
'    InfoText(5).Caption = "Defense: " & SelectPKMN(Battle.PNum, Number).Defense
'    InfoText(6).Caption = "Speed: " & SelectPKMN(Battle.PNum, Number).Speed
'    InfoText(7).Caption = "Sp.Attack: " & SelectPKMN(Battle.PNum, Number).SpecialAttack
'    InfoText(8).Caption = "Sp.Defense: " & SelectPKMN(Battle.PNum, Number).SpecialDefense
''    MoveList.ListItems.Clear
''    MoveList.SmallIcons = MainContainer.Types
''    MoveList.Icons = MainContainer.Types
'    On Error Resume Next
'    For X = 1 To 4
'        imgMoveType(X - 1).Picture = MainContainer.Types.ListImages(Moves(SelectPKMN(Battle.PNum, Number).Move(X)).Type).Picture
'        InfoText(X + 8).Caption = Moves(SelectPKMN(Battle.PNum, Number).Move(X)).Name
'    Next
''    For X = 1 To 4
''        MoveList.ListItems.Add , , , Moves(SelectPKMN(Battle.PNum, Number).Move(X)).Type Moves(SelectPKMN(Battle.PNum, Number).Move(X)).Type
'    Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'OpponentPKMN.ListItems.Clear
    YourPKMN.ListItems.Clear
End Sub

'Private Sub OpponentPKMN_ItemClick(ByVal Item As MSComctlLib.ListItem)
'    Call FillInSecretData(Item.Index)
'End Sub

Private Sub ReadyButton_Click()
    Dim PokeString As String
    PokeString = Selection(1) & Selection(2) & Selection(3)
    Unload Me
    Call Battle.SetStadium(PokeString)
End Sub

Private Sub RemoveButton_Click()
    Dim X As Integer
    Dim SelectedPKMNNo As Integer
    For X = 3 To 1 Step -1
        If Selection(X) > 0 Then
            Selection(X) = 0
            Exit For
        End If
    Next X
    Call RefreshSel
    AddButton.Enabled = True
    ReadyButton.Enabled = False
    RemoveButton.Enabled = (X > 1)
End Sub

'Private Sub SelectedPKMN_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'    If DebugMode Then
'        MsgBox ColumnHeader.Width
'    End If
'End Sub

'Private Sub SelectedPKMN_DblClick()
'    Call RemoveButton_Click
'End Sub

'Private Sub YourPKMN_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'    If DebugMode Then
'        MsgBox ColumnHeader.Width
'    End If
'End Sub

Private Sub YourPKMN_DblClick()
    Call AddButton_Click
End Sub

Private Sub YourPKMN_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Call FillInData(Item.Index)
End Sub

'Sub FillInSecretData(ByVal Number As Integer)
'    Dim X As Integer
'
'    InfoText(0).Caption = "Species: " & SelectPKMN(Battle.ONum, Number).Name
'    If SelectPKMN(Battle.ONum, Number).Type2 = 0 Then
'        InfoText(1).Caption = "Type: " & Element(SelectPKMN(Battle.ONum, Number).Type1)
'    Else
'        InfoText(1).Caption = "Type: " & Element(SelectPKMN(Battle.ONum, Number).Type1) & "/" & Element(SelectPKMN(Battle.ONum, Number).Type2)
'    End If
'    InfoText(2).Caption = "Gender: " & Gender(SelectPKMN(Battle.ONum, Number).Gender)
'    InfoText(3).Caption = "HP: ???"
'    InfoText(4).Caption = "Attack: ???"
'    InfoText(5).Caption = "Defense: ???"
'    InfoText(6).Caption = "Speed: ???"
'    InfoText(7).Caption = "Sp.Attack: ???"
'    InfoText(8).Caption = "Sp.Defense: ???"
''    MoveList.ListItems.Clear
''    MoveList.SmallIcons = MainContainer.Types
''    MoveList.Icons = MainContainer.Types
'    For X = 1 To 4
'        imgMoveType(X - 1).Picture = Nothing
'        InfoText(X + 8).Caption = "???"
'    Next
'End Sub
