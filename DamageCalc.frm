VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form DamageCalc 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Damage Calculator"
   ClientHeight    =   3690
   ClientLeft      =   3075
   ClientTop       =   3855
   ClientWidth     =   8745
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   8745
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      Caption         =   "Attacker"
      Height          =   1935
      Left            =   120
      TabIndex        =   28
      Top             =   0
      Width           =   4575
      Begin VB.CheckBox chkDamageCalc 
         Caption         =   "Same Type Attack Bonus"
         Height          =   195
         Index           =   0
         Left            =   2280
         TabIndex        =   3
         Top             =   240
         Width           =   2175
      End
      Begin VB.TextBox txtDamageCalc 
         Height          =   285
         Index           =   1
         Left            =   960
         MaxLength       =   4
         TabIndex        =   21
         Text            =   "200"
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox txtDamageCalc 
         Height          =   285
         Index           =   0
         Left            =   120
         MaxLength       =   3
         TabIndex        =   20
         Text            =   "100"
         Top             =   480
         Width           =   615
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   0
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox txtDamageCalc 
         Height          =   315
         Index           =   3
         Left            =   120
         MaxLength       =   3
         TabIndex        =   1
         Text            =   "0"
         Top             =   1440
         Width           =   495
      End
      Begin VB.CheckBox chkDamageCalc 
         Caption         =   "Type Boosting Item Bonus"
         Height          =   195
         Index           =   1
         Left            =   2280
         TabIndex        =   4
         Top             =   480
         Width           =   2175
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         ItemData        =   "DamageCalc.frx":0000
         Left            =   2280
         List            =   "DamageCalc.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1320
         Width           =   1695
      End
      Begin VB.CheckBox chkDamageCalc 
         Caption         =   "Low-HP Trait Bonus"
         Height          =   195
         Index           =   3
         Left            =   2280
         TabIndex        =   5
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label lblDamageCalc 
         BackStyle       =   0  'Transparent
         Caption         =   "Attack/Sp. Atk"
         Height          =   255
         Index           =   1
         Left            =   960
         TabIndex        =   32
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblDamageCalc 
         BackStyle       =   0  'Transparent
         Caption         =   "Level"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblDamageCalc 
         BackStyle       =   0  'Transparent
         Caption         =   "Move Power"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   30
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label lblDamageCalc 
         BackStyle       =   0  'Transparent
         Caption         =   "Weather Modifier"
         Height          =   255
         Index           =   13
         Left            =   2280
         TabIndex        =   29
         Top             =   1080
         Width           =   1935
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Defender"
      Height          =   1695
      Left            =   120
      TabIndex        =   24
      Top             =   1920
      Width           =   4575
      Begin VB.TextBox txtDamageCalc 
         Height          =   285
         Index           =   2
         Left            =   120
         MaxLength       =   4
         TabIndex        =   7
         Text            =   "200"
         Top             =   480
         Width           =   615
      End
      Begin VB.OptionButton optReflect 
         Caption         =   "Double Battle"
         Height          =   195
         Index           =   2
         Left            =   2280
         TabIndex        =   11
         Top             =   960
         Width           =   1335
      End
      Begin VB.OptionButton optReflect 
         Caption         =   "Single Battle"
         Height          =   195
         Index           =   1
         Left            =   2280
         TabIndex        =   10
         Top             =   720
         Width           =   1335
      End
      Begin VB.OptionButton optReflect 
         Caption         =   "None"
         Height          =   195
         Index           =   0
         Left            =   2280
         TabIndex        =   9
         Top             =   480
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         ItemData        =   "DamageCalc.frx":003A
         Left            =   120
         List            =   "DamageCalc.frx":004D
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1200
         Width           =   1695
      End
      Begin VB.CheckBox chkDamageCalc 
         Caption         =   "Critical Hit"
         Height          =   195
         Index           =   2
         Left            =   2280
         TabIndex        =   12
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label lblDamageCalc 
         BackStyle       =   0  'Transparent
         Caption         =   "Defense/Sp. Def"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblDamageCalc 
         BackStyle       =   0  'Transparent
         Caption         =   "Light Screen/Reflect"
         Height          =   255
         Index           =   11
         Left            =   2280
         TabIndex        =   26
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label lblDamageCalc 
         BackStyle       =   0  'Transparent
         Caption         =   "Type Effectiveness"
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   25
         Top             =   960
         Width           =   1935
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Results"
      Height          =   3615
      Left            =   4800
      TabIndex        =   2
      Top             =   0
      Width           =   3855
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   3360
         Top             =   3120
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   13
         Top             =   600
         Width           =   1815
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   2895
         Left            =   120
         ScaleHeight     =   2895
         ScaleWidth      =   3615
         TabIndex        =   22
         Top             =   600
         Visible         =   0   'False
         Width           =   3615
         Begin VB.PictureBox Picture2 
            Height          =   240
            Left            =   3000
            ScaleHeight     =   180
            ScaleWidth      =   315
            TabIndex        =   38
            Top             =   1320
            Width           =   375
            Begin VB.Label lblBattleMod 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "0"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   -30
               TabIndex        =   39
               Top             =   0
               Width           =   375
            End
         End
         Begin VB.TextBox txtDamageCalc 
            Height          =   285
            Index           =   4
            Left            =   3000
            MaxLength       =   3
            TabIndex        =   14
            Text            =   "100"
            Top             =   0
            Width           =   495
         End
         Begin VB.PictureBox Picture3 
            BorderStyle     =   0  'None
            Height          =   735
            Left            =   1080
            ScaleHeight     =   735
            ScaleWidth      =   1575
            TabIndex        =   33
            Top             =   1080
            Width           =   1575
            Begin VB.OptionButton Option2 
               Caption         =   "Nature Boost"
               Height          =   255
               Index           =   0
               Left            =   0
               TabIndex        =   17
               Top             =   0
               Width           =   1455
            End
            Begin VB.OptionButton Option2 
               Caption         =   "Nature Neutral"
               Height          =   255
               Index           =   1
               Left            =   0
               TabIndex        =   18
               Top             =   240
               Value           =   -1  'True
               Width           =   1455
            End
            Begin VB.OptionButton Option2 
               Caption         =   "Nature Reduction"
               Height          =   255
               Index           =   2
               Left            =   0
               TabIndex        =   19
               Top             =   480
               Width           =   1575
            End
         End
         Begin NetBattle.ColorProgress DemoBar 
            Height          =   375
            Left            =   0
            TabIndex        =   34
            Top             =   2520
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   661
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Defense"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   15
            Top             =   1080
            Value           =   -1  'True
            Width           =   1455
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Sp. Def"
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   16
            Top             =   1320
            Width           =   1455
         End
         Begin VB.HScrollBar HScroll1 
            Height          =   240
            Left            =   2760
            Max             =   6
            Min             =   -6
            TabIndex        =   40
            Top             =   1320
            Width           =   855
         End
         Begin VB.PictureBox Picture4 
            BorderStyle     =   0  'None
            Height          =   615
            Left            =   0
            ScaleHeight     =   615
            ScaleWidth      =   3495
            TabIndex        =   42
            Top             =   360
            Width           =   3495
            Begin MSComctlLib.Slider Slider1 
               Height          =   255
               Left            =   720
               TabIndex        =   43
               Top             =   360
               Width           =   2295
               _ExtentX        =   4048
               _ExtentY        =   450
               _Version        =   393216
               LargeChange     =   16
               Max             =   255
               SelStart        =   255
               TickFrequency   =   16
               Value           =   255
            End
            Begin MSComctlLib.Slider Slider2 
               Height          =   255
               Left            =   720
               TabIndex        =   44
               Top             =   0
               Width           =   2295
               _ExtentX        =   4048
               _ExtentY        =   450
               _Version        =   393216
               LargeChange     =   16
               Max             =   255
               SelStart        =   255
               TickFrequency   =   16
               Value           =   255
            End
            Begin VB.Label lblDamageCalc 
               BackStyle       =   0  'Transparent
               Caption         =   "255"
               Height          =   255
               Index           =   9
               Left            =   3120
               TabIndex        =   48
               Top             =   360
               Width           =   975
            End
            Begin VB.Label lblDamageCalc 
               BackStyle       =   0  'Transparent
               Caption         =   "255"
               Height          =   255
               Index           =   8
               Left            =   3120
               TabIndex        =   47
               Top             =   0
               Width           =   975
            End
            Begin VB.Label lblDamageCalc 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Def EV"
               Height          =   255
               Index           =   3
               Left            =   -240
               TabIndex        =   46
               Top             =   360
               Width           =   855
            End
            Begin VB.Label lblDamageCalc 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "HP EV"
               Height          =   255
               Index           =   2
               Left            =   -240
               TabIndex        =   45
               Top             =   0
               Width           =   855
            End
         End
         Begin VB.Label lblDamageCalc 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Battle Mod"
            Height          =   255
            Index           =   14
            Left            =   2640
            TabIndex        =   41
            Top             =   1080
            Width           =   1095
         End
         Begin VB.Image Image1 
            Height          =   960
            Left            =   2640
            Tag             =   "0"
            Top             =   1920
            Width           =   960
         End
         Begin VB.Label lblDamageCalc 
            BackStyle       =   0  'Transparent
            Caption         =   "Level"
            Height          =   255
            Index           =   4
            Left            =   2520
            TabIndex        =   36
            Top             =   0
            Width           =   975
         End
         Begin VB.Label lblDamageCalc 
            BackStyle       =   0  'Transparent
            Height          =   255
            Index           =   7
            Left            =   0
            TabIndex        =   35
            Top             =   2280
            Width           =   2535
         End
      End
      Begin VB.Label lblDamageCalc 
         BackStyle       =   0  'Transparent
         Caption         =   "Select a Pokémon from the list to see a simulation."
         Height          =   495
         Index           =   10
         Left            =   120
         TabIndex        =   37
         Top             =   1800
         Width           =   3615
      End
      Begin VB.Label Label6 
         Caption         =   "Damage:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   3615
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileItem 
         Caption         =   "E&xit"
         Index           =   0
      End
   End
   Begin VB.Menu mnuVersion 
      Caption         =   "&Version"
      Begin VB.Menu mnuVersionItem 
         Caption         =   "&Red/Blue/Yellow"
         Index           =   0
      End
      Begin VB.Menu mnuVersionItem 
         Caption         =   "&Gold/Silver/Crystal"
         Index           =   1
      End
      Begin VB.Menu mnuVersionItem 
         Caption         =   "Ruby/&Sapphire"
         Checked         =   -1  'True
         Index           =   2
      End
   End
End
Attribute VB_Name = "DamageCalc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Cycle As Boolean
Dim Min As Long
Dim Max As Long
Dim Mode As Byte

Private Sub chkDamageCalc_Click(Index As Integer)
    Call DoCalc
End Sub

Private Sub Combo1_Click()
    Dim X As Integer
    X = GetPokeNum(Combo1.List(Combo1.ListIndex))
    If X = Image1.Tag Then Exit Sub
    Image1.Tag = X
    Call MainContainer.DoPicture(ChooseImage(BasePKMN(X), 3))
    Image1.Picture = MainContainer.SwapSpace
    DemoBar.Max = GetAdvHP(BasePKMN(X).BaseHP, 31, Slider2.Value, Val(txtDamageCalc(4).Text))
    DemoBar.Value = DemoBar.Max
    DemoBar.RefreshBar
    Call UpdateDefense
    Picture1.Visible = True
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    Dim X As Integer
    Dim Y As Integer
    Dim F As Integer
    Dim B As Boolean
    Dim Temp As String
    If KeyAscii = 8 Then Exit Sub
    If KeyAscii = 13 Then
        KeyAscii = 0
        Combo1.SelStart = Len(Combo1.Text)
        Exit Sub
    End If
    Temp = FutureText(Combo1, KeyAscii)
    If Temp = "" Then Exit Sub
    KeyAscii = 0
    B = False
    With Combo1
        Y = Len(Temp)
        For X = 0 To .ListCount - 1
            If LCase(Left(.List(X), Y)) = LCase(Temp) Then
                .ListIndex = X
                Call Combo1_Click
                .Text = .List(X)
                .ListIndex = X
                .SelStart = Y
                .SelLength = Len(.List(X)) - Y
                B = True
                Exit For
            End If
        Next X
        If Not B Then
            X = .SelStart + 1
            .Text = Temp
            .SelStart = X
        End If
    End With
End Sub

Private Sub Combo2_Click()
    Dim X As Integer
    X = GetMoveNum(Combo2.List(Combo2.ListIndex))
    txtDamageCalc(3).Text = Moves(X).Power
    Select Case Moves(X).Type
    Case 2 To 6, 11, 15, 16
        Option1(1).Value = True
    Case Else
        Option1(0).Value = True
    End Select
    Call DoCalc
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
    Dim X As Integer
    Dim Y As Integer
    Dim F As Integer
    Dim B As Boolean
    Dim Temp As String
    If KeyAscii = 8 Then Exit Sub
    If KeyAscii = 13 Then
        KeyAscii = 0
        Combo2.SelStart = Len(Combo2.Text)
        Exit Sub
    End If
    Temp = FutureText(Combo2, KeyAscii)
    If Temp = "" Then Exit Sub
    KeyAscii = 0
    B = False
    With Combo2
        Y = Len(Temp)
        For X = 0 To .ListCount - 1
            If LCase(Left(.List(X), Y)) = LCase(Temp) Then
                .ListIndex = X
                Call Combo2_Click
                .Text = .List(X)
                .ListIndex = X
                .SelStart = Y
                .SelLength = Len(.List(X)) - Y
                B = True
                Exit For
            End If
        Next X
        If Not B Then
            X = .SelStart + 1
            .Text = Temp
            .SelStart = X
        End If
    End With
End Sub

Private Sub Combo3_Click()
    Call DoCalc
End Sub

Private Sub Combo4_Click()
    Call DoCalc
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim X As Integer
    Dim Y As Boolean
    Call CenterWindow(Me)
    Mode = 2
    Combo3.ListIndex = 2
    Combo4.ListIndex = 1
    Combo2.Visible = False
    For X = 1 To 354
        If Moves(X).Power > 0 Then Combo2.AddItem Moves(X).Name
    Next X
    Combo2.Visible = True
    Combo1.Visible = True
    For X = 1 To 386
        Combo1.AddItem BasePKMN(X).Name
    Next X
    Combo1.Visible = True
    DemoBar.Caption = nbExact
End Sub

Private Sub HScroll1_Change()
    lblBattleMod.Caption = IIf(HScroll1.Value > 0, "+", "") & CStr(HScroll1.Value)
    Call UpdateDefense
End Sub

Private Sub mnuFileItem_Click(Index As Integer)
    Unload Me
End Sub

Private Sub mnuVersionItem_Click(Index As Integer)
    Dim X As Integer
    Dim Y As Integer
    Dim Temp As String
    If Mode = Index Then Exit Sub
    mnuVersionItem(Mode).Checked = False
    mnuVersionItem(Index).Checked = True
    Mode = Index
    chkDamageCalc(1).Enabled = (Mode > 0)
    lblDamageCalc(13).Enabled = (Mode > 0)
    Combo4.Enabled = (Mode > 0)
    chkDamageCalc(3).Enabled = (Mode = 2)
    optReflect(2).Enabled = (Mode = 2)
    Slider1.Enabled = (Mode = 2)
    Slider2.Enabled = (Mode = 2)
    Option2(0).Enabled = (Mode = 2)
    Option2(1).Enabled = (Mode = 2)
    Option2(2).Enabled = (Mode = 2)
    Option1(1).Caption = IIf(Mode = 0, "Special", "Sp. Def")
    Select Case Mode
    Case 0: X = 151
    Case 1: X = 251
    Case 2: X = 386
    End Select
    If Mode < 2 Then
        Option2(1).Value = True
        If optReflect(2).Value Then optReflect(1).Value = True
        chkDamageCalc(3).Value = 0
    End If
    If Mode < 1 Then
        chkDamageCalc(1).Value = 0
        Combo4.ListIndex = 1
    End If
    Combo2.Visible = False
    Temp = Combo2.List(Combo2.ListIndex)
    Combo2.Clear
    For X = 1 To 354
        If Moves(X).Power > 0 Then
            Y = 0
            Select Case Mode
            Case 0: If Moves(X).RBYMove Then Y = 1
            Case 1: If Moves(X).GSCMove Then Y = 1
            Case 2: If Moves(X).AdvMove Then Y = 1
            End Select
            If Y = 1 Then
                Combo2.AddItem Moves(X).Name
                If Moves(X).Name = Temp Then
                    For Y = 0 To Combo2.ListCount - 1
                        If Combo2.List(Y) = Temp Then Combo2.ListIndex = Y
                    Next Y
                End If
            End If
        End If
    Next X
    Combo2.Visible = True
    Select Case Mode
    Case 0: X = 151
    Case 1: X = 251
    Case 2: X = 386
    End Select
    Combo1.Visible = True
    Temp = Combo1.List(Combo1.ListIndex)
    Combo1.Clear
    For X = 1 To X
        Combo1.AddItem BasePKMN(X).Name
        If BasePKMN(X).Name = Temp Then
            For Y = 0 To Combo2.ListCount - 1
                If Combo1.List(Y) = Temp Then Combo1.ListIndex = Y
            Next Y
        End If
    Next X
    If Combo1.ListIndex = -1 Then Combo1.ListIndex = 0
    Combo1.Visible = True
    Call Slider2_Click
    Call UpdateDefense
    Call DoCalc
End Sub

Private Sub Option1_Click(Index As Integer)
    Call UpdateDefense
End Sub

Private Sub Option2_Click(Index As Integer)
    Call UpdateDefense
End Sub

Private Sub optReflect_Click(Index As Integer)
    Call DoCalc
End Sub

Private Sub Slider1_Click()
    Call UpdateDefense
    lblDamageCalc(9).Caption = Slider1.Value
End Sub

Private Sub Slider1_Scroll()
    Call Slider1_Click
End Sub

Private Sub Slider2_Click()
    lblDamageCalc(8).Caption = Slider2.Value
    If Mode = 2 Then
        DemoBar.Max = GetAdvHP(BasePKMN(Image1.Tag).BaseHP, 31, Slider2.Value, Val(txtDamageCalc(4).Text))
    Else
        DemoBar.Max = GetHP(Val(txtDamageCalc(4).Text), BasePKMN(Image1.Tag).BaseHP, 15)
    End If
    DemoBar.RefreshBar
    Call UpdateDefense
End Sub

Private Sub Slider2_Scroll()
    Call Slider2_Click
End Sub

Private Sub Timer1_Timer()
    Cycle = Not Cycle
    DemoBar.Value = DemoBar.Max - Cap(IIf(Cycle, Min, Max), DemoBar.Max)
End Sub

Private Sub txtDamageCalc_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call txtDamageCalc_LostFocus(Index)
        KeyAscii = 0
    End If
End Sub

Private Sub txtDamageCalc_LostFocus(Index As Integer)
    Dim X As Long
    X = Val(txtDamageCalc(Index).Text)
    If X < 1 And Index <> 3 Then X = 1
    If X < 0 And Index = 3 Then X = 0
    If (Index = 0 Or Index = 4) And X > 100 Then X = 100
    txtDamageCalc(Index).Text = CStr(X)
    If Index <> 2 Then Call Slider2_Click
    Call DoCalc
End Sub
Private Sub UpdateDefense()
    Dim Y As Integer
    Dim Z As Integer
    lblDamageCalc(9).Caption = IIf(Slider1.Enabled, Slider1.Value, "---")
    lblDamageCalc(8).Caption = IIf(Slider2.Enabled, Slider2.Value, "---")
    If Image1.Tag = 0 Then Exit Sub
    If Option1(0).Value Then Y = BasePKMN(Image1.Tag).BaseDefense Else Y = BasePKMN(Image1.Tag).BaseSDefense
    If Option2(0).Value Then Z = 1
    If Option2(2).Value Then Z = -1
    If Mode = 2 Then
        Y = GetAdvStat(Y, 31, Slider1.Value, Val(txtDamageCalc(4).Text), Z)
    Else
        Y = GetStat(Val(txtDamageCalc(4).Text), Y, 15)
    End If
    txtDamageCalc(2).Text = CStr(Int(Y * StatChange(HScroll1.Value)))
    Call DoCalc
End Sub
Private Sub DoCalc()
    Dim ATK As Long
    Dim DEF As Long
    Dim Lev As Long
    Dim POW As Single
    Dim TMATCH As Single
    Dim STAB As Single
    Dim RAND As Integer
    Dim DamageTemp As Long
    Dim DT2 As Long
    Dim X As Long
    Dim Temp As String
    If Val(txtDamageCalc(3).Text) = 0 Then Exit Sub
    For X = 1 To 2
        Lev = Val(txtDamageCalc(0).Text)
        ATK = Val(txtDamageCalc(1).Text)
        DEF = Val(txtDamageCalc(2).Text)
        POW = Val(txtDamageCalc(3).Text)
        STAB = IIf(chkDamageCalc(0).Value = 1, 1.5, 1)
        If chkDamageCalc(1).Value = 1 Then POW = POW * 1.1
        If chkDamageCalc(2).Value = 1 Then Lev = Lev * 2
        If Combo4.ListIndex = 0 Then POW = POW * 1.5
        If Combo4.ListIndex = 2 Then POW = POW \ 2
        Select Case Combo3.ListIndex
        Case 0: TMATCH = 4
        Case 1: TMATCH = 2
        Case 2: TMATCH = 1
        Case 3: TMATCH = 0.5
        Case 4: TMATCH = 0.25
        End Select
        RAND = IIf(X = 1, 217, 255)
        DamageTemp = Int(((((((((2 * Lev) \ 5 + 2) * ATK * POW) \ DEF) / 50) + 2) * STAB) * TMATCH) * RAND) \ 255
        DT2 = ((((((((2 * Lev / 5 + 2) * ATK * POW) / DEF) / 50) + 2) * STAB) * TMATCH) * RAND) / 255
        If Abs(DT2 - DamageTemp) > 2 Then Stop
        If chkDamageCalc(3).Value = 1 Then DamageTemp = DamageTemp * 1.5
        If optReflect(1).Value Then DamageTemp = DamageTemp \ 2
        If optReflect(2).Value Then DamageTemp = Int(DamageTemp / 3 * 2)
        If DamageTemp = 0 Then DamageTemp = 1
        If X = 1 Then Min = DamageTemp Else Max = DamageTemp
    Next X

    Label6.Caption = "Damage: " & Min & " ~ " & Max
    If Image1.Tag > 0 Then
        If Min >= DemoBar.Max Then
            lblDamageCalc(7).Caption = "100% Damage"
        Else
            X = Round((Min / DemoBar.Max) * 100)
            If X = 100 Then X = 99
            Temp = "Between " & X
            X = Round((Cap(Max, DemoBar.Max) / DemoBar.Max) * 100)
            If Max < DemoBar.Max And X = 100 Then X = 99
            lblDamageCalc(7).Caption = Temp & "% and " & X & "% Damage"
        End If
        Cycle = False
        Call Timer1_Timer
        Timer1.Enabled = False
        Timer1.Enabled = True
    End If
        
End Sub
