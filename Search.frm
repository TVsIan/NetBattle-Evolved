VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Search 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Search Pokedex"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6315
   Icon            =   "Search.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   6315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Search Results"
      Height          =   3135
      Left            =   120
      TabIndex        =   16
      Top             =   3000
      Width           =   6135
      Begin MSComctlLib.ListView Results 
         Height          =   2775
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   4895
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Name"
            Object.Width           =   1834
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Type1"
            Object.Width           =   1296
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Type2"
            Object.Width           =   1296
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "HP"
            Object.Width           =   873
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Atk"
            Object.Width           =   873
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "Def"
            Object.Width           =   873
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Text            =   "Spd"
            Object.Width           =   873
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   7
            Text            =   "SAtk"
            Object.Width           =   953
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   8
            Text            =   "SDef"
            Object.Width           =   953
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Search Options"
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6135
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   2535
         Left            =   120
         ScaleHeight     =   2535
         ScaleWidth      =   5895
         TabIndex        =   1
         Top             =   240
         Width           =   5895
         Begin MSComctlLib.ListView MoveList 
            Height          =   1815
            Left            =   3600
            TabIndex        =   15
            Top             =   240
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   3201
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            HideColumnHeaders=   -1  'True
            Checkboxes      =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Move"
               Object.Width           =   5080
            EndProperty
         End
         Begin VB.ComboBox cmbSearch 
            Height          =   315
            Index           =   3
            Left            =   1800
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   240
            Width           =   1695
         End
         Begin VB.TextBox NameBox 
            Height          =   315
            Left            =   0
            TabIndex        =   19
            Text            =   "Text1"
            Top             =   240
            Width           =   1695
         End
         Begin VB.ComboBox cmbSearch 
            Height          =   315
            Index           =   1
            Left            =   0
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   960
            Width           =   1695
         End
         Begin VB.ComboBox cmbSearch 
            Height          =   315
            Index           =   4
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   960
            Width           =   1695
         End
         Begin VB.ComboBox cmbSearch 
            Height          =   315
            Index           =   5
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   1680
            Width           =   1695
         End
         Begin VB.ComboBox cmbSearch 
            Height          =   315
            Index           =   2
            Left            =   0
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   1680
            Width           =   1695
         End
         Begin VB.CheckBox SCheck 
            Caption         =   "Name"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   18
            Top             =   0
            Width           =   975
         End
         Begin VB.CommandButton CancelButton 
            Cancel          =   -1  'True
            Caption         =   "&Exit"
            Height          =   375
            Left            =   4800
            TabIndex        =   14
            Top             =   2160
            Width           =   1095
         End
         Begin VB.CommandButton SearchButton 
            Caption         =   "&Search"
            Default         =   -1  'True
            Enabled         =   0   'False
            Height          =   375
            Left            =   3600
            TabIndex        =   13
            Top             =   2160
            Width           =   1095
         End
         Begin VB.CheckBox SCheck 
            Caption         =   "Moves"
            Height          =   255
            Index           =   9
            Left            =   3600
            TabIndex        =   12
            Top             =   0
            Width           =   975
         End
         Begin VB.CheckBox SCheck 
            Caption         =   "2nd Type"
            Height          =   255
            Index           =   2
            Left            =   0
            TabIndex        =   11
            Top             =   1440
            Width           =   975
         End
         Begin VB.CheckBox SCheck 
            Caption         =   "Type"
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   10
            Top             =   720
            Width           =   975
         End
         Begin VB.CheckBox SCheck 
            Caption         =   "Trait"
            Height          =   255
            Index           =   3
            Left            =   1800
            TabIndex        =   8
            Top             =   0
            Width           =   975
         End
         Begin VB.CheckBox SCheck 
            Caption         =   "Egg Group"
            Height          =   255
            Index           =   4
            Left            =   1800
            TabIndex        =   3
            Top             =   720
            Width           =   1215
         End
         Begin VB.CheckBox SCheck 
            Caption         =   "2nd Egg Group"
            Height          =   255
            Index           =   5
            Left            =   1800
            TabIndex        =   2
            Top             =   1440
            Width           =   1575
         End
      End
   End
End
Attribute VB_Name = "Search"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SearchMatches() As Boolean
Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim X As Integer
    
    ReDim SearchMatches(UBound(BasePKMN)) As Boolean
    If BasePKMN(1).No <> 1 Then PokeLoader.Show
    
    cmbSearch(2).AddItem "(None)", 0
    For X = 1 To 17
        cmbSearch(1).AddItem Element(X), X - 1
        cmbSearch(2).AddItem Element(X), X
    Next
    MoveList.SmallIcons = MainContainer.Types
    MoveList.Icons = MainContainer.Types
    For X = 1 To UBound(AttributeText)
        cmbSearch(3).AddItem AttributeText(X)
    Next X
    Call MasterDex.FillSearchEGs
    Call RefreshMode
    Call CopyMoveList
    Call Reset
End Sub

Public Sub DoSearch()
    Dim X As Integer
    Dim Y As Integer
    Dim Z As Integer
    Dim HasMove As Boolean
    Dim TypeSearch As Integer
    Dim MoveSearch(4) As Integer
    Dim NumberString As String
    Dim Index As Integer
    Dim TempItem As ListItem
    Dim TempMove() As Integer
    Dim TempSource() As String
    
    Me.MousePointer = vbHourglass
    Select Case MasterDex.CurrentMode
    Case 0: Y = 151
    Case 1: Y = 251
    Case 2: Y = UBound(BasePKMN)
    End Select
    For X = 1 To UBound(SearchMatches)
        SearchMatches(X) = (X <= Y)
    Next X
    
    If SCheck(0).Value = 1 Then
        For X = 1 To UBound(BasePKMN)
            If InStr(1, UCase(BasePKMN(X).Name), UCase(NameBox.Text)) = 0 Then SearchMatches(X) = False
        Next X
    End If
    
    If SCheck(1).Value = 1 Then
        For X = 1 To UBound(BasePKMN)
            If BasePKMN(X).Type1 <> cmbSearch(1).ListIndex + 1 And BasePKMN(X).Type2 <> cmbSearch(1).ListIndex + 1 Then SearchMatches(X) = False
        Next X
    End If
    If SCheck(2).Value = 1 Then
        For X = 1 To UBound(BasePKMN)
            If BasePKMN(X).Type1 <> cmbSearch(2).ListIndex And BasePKMN(X).Type2 <> cmbSearch(2).ListIndex Then SearchMatches(X) = False
        Next X
    End If
    For Y = 1 To UBound(AttributeText)
        If AttributeText(Y) = cmbSearch(3).List(cmbSearch(3).ListIndex) Then Exit For
    Next Y
    If SCheck(3).Value = 1 Then
        For X = 1 To UBound(BasePKMN)
            If BasePKMN(X).PAtt(0) <> Y And BasePKMN(X).PAtt(1) <> Y Then SearchMatches(X) = False
        Next X
    End If
    If SCheck(4).Value = 1 Then
        For X = 1 To UBound(BasePKMN)
            If BasePKMN(X).EggGroup1 <> cmbSearch(4).ListIndex + 1 And BasePKMN(X).EggGroup2 <> cmbSearch(4).ListIndex + 1 Then SearchMatches(X) = False
        Next X
    End If
    If SCheck(5).Value = 1 Then
        For X = 1 To UBound(BasePKMN)
            If BasePKMN(X).EggGroup1 <> cmbSearch(5).ListIndex And BasePKMN(X).EggGroup2 <> cmbSearch(5).ListIndex Then SearchMatches(X) = False
        Next X
    End If
'    If SCheck(4).Value = 1 Then
'        For X = 1 To UBound(BasePKMN)
'            If BasePKMN(X).Attack < Val(ATKBox) Then SearchMatches(X) = False
'        Next X
'    End If
'    If SCheck(5).Value = 1 Then
'        For X = 1 To UBound(BasePKMN)
'            If BasePKMN(X).Defense < Val(DEFBox) Then SearchMatches(X) = False
'        Next X
'    End If
'    If SCheck(6).Value = 1 Then
'        For X = 1 To UBound(BasePKMN)
'            If BasePKMN(X).Speed < Val(SPDBox) Then SearchMatches(X) = False
'        Next X
'    End If
'    If SCheck(7).Value = 1 Then
'        For X = 1 To UBound(BasePKMN)
'            If BasePKMN(X).SpecialAttack < Val(SATKBox) Then SearchMatches(X) = False
'        Next X
'    End If
'    If SCheck(8).Value = 1 Then
'        For X = 1 To UBound(BasePKMN)
'            If BasePKMN(X).SpecialDefense < Val(SDEFBox) Then SearchMatches(X) = False
'        Next X
'    End If
    If SCheck(9).Value = 1 Then
        Y = 1
        For X = 1 To MoveList.ListItems.count
            If MoveList.ListItems(X).Checked Then
                MoveSearch(Y) = Val(Right(MoveList.ListItems(X).Key, 3))
                Y = Y + 1
            End If
        Next
        Select Case MasterDex.CurrentMode
        Case 0: X = 151
        Case 1: X = 251
        Case 2: X = UBound(BasePKMN)
        End Select
        For X = 1 To X
            Select Case MasterDex.CurrentMode
            Case 0: Z = 0
            Case 1: Z = 1
            Case 2: Z = 3
            End Select
            Call MakeMoveArray(X, Z, TempMove, TempSource)
            For Y = 1 To 4
                If MoveSearch(Y) > 0 Then
                    HasMove = False
                    For Z = 1 To UBound(TempMove)
                        If TempMove(Z) = MoveSearch(Y) Then HasMove = True: Exit For
                    Next Z
                    If HasMove = False Then SearchMatches(X) = False
                End If
            Next
        Next
    End If
        
    Results.ListItems.Clear
    Results.Sorted = False
    Index = 1
    For X = 1 To UBound(SearchMatches)
        If SearchMatches(X) Then
            Set TempItem = Results.ListItems.Add(, "#" & Format(X, "000"), BasePKMN(X).Name)
            'Results.ListItems(Index).SubItems(1) = BasePKMN(X).Name
            TempItem.SubItems(1) = Element(BasePKMN(X).Type1)
            TempItem.SubItems(2) = Element(BasePKMN(X).Type2)
            TempItem.SubItems(3) = BasePKMN(X).BaseHP
            TempItem.SubItems(4) = BasePKMN(X).BaseAttack
            TempItem.SubItems(5) = BasePKMN(X).BaseDefense
            TempItem.SubItems(6) = BasePKMN(X).BaseSpeed
            TempItem.SubItems(7) = BasePKMN(X).BaseSAttack
            TempItem.SubItems(8) = BasePKMN(X).BaseSDefense
        End If
    Next
'    Results.SortKey = 0
'    Results.Sorted = True
    Me.MousePointer = vbNormal
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MasterDex.SearchOpen = False
End Sub

Private Sub MoveList_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Dim CheckedItems As Integer
    Dim CurrentItem As Integer
    Dim X As Integer

    For X = 1 To MoveList.ListItems.count
        If MoveList.ListItems(X).Text = Item Then CurrentItem = X
        If MoveList.ListItems(X).Checked Then CheckedItems = CheckedItems + 1
    Next

    If CheckedItems <= 4 Then
        Exit Sub
    Else
        MsgBox "You can only search for four moves at a time.", vbInformation, "Error"
        MoveList.ListItems(CurrentItem).Checked = False
    End If
End Sub
Public Sub Reset()
    Dim X As Integer
    On Error Resume Next
    For X = 0 To 9
        SCheck(X).Value = 0
        cmbSearch(X).ListIndex = 0
    Next X
    NameBox.Text = ""
    For X = 1 To MoveList.ListItems.count
        MoveList.ListItems(X).Checked = False
    Next X
End Sub
Private Sub Results_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Dim Numbers As Boolean
    Numbers = (ColumnHeader.Index >= 4)
    If Results.SortKey = ColumnHeader.Index - 1 Then
        If Results.SortOrder = lvwAscending Then Results.SortOrder = lvwDescending Else Results.SortOrder = lvwAscending
    Else
        Results.SortKey = ColumnHeader.Index - 1
        Results.SortOrder = IIf(Numbers, lvwDescending, lvwAscending)
    End If
    If Numbers Then
        Call ListViewNumberSort(Results, ColumnHeader.Index)
    Else
        Results.Sorted = True
        Results.Sorted = False
    End If

End Sub

Private Sub Results_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Call MasterDex.DoPoke(Val(Right(Item.Key, 3)))
End Sub

Private Sub SCheck_Click(Index As Integer)
    Dim X As Integer
    Dim Checked As Integer
    On Error Resume Next
    For X = 0 To SCheck.UBound
        Checked = Checked + SCheck(X).Value
    Next
    
    If Checked > 0 Then
        SearchButton.Enabled = True
    Else
        SearchButton.Enabled = False
    End If
End Sub

Private Sub SearchButton_Click()
    Call DoSearch
End Sub
Public Sub CopyMoveList()
    Dim X As Integer
    SetRedraw MoveList.hWnd, False
    MoveList.Visible = False
    MoveList.ListItems.Clear
    For X = 1 To MasterDex.MoveList.ListItems.count
        With MasterDex.MoveList.ListItems(X)
            MoveList.ListItems.Add , .Key, .Text, .Icon, .SmallIcon
        End With
    Next X
    MoveList.Visible = True
    SetRedraw MoveList.hWnd, True
End Sub
Public Sub RefreshMode()
    SCheck(3).Enabled = (MasterDex.CurrentMode = 2)
    cmbSearch(3).Enabled = SCheck(3).Enabled
    If Not SCheck(3).Enabled Then SCheck(3).Value = 0
    'DEBUG: Uncomment once Egg Groups are implemented
'    SCheck(4).Enabled = (MasterDex.CurrentMode > 0)
'    cmbSearch(4).Enabled = SCheck(4).Enabled
'    If Not SCheck(4).Enabled Then SCheck(4).Value = 0
'    SCheck(5).Enabled = (MasterDex.CurrentMode > 0)
'    cmbSearch(5).Enabled = SCheck(5).Enabled
'    If Not SCheck(5).Enabled Then SCheck(5).Value = 0
End Sub
