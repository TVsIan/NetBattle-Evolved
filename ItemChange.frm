VERSION 5.00
Begin VB.Form ItemChange 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Change Items"
   ClientHeight    =   2535
   ClientLeft      =   2760
   ClientTop       =   3630
   ClientWidth     =   4935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox ItemPick 
      Height          =   315
      Index           =   5
      Left            =   2520
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   1560
      Width           =   2295
   End
   Begin VB.ComboBox ItemPick 
      Height          =   315
      Index           =   4
      Left            =   2520
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   960
      Width           =   2295
   End
   Begin VB.ComboBox ItemPick 
      Height          =   315
      Index           =   3
      Left            =   2520
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   360
      Width           =   2295
   End
   Begin VB.ComboBox ItemPick 
      Height          =   315
      Index           =   2
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1560
      Width           =   2295
   End
   Begin VB.ComboBox ItemPick 
      Height          =   315
      Index           =   1
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   960
      Width           =   2295
   End
   Begin VB.ComboBox ItemPick 
      Height          =   315
      Index           =   0
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   360
      Width           =   2295
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2280
      TabIndex        =   0
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label PokeName 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      Height          =   255
      Index           =   5
      Left            =   2520
      TabIndex        =   12
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Label PokeName 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      Height          =   255
      Index           =   4
      Left            =   2520
      TabIndex        =   10
      Top             =   720
      Width           =   2175
   End
   Begin VB.Label PokeName 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      Height          =   255
      Index           =   3
      Left            =   2520
      TabIndex        =   8
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label PokeName 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Label PokeName 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   2175
   End
   Begin VB.Label PokeName 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "ItemChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim X As Long
    Dim Y As Byte
    Dim Temp As String
    Dim TItem() As String
    Select Case CompatVersion(PKMN(1).GameVersion)
    Case nbRBYBattle
        ReDim TItem(0)
        TItem(0) = "N/A"
    Case nbGSCBattle
        TItem = Item
        ReDim Preserve TItem(41)
        TItem(0) = ""
        Call SortStringArray(TItem)
        TItem(0) = Item(0)
    Case nbAdvBattle
        Y = 1
        For X = 1 To UBound(Item)
            If Item(X) <> "" Then
                If AdvItem(X) Then
                    ReDim Preserve TItem(Y)
                    TItem(Y) = Item(X)
                    Y = Y + 1
                End If
            End If
        Next X
        TItem(0) = ""
        Call SortStringArray(TItem)
        TItem(0) = Item(0)
    End Select
    
    For Y = 0 To UBound(TItem)
        For X = 0 To 5
            ItemPick(X).AddItem TItem(Y)
            If Item(PKMN(X + 1).Item) = TItem(Y) Then ItemPick(X).ListIndex = Y
            PokeName(X) = PKMN(X + 1).Name & IIf(PKMN(X + 1).Nickname <> "", " (" & PKMN(X + 1).Nickname & ")", "")
        Next X
    Next Y
        
End Sub

Private Sub OKButton_Click()
    Dim X As Byte
    Dim Y As Long
    Dim Changed As Boolean
    For X = 0 To 5
        For Y = 1 To UBound(Item)
            If Item(Y) = ItemPick(X).List(ItemPick(X).ListIndex) Then Exit For
        Next Y
        If PKMN(X + 1).Item <> Y Then
            Changed = True
            PKMN(X + 1).Item = Y
        End If
    Next X
    If TeamChangeFromMS And Changed Then MasterServer.TeamChanged = True
    Unload Me
End Sub
