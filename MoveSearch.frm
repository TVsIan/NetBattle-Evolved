VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form MoveSearch 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Search Moves"
   ClientHeight    =   4995
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   3975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox MoveList 
      Height          =   1620
      Left            =   120
      TabIndex        =   22
      Top             =   3240
      Width           =   3735
   End
   Begin VB.CheckBox SCheck 
      Caption         =   "Substitute Blocks"
      Height          =   255
      Index           =   8
      Left            =   2040
      TabIndex        =   8
      Top             =   1800
      Width           =   1695
   End
   Begin VB.CheckBox SCheck 
      Caption         =   "Self-Affecting"
      Height          =   255
      Index           =   7
      Left            =   240
      TabIndex        =   7
      Top             =   1800
      Width           =   1695
   End
   Begin VB.CheckBox SCheck 
      Caption         =   "King's Rock"
      Height          =   255
      Index           =   6
      Left            =   2040
      TabIndex        =   6
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Search Options"
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3735
      Begin VB.TextBox PowBox 
         Height          =   285
         Left            =   480
         TabIndex        =   15
         Top             =   1080
         Width           =   720
      End
      Begin VB.TextBox AccBox 
         Height          =   285
         Left            =   1440
         TabIndex        =   14
         Top             =   1080
         Width           =   720
      End
      Begin VB.TextBox PPBox 
         Height          =   285
         Left            =   2400
         TabIndex        =   13
         Top             =   1080
         Width           =   720
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   120
         ScaleHeight     =   495
         ScaleWidth      =   3495
         TabIndex        =   19
         Top             =   2400
         Width           =   3495
         Begin VB.CommandButton SearchButton 
            Caption         =   "&Search"
            Default         =   -1  'True
            Height          =   375
            Left            =   120
            TabIndex        =   21
            Top             =   120
            Width           =   1455
         End
         Begin VB.CommandButton CancelButton 
            Cancel          =   -1  'True
            Caption         =   "&Cancel"
            Height          =   375
            Left            =   1920
            TabIndex        =   20
            Top             =   120
            Width           =   1455
         End
      End
      Begin VB.CheckBox SCheck 
         Caption         =   "Power"
         Height          =   255
         Index           =   2
         Left            =   480
         TabIndex        =   18
         Top             =   840
         Width           =   855
      End
      Begin VB.CheckBox SCheck 
         Caption         =   "Acc."
         Height          =   255
         Index           =   3
         Left            =   1440
         TabIndex        =   17
         Top             =   840
         Width           =   735
      End
      Begin VB.CheckBox SCheck 
         Caption         =   "PP"
         Height          =   255
         Index           =   4
         Left            =   2400
         TabIndex        =   16
         Top             =   840
         Width           =   735
      End
      Begin VB.CheckBox SCheck 
         Caption         =   "Hits All"
         Height          =   255
         Index           =   12
         Left            =   1920
         TabIndex        =   12
         Top             =   2160
         Width           =   1695
      End
      Begin VB.CheckBox SCheck 
         Caption         =   "Contact Move"
         Height          =   255
         Index           =   10
         Left            =   1920
         TabIndex        =   11
         Top             =   1920
         Width           =   1695
      End
      Begin VB.CheckBox SCheck 
         Caption         =   "Hits Both"
         Height          =   255
         Index           =   11
         Left            =   120
         TabIndex        =   10
         Top             =   2160
         Width           =   1695
      End
      Begin VB.CheckBox SCheck 
         Caption         =   "Sound Move"
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   9
         Top             =   1920
         Width           =   1695
      End
      Begin VB.CheckBox SCheck 
         Caption         =   "Bright Powder"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   5
         Top             =   1440
         Width           =   1695
      End
      Begin MSComctlLib.ImageCombo Type1 
         Height          =   330
         Left            =   1920
         TabIndex        =   3
         Top             =   480
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
      End
      Begin VB.CheckBox SCheck 
         Caption         =   "Type"
         Height          =   255
         Index           =   1
         Left            =   1920
         TabIndex        =   4
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   1695
      End
      Begin VB.CheckBox SCheck 
         Caption         =   "Name"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1575
      End
   End
End
Attribute VB_Name = "MoveSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim FoundMoves() As Boolean

Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim X As Integer
    
    Type1.ImageList = MainContainer.Types
    For X = 1 To 17
        Type1.ComboItems.Add X, , Element(X), X
    Next
End Sub

Private Sub MoveList_Click()
    If MoveList.ListCount = 0 Then Exit Sub
    If MoveList.ListIndex < 0 Then Exit Sub
    Call MoveDex.ChangeMe(MoveList.List(MoveList.ListIndex))
End Sub

Private Sub SearchButton_Click()
    Dim X As Integer
    Dim SearchType As Integer
    
    MoveList.Clear
    ReDim FoundMoves(UBound(Moves)) As Boolean
    For X = 1 To UBound(Moves)
        FoundMoves(X) = True
    Next
    If SCheck(0).Value = 1 Then
        For X = 1 To UBound(Moves)
            If InStr(1, UCase(Moves(X).Name), UCase(Text1.Text)) = 0 Then FoundMoves(X) = False
        Next
    End If
    If SCheck(1).Value = 1 Then
        For X = 1 To 17
            If Type1.ComboItems(X).Selected Then SearchType = X
        Next
        For X = 1 To UBound(Moves)
            If Moves(X).Type <> SearchType Then FoundMoves(X) = False
        Next
    End If
    If SCheck(2).Value = 1 Then
        For X = 1 To UBound(Moves)
            If Moves(X).Power < Val(PowBox.Text) Then FoundMoves(X) = False
        Next
    End If
    If SCheck(3).Value = 1 Then
        For X = 1 To UBound(Moves)
            If Val(AccBox.Text) = 100 Then
                If Moves(X).Accuracy < 100 And Moves(X).Accuracy <> 0 Then FoundMoves(X) = False
            Else
                If Moves(X).Accuracy < Val(AccBox.Text) Then FoundMoves(X) = False
            End If
        Next
    End If
    If SCheck(4).Value = 1 Then
        For X = 1 To UBound(Moves)
            If Moves(X).PP < Val(PPBox.Text) Then FoundMoves(X) = False
        Next
    End If
    If SCheck(5).Value = 1 Then
        For X = 1 To UBound(Moves)
            If Not Moves(X).BrightPowder Then FoundMoves(X) = False
        Next
    End If
    If SCheck(6).Value = 1 Then
        For X = 1 To UBound(Moves)
            If Not Moves(X).KingsRock Then FoundMoves(X) = False
        Next
    End If
    If SCheck(7).Value = 1 Then
        For X = 1 To UBound(Moves)
            If Not Moves(X).SelfMove Then FoundMoves(X) = False
        Next
    End If
    If SCheck(8).Value = 1 Then
        For X = 1 To UBound(Moves)
            If Not Moves(X).SubstituteBlocks Then FoundMoves(X) = False
        Next
    End If
    If SCheck(9).Value = 1 Then
        For X = 1 To UBound(Moves)
            If Not Moves(X).SoundMove Then FoundMoves(X) = False
        Next
    End If
    If SCheck(10).Value = 1 Then
        For X = 1 To UBound(Moves)
            If Not Moves(X).PhysMove Then FoundMoves(X) = False
        Next
    End If
    If SCheck(11).Value = 1 Then
        For X = 1 To UBound(Moves)
            If Not Moves(X).HitsAll Then FoundMoves(X) = False
        Next
    End If
    If SCheck(12).Value = 1 Then
        For X = 1 To UBound(Moves)
            If Not Moves(X).HitsTeam Then FoundMoves(X) = False
        Next
    End If
    For X = 1 To UBound(Moves)
        If FoundMoves(X) Then MoveList.AddItem Moves(X).Name
    Next
End Sub
