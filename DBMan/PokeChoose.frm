VERSION 5.00
Begin VB.Form PokeChoose 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Jump"
   ClientHeight    =   1455
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   2520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   2520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   960
      Width           =   1095
   End
   Begin VB.ComboBox PokeCombo 
      Height          =   315
      ItemData        =   "PokeChoose.frx":0000
      Left            =   120
      List            =   "PokeChoose.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   480
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Choose a Pokemon to jump to."
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "PokeChoose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim X As Integer
    For X = 1 To UBound(BasePKMN)
        If PokeCombo.Text = BasePKMN(X).Name & " - " & X Then Exit For
    Next X
    
    Me.Hide
    EditHelp.Enabled = True
    EditHelp.SetFocus
    If X < 1 Or X > UBound(BasePKMN) Then Exit Sub
    Call EditHelp.SetPoke(X)
End Sub

Private Sub Command2_Click()
    Me.Hide
    EditHelp.Enabled = True
    EditHelp.SetFocus
End Sub

