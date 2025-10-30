VERSION 5.00
Begin VB.Form TempbanSet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "TempBan"
   ClientHeight    =   1260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   1770
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1260
   ScaleWidth      =   1770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   960
      TabIndex        =   4
      Top             =   840
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   480
      MaxLength       =   4
      TabIndex        =   0
      Text            =   "15"
      Top             =   360
      Width           =   495
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   285
      Left            =   240
      Max             =   1440
      Min             =   1
      TabIndex        =   1
      Top             =   360
      Value           =   1
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Enter duration:"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "min."
      Height          =   255
      Left            =   1320
      TabIndex        =   2
      Top             =   420
      Width           =   495
   End
End
Attribute VB_Name = "TempbanSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    TempbanDuration = Val(Text1.Text)
    Unload Me
End Sub

Private Sub Command2_Click()
    TempbanDuration = 0
    Unload Me
End Sub

Private Sub Form_Load()
    Text1.Text = "15"
    Text1.SelLength = 2
    TempbanDuration = 0
    Call Text1_Change
End Sub

Private Sub HScroll1_Change()
    Text1.Text = HScroll1.Value
End Sub

Private Sub Text1_Change()
    Dim X As Long
    X = Val(Text1.Text)
    If X > 1440 Then X = 1440
    If X < 1 Then X = 1
    If Text1.Text <> CStr(X) Then Text1.Text = CStr(X)
    HScroll1.Value = X
End Sub
