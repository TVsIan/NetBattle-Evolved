VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form MoveEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Move Editor"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4335
   Icon            =   "MoveEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   4335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "&Add"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1560
      TabIndex        =   37
      Top             =   6240
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Done"
      Height          =   375
      Left            =   3000
      TabIndex        =   36
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Save Current"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   35
      Top             =   6240
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Machine"
      Height          =   975
      Left            =   840
      TabIndex        =   28
      Top             =   4680
      Width           =   2655
      Begin VB.TextBox Text9 
         DataField       =   "AdvTM"
         DataSource      =   "MoveData"
         Height          =   285
         Left            =   1800
         TabIndex        =   33
         Text            =   "Text2"
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox Text8 
         DataField       =   "GSTM"
         DataSource      =   "MoveData"
         Height          =   285
         Left            =   960
         TabIndex        =   31
         Text            =   "Text2"
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox Text7 
         DataField       =   "RBYTM"
         DataSource      =   "MoveData"
         Height          =   285
         Left            =   120
         TabIndex        =   29
         Text            =   "Text2"
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label11 
         Caption         =   "Advance"
         Height          =   255
         Left            =   1800
         TabIndex        =   34
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label10 
         Caption         =   "GSC"
         Height          =   255
         Left            =   960
         TabIndex        =   32
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label9 
         Caption         =   "RBY"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.CheckBox Check12 
      Caption         =   "Affects User"
      DataField       =   "AffectsSelf"
      DataSource      =   "MoveData"
      Height          =   255
      Left            =   120
      TabIndex        =   27
      Top             =   4320
      Width           =   1815
   End
   Begin VB.CheckBox Check11 
      Caption         =   "Substitute Blocks"
      DataField       =   "BlockSubstitute"
      DataSource      =   "MoveData"
      Height          =   255
      Left            =   2160
      TabIndex        =   26
      Top             =   4320
      Width           =   2055
   End
   Begin VB.CheckBox Check10 
      Caption         =   "Contact Move"
      DataField       =   "PhysMove"
      DataSource      =   "MoveData"
      Height          =   255
      Left            =   2160
      TabIndex        =   25
      Top             =   3960
      Width           =   2055
   End
   Begin VB.CheckBox Check9 
      Caption         =   "Sound Move"
      DataField       =   "SoundMove"
      DataSource      =   "MoveData"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   3960
      Width           =   1815
   End
   Begin VB.CheckBox Check8 
      Caption         =   "Hits all Pokémon"
      DataField       =   "HitsAll"
      DataSource      =   "MoveData"
      Height          =   255
      Left            =   2160
      TabIndex        =   23
      Top             =   3600
      Width           =   2055
   End
   Begin VB.CheckBox Check7 
      Caption         =   "Exists in Advance"
      DataField       =   "AdvCompatible"
      DataSource      =   "MoveData"
      Height          =   255
      Left            =   2160
      TabIndex        =   22
      Top             =   3240
      Width           =   2055
   End
   Begin VB.CheckBox Check6 
      Caption         =   "Exists in RBY"
      DataField       =   "RBYCompatible"
      DataSource      =   "MoveData"
      Height          =   255
      Left            =   2160
      TabIndex        =   21
      Top             =   2880
      Width           =   2055
   End
   Begin VB.CheckBox Check5 
      Caption         =   "Bright Powder Effect"
      DataField       =   "BrightPowder"
      DataSource      =   "MoveData"
      Height          =   255
      Left            =   2160
      TabIndex        =   20
      Top             =   2520
      Width           =   2055
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Hits Both Opponents"
      DataField       =   "HitsBoth"
      DataSource      =   "MoveData"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   3600
      Width           =   1815
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Exists in GSC"
      DataField       =   "GSCCompatible"
      DataSource      =   "MoveData"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   3240
      Width           =   1815
   End
   Begin VB.CheckBox Check2 
      Caption         =   "King's Rock Effect"
      DataField       =   "KingsRock"
      DataSource      =   "MoveData"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   2880
      Width           =   1815
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Effect is Working"
      DataField       =   "Works Properly"
      DataSource      =   "MoveData"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   2520
      Width           =   1815
   End
   Begin VB.TextBox Text6 
      DataField       =   "Special"
      DataSource      =   "MoveData"
      Height          =   285
      Left            =   3480
      TabIndex        =   13
      Text            =   "Text2"
      Top             =   2040
      Width           =   735
   End
   Begin VB.TextBox Text5 
      DataField       =   "Percent"
      DataSource      =   "MoveData"
      Height          =   285
      Left            =   2640
      TabIndex        =   11
      Text            =   "Text2"
      Top             =   2040
      Width           =   735
   End
   Begin VB.TextBox Text4 
      DataField       =   "PP"
      DataSource      =   "MoveData"
      Height          =   285
      Left            =   1800
      TabIndex        =   6
      Text            =   "Text2"
      Top             =   2040
      Width           =   735
   End
   Begin VB.TextBox Text3 
      DataField       =   "Accuracy"
      DataSource      =   "MoveData"
      Height          =   285
      Left            =   960
      TabIndex        =   4
      Text            =   "Text2"
      Top             =   2040
      Width           =   735
   End
   Begin VB.TextBox Text2 
      DataField       =   "Power"
      DataSource      =   "MoveData"
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   2040
      Width           =   735
   End
   Begin VB.TextBox Text1 
      DataField       =   "Description"
      DataSource      =   "MoveData"
      Height          =   735
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "MoveEdit.frx":27A2
      Top             =   960
      Width           =   4095
   End
   Begin VB.TextBox MoveName 
      DataField       =   "Name"
      DataSource      =   "MoveData"
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   360
      Width           =   2055
   End
   Begin MSAdodcLib.Adodc MoveData 
      Height          =   330
      Left            =   120
      Top             =   5760
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSComctlLib.ImageCombo Type1 
      Height          =   330
      Left            =   2280
      TabIndex        =   10
      Top             =   360
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Locked          =   -1  'True
      ImageList       =   "Types"
   End
   Begin MSComctlLib.ImageList Types 
      Left            =   4680
      Top             =   480
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
            Picture         =   "MoveEdit.frx":27A8
            Key             =   ""
            Object.Tag             =   "Normal"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MoveEdit.frx":2D42
            Key             =   ""
            Object.Tag             =   "Fire"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MoveEdit.frx":32DC
            Key             =   ""
            Object.Tag             =   "Water"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MoveEdit.frx":3876
            Key             =   ""
            Object.Tag             =   "Electric"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MoveEdit.frx":3E10
            Key             =   ""
            Object.Tag             =   "Grass"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MoveEdit.frx":43AA
            Key             =   ""
            Object.Tag             =   "Ice"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MoveEdit.frx":4944
            Key             =   ""
            Object.Tag             =   "Fighting"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MoveEdit.frx":4EDE
            Key             =   ""
            Object.Tag             =   "Poison"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MoveEdit.frx":5478
            Key             =   ""
            Object.Tag             =   "Ground"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MoveEdit.frx":5A12
            Key             =   ""
            Object.Tag             =   "Flying"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MoveEdit.frx":5FAC
            Key             =   ""
            Object.Tag             =   "Psychic"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MoveEdit.frx":6546
            Key             =   ""
            Object.Tag             =   "Bug"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MoveEdit.frx":6AE0
            Key             =   ""
            Object.Tag             =   "Rock"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MoveEdit.frx":707A
            Key             =   ""
            Object.Tag             =   "Ghost"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MoveEdit.frx":7614
            Key             =   ""
            Object.Tag             =   "Dragon"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MoveEdit.frx":7BAE
            Key             =   ""
            Object.Tag             =   "Dark"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MoveEdit.frx":8148
            Key             =   ""
            Object.Tag             =   "Steel"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label12 
      Caption         =   "TypeTemp"
      DataField       =   "Type"
      DataSource      =   "MoveData"
      Height          =   375
      Left            =   4680
      TabIndex        =   38
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label8 
      Caption         =   "Description"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   720
      Width           =   3855
   End
   Begin VB.Label Label7 
      Caption         =   "Effect"
      Height          =   255
      Left            =   3480
      TabIndex        =   14
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Label6 
      Caption         =   "Sp. %"
      Height          =   255
      Left            =   2640
      TabIndex        =   12
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "Type"
      Height          =   255
      Left            =   2280
      TabIndex        =   9
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "Name"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "PP"
      Height          =   255
      Left            =   1800
      TabIndex        =   7
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Acc."
      Height          =   255
      Left            =   960
      TabIndex        =   5
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Power"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   495
   End
End
Attribute VB_Name = "MoveEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim Position As Integer
    
    Position = MoveData.Recordset.AbsolutePosition
    MoveData.Recordset.Save
    MoveData.Recordset.AbsolutePosition = Position
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Command3_Click()
    MoveData.Recordset.AddNew
End Sub

Private Sub Form_Load()
    Dim X As Integer
    
    For X = 1 To 17
        Type1.ComboItems.Add X, , Element(X), X
    Next
    With MoveData
        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & SlashPath & "PokeDB.mdb" & ";Mode=ReadWrite;Persist Security Info=False;Jet OLEDB:Database Password=ginyu4ce"
        .RecordSource = "Moves"
        .Refresh
        .Recordset.StayInSync = True
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DBManMain.Command2.Enabled = True
End Sub

Private Sub Label12_Change()
    If Val(Label12.Caption) > 0 Then
        If Type1.ComboItems.Item(Val(Label12.Caption)).Selected Then Exit Sub
    Else
        If Type1.ComboItems.Item(1).Selected Then Exit Sub
    End If
    If Val(Label12.Caption) > 0 Then
        Type1.ComboItems.Item(Val(Label12.Caption)).Selected = True
    Else
        Type1.ComboItems.Item(1).Selected = True
    End If
End Sub


'Private Sub MoveData_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
'    MoveData.Caption = MoveData.Recordset.AbsolutePosition & "/" & MoveData.Recordset.RecordCount
'    If MoveData.Recordset.AbsolutePosition < MoveData.Recordset.RecordCount Then Command3.Enabled = False Else Command3.Enabled = True
'End Sub
'

Private Sub Type1_Change()
    Dim X As Integer
    
    For X = 1 To 17
        If Type1.ComboItems.Item(X).Selected Then Label12.Caption = X
    Next X
End Sub
