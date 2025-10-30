VERSION 5.00
Begin VB.Form BatEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Battle Chart Editor"
   ClientHeight    =   7890
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6750
   Icon            =   "BatEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7890
   ScaleWidth      =   6750
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture2 
      BackColor       =   &H80000005&
      Height          =   1035
      Left            =   120
      ScaleHeight     =   975
      ScaleWidth      =   2355
      TabIndex        =   4
      Top             =   6720
      Width           =   2415
      Begin VB.Image DispImg 
         Height          =   240
         Index           =   0
         Left            =   60
         Picture         =   "BatEdit.frx":0442
         Top             =   60
         Width           =   240
      End
      Begin VB.Image DispImg 
         Height          =   240
         Index           =   1
         Left            =   60
         Picture         =   "BatEdit.frx":0546
         Top             =   360
         Width           =   240
      End
      Begin VB.Image DispImg 
         Height          =   240
         Index           =   2
         Left            =   60
         Picture         =   "BatEdit.frx":062B
         Top             =   660
         Width           =   240
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Attack will do no damage"
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   60
         Width           =   2055
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Attack will do 1/2 damage"
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Attack will do 2x damage"
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   660
         Width           =   2055
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Done"
      Height          =   495
      Left            =   5340
      TabIndex        =   3
      Top             =   7320
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000005&
      Height          =   6555
      Left            =   120
      ScaleHeight     =   6495
      ScaleWidth      =   6495
      TabIndex        =   0
      Top             =   120
      Width           =   6555
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   0
         Left            =   420
         Picture         =   "BatEdit.frx":06B4
         Top             =   420
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   1
         Left            =   780
         Picture         =   "BatEdit.frx":0C3E
         Top             =   420
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   2
         Left            =   1140
         Picture         =   "BatEdit.frx":11C8
         Top             =   420
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   3
         Left            =   1500
         Picture         =   "BatEdit.frx":1752
         Top             =   420
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   4
         Left            =   1860
         Picture         =   "BatEdit.frx":1CDC
         Top             =   420
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   5
         Left            =   2220
         Picture         =   "BatEdit.frx":2266
         Top             =   420
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   6
         Left            =   2580
         Picture         =   "BatEdit.frx":27F0
         Top             =   420
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   7
         Left            =   2940
         Picture         =   "BatEdit.frx":2D7A
         Top             =   420
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   8
         Left            =   3300
         Picture         =   "BatEdit.frx":3304
         Top             =   420
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   9
         Left            =   3660
         Picture         =   "BatEdit.frx":388E
         Top             =   420
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   10
         Left            =   4020
         Picture         =   "BatEdit.frx":3E18
         Top             =   420
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   11
         Left            =   4380
         Picture         =   "BatEdit.frx":43A2
         Top             =   420
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   12
         Left            =   4740
         Picture         =   "BatEdit.frx":492C
         Top             =   420
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   13
         Left            =   5100
         Picture         =   "BatEdit.frx":4EB6
         Top             =   420
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   14
         Left            =   5460
         Picture         =   "BatEdit.frx":5440
         Top             =   420
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   15
         Left            =   5820
         Picture         =   "BatEdit.frx":59CA
         Top             =   420
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   16
         Left            =   6180
         Picture         =   "BatEdit.frx":5F54
         Top             =   420
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   17
         Left            =   420
         Picture         =   "BatEdit.frx":64DE
         Top             =   780
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   18
         Left            =   780
         Picture         =   "BatEdit.frx":6A68
         Top             =   780
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   19
         Left            =   1140
         Picture         =   "BatEdit.frx":6FF2
         Top             =   780
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   20
         Left            =   1500
         Picture         =   "BatEdit.frx":757C
         Top             =   780
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   21
         Left            =   1860
         Picture         =   "BatEdit.frx":7B06
         Top             =   780
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   22
         Left            =   2220
         Picture         =   "BatEdit.frx":8090
         Top             =   780
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   23
         Left            =   2580
         Picture         =   "BatEdit.frx":861A
         Top             =   780
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   24
         Left            =   2940
         Picture         =   "BatEdit.frx":8BA4
         Top             =   780
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   25
         Left            =   3300
         Picture         =   "BatEdit.frx":912E
         Top             =   780
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   26
         Left            =   3660
         Picture         =   "BatEdit.frx":96B8
         Top             =   780
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   27
         Left            =   4020
         Picture         =   "BatEdit.frx":9C42
         Top             =   780
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   28
         Left            =   4380
         Picture         =   "BatEdit.frx":A1CC
         Top             =   780
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   29
         Left            =   4740
         Picture         =   "BatEdit.frx":A756
         Top             =   780
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   30
         Left            =   5100
         Picture         =   "BatEdit.frx":ACE0
         Top             =   780
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   31
         Left            =   5460
         Picture         =   "BatEdit.frx":B26A
         Top             =   780
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   32
         Left            =   5820
         Picture         =   "BatEdit.frx":B7F4
         Top             =   780
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   33
         Left            =   6180
         Picture         =   "BatEdit.frx":BD7E
         Top             =   780
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   34
         Left            =   420
         Picture         =   "BatEdit.frx":C308
         Top             =   1140
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   35
         Left            =   780
         Picture         =   "BatEdit.frx":C892
         Top             =   1140
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   36
         Left            =   1140
         Picture         =   "BatEdit.frx":CE1C
         Top             =   1140
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   37
         Left            =   1500
         Picture         =   "BatEdit.frx":D3A6
         Top             =   1140
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   38
         Left            =   1860
         Picture         =   "BatEdit.frx":D930
         Top             =   1140
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   39
         Left            =   2220
         Picture         =   "BatEdit.frx":DEBA
         Top             =   1140
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   40
         Left            =   2580
         Picture         =   "BatEdit.frx":E444
         Top             =   1140
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   41
         Left            =   2940
         Picture         =   "BatEdit.frx":E9CE
         Top             =   1140
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   42
         Left            =   3300
         Picture         =   "BatEdit.frx":EF58
         Top             =   1140
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   43
         Left            =   3660
         Picture         =   "BatEdit.frx":F4E2
         Top             =   1140
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   44
         Left            =   4020
         Picture         =   "BatEdit.frx":FA6C
         Top             =   1140
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   45
         Left            =   4380
         Picture         =   "BatEdit.frx":FFF6
         Top             =   1140
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   46
         Left            =   4740
         Picture         =   "BatEdit.frx":10580
         Top             =   1140
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   47
         Left            =   5100
         Picture         =   "BatEdit.frx":10B0A
         Top             =   1140
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   48
         Left            =   5460
         Picture         =   "BatEdit.frx":11094
         Top             =   1140
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   49
         Left            =   5820
         Picture         =   "BatEdit.frx":1161E
         Top             =   1140
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   50
         Left            =   6180
         Picture         =   "BatEdit.frx":11BA8
         Top             =   1140
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   51
         Left            =   420
         Picture         =   "BatEdit.frx":12132
         Top             =   1500
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   52
         Left            =   780
         Picture         =   "BatEdit.frx":126BC
         Top             =   1500
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   53
         Left            =   1140
         Picture         =   "BatEdit.frx":12C46
         Top             =   1500
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   54
         Left            =   1500
         Picture         =   "BatEdit.frx":131D0
         Top             =   1500
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   55
         Left            =   1860
         Picture         =   "BatEdit.frx":1375A
         Top             =   1500
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   56
         Left            =   2220
         Picture         =   "BatEdit.frx":13CE4
         Top             =   1500
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   57
         Left            =   2580
         Picture         =   "BatEdit.frx":1426E
         Top             =   1500
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   58
         Left            =   2940
         Picture         =   "BatEdit.frx":147F8
         Top             =   1500
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   59
         Left            =   3300
         Picture         =   "BatEdit.frx":14D82
         Top             =   1500
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   60
         Left            =   3660
         Picture         =   "BatEdit.frx":1530C
         Top             =   1500
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   61
         Left            =   4020
         Picture         =   "BatEdit.frx":15896
         Top             =   1500
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   62
         Left            =   4380
         Picture         =   "BatEdit.frx":15E20
         Top             =   1500
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   63
         Left            =   4740
         Picture         =   "BatEdit.frx":163AA
         Top             =   1500
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   64
         Left            =   5100
         Picture         =   "BatEdit.frx":16934
         Top             =   1500
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   65
         Left            =   5460
         Picture         =   "BatEdit.frx":16EBE
         Top             =   1500
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   66
         Left            =   5820
         Picture         =   "BatEdit.frx":17448
         Top             =   1500
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   67
         Left            =   6180
         Picture         =   "BatEdit.frx":179D2
         Top             =   1500
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   68
         Left            =   420
         Picture         =   "BatEdit.frx":17F5C
         Top             =   1860
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   69
         Left            =   780
         Picture         =   "BatEdit.frx":184E6
         Top             =   1860
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   70
         Left            =   1140
         Picture         =   "BatEdit.frx":18A70
         Top             =   1860
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   71
         Left            =   1500
         Picture         =   "BatEdit.frx":18FFA
         Top             =   1860
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   72
         Left            =   1860
         Picture         =   "BatEdit.frx":19584
         Top             =   1860
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   73
         Left            =   2220
         Picture         =   "BatEdit.frx":19B0E
         Top             =   1860
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   74
         Left            =   2580
         Picture         =   "BatEdit.frx":1A098
         Top             =   1860
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   75
         Left            =   2940
         Picture         =   "BatEdit.frx":1A622
         Top             =   1860
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   76
         Left            =   3300
         Picture         =   "BatEdit.frx":1ABAC
         Top             =   1860
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   77
         Left            =   3660
         Picture         =   "BatEdit.frx":1B136
         Top             =   1860
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   78
         Left            =   4020
         Picture         =   "BatEdit.frx":1B6C0
         Top             =   1860
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   79
         Left            =   4380
         Picture         =   "BatEdit.frx":1BC4A
         Top             =   1860
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   80
         Left            =   4740
         Picture         =   "BatEdit.frx":1C1D4
         Top             =   1860
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   81
         Left            =   5100
         Picture         =   "BatEdit.frx":1C75E
         Top             =   1860
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   82
         Left            =   5460
         Picture         =   "BatEdit.frx":1CCE8
         Top             =   1860
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   83
         Left            =   5820
         Picture         =   "BatEdit.frx":1D272
         Top             =   1860
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   84
         Left            =   6180
         Picture         =   "BatEdit.frx":1D7FC
         Top             =   1860
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   85
         Left            =   420
         Picture         =   "BatEdit.frx":1DD86
         Top             =   2220
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   86
         Left            =   780
         Picture         =   "BatEdit.frx":1E310
         Top             =   2220
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   87
         Left            =   1140
         Picture         =   "BatEdit.frx":1E89A
         Top             =   2220
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   88
         Left            =   1500
         Picture         =   "BatEdit.frx":1EE24
         Top             =   2220
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   89
         Left            =   1860
         Picture         =   "BatEdit.frx":1F3AE
         Top             =   2220
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   90
         Left            =   2220
         Picture         =   "BatEdit.frx":1F938
         Top             =   2220
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   91
         Left            =   2580
         Picture         =   "BatEdit.frx":1FEC2
         Top             =   2220
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   92
         Left            =   2940
         Picture         =   "BatEdit.frx":2044C
         Top             =   2220
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   93
         Left            =   3300
         Picture         =   "BatEdit.frx":209D6
         Top             =   2220
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   94
         Left            =   3660
         Picture         =   "BatEdit.frx":20F60
         Top             =   2220
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   95
         Left            =   4020
         Picture         =   "BatEdit.frx":214EA
         Top             =   2220
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   96
         Left            =   4380
         Picture         =   "BatEdit.frx":21A74
         Top             =   2220
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   97
         Left            =   4740
         Picture         =   "BatEdit.frx":21FFE
         Top             =   2220
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   98
         Left            =   5100
         Picture         =   "BatEdit.frx":22588
         Top             =   2220
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   99
         Left            =   5460
         Picture         =   "BatEdit.frx":22B12
         Top             =   2220
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   100
         Left            =   5820
         Picture         =   "BatEdit.frx":2309C
         Top             =   2220
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   101
         Left            =   6180
         Picture         =   "BatEdit.frx":23626
         Top             =   2220
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   102
         Left            =   420
         Picture         =   "BatEdit.frx":23BB0
         Top             =   2580
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   103
         Left            =   780
         Picture         =   "BatEdit.frx":2413A
         Top             =   2580
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   104
         Left            =   1140
         Picture         =   "BatEdit.frx":246C4
         Top             =   2580
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   105
         Left            =   1500
         Picture         =   "BatEdit.frx":24C4E
         Top             =   2580
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   106
         Left            =   1860
         Picture         =   "BatEdit.frx":251D8
         Top             =   2580
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   107
         Left            =   2220
         Picture         =   "BatEdit.frx":25762
         Top             =   2580
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   108
         Left            =   2580
         Picture         =   "BatEdit.frx":25CEC
         Top             =   2580
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   109
         Left            =   2940
         Picture         =   "BatEdit.frx":26276
         Top             =   2580
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   110
         Left            =   3300
         Picture         =   "BatEdit.frx":26800
         Top             =   2580
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   111
         Left            =   3660
         Picture         =   "BatEdit.frx":26D8A
         Top             =   2580
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   112
         Left            =   4020
         Picture         =   "BatEdit.frx":27314
         Top             =   2580
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   113
         Left            =   4380
         Picture         =   "BatEdit.frx":2789E
         Top             =   2580
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   114
         Left            =   4740
         Picture         =   "BatEdit.frx":27E28
         Top             =   2580
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   115
         Left            =   5100
         Picture         =   "BatEdit.frx":283B2
         Top             =   2580
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   116
         Left            =   5460
         Picture         =   "BatEdit.frx":2893C
         Top             =   2580
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   117
         Left            =   5820
         Picture         =   "BatEdit.frx":28EC6
         Top             =   2580
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   118
         Left            =   6180
         Picture         =   "BatEdit.frx":29450
         Top             =   2580
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   119
         Left            =   420
         Picture         =   "BatEdit.frx":299DA
         Top             =   2940
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   120
         Left            =   780
         Picture         =   "BatEdit.frx":29F64
         Top             =   2940
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   121
         Left            =   1140
         Picture         =   "BatEdit.frx":2A4EE
         Top             =   2940
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   122
         Left            =   1500
         Picture         =   "BatEdit.frx":2AA78
         Top             =   2940
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   123
         Left            =   1860
         Picture         =   "BatEdit.frx":2B002
         Top             =   2940
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   124
         Left            =   2220
         Picture         =   "BatEdit.frx":2B58C
         Top             =   2940
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   125
         Left            =   2580
         Picture         =   "BatEdit.frx":2BB16
         Top             =   2940
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   126
         Left            =   2940
         Picture         =   "BatEdit.frx":2C0A0
         Top             =   2940
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   127
         Left            =   3300
         Picture         =   "BatEdit.frx":2C62A
         Top             =   2940
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   128
         Left            =   3660
         Picture         =   "BatEdit.frx":2CBB4
         Top             =   2940
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   129
         Left            =   4020
         Picture         =   "BatEdit.frx":2D13E
         Top             =   2940
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   130
         Left            =   4380
         Picture         =   "BatEdit.frx":2D6C8
         Top             =   2940
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   131
         Left            =   4740
         Picture         =   "BatEdit.frx":2DC52
         Top             =   2940
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   132
         Left            =   5100
         Picture         =   "BatEdit.frx":2E1DC
         Top             =   2940
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   133
         Left            =   5460
         Picture         =   "BatEdit.frx":2E766
         Top             =   2940
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   134
         Left            =   5820
         Picture         =   "BatEdit.frx":2ECF0
         Top             =   2940
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   135
         Left            =   6180
         Picture         =   "BatEdit.frx":2F27A
         Top             =   2940
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   136
         Left            =   420
         Picture         =   "BatEdit.frx":2F804
         Top             =   3300
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   137
         Left            =   780
         Picture         =   "BatEdit.frx":2FD8E
         Top             =   3300
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   138
         Left            =   1140
         Picture         =   "BatEdit.frx":30318
         Top             =   3300
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   139
         Left            =   1500
         Picture         =   "BatEdit.frx":308A2
         Top             =   3300
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   140
         Left            =   1860
         Picture         =   "BatEdit.frx":30E2C
         Top             =   3300
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   141
         Left            =   2220
         Picture         =   "BatEdit.frx":313B6
         Top             =   3300
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   142
         Left            =   2580
         Picture         =   "BatEdit.frx":31940
         Top             =   3300
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   143
         Left            =   2940
         Picture         =   "BatEdit.frx":31ECA
         Top             =   3300
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   144
         Left            =   3300
         Picture         =   "BatEdit.frx":32454
         Top             =   3300
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   145
         Left            =   3660
         Picture         =   "BatEdit.frx":329DE
         Top             =   3300
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   146
         Left            =   4020
         Picture         =   "BatEdit.frx":32F68
         Top             =   3300
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   147
         Left            =   4380
         Picture         =   "BatEdit.frx":334F2
         Top             =   3300
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   148
         Left            =   4740
         Picture         =   "BatEdit.frx":33A7C
         Top             =   3300
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   149
         Left            =   5100
         Picture         =   "BatEdit.frx":34006
         Top             =   3300
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   150
         Left            =   5460
         Picture         =   "BatEdit.frx":34590
         Top             =   3300
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   151
         Left            =   5820
         Picture         =   "BatEdit.frx":34B1A
         Top             =   3300
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   152
         Left            =   6180
         Picture         =   "BatEdit.frx":350A4
         Top             =   3300
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   153
         Left            =   420
         Picture         =   "BatEdit.frx":3562E
         Top             =   3660
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   154
         Left            =   780
         Picture         =   "BatEdit.frx":35BB8
         Top             =   3660
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   155
         Left            =   1140
         Picture         =   "BatEdit.frx":36142
         Top             =   3660
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   156
         Left            =   1500
         Picture         =   "BatEdit.frx":366CC
         Top             =   3660
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   157
         Left            =   1860
         Picture         =   "BatEdit.frx":36C56
         Top             =   3660
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   158
         Left            =   2220
         Picture         =   "BatEdit.frx":371E0
         Top             =   3660
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   159
         Left            =   2580
         Picture         =   "BatEdit.frx":3776A
         Top             =   3660
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   160
         Left            =   2940
         Picture         =   "BatEdit.frx":37CF4
         Top             =   3660
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   161
         Left            =   3300
         Picture         =   "BatEdit.frx":3827E
         Top             =   3660
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   162
         Left            =   3660
         Picture         =   "BatEdit.frx":38808
         Top             =   3660
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   163
         Left            =   4020
         Picture         =   "BatEdit.frx":38D92
         Top             =   3660
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   164
         Left            =   4380
         Picture         =   "BatEdit.frx":3931C
         Top             =   3660
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   165
         Left            =   4740
         Picture         =   "BatEdit.frx":398A6
         Top             =   3660
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   166
         Left            =   5100
         Picture         =   "BatEdit.frx":39E30
         Top             =   3660
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   167
         Left            =   5460
         Picture         =   "BatEdit.frx":3A3BA
         Top             =   3660
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   168
         Left            =   5820
         Picture         =   "BatEdit.frx":3A944
         Top             =   3660
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   169
         Left            =   6180
         Picture         =   "BatEdit.frx":3AECE
         Top             =   3660
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   170
         Left            =   420
         Picture         =   "BatEdit.frx":3B458
         Top             =   4020
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   171
         Left            =   780
         Picture         =   "BatEdit.frx":3B9E2
         Top             =   4020
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   172
         Left            =   1140
         Picture         =   "BatEdit.frx":3BF6C
         Top             =   4020
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   173
         Left            =   1500
         Picture         =   "BatEdit.frx":3C4F6
         Top             =   4020
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   174
         Left            =   1860
         Picture         =   "BatEdit.frx":3CA80
         Top             =   4020
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   175
         Left            =   2220
         Picture         =   "BatEdit.frx":3D00A
         Top             =   4020
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   176
         Left            =   2580
         Picture         =   "BatEdit.frx":3D594
         Top             =   4020
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   177
         Left            =   2940
         Picture         =   "BatEdit.frx":3DB1E
         Top             =   4020
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   178
         Left            =   3300
         Picture         =   "BatEdit.frx":3E0A8
         Top             =   4020
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   179
         Left            =   3660
         Picture         =   "BatEdit.frx":3E632
         Top             =   4020
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   180
         Left            =   4020
         Picture         =   "BatEdit.frx":3EBBC
         Top             =   4020
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   181
         Left            =   4380
         Picture         =   "BatEdit.frx":3F146
         Top             =   4020
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   182
         Left            =   4740
         Picture         =   "BatEdit.frx":3F6D0
         Top             =   4020
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   183
         Left            =   5100
         Picture         =   "BatEdit.frx":3FC5A
         Top             =   4020
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   184
         Left            =   5460
         Picture         =   "BatEdit.frx":401E4
         Top             =   4020
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   185
         Left            =   5820
         Picture         =   "BatEdit.frx":4076E
         Top             =   4020
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   186
         Left            =   6180
         Picture         =   "BatEdit.frx":40CF8
         Top             =   4020
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   187
         Left            =   420
         Picture         =   "BatEdit.frx":41282
         Top             =   4380
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   188
         Left            =   780
         Picture         =   "BatEdit.frx":4180C
         Top             =   4380
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   189
         Left            =   1140
         Picture         =   "BatEdit.frx":41D96
         Top             =   4380
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   190
         Left            =   1500
         Picture         =   "BatEdit.frx":42320
         Top             =   4380
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   191
         Left            =   1860
         Picture         =   "BatEdit.frx":428AA
         Top             =   4380
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   192
         Left            =   2220
         Picture         =   "BatEdit.frx":42E34
         Top             =   4380
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   193
         Left            =   2580
         Picture         =   "BatEdit.frx":433BE
         Top             =   4380
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   194
         Left            =   2940
         Picture         =   "BatEdit.frx":43948
         Top             =   4380
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   195
         Left            =   3300
         Picture         =   "BatEdit.frx":43ED2
         Top             =   4380
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   196
         Left            =   3660
         Picture         =   "BatEdit.frx":4445C
         Top             =   4380
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   197
         Left            =   4020
         Picture         =   "BatEdit.frx":449E6
         Top             =   4380
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   198
         Left            =   4380
         Picture         =   "BatEdit.frx":44F70
         Top             =   4380
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   199
         Left            =   4740
         Picture         =   "BatEdit.frx":454FA
         Top             =   4380
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   200
         Left            =   5100
         Picture         =   "BatEdit.frx":45A84
         Top             =   4380
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   201
         Left            =   5460
         Picture         =   "BatEdit.frx":4600E
         Top             =   4380
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   202
         Left            =   5820
         Picture         =   "BatEdit.frx":46598
         Top             =   4380
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   203
         Left            =   6180
         Picture         =   "BatEdit.frx":46B22
         Top             =   4380
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   204
         Left            =   420
         Picture         =   "BatEdit.frx":470AC
         Top             =   4740
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   205
         Left            =   780
         Picture         =   "BatEdit.frx":47636
         Top             =   4740
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   206
         Left            =   1140
         Picture         =   "BatEdit.frx":47BC0
         Top             =   4740
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   207
         Left            =   1500
         Picture         =   "BatEdit.frx":4814A
         Top             =   4740
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   208
         Left            =   1860
         Picture         =   "BatEdit.frx":486D4
         Top             =   4740
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   209
         Left            =   2220
         Picture         =   "BatEdit.frx":48C5E
         Top             =   4740
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   210
         Left            =   2580
         Picture         =   "BatEdit.frx":491E8
         Top             =   4740
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   211
         Left            =   2940
         Picture         =   "BatEdit.frx":49772
         Top             =   4740
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   212
         Left            =   3300
         Picture         =   "BatEdit.frx":49CFC
         Top             =   4740
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   213
         Left            =   3660
         Picture         =   "BatEdit.frx":4A286
         Top             =   4740
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   214
         Left            =   4020
         Picture         =   "BatEdit.frx":4A810
         Top             =   4740
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   215
         Left            =   4380
         Picture         =   "BatEdit.frx":4AD9A
         Top             =   4740
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   216
         Left            =   4740
         Picture         =   "BatEdit.frx":4B324
         Top             =   4740
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   217
         Left            =   5100
         Picture         =   "BatEdit.frx":4B8AE
         Top             =   4740
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   218
         Left            =   5460
         Picture         =   "BatEdit.frx":4BE38
         Top             =   4740
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   219
         Left            =   5820
         Picture         =   "BatEdit.frx":4C3C2
         Top             =   4740
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   220
         Left            =   6180
         Picture         =   "BatEdit.frx":4C94C
         Top             =   4740
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   221
         Left            =   420
         Picture         =   "BatEdit.frx":4CED6
         Top             =   5100
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   222
         Left            =   780
         Picture         =   "BatEdit.frx":4D460
         Top             =   5100
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   223
         Left            =   1140
         Picture         =   "BatEdit.frx":4D9EA
         Top             =   5100
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   224
         Left            =   1500
         Picture         =   "BatEdit.frx":4DF74
         Top             =   5100
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   225
         Left            =   1860
         Picture         =   "BatEdit.frx":4E4FE
         Top             =   5100
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   226
         Left            =   2220
         Picture         =   "BatEdit.frx":4EA88
         Top             =   5100
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   227
         Left            =   2580
         Picture         =   "BatEdit.frx":4F012
         Top             =   5100
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   228
         Left            =   2940
         Picture         =   "BatEdit.frx":4F59C
         Top             =   5100
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   229
         Left            =   3300
         Picture         =   "BatEdit.frx":4FB26
         Top             =   5100
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   230
         Left            =   3660
         Picture         =   "BatEdit.frx":500B0
         Top             =   5100
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   231
         Left            =   4020
         Picture         =   "BatEdit.frx":5063A
         Top             =   5100
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   232
         Left            =   4380
         Picture         =   "BatEdit.frx":50BC4
         Top             =   5100
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   233
         Left            =   4740
         Picture         =   "BatEdit.frx":5114E
         Top             =   5100
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   234
         Left            =   5100
         Picture         =   "BatEdit.frx":516D8
         Top             =   5100
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   235
         Left            =   5460
         Picture         =   "BatEdit.frx":51C62
         Top             =   5100
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   236
         Left            =   5820
         Picture         =   "BatEdit.frx":521EC
         Top             =   5100
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   237
         Left            =   6180
         Picture         =   "BatEdit.frx":52776
         Top             =   5100
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   238
         Left            =   420
         Picture         =   "BatEdit.frx":52D00
         Top             =   5460
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   239
         Left            =   780
         Picture         =   "BatEdit.frx":5328A
         Top             =   5460
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   240
         Left            =   1140
         Picture         =   "BatEdit.frx":53814
         Top             =   5460
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   241
         Left            =   1500
         Picture         =   "BatEdit.frx":53D9E
         Top             =   5460
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   242
         Left            =   1860
         Picture         =   "BatEdit.frx":54328
         Top             =   5460
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   243
         Left            =   2220
         Picture         =   "BatEdit.frx":548B2
         Top             =   5460
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   244
         Left            =   2580
         Picture         =   "BatEdit.frx":54E3C
         Top             =   5460
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   245
         Left            =   2940
         Picture         =   "BatEdit.frx":553C6
         Top             =   5460
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   246
         Left            =   3300
         Picture         =   "BatEdit.frx":55950
         Top             =   5460
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   247
         Left            =   3660
         Picture         =   "BatEdit.frx":55EDA
         Top             =   5460
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   248
         Left            =   4020
         Picture         =   "BatEdit.frx":56464
         Top             =   5460
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   249
         Left            =   4380
         Picture         =   "BatEdit.frx":569EE
         Top             =   5460
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   250
         Left            =   4740
         Picture         =   "BatEdit.frx":56F78
         Top             =   5460
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   251
         Left            =   5100
         Picture         =   "BatEdit.frx":57502
         Top             =   5460
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   252
         Left            =   5460
         Picture         =   "BatEdit.frx":57A8C
         Top             =   5460
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   253
         Left            =   5820
         Picture         =   "BatEdit.frx":58016
         Top             =   5460
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   254
         Left            =   6180
         Picture         =   "BatEdit.frx":585A0
         Top             =   5460
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   255
         Left            =   420
         Picture         =   "BatEdit.frx":58B2A
         Top             =   5820
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   256
         Left            =   780
         Picture         =   "BatEdit.frx":590B4
         Top             =   5820
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   257
         Left            =   1140
         Picture         =   "BatEdit.frx":5963E
         Top             =   5820
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   258
         Left            =   1500
         Picture         =   "BatEdit.frx":59BC8
         Top             =   5820
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   259
         Left            =   1860
         Picture         =   "BatEdit.frx":5A152
         Top             =   5820
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   260
         Left            =   2220
         Picture         =   "BatEdit.frx":5A6DC
         Top             =   5820
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   261
         Left            =   2580
         Picture         =   "BatEdit.frx":5AC66
         Top             =   5820
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   262
         Left            =   2940
         Picture         =   "BatEdit.frx":5B1F0
         Top             =   5820
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   263
         Left            =   3300
         Picture         =   "BatEdit.frx":5B77A
         Top             =   5820
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   264
         Left            =   3660
         Picture         =   "BatEdit.frx":5BD04
         Top             =   5820
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   265
         Left            =   4020
         Picture         =   "BatEdit.frx":5C28E
         Top             =   5820
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   266
         Left            =   4380
         Picture         =   "BatEdit.frx":5C818
         Top             =   5820
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   267
         Left            =   4740
         Picture         =   "BatEdit.frx":5CDA2
         Top             =   5820
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   268
         Left            =   5100
         Picture         =   "BatEdit.frx":5D32C
         Top             =   5820
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   269
         Left            =   5460
         Picture         =   "BatEdit.frx":5D8B6
         Top             =   5820
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   270
         Left            =   5820
         Picture         =   "BatEdit.frx":5DE40
         Top             =   5820
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   271
         Left            =   6180
         Picture         =   "BatEdit.frx":5E3CA
         Top             =   5820
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   272
         Left            =   420
         Picture         =   "BatEdit.frx":5E954
         Top             =   6180
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   273
         Left            =   780
         Picture         =   "BatEdit.frx":5EEDE
         Top             =   6180
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   274
         Left            =   1140
         Picture         =   "BatEdit.frx":5F468
         Top             =   6180
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   275
         Left            =   1500
         Picture         =   "BatEdit.frx":5F9F2
         Top             =   6180
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   276
         Left            =   1860
         Picture         =   "BatEdit.frx":5FF7C
         Top             =   6180
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   277
         Left            =   2220
         Picture         =   "BatEdit.frx":60506
         Top             =   6180
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   278
         Left            =   2580
         Picture         =   "BatEdit.frx":60A90
         Top             =   6180
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   279
         Left            =   2940
         Picture         =   "BatEdit.frx":6101A
         Top             =   6180
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   280
         Left            =   3300
         Picture         =   "BatEdit.frx":615A4
         Top             =   6180
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   281
         Left            =   3660
         Picture         =   "BatEdit.frx":61B2E
         Top             =   6180
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   282
         Left            =   4020
         Picture         =   "BatEdit.frx":620B8
         Top             =   6180
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   283
         Left            =   4380
         Picture         =   "BatEdit.frx":62642
         Top             =   6180
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   284
         Left            =   4740
         Picture         =   "BatEdit.frx":62BCC
         Top             =   6180
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   285
         Left            =   5100
         Picture         =   "BatEdit.frx":63156
         Top             =   6180
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   286
         Left            =   5460
         Picture         =   "BatEdit.frx":636E0
         Top             =   6180
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   287
         Left            =   5820
         Picture         =   "BatEdit.frx":63C6A
         Top             =   6180
         Width           =   240
      End
      Begin VB.Image ChartPic 
         Height          =   240
         Index           =   288
         Left            =   6180
         Picture         =   "BatEdit.frx":641F4
         Top             =   6180
         Width           =   240
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   6960
         X2              =   0
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   6960
         X2              =   0
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   6900
         X2              =   -60
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line1 
         Index           =   3
         X1              =   6960
         X2              =   0
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Line Line1 
         Index           =   4
         X1              =   6960
         X2              =   0
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Line Line1 
         Index           =   5
         X1              =   6900
         X2              =   -60
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Line Line1 
         Index           =   6
         X1              =   6900
         X2              =   -60
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Line Line1 
         Index           =   7
         X1              =   6900
         X2              =   -60
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Line Line1 
         Index           =   8
         X1              =   6900
         X2              =   -60
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Line Line1 
         Index           =   9
         X1              =   6960
         X2              =   0
         Y1              =   3600
         Y2              =   3600
      End
      Begin VB.Line Line1 
         Index           =   10
         X1              =   6960
         X2              =   0
         Y1              =   3240
         Y2              =   3240
      End
      Begin VB.Line Line1 
         Index           =   11
         X1              =   6960
         X2              =   0
         Y1              =   3960
         Y2              =   3960
      End
      Begin VB.Line Line1 
         Index           =   12
         X1              =   6660
         X2              =   -300
         Y1              =   4320
         Y2              =   4320
      End
      Begin VB.Line Line1 
         Index           =   13
         X1              =   6960
         X2              =   0
         Y1              =   5040
         Y2              =   5040
      End
      Begin VB.Line Line1 
         Index           =   14
         X1              =   6480
         X2              =   -480
         Y1              =   4680
         Y2              =   4680
      End
      Begin VB.Line Line1 
         Index           =   15
         X1              =   6960
         X2              =   0
         Y1              =   5400
         Y2              =   5400
      End
      Begin VB.Line Line1 
         Index           =   16
         X1              =   6960
         X2              =   0
         Y1              =   5760
         Y2              =   5760
      End
      Begin VB.Line Line1 
         Index           =   17
         X1              =   6900
         X2              =   -60
         Y1              =   6120
         Y2              =   6120
      End
      Begin VB.Line Line1 
         Index           =   18
         X1              =   6780
         X2              =   -180
         Y1              =   6480
         Y2              =   6480
      End
      Begin VB.Line Line2 
         Index           =   0
         X1              =   360
         X2              =   360
         Y1              =   0
         Y2              =   6540
      End
      Begin VB.Line Line3 
         Index           =   0
         X1              =   720
         X2              =   720
         Y1              =   0
         Y2              =   6540
      End
      Begin VB.Line Line4 
         Index           =   0
         X1              =   1080
         X2              =   1080
         Y1              =   0
         Y2              =   6540
      End
      Begin VB.Line Line2 
         Index           =   1
         X1              =   1440
         X2              =   1440
         Y1              =   0
         Y2              =   6540
      End
      Begin VB.Line Line3 
         Index           =   1
         X1              =   1800
         X2              =   1800
         Y1              =   0
         Y2              =   6540
      End
      Begin VB.Line Line4 
         Index           =   1
         X1              =   2160
         X2              =   2160
         Y1              =   0
         Y2              =   6540
      End
      Begin VB.Line Line2 
         Index           =   2
         X1              =   2520
         X2              =   2520
         Y1              =   0
         Y2              =   6540
      End
      Begin VB.Line Line3 
         Index           =   2
         X1              =   2880
         X2              =   2880
         Y1              =   0
         Y2              =   6540
      End
      Begin VB.Line Line4 
         Index           =   2
         X1              =   3240
         X2              =   3240
         Y1              =   0
         Y2              =   6540
      End
      Begin VB.Line Line2 
         Index           =   3
         X1              =   3600
         X2              =   3600
         Y1              =   0
         Y2              =   6540
      End
      Begin VB.Line Line3 
         Index           =   3
         X1              =   3960
         X2              =   3960
         Y1              =   0
         Y2              =   6540
      End
      Begin VB.Line Line4 
         Index           =   3
         X1              =   4320
         X2              =   4320
         Y1              =   0
         Y2              =   6540
      End
      Begin VB.Line Line2 
         Index           =   4
         X1              =   4680
         X2              =   4680
         Y1              =   0
         Y2              =   6540
      End
      Begin VB.Line Line3 
         Index           =   4
         X1              =   5040
         X2              =   5040
         Y1              =   0
         Y2              =   6540
      End
      Begin VB.Line Line4 
         Index           =   4
         X1              =   5400
         X2              =   5400
         Y1              =   0
         Y2              =   6540
      End
      Begin VB.Line Line2 
         Index           =   5
         X1              =   5760
         X2              =   5760
         Y1              =   0
         Y2              =   6540
      End
      Begin VB.Line Line3 
         Index           =   5
         X1              =   6120
         X2              =   6120
         Y1              =   0
         Y2              =   6540
      End
      Begin VB.Line Line4 
         Index           =   5
         X1              =   6480
         X2              =   6480
         Y1              =   0
         Y2              =   6660
      End
      Begin VB.Image Attacker 
         Height          =   240
         Index           =   0
         Left            =   60
         Picture         =   "BatEdit.frx":6477E
         Top             =   420
         Width           =   240
      End
      Begin VB.Image Attacker 
         Height          =   240
         Index           =   1
         Left            =   60
         Picture         =   "BatEdit.frx":64D08
         Top             =   780
         Width           =   240
      End
      Begin VB.Image Attacker 
         Height          =   240
         Index           =   2
         Left            =   60
         Picture         =   "BatEdit.frx":65292
         Top             =   1140
         Width           =   240
      End
      Begin VB.Image Attacker 
         Height          =   240
         Index           =   3
         Left            =   60
         Picture         =   "BatEdit.frx":6581C
         Top             =   1500
         Width           =   240
      End
      Begin VB.Image Attacker 
         Height          =   240
         Index           =   4
         Left            =   60
         Picture         =   "BatEdit.frx":65DA6
         Top             =   1860
         Width           =   240
      End
      Begin VB.Image Attacker 
         Height          =   240
         Index           =   5
         Left            =   60
         Picture         =   "BatEdit.frx":66330
         Top             =   2220
         Width           =   240
      End
      Begin VB.Image Attacker 
         Height          =   240
         Index           =   6
         Left            =   60
         Picture         =   "BatEdit.frx":668BA
         Top             =   2580
         Width           =   240
      End
      Begin VB.Image Attacker 
         Height          =   240
         Index           =   7
         Left            =   60
         Picture         =   "BatEdit.frx":66E44
         Top             =   2940
         Width           =   240
      End
      Begin VB.Image Attacker 
         Height          =   240
         Index           =   8
         Left            =   60
         Picture         =   "BatEdit.frx":673CE
         Top             =   3300
         Width           =   240
      End
      Begin VB.Image Attacker 
         Height          =   240
         Index           =   9
         Left            =   60
         Picture         =   "BatEdit.frx":67958
         Top             =   3660
         Width           =   240
      End
      Begin VB.Image Attacker 
         Height          =   240
         Index           =   10
         Left            =   60
         Picture         =   "BatEdit.frx":67EE2
         Top             =   4020
         Width           =   240
      End
      Begin VB.Image Attacker 
         Height          =   240
         Index           =   11
         Left            =   60
         Picture         =   "BatEdit.frx":6846C
         Top             =   4380
         Width           =   240
      End
      Begin VB.Image Attacker 
         Height          =   240
         Index           =   12
         Left            =   60
         Picture         =   "BatEdit.frx":689F6
         Top             =   4740
         Width           =   240
      End
      Begin VB.Image Attacker 
         Height          =   240
         Index           =   13
         Left            =   60
         Picture         =   "BatEdit.frx":68F80
         Top             =   5100
         Width           =   240
      End
      Begin VB.Image Attacker 
         Height          =   240
         Index           =   14
         Left            =   60
         Picture         =   "BatEdit.frx":6950A
         Top             =   5460
         Width           =   240
      End
      Begin VB.Image Attacker 
         Height          =   240
         Index           =   15
         Left            =   60
         Picture         =   "BatEdit.frx":69A94
         Top             =   5820
         Width           =   240
      End
      Begin VB.Image Attacker 
         Height          =   240
         Index           =   16
         Left            =   60
         Picture         =   "BatEdit.frx":6A01E
         Top             =   6180
         Width           =   240
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "A"
         Height          =   195
         Index           =   0
         Left            =   60
         TabIndex        =   2
         ToolTipText     =   "Attacker"
         Top             =   180
         Width           =   135
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "D"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   1
         ToolTipText     =   "Defender"
         Top             =   0
         Width           =   195
      End
      Begin VB.Line Line5 
         X1              =   360
         X2              =   0
         Y1              =   360
         Y2              =   0
      End
      Begin VB.Image Defender 
         Height          =   240
         Index           =   0
         Left            =   420
         Picture         =   "BatEdit.frx":6A5A8
         Top             =   60
         Width           =   240
      End
      Begin VB.Image Defender 
         Height          =   240
         Index           =   1
         Left            =   780
         Picture         =   "BatEdit.frx":6AB32
         Top             =   60
         Width           =   240
      End
      Begin VB.Image Defender 
         Height          =   240
         Index           =   2
         Left            =   1140
         Picture         =   "BatEdit.frx":6B0BC
         Top             =   60
         Width           =   240
      End
      Begin VB.Image Defender 
         Height          =   240
         Index           =   3
         Left            =   1500
         Picture         =   "BatEdit.frx":6B646
         Top             =   60
         Width           =   240
      End
      Begin VB.Image Defender 
         Height          =   240
         Index           =   4
         Left            =   1860
         Picture         =   "BatEdit.frx":6BBD0
         Top             =   60
         Width           =   240
      End
      Begin VB.Image Defender 
         Height          =   240
         Index           =   5
         Left            =   2220
         Picture         =   "BatEdit.frx":6C15A
         Top             =   60
         Width           =   240
      End
      Begin VB.Image Defender 
         Height          =   240
         Index           =   6
         Left            =   2580
         Picture         =   "BatEdit.frx":6C6E4
         Top             =   60
         Width           =   240
      End
      Begin VB.Image Defender 
         Height          =   240
         Index           =   7
         Left            =   2940
         Picture         =   "BatEdit.frx":6CC6E
         Top             =   60
         Width           =   240
      End
      Begin VB.Image Defender 
         Height          =   240
         Index           =   8
         Left            =   3300
         Picture         =   "BatEdit.frx":6D1F8
         Top             =   60
         Width           =   240
      End
      Begin VB.Image Defender 
         Height          =   240
         Index           =   9
         Left            =   3660
         Picture         =   "BatEdit.frx":6D782
         Top             =   60
         Width           =   240
      End
      Begin VB.Image Defender 
         Height          =   240
         Index           =   10
         Left            =   4020
         Picture         =   "BatEdit.frx":6DD0C
         Top             =   60
         Width           =   240
      End
      Begin VB.Image Defender 
         Height          =   240
         Index           =   11
         Left            =   4380
         Picture         =   "BatEdit.frx":6E296
         Top             =   60
         Width           =   240
      End
      Begin VB.Image Defender 
         Height          =   240
         Index           =   12
         Left            =   4740
         Picture         =   "BatEdit.frx":6E820
         Top             =   60
         Width           =   240
      End
      Begin VB.Image Defender 
         Height          =   240
         Index           =   13
         Left            =   5100
         Picture         =   "BatEdit.frx":6EDAA
         Top             =   60
         Width           =   240
      End
      Begin VB.Image Defender 
         Height          =   240
         Index           =   14
         Left            =   5460
         Picture         =   "BatEdit.frx":6F334
         Top             =   60
         Width           =   240
      End
      Begin VB.Image Defender 
         Height          =   240
         Index           =   15
         Left            =   5820
         Picture         =   "BatEdit.frx":6F8BE
         Top             =   60
         Width           =   240
      End
      Begin VB.Image Defender 
         Height          =   240
         Index           =   16
         Left            =   6180
         Picture         =   "BatEdit.frx":6FE48
         Top             =   60
         Width           =   240
      End
   End
End
Attribute VB_Name = "BatEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'BatEdit.frm
'Battle chart editor
'Use left/right click to change each value
'Done saves to the database
Option Explicit

Private Sub ChartPic_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim XP As Integer
    Dim YP As Integer
    Dim Temp As Integer
    
    'This one changes the values when clicked.
    XP = 1
    Temp = Index
    
    While Temp > 16
        Temp = Temp - 17
        XP = XP + 1
    Wend
    YP = Temp + 1
    Select Case Button
        Case vbLeftButton
            Select Case BattleMatrix(XP, YP)
                Case 0
                    BattleMatrix(XP, YP) = 0.5
                Case 0.5
                    BattleMatrix(XP, YP) = 1
                Case 1
                    BattleMatrix(XP, YP) = 2
                Case 2
                    BattleMatrix(XP, YP) = 0
            End Select
        Case vbRightButton
            Select Case BattleMatrix(XP, YP)
                Case 0
                    BattleMatrix(XP, YP) = 2
                Case 0.5
                    BattleMatrix(XP, YP) = 0
                Case 1
                    BattleMatrix(XP, YP) = 0.5
                Case 2
                    BattleMatrix(XP, YP) = 1
            End Select
    End Select
    Call RefreshListing
End Sub

Private Sub Command1_Click()
    Dim X As Integer
    Dim Y As Integer
    
    'Save values to the database.
    For X = 1 To 17
        For Y = 1 To 17
            PokeData.Execute "UPDATE BattleChart SET " & Right(Str(Y), Len(Str(Y)) - 1) & "=" & BattleMatrix(X, Y) & " WHERE ID=" & X
        Next
    Next
    Unload Me
End Sub

Private Sub Form_Load()
    Call LoadFromDB
    Call RefreshListing
End Sub

Private Sub RefreshListing()
    Dim X As Integer
    Dim Y As Integer
    Dim CurrentIcon As Integer
    
    'Refresh the listing to current values (in the BattleMatrix array).
    CurrentIcon = 0

    For X = 1 To 17
        Attacker(X - 1).Picture = DBManMain.Types.ListImages(X).Picture
        Defender(X - 1).Picture = DBManMain.Types.ListImages(X).Picture
        For Y = 1 To 17
            Select Case BattleMatrix(X, Y)
                Case 0
                    ChartPic(CurrentIcon).Picture = DispImg(0).Picture
                    ChartPic(CurrentIcon).Visible = True
                    ChartPic(CurrentIcon).ToolTipText = "Immune"
                Case 0.5
                    ChartPic(CurrentIcon).Picture = DispImg(1).Picture
                    ChartPic(CurrentIcon).Visible = True
                    ChartPic(CurrentIcon).ToolTipText = "Strong"
                Case 1
                    ChartPic(CurrentIcon).Picture = LoadPicture()
                    ChartPic(CurrentIcon).Visible = True
                    ChartPic(CurrentIcon).ToolTipText = "Normal"
                Case 2
                    ChartPic(CurrentIcon).Picture = DispImg(2).Picture
                    ChartPic(CurrentIcon).Visible = True
                    ChartPic(CurrentIcon).ToolTipText = "Weak"
            End Select
            CurrentIcon = CurrentIcon + 1
        Next
    Next
End Sub

Private Sub LoadFromDB()
    Dim QueryResults As ADODB.Recordset
    Dim CurrentRecord As Integer
    Dim X As Integer

    Set QueryResults = New ADODB.Recordset
    QueryResults.Open "SELECT * FROM BattleChart WHERE ID > 0 ORDER BY ID ASC", PokeData, adOpenStatic, adLockReadOnly, adCmdText
    QueryResults.MoveLast
    QueryResults.MoveFirst
    While Not QueryResults.EOF
        CurrentRecord = QueryResults("ID")
        For X = 1 To 17
            BattleMatrix(CurrentRecord, X) = QueryResults(X)
        Next
        QueryResults.MoveNext
    Wend
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DBManMain.Command3.Enabled = True
End Sub
