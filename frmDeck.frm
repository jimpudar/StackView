VERSION 5.00
Begin VB.Form frmDeck 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Deck"
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10830
   Icon            =   "frmDeck.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   10830
   Visible         =   0   'False
   WhatsThisHelp   =   -1  'True
   Begin VB.Label PileLabel 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   8
      Left            =   1755
      TabIndex        =   111
      Top             =   5625
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label PileLabel 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   7
      Left            =   1560
      TabIndex        =   110
      Top             =   5610
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label PileLabel 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   6
      Left            =   1365
      TabIndex        =   109
      Top             =   5640
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label PileLabel 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   5
      Left            =   1140
      TabIndex        =   108
      Top             =   5610
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label PileLabel 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   4
      Left            =   915
      TabIndex        =   107
      Top             =   5595
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label PileLabel 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   705
      TabIndex        =   106
      Top             =   5580
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label PileLabel 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   495
      TabIndex        =   105
      Top             =   5580
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label PileLabel 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   270
      TabIndex        =   104
      Top             =   5595
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label Ordinal 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "52"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   225
      Index           =   52
      Left            =   2145
      TabIndex        =   103
      Top             =   5385
      Width           =   150
   End
   Begin VB.Label Ordinal 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "51"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   225
      Index           =   51
      Left            =   2340
      TabIndex        =   102
      Top             =   5400
      Width           =   150
   End
   Begin VB.Label Ordinal 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "50"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   225
      Index           =   50
      Left            =   2595
      TabIndex        =   101
      Top             =   5370
      Width           =   150
   End
   Begin VB.Label Ordinal 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "49"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   225
      Index           =   49
      Left            =   2850
      TabIndex        =   100
      Top             =   5355
      Width           =   150
   End
   Begin VB.Label Ordinal 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "48"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   225
      Index           =   48
      Left            =   3090
      TabIndex        =   99
      Top             =   5385
      Width           =   150
   End
   Begin VB.Label Ordinal 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "47"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   225
      Index           =   47
      Left            =   3330
      TabIndex        =   98
      Top             =   5430
      Width           =   150
   End
   Begin VB.Label Ordinal 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "46"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   225
      Index           =   46
      Left            =   3525
      TabIndex        =   97
      Top             =   5400
      Width           =   150
   End
   Begin VB.Label Ordinal 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "45"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   225
      Index           =   45
      Left            =   3720
      TabIndex        =   96
      Top             =   5325
      Width           =   150
   End
   Begin VB.Label Ordinal 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "44"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   225
      Index           =   44
      Left            =   3915
      TabIndex        =   95
      Top             =   5340
      Width           =   150
   End
   Begin VB.Label Ordinal 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "43"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   225
      Index           =   43
      Left            =   4110
      TabIndex        =   94
      Top             =   5340
      Width           =   150
   End
   Begin VB.Label Ordinal 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "42"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   225
      Index           =   42
      Left            =   4290
      TabIndex        =   93
      Top             =   5355
      Width           =   150
   End
   Begin VB.Label Ordinal 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "41"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   225
      Index           =   41
      Left            =   4440
      TabIndex        =   92
      Top             =   5385
      Width           =   150
   End
   Begin VB.Label Ordinal 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "40"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   225
      Index           =   40
      Left            =   4635
      TabIndex        =   91
      Top             =   5430
      Width           =   150
   End
   Begin VB.Label Ordinal 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "39"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   225
      Index           =   39
      Left            =   4785
      TabIndex        =   90
      Top             =   5460
      Width           =   150
   End
   Begin VB.Label Ordinal 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "38"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   225
      Index           =   38
      Left            =   4950
      TabIndex        =   89
      Top             =   5490
      Width           =   150
   End
   Begin VB.Label Ordinal 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "37"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   225
      Index           =   37
      Left            =   5070
      TabIndex        =   88
      Top             =   5370
      Width           =   150
   End
   Begin VB.Label Ordinal 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "36"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   225
      Index           =   36
      Left            =   5265
      TabIndex        =   87
      Top             =   5445
      Width           =   150
   End
   Begin VB.Label Ordinal 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "35"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   225
      Index           =   35
      Left            =   5430
      TabIndex        =   86
      Top             =   5445
      Width           =   150
   End
   Begin VB.Label Ordinal 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "34"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   225
      Index           =   34
      Left            =   5625
      TabIndex        =   85
      Top             =   5460
      Width           =   150
   End
   Begin VB.Label Ordinal 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "33"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   225
      Index           =   33
      Left            =   5835
      TabIndex        =   84
      Top             =   5385
      Width           =   150
   End
   Begin VB.Label Ordinal 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "32"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   225
      Index           =   32
      Left            =   6015
      TabIndex        =   83
      Top             =   5430
      Width           =   150
   End
   Begin VB.Label Ordinal 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "31"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   225
      Index           =   31
      Left            =   6195
      TabIndex        =   82
      Top             =   5385
      Width           =   150
   End
   Begin VB.Label Ordinal 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "30"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   225
      Index           =   30
      Left            =   6360
      TabIndex        =   81
      Top             =   5415
      Width           =   150
   End
   Begin VB.Label Ordinal 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "29"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   225
      Index           =   29
      Left            =   6570
      TabIndex        =   80
      Top             =   5430
      Width           =   150
   End
   Begin VB.Label Ordinal 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "28"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   225
      Index           =   28
      Left            =   6705
      TabIndex        =   79
      Top             =   5370
      Width           =   150
   End
   Begin VB.Label Ordinal 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "27"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   225
      Index           =   27
      Left            =   6885
      TabIndex        =   78
      Top             =   5490
      Width           =   150
   End
   Begin VB.Label Ordinal 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "26"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   225
      Index           =   26
      Left            =   7020
      TabIndex        =   77
      Top             =   5475
      Width           =   150
   End
   Begin VB.Label Ordinal 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "25"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   225
      Index           =   25
      Left            =   7230
      TabIndex        =   76
      Top             =   5505
      Width           =   150
   End
   Begin VB.Label Ordinal 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "24"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   225
      Index           =   24
      Left            =   7410
      TabIndex        =   75
      Top             =   5475
      Width           =   150
   End
   Begin VB.Label Ordinal 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "23"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   225
      Index           =   23
      Left            =   7560
      TabIndex        =   74
      Top             =   5445
      Width           =   150
   End
   Begin VB.Label Ordinal 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "22"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   225
      Index           =   22
      Left            =   7725
      TabIndex        =   73
      Top             =   5430
      Width           =   150
   End
   Begin VB.Label Ordinal 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "21"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   225
      Index           =   21
      Left            =   7920
      TabIndex        =   72
      Top             =   5385
      Width           =   150
   End
   Begin VB.Label Ordinal 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "20"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   225
      Index           =   20
      Left            =   8100
      TabIndex        =   71
      Top             =   5325
      Width           =   150
   End
   Begin VB.Label Ordinal 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "19"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   225
      Index           =   19
      Left            =   8280
      TabIndex        =   70
      Top             =   5385
      Width           =   150
   End
   Begin VB.Label Ordinal 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "18"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   225
      Index           =   18
      Left            =   8445
      TabIndex        =   69
      Top             =   5475
      Width           =   150
   End
   Begin VB.Label Ordinal 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "17"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   225
      Index           =   17
      Left            =   8610
      TabIndex        =   68
      Top             =   5430
      Width           =   150
   End
   Begin VB.Label Ordinal 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "16"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   225
      Index           =   16
      Left            =   2370
      TabIndex        =   67
      Top             =   5685
      Width           =   150
   End
   Begin VB.Label Ordinal 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "15"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   225
      Index           =   15
      Left            =   2520
      TabIndex        =   66
      Top             =   5655
      Width           =   150
   End
   Begin VB.Label Ordinal 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "14"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   225
      Index           =   14
      Left            =   2760
      TabIndex        =   65
      Top             =   5670
      Width           =   150
   End
   Begin VB.Label Ordinal 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "13"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   225
      Index           =   13
      Left            =   2910
      TabIndex        =   64
      Top             =   5670
      Width           =   150
   End
   Begin VB.Label Ordinal 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   225
      Index           =   12
      Left            =   3090
      TabIndex        =   63
      Top             =   5700
      Width           =   150
   End
   Begin VB.Label Ordinal 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "11"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   225
      Index           =   11
      Left            =   3285
      TabIndex        =   62
      Top             =   5625
      Width           =   150
   End
   Begin VB.Label Ordinal 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   225
      Index           =   10
      Left            =   3465
      TabIndex        =   61
      Top             =   5670
      Width           =   150
   End
   Begin VB.Label Ordinal 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   225
      Index           =   9
      Left            =   3690
      TabIndex        =   60
      Top             =   5655
      Width           =   90
   End
   Begin VB.Label Ordinal 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   225
      Index           =   8
      Left            =   3885
      TabIndex        =   59
      Top             =   5655
      Width           =   90
   End
   Begin VB.Label Ordinal 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   225
      Index           =   7
      Left            =   4035
      TabIndex        =   58
      Top             =   5610
      Width           =   90
   End
   Begin VB.Label Ordinal 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   225
      Index           =   6
      Left            =   4155
      TabIndex        =   57
      Top             =   5655
      Width           =   90
   End
   Begin VB.Label Ordinal 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   225
      Index           =   5
      Left            =   4305
      TabIndex        =   56
      Top             =   5715
      Width           =   90
   End
   Begin VB.Label Ordinal 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   225
      Index           =   4
      Left            =   4425
      TabIndex        =   55
      Top             =   5685
      Width           =   90
   End
   Begin VB.Label Ordinal 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   225
      Index           =   3
      Left            =   4575
      TabIndex        =   54
      Top             =   5715
      Width           =   90
   End
   Begin VB.Label Ordinal 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   225
      Index           =   2
      Left            =   4695
      TabIndex        =   53
      Top             =   5745
      Width           =   90
   End
   Begin VB.Label Ordinal 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   225
      Index           =   1
      Left            =   4800
      TabIndex        =   52
      Top             =   5715
      Width           =   90
   End
   Begin VB.Image SelCardKD 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   8820
      Picture         =   "frmDeck.frx":1CCA
      Tag             =   "CardSelected"
      Top             =   2835
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image SelCardQD 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   8820
      Picture         =   "frmDeck.frx":4236
      Tag             =   "CardSelected"
      Top             =   2805
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image SelCardJD 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   8820
      Picture         =   "frmDeck.frx":681E
      Tag             =   "CardSelected"
      Top             =   2805
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image SelCard10D 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   8820
      Picture         =   "frmDeck.frx":8DF3
      Tag             =   "CardSelected"
      Top             =   2805
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image SelCard9D 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   8820
      Picture         =   "frmDeck.frx":ADAF
      Tag             =   "CardSelected"
      Top             =   2820
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image SelCard8D 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   8820
      Picture         =   "frmDeck.frx":CBD0
      Tag             =   "CardSelected"
      Top             =   2820
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image SelCard7D 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   8820
      Picture         =   "frmDeck.frx":E8B7
      Tag             =   "CardSelected"
      Top             =   2820
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image SelCard6D 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   8820
      Picture         =   "frmDeck.frx":1030A
      Tag             =   "CardSelected"
      Top             =   2805
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image SelCard5D 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   8820
      Picture         =   "frmDeck.frx":11B5B
      Tag             =   "CardSelected"
      Top             =   2805
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image SelCard4D 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   8820
      Picture         =   "frmDeck.frx":131BE
      Tag             =   "CardSelected"
      Top             =   2805
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image SelCard3D 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   8820
      Picture         =   "frmDeck.frx":1455D
      Tag             =   "CardSelected"
      Top             =   2805
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image SelCard2D 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   8835
      Picture         =   "frmDeck.frx":15931
      Tag             =   "CardSelected"
      Top             =   2805
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image SelCardAD 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   8820
      Picture         =   "frmDeck.frx":16A62
      Tag             =   "CardSelected"
      Top             =   2820
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image SelCardKS 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   8820
      Picture         =   "frmDeck.frx":17A94
      Tag             =   "CardSelected"
      Top             =   2805
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image SelCardQS 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   8820
      Picture         =   "frmDeck.frx":19F81
      Tag             =   "CardSelected"
      Top             =   2820
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image SelCardJS 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   8835
      Picture         =   "frmDeck.frx":1C413
      Tag             =   "CardSelected"
      Top             =   2805
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image SelCard10S 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   8820
      Picture         =   "frmDeck.frx":1E6C8
      Tag             =   "CardSelected"
      Top             =   2790
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image SelCard9S 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   8805
      Picture         =   "frmDeck.frx":203F7
      Tag             =   "CardSelected"
      Top             =   2805
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image SelCard8S 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   8820
      Picture         =   "frmDeck.frx":21FAD
      Tag             =   "CardSelected"
      Top             =   2805
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image SelCard7S 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   8820
      Picture         =   "frmDeck.frx":23A21
      Tag             =   "CardSelected"
      Top             =   2820
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image SelCard6S 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   8820
      Picture         =   "frmDeck.frx":2523D
      Tag             =   "CardSelected"
      Top             =   2820
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image SelCard5S 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   8835
      Picture         =   "frmDeck.frx":268F4
      Tag             =   "CardSelected"
      Top             =   2805
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image SelCard4S 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   8820
      Picture         =   "frmDeck.frx":27EF8
      Tag             =   "CardSelected"
      Top             =   2820
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image SelCard3S 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   8835
      Picture         =   "frmDeck.frx":292E8
      Tag             =   "CardSelected"
      Top             =   2805
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image SelCard2S 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   8835
      Picture         =   "frmDeck.frx":2A410
      Tag             =   "CardSelected"
      Top             =   2805
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image SelCardAS 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   8835
      Picture         =   "frmDeck.frx":2B405
      Tag             =   "CardSelected"
      Top             =   2805
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image SelCardKH 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   8835
      Picture         =   "frmDeck.frx":2C51E
      Tag             =   "CardSelected"
      Top             =   2805
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image SelCardQH 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   8835
      Picture         =   "frmDeck.frx":2EA47
      Tag             =   "CardSelected"
      Top             =   2820
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image SelCardJH 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   8820
      Picture         =   "frmDeck.frx":30F2D
      Tag             =   "CardSelected"
      Top             =   2820
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image SelCard10H 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   8820
      Picture         =   "frmDeck.frx":334C6
      Tag             =   "CardSelected"
      Top             =   2805
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image SelCard9H 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   8850
      Picture         =   "frmDeck.frx":35726
      Tag             =   "CardSelected"
      Top             =   2820
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image SelCard8H 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   8835
      Picture         =   "frmDeck.frx":3777B
      Tag             =   "CardSelected"
      Top             =   2790
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image SelCard7H 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   8835
      Picture         =   "frmDeck.frx":39645
      Tag             =   "CardSelected"
      Top             =   2805
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image SelCard6H 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   8835
      Picture         =   "frmDeck.frx":3B20E
      Tag             =   "CardSelected"
      Top             =   2835
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image SelCard5H 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   8835
      Picture         =   "frmDeck.frx":3CBCA
      Tag             =   "CardSelected"
      Top             =   2820
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image SelCard4H 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   8835
      Picture         =   "frmDeck.frx":3E363
      Tag             =   "CardSelected"
      Top             =   2820
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image SelCard3H 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   8835
      Picture         =   "frmDeck.frx":3F83A
      Tag             =   "CardSelected"
      Top             =   2820
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image SelCard2H 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   8835
      Picture         =   "frmDeck.frx":40C89
      Tag             =   "CardSelected"
      Top             =   2820
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image SelCardAH 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   8850
      Picture         =   "frmDeck.frx":41F01
      Tag             =   "CardSelected"
      Top             =   2820
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image SelCardKC 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   8850
      Picture         =   "frmDeck.frx":42EEB
      Tag             =   "CardSelected"
      Top             =   2820
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image SelCardQC 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   8835
      Picture         =   "frmDeck.frx":4502B
      Tag             =   "CardSelected"
      Top             =   2820
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image SelCardJC 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   8820
      Picture         =   "frmDeck.frx":472DB
      Tag             =   "CardSelected"
      Top             =   2820
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image SelCard10C 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   8835
      Picture         =   "frmDeck.frx":4947E
      Tag             =   "CardSelected"
      Top             =   2805
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image SelCard9C 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   8835
      Picture         =   "frmDeck.frx":4B5A7
      Tag             =   "CardSelected"
      Top             =   2805
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image SelCard8C 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   8835
      Picture         =   "frmDeck.frx":4D548
      Tag             =   "CardSelected"
      Top             =   2820
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image SelCard7C 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   8835
      Picture         =   "frmDeck.frx":4F379
      Tag             =   "CardSelected"
      Top             =   2820
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image SelCard6C 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   8835
      Picture         =   "frmDeck.frx":50E1C
      Tag             =   "CardSelected"
      Top             =   2820
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image SelCard5C 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   8820
      Picture         =   "frmDeck.frx":528B7
      Tag             =   "CardSelected"
      Top             =   2805
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image SelCard4C 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   8820
      Picture         =   "frmDeck.frx":53F58
      Tag             =   "CardSelected"
      Top             =   2835
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image SelCard3C 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   8835
      Picture         =   "frmDeck.frx":55425
      Tag             =   "CardSelected"
      Top             =   2835
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image SelCard2C 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   8850
      Picture         =   "frmDeck.frx":5DC67
      Tag             =   "CardSelected"
      Top             =   2820
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image SelCardAC 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   8835
      Picture         =   "frmDeck.frx":5EC94
      Tag             =   "CardSelected"
      Top             =   2820
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image BackSelected 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   51
      Left            =   8400
      Picture         =   "frmDeck.frx":5FACB
      Tag             =   "BackSelected"
      Top             =   720
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image BackSelected 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   50
      Left            =   8400
      Picture         =   "frmDeck.frx":62B5E
      Tag             =   "BackSelected"
      Top             =   720
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image BackSelected 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   49
      Left            =   8400
      Picture         =   "frmDeck.frx":65BF1
      Tag             =   "BackSelected"
      Top             =   720
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image BackSelected 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   48
      Left            =   8400
      Picture         =   "frmDeck.frx":68C84
      Tag             =   "BackSelected"
      Top             =   720
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image BackSelected 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   47
      Left            =   8400
      Picture         =   "frmDeck.frx":6BD17
      Tag             =   "BackSelected"
      Top             =   720
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image BackSelected 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   46
      Left            =   8400
      Picture         =   "frmDeck.frx":6EDAA
      Tag             =   "BackSelected"
      Top             =   720
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image BackSelected 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   45
      Left            =   8400
      Picture         =   "frmDeck.frx":71E3D
      Tag             =   "BackSelected"
      Top             =   720
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image BackSelected 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   44
      Left            =   8400
      Picture         =   "frmDeck.frx":74ED0
      Tag             =   "BackSelected"
      Top             =   720
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image BackSelected 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   43
      Left            =   8400
      Picture         =   "frmDeck.frx":77F63
      Tag             =   "BackSelected"
      Top             =   720
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image BackSelected 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   42
      Left            =   8400
      Picture         =   "frmDeck.frx":7AFF6
      Tag             =   "BackSelected"
      Top             =   720
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image BackSelected 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   41
      Left            =   8400
      Picture         =   "frmDeck.frx":7E089
      Tag             =   "BackSelected"
      Top             =   720
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image BackSelected 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   40
      Left            =   8400
      Picture         =   "frmDeck.frx":8111C
      Tag             =   "BackSelected"
      Top             =   720
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image BackSelected 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   39
      Left            =   8400
      Picture         =   "frmDeck.frx":841AF
      Tag             =   "BackSelected"
      Top             =   720
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image BackSelected 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   38
      Left            =   8400
      Picture         =   "frmDeck.frx":87242
      Tag             =   "BackSelected"
      Top             =   720
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image BackSelected 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   37
      Left            =   8400
      Picture         =   "frmDeck.frx":8A2D5
      Tag             =   "BackSelected"
      Top             =   720
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image BackSelected 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   36
      Left            =   8400
      Picture         =   "frmDeck.frx":8D368
      Tag             =   "BackSelected"
      Top             =   720
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image BackSelected 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   35
      Left            =   8400
      Picture         =   "frmDeck.frx":903FB
      Tag             =   "BackSelected"
      Top             =   720
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image BackSelected 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   34
      Left            =   8400
      Picture         =   "frmDeck.frx":9348E
      Tag             =   "BackSelected"
      Top             =   720
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image BackSelected 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   33
      Left            =   8400
      Picture         =   "frmDeck.frx":96521
      Tag             =   "BackSelected"
      Top             =   720
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image BackSelected 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   32
      Left            =   8400
      Picture         =   "frmDeck.frx":995B4
      Tag             =   "BackSelected"
      Top             =   720
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image BackSelected 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   31
      Left            =   8400
      Picture         =   "frmDeck.frx":9C647
      Tag             =   "BackSelected"
      Top             =   720
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image BackSelected 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   30
      Left            =   8400
      Picture         =   "frmDeck.frx":9F6DA
      Tag             =   "BackSelected"
      Top             =   720
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image BackSelected 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   29
      Left            =   8400
      Picture         =   "frmDeck.frx":A276D
      Tag             =   "BackSelected"
      Top             =   720
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image BackSelected 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   28
      Left            =   8400
      Picture         =   "frmDeck.frx":A5800
      Tag             =   "BackSelected"
      Top             =   720
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image BackSelected 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   27
      Left            =   8400
      Picture         =   "frmDeck.frx":A8893
      Tag             =   "BackSelected"
      Top             =   720
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image BackSelected 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   26
      Left            =   8400
      Picture         =   "frmDeck.frx":AB926
      Tag             =   "BackSelected"
      Top             =   720
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image BackSelected 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   25
      Left            =   8400
      Picture         =   "frmDeck.frx":AE9B9
      Tag             =   "BackSelected"
      Top             =   720
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image BackSelected 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   24
      Left            =   8400
      Picture         =   "frmDeck.frx":B1A4C
      Tag             =   "BackSelected"
      Top             =   720
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image BackSelected 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   23
      Left            =   8400
      Picture         =   "frmDeck.frx":B4ADF
      Tag             =   "BackSelected"
      Top             =   720
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image BackSelected 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   22
      Left            =   8400
      Picture         =   "frmDeck.frx":B7B72
      Tag             =   "BackSelected"
      Top             =   720
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image BackSelected 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   21
      Left            =   8400
      Picture         =   "frmDeck.frx":BAC05
      Tag             =   "BackSelected"
      Top             =   720
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image BackSelected 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   20
      Left            =   8400
      Picture         =   "frmDeck.frx":BDC98
      Tag             =   "BackSelected"
      Top             =   720
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image BackSelected 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   19
      Left            =   8400
      Picture         =   "frmDeck.frx":C0D2B
      Tag             =   "BackSelected"
      Top             =   720
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image BackSelected 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   18
      Left            =   8400
      Picture         =   "frmDeck.frx":C3DBE
      Tag             =   "BackSelected"
      Top             =   720
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image BackSelected 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   17
      Left            =   8400
      Picture         =   "frmDeck.frx":C6E51
      Tag             =   "BackSelected"
      Top             =   720
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image BackSelected 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   16
      Left            =   8400
      Picture         =   "frmDeck.frx":C9EE4
      Tag             =   "BackSelected"
      Top             =   720
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image BackSelected 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   15
      Left            =   8400
      Picture         =   "frmDeck.frx":CCF77
      Tag             =   "BackSelected"
      Top             =   720
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image BackSelected 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   14
      Left            =   8400
      Picture         =   "frmDeck.frx":D000A
      Tag             =   "BackSelected"
      Top             =   720
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image BackSelected 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   13
      Left            =   8400
      Picture         =   "frmDeck.frx":D309D
      Tag             =   "BackSelected"
      Top             =   720
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image BackSelected 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   12
      Left            =   8400
      Picture         =   "frmDeck.frx":D6130
      Tag             =   "BackSelected"
      Top             =   720
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image BackSelected 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   11
      Left            =   8400
      Picture         =   "frmDeck.frx":D91C3
      Tag             =   "BackSelected"
      Top             =   720
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image BackSelected 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   10
      Left            =   8400
      Picture         =   "frmDeck.frx":DC256
      Tag             =   "BackSelected"
      Top             =   720
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image BackSelected 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   9
      Left            =   8400
      Picture         =   "frmDeck.frx":DF2E9
      Tag             =   "BackSelected"
      Top             =   720
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image BackSelected 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   8
      Left            =   8400
      Picture         =   "frmDeck.frx":E237C
      Tag             =   "BackSelected"
      Top             =   720
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image BackSelected 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   7
      Left            =   8400
      Picture         =   "frmDeck.frx":E540F
      Tag             =   "BackSelected"
      Top             =   720
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image BackSelected 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   6
      Left            =   8400
      Picture         =   "frmDeck.frx":E84A2
      Tag             =   "BackSelected"
      Top             =   720
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image BackSelected 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   5
      Left            =   8400
      Picture         =   "frmDeck.frx":EB535
      Tag             =   "BackSelected"
      Top             =   720
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image BackSelected 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   4
      Left            =   8400
      Picture         =   "frmDeck.frx":EE5C8
      Tag             =   "BackSelected"
      Top             =   720
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image BackSelected 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   3
      Left            =   8400
      Picture         =   "frmDeck.frx":F165B
      Tag             =   "BackSelected"
      Top             =   720
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image BackSelected 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   2
      Left            =   8400
      Picture         =   "frmDeck.frx":F46EE
      Tag             =   "BackSelected"
      Top             =   720
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image BackSelected 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   1
      Left            =   8400
      Picture         =   "frmDeck.frx":F7781
      Tag             =   "BackSelected"
      Top             =   720
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image BackSelected 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   0
      Left            =   8400
      Picture         =   "frmDeck.frx":FA814
      Tag             =   "BackSelected"
      Top             =   735
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image Back 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   51
      Left            =   9330
      Picture         =   "frmDeck.frx":FD8A7
      Tag             =   "Back"
      Top             =   990
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image Back 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   50
      Left            =   9435
      Picture         =   "frmDeck.frx":10194B
      Tag             =   "Back"
      Top             =   990
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image Back 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   49
      Left            =   9390
      Picture         =   "frmDeck.frx":1059EF
      Tag             =   "Back"
      Top             =   915
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image Back 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   48
      Left            =   9375
      Picture         =   "frmDeck.frx":109A93
      Tag             =   "Back"
      Top             =   945
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image Back 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   47
      Left            =   9420
      Picture         =   "frmDeck.frx":10DB37
      Tag             =   "Back"
      Top             =   975
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image Back 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   46
      Left            =   9375
      Picture         =   "frmDeck.frx":111BDB
      Tag             =   "Back"
      Top             =   975
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image Back 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   45
      Left            =   9405
      Picture         =   "frmDeck.frx":115C7F
      Tag             =   "Back"
      Top             =   900
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image Back 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   44
      Left            =   9390
      Picture         =   "frmDeck.frx":119D23
      Tag             =   "Back"
      Top             =   915
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image Back 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   43
      Left            =   9390
      Picture         =   "frmDeck.frx":11DDC7
      Tag             =   "Back"
      Top             =   1005
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image Back 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   42
      Left            =   9390
      Picture         =   "frmDeck.frx":121E6B
      Tag             =   "Back"
      Top             =   930
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image Back 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   41
      Left            =   9390
      Picture         =   "frmDeck.frx":125F0F
      Tag             =   "Back"
      Top             =   885
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image Back 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   40
      Left            =   9375
      Picture         =   "frmDeck.frx":129FB3
      Tag             =   "Back"
      Top             =   915
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image Back 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   39
      Left            =   9330
      Picture         =   "frmDeck.frx":12E057
      Tag             =   "Back"
      Top             =   945
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image Back 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   38
      Left            =   9405
      Picture         =   "frmDeck.frx":1320FB
      Tag             =   "Back"
      Top             =   945
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image Back 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   37
      Left            =   9405
      Picture         =   "frmDeck.frx":13619F
      Tag             =   "Back"
      Top             =   960
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image Back 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   36
      Left            =   9420
      Picture         =   "frmDeck.frx":13A243
      Tag             =   "Back"
      Top             =   960
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image Back 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   35
      Left            =   9405
      Picture         =   "frmDeck.frx":13E2E7
      Tag             =   "Back"
      Top             =   1005
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image Back 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   34
      Left            =   9345
      Picture         =   "frmDeck.frx":14238B
      Tag             =   "Back"
      Top             =   870
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image Back 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   33
      Left            =   9300
      Picture         =   "frmDeck.frx":14642F
      Tag             =   "Back"
      Top             =   915
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image Back 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   32
      Left            =   9345
      Picture         =   "frmDeck.frx":14A4D3
      Tag             =   "Back"
      Top             =   945
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image Back 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   31
      Left            =   9315
      Picture         =   "frmDeck.frx":14E577
      Tag             =   "Back"
      Top             =   795
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image Back 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   30
      Left            =   9390
      Picture         =   "frmDeck.frx":15261B
      Tag             =   "Back"
      Top             =   870
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image Back 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   29
      Left            =   9405
      Picture         =   "frmDeck.frx":1566BF
      Tag             =   "Back"
      Top             =   915
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image Back 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   28
      Left            =   9360
      Picture         =   "frmDeck.frx":15A763
      Tag             =   "Back"
      Top             =   975
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image Back 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   27
      Left            =   9345
      Picture         =   "frmDeck.frx":15E807
      Tag             =   "Back"
      Top             =   885
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image Back 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   26
      Left            =   9345
      Picture         =   "frmDeck.frx":1628AB
      Tag             =   "Back"
      Top             =   930
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image Back 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   25
      Left            =   9405
      Picture         =   "frmDeck.frx":16694F
      Tag             =   "Back"
      Top             =   945
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image Back 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   24
      Left            =   9285
      Picture         =   "frmDeck.frx":16A9F3
      Tag             =   "Back"
      Top             =   915
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image Back 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   23
      Left            =   9330
      Picture         =   "frmDeck.frx":16EA97
      Tag             =   "Back"
      Top             =   930
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image Back 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   22
      Left            =   9345
      Picture         =   "frmDeck.frx":172B3B
      Tag             =   "Back"
      Top             =   960
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image Back 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   21
      Left            =   9315
      Picture         =   "frmDeck.frx":176BDF
      Tag             =   "Back"
      Top             =   885
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image Back 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   20
      Left            =   9345
      Picture         =   "frmDeck.frx":17AC83
      Tag             =   "Back"
      Top             =   915
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image Back 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   19
      Left            =   9345
      Picture         =   "frmDeck.frx":17ED27
      Tag             =   "Back"
      Top             =   1050
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image Back 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   18
      Left            =   9300
      Picture         =   "frmDeck.frx":182DCB
      Tag             =   "Back"
      Top             =   915
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image Back 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   17
      Left            =   9270
      Picture         =   "frmDeck.frx":186E6F
      Tag             =   "Back"
      Top             =   900
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image Back 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   16
      Left            =   9330
      Picture         =   "frmDeck.frx":18AF13
      Tag             =   "Back"
      Top             =   930
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image Back 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   15
      Left            =   9270
      Picture         =   "frmDeck.frx":18EFB7
      Tag             =   "Back"
      Top             =   945
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image Back 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   14
      Left            =   9405
      Picture         =   "frmDeck.frx":19305B
      Tag             =   "Back"
      Top             =   945
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image Back 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   13
      Left            =   9330
      Picture         =   "frmDeck.frx":1970FF
      Tag             =   "Back"
      Top             =   930
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image Back 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   12
      Left            =   9315
      Picture         =   "frmDeck.frx":19B1A3
      Tag             =   "Back"
      Top             =   885
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image Back 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   11
      Left            =   9300
      Picture         =   "frmDeck.frx":19F247
      Tag             =   "Back"
      Top             =   840
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image Back 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   10
      Left            =   9345
      Picture         =   "frmDeck.frx":1A32EB
      Tag             =   "Back"
      Top             =   885
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image Back 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   9
      Left            =   9270
      Picture         =   "frmDeck.frx":1A738F
      Tag             =   "Back"
      Top             =   915
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image Back 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   8
      Left            =   9360
      Picture         =   "frmDeck.frx":1AB433
      Tag             =   "Back"
      Top             =   900
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image Back 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   7
      Left            =   9345
      Picture         =   "frmDeck.frx":1AF4D7
      Tag             =   "Back"
      Top             =   945
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image Back 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   6
      Left            =   9330
      Picture         =   "frmDeck.frx":1B357B
      Tag             =   "Back"
      Top             =   900
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image Back 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   5
      Left            =   9300
      Picture         =   "frmDeck.frx":1B761F
      Tag             =   "Back"
      Top             =   1050
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image Back 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   4
      Left            =   9285
      Picture         =   "frmDeck.frx":1BB6C3
      Tag             =   "Back"
      Top             =   885
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image Back 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   3
      Left            =   9330
      Picture         =   "frmDeck.frx":1BF767
      Tag             =   "Back"
      Top             =   870
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image Back 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   2
      Left            =   9300
      Picture         =   "frmDeck.frx":1C380B
      Tag             =   "Back"
      Top             =   810
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image Back 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   0
      Left            =   9315
      Picture         =   "frmDeck.frx":1C78AF
      Tag             =   "Back"
      Top             =   450
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image Back 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   1
      Left            =   9390
      Picture         =   "frmDeck.frx":1CB953
      Tag             =   "Back"
      Top             =   930
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Label Position52 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "52"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   6480
      TabIndex        =   51
      Tag             =   "Position"
      Top             =   2520
      Width           =   150
   End
   Begin VB.Label Position51 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "51"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   6240
      TabIndex        =   50
      Tag             =   "Position"
      Top             =   2505
      Width           =   150
   End
   Begin VB.Label Position50 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "50"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5970
      TabIndex        =   49
      Tag             =   "Position"
      Top             =   2520
      Width           =   150
   End
   Begin VB.Label Position49 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "49"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5730
      TabIndex        =   48
      Tag             =   "Position"
      Top             =   2520
      Width           =   150
   End
   Begin VB.Label Position48 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "48"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5475
      TabIndex        =   47
      Tag             =   "Position"
      Top             =   2505
      Width           =   150
   End
   Begin VB.Label Position47 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "47"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5235
      TabIndex        =   46
      Tag             =   "Position"
      Top             =   2535
      Width           =   150
   End
   Begin VB.Label Position46 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "46"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5025
      TabIndex        =   45
      Tag             =   "Position"
      Top             =   2520
      Width           =   150
   End
   Begin VB.Label Position45 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "45"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   4785
      TabIndex        =   44
      Tag             =   "Position"
      Top             =   2535
      Width           =   150
   End
   Begin VB.Label Position44 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "44"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   4545
      TabIndex        =   43
      Tag             =   "Position"
      Top             =   2535
      Width           =   150
   End
   Begin VB.Label Position43 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "43"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   4290
      TabIndex        =   42
      Tag             =   "Position"
      Top             =   2520
      Width           =   150
   End
   Begin VB.Label Position42 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "42"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   4065
      TabIndex        =   41
      Tag             =   "Position"
      Top             =   2520
      Width           =   150
   End
   Begin VB.Label Position41 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "41"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3810
      TabIndex        =   40
      Tag             =   "Position"
      Top             =   2535
      Width           =   150
   End
   Begin VB.Label Position40 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "40"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3585
      TabIndex        =   39
      Tag             =   "Position"
      Top             =   2520
      Width           =   150
   End
   Begin VB.Label Position39 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "39"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3345
      TabIndex        =   38
      Tag             =   "Position"
      Top             =   2535
      Width           =   150
   End
   Begin VB.Label Position38 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "38"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3090
      TabIndex        =   37
      Tag             =   "Position"
      Top             =   2520
      Width           =   150
   End
   Begin VB.Label Position37 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "37"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2865
      TabIndex        =   36
      Tag             =   "Position"
      Top             =   2535
      Width           =   150
   End
   Begin VB.Label Position36 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "36"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2625
      TabIndex        =   35
      Tag             =   "Position"
      Top             =   2550
      Width           =   150
   End
   Begin VB.Label Position35 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "35"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2385
      TabIndex        =   34
      Tag             =   "Position"
      Top             =   2535
      Width           =   150
   End
   Begin VB.Label Position34 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "34"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2145
      TabIndex        =   33
      Tag             =   "Position"
      Top             =   2520
      Width           =   150
   End
   Begin VB.Label Position33 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "33"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1890
      TabIndex        =   32
      Tag             =   "Position"
      Top             =   2535
      Width           =   150
   End
   Begin VB.Label Position32 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "32"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1665
      TabIndex        =   31
      Tag             =   "Position"
      Top             =   2535
      Width           =   150
   End
   Begin VB.Label Position31 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "31"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1440
      TabIndex        =   30
      Tag             =   "Position"
      Top             =   2550
      Width           =   150
   End
   Begin VB.Label Position30 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "30"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1185
      TabIndex        =   29
      Tag             =   "Position"
      Top             =   2550
      Width           =   150
   End
   Begin VB.Label Position29 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "29"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   945
      TabIndex        =   28
      Tag             =   "Position"
      Top             =   2550
      Width           =   150
   End
   Begin VB.Label Position28 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "28"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   705
      TabIndex        =   27
      Tag             =   "Position"
      Top             =   2565
      Width           =   150
   End
   Begin VB.Label Position27 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "27"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   465
      TabIndex        =   26
      Tag             =   "Position"
      Top             =   2550
      Width           =   150
   End
   Begin VB.Label Position26 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "26"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   6450
      TabIndex        =   25
      Tag             =   "Position"
      Top             =   180
      Width           =   150
   End
   Begin VB.Label Position25 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "25"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   6210
      TabIndex        =   24
      Tag             =   "Position"
      Top             =   195
      Width           =   150
   End
   Begin VB.Label Position24 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "24"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5970
      TabIndex        =   23
      Tag             =   "Position"
      Top             =   210
      Width           =   150
   End
   Begin VB.Label Position23 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "23"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5760
      TabIndex        =   22
      Tag             =   "Position"
      Top             =   195
      Width           =   150
   End
   Begin VB.Label Position22 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "22"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5505
      TabIndex        =   21
      Tag             =   "Position"
      Top             =   180
      Width           =   150
   End
   Begin VB.Label Position21 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "21"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5250
      TabIndex        =   20
      Tag             =   "Position"
      Top             =   210
      Width           =   150
   End
   Begin VB.Label Position20 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "20"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   4995
      TabIndex        =   19
      Tag             =   "Position"
      Top             =   210
      Width           =   150
   End
   Begin VB.Label Position19 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "19"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   4785
      TabIndex        =   18
      Tag             =   "Position"
      Top             =   210
      Width           =   150
   End
   Begin VB.Label Position18 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "18"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   4530
      TabIndex        =   17
      Tag             =   "Position"
      Top             =   210
      Width           =   150
   End
   Begin VB.Label Position17 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "17"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   4290
      TabIndex        =   16
      Tag             =   "Position"
      Top             =   225
      Width           =   150
   End
   Begin VB.Label Position16 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "16"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   4050
      TabIndex        =   15
      Tag             =   "Position"
      Top             =   210
      Width           =   150
   End
   Begin VB.Label Position15 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "15"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3810
      TabIndex        =   14
      Tag             =   "Position"
      Top             =   210
      Width           =   150
   End
   Begin VB.Label Position14 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "14"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3585
      TabIndex        =   13
      Tag             =   "Position"
      Top             =   210
      Width           =   150
   End
   Begin VB.Label Position13 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "13"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3330
      TabIndex        =   12
      Tag             =   "Position"
      Top             =   210
      Width           =   150
   End
   Begin VB.Label Position12 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3090
      TabIndex        =   11
      Tag             =   "Position"
      Top             =   225
      Width           =   150
   End
   Begin VB.Label Position11 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "11"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2865
      TabIndex        =   10
      Tag             =   "Position"
      Top             =   210
      Width           =   150
   End
   Begin VB.Label Position10 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2625
      TabIndex        =   9
      Tag             =   "Position"
      Top             =   210
      Width           =   150
   End
   Begin VB.Label Position9 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2385
      TabIndex        =   8
      Tag             =   "Position"
      Top             =   210
      Width           =   90
   End
   Begin VB.Label Position8 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2130
      TabIndex        =   7
      Tag             =   "Position"
      Top             =   225
      Width           =   90
   End
   Begin VB.Label Position7 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1920
      TabIndex        =   6
      Tag             =   "Position"
      Top             =   210
      Width           =   90
   End
   Begin VB.Label Position6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1695
      TabIndex        =   5
      Tag             =   "Position"
      Top             =   210
      Width           =   90
   End
   Begin VB.Label Position5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1440
      TabIndex        =   4
      Tag             =   "Position"
      Top             =   225
      Width           =   90
   End
   Begin VB.Label Position4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1185
      TabIndex        =   3
      Tag             =   "Position"
      Top             =   240
      Width           =   90
   End
   Begin VB.Label Position3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   945
      TabIndex        =   2
      Tag             =   "Position"
      Top             =   225
      Width           =   90
   End
   Begin VB.Label Position2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   720
      TabIndex        =   1
      Tag             =   "Position"
      Top             =   225
      Width           =   90
   End
   Begin VB.Label Position1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   465
      TabIndex        =   0
      Tag             =   "Position"
      Top             =   210
      Width           =   90
   End
   Begin VB.Image CardAS 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   6405
      Picture         =   "frmDeck.frx":1CF9F7
      Tag             =   "Card"
      Top             =   2850
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image CardAH 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   6165
      Picture         =   "frmDeck.frx":1D0300
      Tag             =   "Card"
      Top             =   2850
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image CardAD 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   5925
      Picture         =   "frmDeck.frx":1D0AF9
      Tag             =   "Card"
      Top             =   2850
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image CardAC 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   5685
      Picture         =   "frmDeck.frx":1D12FA
      Tag             =   "Card"
      Top             =   2850
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image CardKS 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   5445
      Picture         =   "frmDeck.frx":1D1AEF
      Tag             =   "Card"
      Top             =   2850
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image CardKH 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   5205
      Picture         =   "frmDeck.frx":1D2B3D
      Tag             =   "Card"
      Top             =   2850
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image CardKD 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   4965
      Picture         =   "frmDeck.frx":1D3B92
      Tag             =   "Card"
      Top             =   2850
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image CardKC 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   4725
      Picture         =   "frmDeck.frx":1D4BBE
      Tag             =   "Card"
      Top             =   2850
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image CardQS 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   4485
      Picture         =   "frmDeck.frx":1D5B0A
      Tag             =   "Card"
      Top             =   2850
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image CardQH 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   4245
      Picture         =   "frmDeck.frx":1D6B09
      Tag             =   "Card"
      Top             =   2850
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image CardQD 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   4005
      Picture         =   "frmDeck.frx":1D7B28
      Tag             =   "Card"
      Top             =   2850
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image CardQC 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   3765
      Picture         =   "frmDeck.frx":1D8B96
      Tag             =   "Card"
      Top             =   2850
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image CardJS 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   3525
      Picture         =   "frmDeck.frx":1D9AFE
      Tag             =   "Card"
      Top             =   2850
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image CardJH 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   3285
      Picture         =   "frmDeck.frx":1DAA78
      Tag             =   "Card"
      Top             =   2850
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image CardJD 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   3045
      Picture         =   "frmDeck.frx":1DBA82
      Tag             =   "Card"
      Top             =   2850
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image CardJC 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   2805
      Picture         =   "frmDeck.frx":1DCAC0
      Tag             =   "Card"
      Top             =   2850
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image Card10S 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   2565
      Picture         =   "frmDeck.frx":1DDA06
      Tag             =   "Card"
      Top             =   2850
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image Card10H 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   2325
      Picture         =   "frmDeck.frx":1DE778
      Tag             =   "Card"
      Top             =   2850
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image Card10D 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   2085
      Picture         =   "frmDeck.frx":1DF4E9
      Tag             =   "Card"
      Top             =   2850
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image Card10C 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   1845
      Picture         =   "frmDeck.frx":1E0177
      Tag             =   "Card"
      Top             =   2850
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image Card9S 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   1605
      Picture         =   "frmDeck.frx":1E107C
      Tag             =   "Card"
      Top             =   2850
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image Card9H 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   1365
      Picture         =   "frmDeck.frx":1E1D4C
      Tag             =   "Card"
      Top             =   2850
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image Card9D 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   1125
      Picture         =   "frmDeck.frx":1E2A38
      Tag             =   "Card"
      Top             =   2850
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image Card9C 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   885
      Picture         =   "frmDeck.frx":1E366A
      Tag             =   "Card"
      Top             =   2850
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image Card8S 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   645
      Picture         =   "frmDeck.frx":1E44C3
      Tag             =   "Card"
      Top             =   2850
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image Card8H 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   405
      Picture         =   "frmDeck.frx":1E5112
      Tag             =   "Card"
      Top             =   2850
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image Card8D 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   6405
      Picture         =   "frmDeck.frx":1E5D60
      Tag             =   "Card"
      Top             =   495
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image Card8C 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   6165
      Picture         =   "frmDeck.frx":1E690D
      Tag             =   "Card"
      Top             =   495
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image Card7S 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   5925
      Picture         =   "frmDeck.frx":1E76D3
      Tag             =   "Card"
      Top             =   495
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image Card7H 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   5685
      Picture         =   "frmDeck.frx":1E8246
      Tag             =   "Card"
      Top             =   495
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image Card7D 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   5445
      Picture         =   "frmDeck.frx":1E8DB5
      Tag             =   "Card"
      Top             =   495
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image Card7C 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   5205
      Picture         =   "frmDeck.frx":1E98A9
      Tag             =   "Card"
      Top             =   495
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image Card6S 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   4965
      Picture         =   "frmDeck.frx":1EA52B
      Tag             =   "Card"
      Top             =   495
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image Card6H 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   4725
      Picture         =   "frmDeck.frx":1EB01F
      Tag             =   "Card"
      Top             =   495
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image Card6D 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   4485
      Picture         =   "frmDeck.frx":1EBB0A
      Tag             =   "Card"
      Top             =   495
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image Card6C 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   4245
      Picture         =   "frmDeck.frx":1EC57C
      Tag             =   "Card"
      Top             =   495
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image Card5S 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   4005
      Picture         =   "frmDeck.frx":1ED1DE
      Tag             =   "Card"
      Top             =   495
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image Card5H 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   3765
      Picture         =   "frmDeck.frx":1EDC8D
      Tag             =   "Card"
      Top             =   495
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image Card5D 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   3525
      Picture         =   "frmDeck.frx":1EE6B8
      Tag             =   "Card"
      Top             =   495
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image Card5C 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   3285
      Picture         =   "frmDeck.frx":1EF09A
      Tag             =   "Card"
      Top             =   495
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image Card4S 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   3045
      Picture         =   "frmDeck.frx":1EFBAB
      Tag             =   "Card"
      Top             =   495
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image Card4H 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   2805
      Picture         =   "frmDeck.frx":1F0599
      Tag             =   "Card"
      Top             =   495
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image Card4D 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   2565
      Picture         =   "frmDeck.frx":1F0F11
      Tag             =   "Card"
      Top             =   495
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image Card4C 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   2325
      Picture         =   "frmDeck.frx":1F1841
      Tag             =   "Card"
      Top             =   495
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image Card3S 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   2085
      Picture         =   "frmDeck.frx":1F227D
      Tag             =   "Card"
      Top             =   495
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image Card3H 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   1845
      Picture         =   "frmDeck.frx":1F2B82
      Tag             =   "Card"
      Top             =   495
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image Card3D 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   1605
      Picture         =   "frmDeck.frx":1F34B3
      Tag             =   "Card"
      Top             =   495
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image Card3C 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   1365
      Picture         =   "frmDeck.frx":1F3DA3
      Tag             =   "Card"
      Top             =   495
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image Card2S 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   1125
      Picture         =   "frmDeck.frx":1F46E9
      Tag             =   "Card"
      Top             =   495
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image Card2H 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   885
      Picture         =   "frmDeck.frx":1F4F83
      Tag             =   "Card"
      Top             =   495
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image Card2D 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   645
      Picture         =   "frmDeck.frx":1F580F
      Tag             =   "Card"
      Top             =   495
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image Card2C 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   405
      Picture         =   "frmDeck.frx":1F6051
      Tag             =   "Card"
      Top             =   495
      Visible         =   0   'False
      Width           =   1350
   End
End
Attribute VB_Name = "frmDeck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub DisplayPiles()
PokerCardsDealt = 0
PilesShown = 1
GilbreathShown = False
GilbreathActive = False
For i% = 1 To 8
    PileLabel(i%).Visible = False
Next i%
For Each Ctrl In Controls
    If Ctrl.Tag = "Card" Or Ctrl.Tag = "Position" Or Ctrl.Tag = "Back" _
        Or Ctrl.Tag = "BackSelected" Or Ctrl.Tag = "CardSelected" Then
        Ctrl.Visible = False
    End If
Next Ctrl
For i% = 1 To 52
    Deck(3, i%) = "Card" & Deck(2, i%)
    Deck(5, i%) = "Position" & Deck(1, i%)
    Back(i% - 1).ToolTipText = Deck(2, i%)
    BackSelected(i% - 1).ToolTipText = Deck(2, i%)
Next i%
'transfer the deck to the PileDeck array
For m% = 1 To DeckCount
    For z% = 1 To DeckProperties
        PileDeck(z%, m%) = Deck(z%, m%)
    Next z%
Next m%
'code for turning over piles for view from abpve
If frmPiles.ViewPilesAbove Then
    For k% = 1 To NumPiles
        For j% = PileTable(k%, 1) To PileTable(k%, 2)
            For p% = 1 To DeckProperties
                ChangedDeck(p%, j%) = PileDeck(p%, PileTable(k%, 2) - (j% - PileTable(k%, 1)))
            Next p%
        Next j%
    Next k%
    For m% = 1 To DeckCount
        For z% = 1 To DeckProperties
            PileDeck(z%, m%) = ChangedDeck(z%, m%)
        Next z%
    Next m%
    For i% = 1 To DeckCount
        'TOOLTIPXXX
        Back(i% - 1).ToolTipText = PileDeck(2, i%)
        BackSelected(i% - 1).ToolTipText = PileDeck(2, i%)
        'TOOLTIPXXX
        PileDeck(6, i%) = Not PileDeck(6, i%)
    Next i%
End If
'need code to set the size of window correctly based on piles
Me.Width = MaxWidth
Me.Height = MaxHeight
'NEW SECTION
For k% = 1 To NumPiles
        For j% = PileTable(k%, 1) To PileTable(k%, 2)
            For Each Ctrl In Controls
                If Ctrl.Tag = "Card" Then
                    If PileDeck(3, j%) = Ctrl.Name Then
                        Ctrl.Top = PileLocations(k%, 1)
                        Ctrl.Left = PileLocations(k%, 2) + 250 * (j% - PileTable(k%, 1))
                        Ctrl.ZOrder
                        If frmStackView.HighlightSelectionsCheck = 1 And _
                            PileDeck(4, j%) = "Selected" Then
                            If PileDeck(6, j%) = True Then
                                BackSelected(j% - 1).Top = PileLocations(k%, 1)
                                BackSelected(j% - 1).Left = PileLocations(k%, 2) + 250 * (j% - PileTable(k%, 1))
                                BackSelected(j% - 1).ZOrder
                            Else
                                For Each NewCtrl In Controls
                                    If NewCtrl.Tag = "CardSelected" Then
                                        If PileDeck(3, j%) = Mid(NewCtrl.Name, 4) Then
                                            NewCtrl.Top = PileLocations(k%, 1)
                                            NewCtrl.Left = PileLocations(k%, 2) + 250 * (j% - PileTable(k%, 1))
                                            NewCtrl.ZOrder
                                        End If
                                    End If
                                Next NewCtrl
                            End If
                        Else
                            If PileDeck(6, j%) = True Then
                                Back(j% - 1).Top = PileLocations(k%, 1)
                                Back(j% - 1).Left = PileLocations(k%, 2) + 250 * (j% - PileTable(k%, 1))
                                Back(j% - 1).ZOrder
                            End If
                        End If
                    End If
                End If
            Next Ctrl
            For Each Ctrl In Controls
                If Ctrl.Tag = "Position" Then
                    If PileDeck(5, j%) = Ctrl.Name Then
                        Ctrl.Top = PileLocations(k%, 1) - 240
                        Ctrl.Left = PileLocations(k%, 2) + 250 * (j% - PileTable(k%, 1)) + 50
                        Ctrl.ZOrder
                        If frmStackView.HighlightSelectionsCheck = 1 Then
                            If PileDeck(4, j%) = "Selected" Then
                                Ctrl.Font.Bold = True
                                Ctrl.Font.Size = 12
                                Ctrl.Top = PileLocations(k%, 1) - 340
                                Ctrl.Left = PileLocations(k%, 2) + 250 * (j% - PileTable(k%, 1)) - 10 + 50
                            Else
                                Ctrl.Font.Bold = False
                                Ctrl.Font.Size = 8
                            End If
                        Else
                            Ctrl.Font.Bold = False
                            Ctrl.Font.Size = 8
                        End If
                    End If
                End If
            Next Ctrl
            'adjust card positions based on viewing direction
            If frmPiles.ViewPilesAbove Then
                If frmStackView.CountFromBack = True Then
                    Ordinal(PileTable(k%, 2) - (j% - PileTable(k%, 1))).Top = _
                        PileLocations(k%, 1) + 1925
                    Ordinal(PileTable(k%, 2) - (j% - PileTable(k%, 1))).Left = _
                        PileLocations(k%, 2) + 250 * (j% - PileTable(k%, 1)) + 50
                ElseIf frmStackView.CountFromFace = True Then
                    Ordinal(53 - (PileTable(k%, 2) - (j% - PileTable(k%, 1)))).Top = _
                        PileLocations(k%, 1) + 1925
                    Ordinal(53 - (PileTable(k%, 2) - (j% - PileTable(k%, 1)))).Left = _
                        PileLocations(k%, 2) + 250 * (j% - PileTable(k%, 1)) + 50
                End If
                If frmStackView.ShowPositionValues = 1 Then
                    Ordinal(j%).Visible = True
                Else
                    Ordinal(j%).Visible = False
                End If
            Else
                If frmStackView.CountFromBack = True Then
                    Ordinal(j%).Top = PileLocations(k%, 1) + 1925
                    Ordinal(j%).Left = PileLocations(k%, 2) + 250 * (j% - PileTable(k%, 1)) + 50
                ElseIf frmStackView.CountFromFace = True Then
                    Ordinal(53 - j%).Top = PileLocations(k%, 1) + 1925
                    Ordinal(53 - j%).Left = PileLocations(k%, 2) + 250 * (j% - PileTable(k%, 1)) + 50
                End If
                If frmStackView.ShowPositionValues = 1 Then
                    Ordinal(j%).Visible = True
                Else
                    Ordinal(j%).Visible = False
                End If
            End If
        Next j%
    PileLabel(k%).Top = PileLocations(k%, 1) + 825
    PileLabel(k%).Left = PileLocations(k%, 2) - 150
    PileLabel(k%).Visible = True
Next k%
'show the right cards
For j% = 1 To 52
    For Each Ctrl In Controls
        If Ctrl.Tag = "Card" Then
            If PileDeck(3, j%) = Ctrl.Name Then
                If frmStackView.HighlightSelectionsCheck = 1 And _
                    PileDeck(4, j%) = "Selected" Then
                    If PileDeck(6, j%) = True Then
                        Ctrl.Visible = False
                        BackSelected(j% - 1).Visible = True
                    Else
                        BackSelected(j% - 1).Visible = False
                        Ctrl.Visible = False
                        For Each NewCtrl In Controls
                            If NewCtrl.Tag = "CardSelected" Then
                                If PileDeck(3, j%) = Mid(NewCtrl.Name, 4) Then
                                    NewCtrl.Visible = True
                                End If
                            End If
                        Next NewCtrl
                    End If
                Else
                    If PileDeck(6, j%) = True Then
                        Ctrl.Visible = False
                        Back(j% - 1).Visible = True
                    Else
                        Back(j% - 1).Visible = False
                        Ctrl.Visible = True
                    End If
                End If
            End If
        End If
    Next Ctrl
Next j%
For Each Ctrl In Controls
    If Ctrl.Tag = "Position" Then
        If frmStackView.ShowIndexValues = 1 Then
            Ctrl.Visible = True
        Else
            Ctrl.Visible = False
        End If
    End If
Next Ctrl
'ShuffleMeter Setting
If frmMain.mnuShuffleMeter.Checked = True Then
    frmShuffleMeter.SetShuffleMeterParameters
End If
End Sub

Public Sub DisplayPilesKeepGilbreathActive()
PokerCardsDealt = 0
PilesShown = 1
GilbreathShown = False
For i% = 1 To 8
    PileLabel(i%).Visible = False
Next i%
For Each Ctrl In Controls
    If Ctrl.Tag = "Card" Or Ctrl.Tag = "Position" Or Ctrl.Tag = "Back" _
        Or Ctrl.Tag = "BackSelected" Or Ctrl.Tag = "CardSelected" Then
        Ctrl.Visible = False
    End If
Next Ctrl
For i% = 1 To 52
    Deck(3, i%) = "Card" & Deck(2, i%)
    Deck(5, i%) = "Position" & Deck(1, i%)
    Back(i% - 1).ToolTipText = Deck(2, i%)
    BackSelected(i% - 1).ToolTipText = Deck(2, i%)
Next i%
'transfer the deck to the PileDeck array
For m% = 1 To DeckCount
    For z% = 1 To DeckProperties
        PileDeck(z%, m%) = Deck(z%, m%)
    Next z%
Next m%
'code for turning over piles for view from abpve
If frmPiles.ViewPilesAbove Then
    For k% = 1 To NumPiles
        For j% = PileTable(k%, 1) To PileTable(k%, 2)
            For p% = 1 To DeckProperties
                ChangedDeck(p%, j%) = PileDeck(p%, PileTable(k%, 2) - (j% - PileTable(k%, 1)))
            Next p%
        Next j%
    Next k%
    For m% = 1 To DeckCount
        For z% = 1 To DeckProperties
            PileDeck(z%, m%) = ChangedDeck(z%, m%)
        Next z%
    Next m%
    For i% = 1 To DeckCount
        'TOOLTIPXXX
        Back(i% - 1).ToolTipText = PileDeck(2, i%)
        BackSelected(i% - 1).ToolTipText = PileDeck(2, i%)
        'TOOLTIPXXX
        PileDeck(6, i%) = Not PileDeck(6, i%)
    Next i%
End If
'need code to set the size of window correctly based on piles
Me.Width = MaxWidth
Me.Height = MaxHeight
'NEW SECTION
For k% = 1 To NumPiles
        For j% = PileTable(k%, 1) To PileTable(k%, 2)
            For Each Ctrl In Controls
                If Ctrl.Tag = "Card" Then
                    If PileDeck(3, j%) = Ctrl.Name Then
                        Ctrl.Top = PileLocations(k%, 1)
                        Ctrl.Left = PileLocations(k%, 2) + 250 * (j% - PileTable(k%, 1))
                        Ctrl.ZOrder
                        If frmStackView.HighlightSelectionsCheck = 1 And _
                            PileDeck(4, j%) = "Selected" Then
                            If PileDeck(6, j%) = True Then
                                BackSelected(j% - 1).Top = PileLocations(k%, 1)
                                BackSelected(j% - 1).Left = PileLocations(k%, 2) + 250 * (j% - PileTable(k%, 1))
                                BackSelected(j% - 1).ZOrder
                            Else
                                For Each NewCtrl In Controls
                                    If NewCtrl.Tag = "CardSelected" Then
                                        If PileDeck(3, j%) = Mid(NewCtrl.Name, 4) Then
                                            NewCtrl.Top = PileLocations(k%, 1)
                                            NewCtrl.Left = PileLocations(k%, 2) + 250 * (j% - PileTable(k%, 1))
                                            NewCtrl.ZOrder
                                        End If
                                    End If
                                Next NewCtrl
                            End If
                        Else
                            If PileDeck(6, j%) = True Then
                                Back(j% - 1).Top = PileLocations(k%, 1)
                                Back(j% - 1).Left = PileLocations(k%, 2) + 250 * (j% - PileTable(k%, 1))
                                Back(j% - 1).ZOrder
                            End If
                        End If
                    End If
                End If
            Next Ctrl
            For Each Ctrl In Controls
                If Ctrl.Tag = "Position" Then
                    If PileDeck(5, j%) = Ctrl.Name Then
                        Ctrl.Top = PileLocations(k%, 1) - 240
                        Ctrl.Left = PileLocations(k%, 2) + 250 * (j% - PileTable(k%, 1)) + 50
                        Ctrl.ZOrder
                        If frmStackView.HighlightSelectionsCheck = 1 Then
                            If PileDeck(4, j%) = "Selected" Then
                                Ctrl.Font.Bold = True
                                Ctrl.Font.Size = 12
                                Ctrl.Top = PileLocations(k%, 1) - 340
                                Ctrl.Left = PileLocations(k%, 2) + 250 * (j% - PileTable(k%, 1)) - 10 + 50
                            Else
                                Ctrl.Font.Bold = False
                                Ctrl.Font.Size = 8
                            End If
                        Else
                            Ctrl.Font.Bold = False
                            Ctrl.Font.Size = 8
                        End If
                    End If
                End If
            Next Ctrl
            'adjust card positions based on viewing direction
            If frmPiles.ViewPilesAbove Then
                If frmStackView.CountFromBack = True Then
                    Ordinal(PileTable(k%, 2) - (j% - PileTable(k%, 1))).Top = _
                        PileLocations(k%, 1) + 1925
                    Ordinal(PileTable(k%, 2) - (j% - PileTable(k%, 1))).Left = _
                        PileLocations(k%, 2) + 250 * (j% - PileTable(k%, 1)) + 50
                ElseIf frmStackView.CountFromFace = True Then
                    Ordinal(53 - (PileTable(k%, 2) - (j% - PileTable(k%, 1)))).Top = _
                        PileLocations(k%, 1) + 1925
                    Ordinal(53 - (PileTable(k%, 2) - (j% - PileTable(k%, 1)))).Left = _
                        PileLocations(k%, 2) + 250 * (j% - PileTable(k%, 1)) + 50
                End If
                If frmStackView.ShowPositionValues = 1 Then
                    Ordinal(j%).Visible = True
                Else
                    Ordinal(j%).Visible = False
                End If
            Else
                If frmStackView.CountFromBack = True Then
                    Ordinal(j%).Top = PileLocations(k%, 1) + 1925
                    Ordinal(j%).Left = PileLocations(k%, 2) + 250 * (j% - PileTable(k%, 1)) + 50
                ElseIf frmStackView.CountFromFace = True Then
                    Ordinal(53 - j%).Top = PileLocations(k%, 1) + 1925
                    Ordinal(53 - j%).Left = PileLocations(k%, 2) + 250 * (j% - PileTable(k%, 1)) + 50
                End If
                If frmStackView.ShowPositionValues = 1 Then
                    Ordinal(j%).Visible = True
                Else
                    Ordinal(j%).Visible = False
                End If
            End If
        Next j%
    PileLabel(k%).Top = PileLocations(k%, 1) + 825
    PileLabel(k%).Left = PileLocations(k%, 2) - 150
    PileLabel(k%).Visible = True
Next k%
'show the right cards
For j% = 1 To 52
    For Each Ctrl In Controls
        If Ctrl.Tag = "Card" Then
            If PileDeck(3, j%) = Ctrl.Name Then
                If frmStackView.HighlightSelectionsCheck = 1 And _
                    PileDeck(4, j%) = "Selected" Then
                    If PileDeck(6, j%) = True Then
                        Ctrl.Visible = False
                        BackSelected(j% - 1).Visible = True
                    Else
                        BackSelected(j% - 1).Visible = False
                        Ctrl.Visible = False
                        For Each NewCtrl In Controls
                            If NewCtrl.Tag = "CardSelected" Then
                                If PileDeck(3, j%) = Mid(NewCtrl.Name, 4) Then
                                    NewCtrl.Visible = True
                                End If
                            End If
                        Next NewCtrl
                    End If
                Else
                    If PileDeck(6, j%) = True Then
                        Ctrl.Visible = False
                        Back(j% - 1).Visible = True
                    Else
                        Back(j% - 1).Visible = False
                        Ctrl.Visible = True
                    End If
                End If
            End If
        End If
    Next Ctrl
Next j%
For Each Ctrl In Controls
    If Ctrl.Tag = "Position" Then
        If frmStackView.ShowIndexValues = 1 Then
            Ctrl.Visible = True
        Else
            Ctrl.Visible = False
        End If
    End If
Next Ctrl
'ShuffleMeter Setting
If frmMain.mnuShuffleMeter.Checked = True Then
    frmShuffleMeter.SetShuffleMeterParameters
End If
End Sub


Public Sub DisplayPilesGilbreath()
PokerCardsDealt = 0
PilesShown = 1
GilbreathShown = True
For i% = 1 To 8
    PileLabel(i%).Visible = False
Next i%
For Each Ctrl In Controls
    If Ctrl.Tag = "Card" Or Ctrl.Tag = "Position" Or Ctrl.Tag = "Back" _
        Or Ctrl.Tag = "BackSelected" Or Ctrl.Tag = "CardSelected" Then
        Ctrl.Visible = False
    End If
Next Ctrl
For i% = 1 To 52
    Deck(3, i%) = "Card" & Deck(2, i%)
    Deck(5, i%) = "Position" & Deck(1, i%)
    Back(i% - 1).ToolTipText = Deck(2, i%)
    BackSelected(i% - 1).ToolTipText = Deck(2, i%)
Next i%
'transfer the deck to the PileDeck array
For m% = 1 To DeckCount
    For z% = 1 To DeckProperties
        PileDeck(z%, m%) = Deck(z%, m%)
    Next z%
Next m%
'code for turning over piles for view from abpve
If frmPiles.ViewPilesAbove Then
    For k% = 1 To NumPiles
        For j% = PileTable(k%, 1) To PileTable(k%, 2)
            For p% = 1 To DeckProperties
                ChangedDeck(p%, j%) = PileDeck(p%, PileTable(k%, 2) - (j% - PileTable(k%, 1)))
            Next p%
        Next j%
    Next k%
    For m% = 1 To DeckCount
        For z% = 1 To DeckProperties
            PileDeck(z%, m%) = ChangedDeck(z%, m%)
        Next z%
    Next m%
    For i% = 1 To DeckCount
        'TOOLTIPXXX
        Back(i% - 1).ToolTipText = PileDeck(2, i%)
        BackSelected(i% - 1).ToolTipText = PileDeck(2, i%)
        'TOOLTIPXXX
        PileDeck(6, i%) = Not PileDeck(6, i%)
    Next i%
End If
'need code to set the size of window correctly based on piles
Me.Width = MaxWidth
Me.Height = MaxHeight
'NEW SECTION
For k% = 1 To NumPiles
        GilbreathOffset = 0
        If k% = GilbreathPileNum Then
            GilbreathOffset = 400
        End If
        For j% = PileTable(k%, 1) To PileTable(k%, 2)
            For Each Ctrl In Controls
                If Ctrl.Tag = "Card" Then
                    If PileDeck(3, j%) = Ctrl.Name Then
                        If GilbreathDeck(j%) Then
                            Ctrl.Top = PileLocations(k%, 1) - GilbreathOffset
                        Else
                            Ctrl.Top = PileLocations(k%, 1)
                        End If
                        Ctrl.Left = PileLocations(k%, 2) + 250 * (j% - PileTable(k%, 1))
                        Ctrl.ZOrder
                        If frmStackView.HighlightSelectionsCheck = 1 And _
                            PileDeck(4, j%) = "Selected" Then
                            If PileDeck(6, j%) = True Then
                                If GilbreathDeck(j%) Then
                                    BackSelected(j% - 1).Top = PileLocations(k%, 1) - GilbreathOffset
                                Else
                                    BackSelected(j% - 1).Top = PileLocations(k%, 1)
                                End If
                                BackSelected(j% - 1).Left = PileLocations(k%, 2) + 250 * (j% - PileTable(k%, 1))
                                BackSelected(j% - 1).ZOrder
                            Else
                                For Each NewCtrl In Controls
                                    If NewCtrl.Tag = "CardSelected" Then
                                        If PileDeck(3, j%) = Mid(NewCtrl.Name, 4) Then
                                            If GilbreathDeck(j%) Then
                                                NewCtrl.Top = PileLocations(k%, 1) - GilbreathOffset
                                            Else
                                                NewCtrl.Top = PileLocations(k%, 1)
                                            End If
                                            NewCtrl.Left = PileLocations(k%, 2) + 250 * (j% - PileTable(k%, 1))
                                            NewCtrl.ZOrder
                                        End If
                                    End If
                                Next NewCtrl
                            End If
                        Else
                            If PileDeck(6, j%) = True Then
                                If GilbreathDeck(j%) Then
                                    Back(j% - 1).Top = PileLocations(k%, 1) - GilbreathOffset
                                Else
                                    Back(j% - 1).Top = PileLocations(k%, 1)
                                End If
                                Back(j% - 1).Left = PileLocations(k%, 2) + 250 * (j% - PileTable(k%, 1))
                                Back(j% - 1).ZOrder
                            End If
                        End If
                    End If
                End If
            Next Ctrl
            For Each Ctrl In Controls
                If Ctrl.Tag = "Position" Then
                    If PileDeck(5, j%) = Ctrl.Name Then
                        Ctrl.Top = PileLocations(k%, 1) - 240 - GilbreathOffset
                        Ctrl.Left = PileLocations(k%, 2) + 250 * (j% - PileTable(k%, 1)) + 50
                        Ctrl.ZOrder
                        If frmStackView.HighlightSelectionsCheck = 1 Then
                            If PileDeck(4, j%) = "Selected" Then
                                Ctrl.Font.Bold = True
                                Ctrl.Font.Size = 12
                                Ctrl.Top = PileLocations(k%, 1) - 340 - GilbreathOffset
                                Ctrl.Left = PileLocations(k%, 2) + 250 * (j% - PileTable(k%, 1)) - 10 + 50
                            Else
                                Ctrl.Font.Bold = False
                                Ctrl.Font.Size = 8
                            End If
                        Else
                            Ctrl.Font.Bold = False
                            Ctrl.Font.Size = 8
                        End If
                    End If
                End If
            Next Ctrl
            'adjust card positions based on viewing direction
            If frmPiles.ViewPilesAbove Then
                If frmStackView.CountFromBack = True Then
                    Ordinal(PileTable(k%, 2) - (j% - PileTable(k%, 1))).Top = _
                        PileLocations(k%, 1) + 1925
                    Ordinal(PileTable(k%, 2) - (j% - PileTable(k%, 1))).Left = _
                        PileLocations(k%, 2) + 250 * (j% - PileTable(k%, 1)) + 50
                ElseIf frmStackView.CountFromFace = True Then
                    Ordinal(53 - (PileTable(k%, 2) - (j% - PileTable(k%, 1)))).Top = _
                        PileLocations(k%, 1) + 1925
                    Ordinal(53 - (PileTable(k%, 2) - (j% - PileTable(k%, 1)))).Left = _
                        PileLocations(k%, 2) + 250 * (j% - PileTable(k%, 1)) + 50
                End If
                If frmStackView.ShowPositionValues = 1 Then
                    Ordinal(j%).Visible = True
                Else
                    Ordinal(j%).Visible = False
                End If
            Else
                If frmStackView.CountFromBack = True Then
                    Ordinal(j%).Top = PileLocations(k%, 1) + 1925
                    Ordinal(j%).Left = PileLocations(k%, 2) + 250 * (j% - PileTable(k%, 1)) + 50
                ElseIf frmStackView.CountFromFace = True Then
                    Ordinal(53 - j%).Top = PileLocations(k%, 1) + 1925
                    Ordinal(53 - j%).Left = PileLocations(k%, 2) + 250 * (j% - PileTable(k%, 1)) + 50
                End If
                If frmStackView.ShowPositionValues = 1 Then
                    Ordinal(j%).Visible = True
                Else
                    Ordinal(j%).Visible = False
                End If
            End If
        Next j%
    PileLabel(k%).Top = PileLocations(k%, 1) + 825
    PileLabel(k%).Left = PileLocations(k%, 2) - 150
    PileLabel(k%).Visible = True
Next k%
'show the right cards
For j% = 1 To 52
    For Each Ctrl In Controls
        If Ctrl.Tag = "Card" Then
            If PileDeck(3, j%) = Ctrl.Name Then
                If frmStackView.HighlightSelectionsCheck = 1 And _
                    PileDeck(4, j%) = "Selected" Then
                    If PileDeck(6, j%) = True Then
                        Ctrl.Visible = False
                        BackSelected(j% - 1).Visible = True
                    Else
                        BackSelected(j% - 1).Visible = False
                        Ctrl.Visible = False
                        For Each NewCtrl In Controls
                            If NewCtrl.Tag = "CardSelected" Then
                                If PileDeck(3, j%) = Mid(NewCtrl.Name, 4) Then
                                    NewCtrl.Visible = True
                                End If
                            End If
                        Next NewCtrl
                    End If
                Else
                    If PileDeck(6, j%) = True Then
                        Ctrl.Visible = False
                        Back(j% - 1).Visible = True
                    Else
                        Back(j% - 1).Visible = False
                        Ctrl.Visible = True
                    End If
                End If
            End If
        End If
    Next Ctrl
Next j%
For Each Ctrl In Controls
    If Ctrl.Tag = "Position" Then
        If frmStackView.ShowIndexValues = 1 Then
            Ctrl.Visible = True
        Else
            Ctrl.Visible = False
        End If
    End If
Next Ctrl
GilbreathActive = True
If frmMain.mnuShuffleMeter.Checked = True Then
    frmShuffleMeter.SetShuffleMeterParameters
End If
End Sub


Public Sub DisplayCards()
If StackViewLoading Then
    Me.Visible = False
Else
    Me.Visible = True
End If
PokerCardsDealt = 0
PilesShown = 0
GilbreathShown = False
GilbreathActive = False
For i% = 1 To 8
    PileLabel(i%).Visible = False
Next i%
Me.Width = 8265
Me.Height = 6095
For Each Ctrl In Controls
    If Ctrl.Tag = "Card" Or Ctrl.Tag = "Position" Or Ctrl.Tag = "Back" _
        Or Ctrl.Tag = "BackSelected" Or Ctrl.Tag = "CardSelected" Then
        Ctrl.Visible = False
    End If
Next Ctrl
For i% = 1 To 52
    Deck(3, i%) = "Card" & Deck(2, i%)
    Deck(5, i%) = "Position" & Deck(1, i%)
    'Back(i% - 1).ToolTipText = Deck(2, i%)
    'BackSelected(i% - 1).ToolTipText = Deck(2, i%)
Next i%
'if the StackView window is not open, it will open, and I need to
'correct the View Menu checkmark on the Control option.
If frmMain.mnuControl.Checked = False Then
    frmMain.mnuControl.Checked = True
End If
'transfer cards to DisplayDeck
If frmStackView.ViewDeckAbove Then
    For i% = 1 To 26
        For p% = 1 To DeckProperties
            ChangedDeck(p%, i%) = Deck(p%, 26 - i% + 1)
        Next p%
    Next i%
    For i% = 27 To 52
        For p% = 1 To DeckProperties
            ChangedDeck(p%, i%) = Deck(p%, 52 - i% + 27)
        Next p%
    Next i%
    For i% = 1 To DeckCount
        For p% = 1 To DeckProperties
            DisplayDeck(p%, i%) = ChangedDeck(p%, i%)
        Next p%
    Next i%
    For i% = 1 To DeckCount
        Back(i% - 1).ToolTipText = DisplayDeck(2, i%)
        BackSelected(i% - 1).ToolTipText = DisplayDeck(2, i%)
        DisplayDeck(6, i%) = Not DisplayDeck(6, i%)
    Next i%
ElseIf frmStackView.ViewDeckBeneath Then
    For i% = 1 To DeckCount
        For p% = 1 To DeckProperties
            DisplayDeck(p%, i%) = Deck(p%, i%)
        Next p%
    Next i%
    For i% = 1 To DeckCount
        Back(i% - 1).ToolTipText = DisplayDeck(2, i%)
        BackSelected(i% - 1).ToolTipText = DisplayDeck(2, i%)
    Next i%
End If
'NEW SECTION
For j% = 1 To 26
    For Each Ctrl In Controls
        If Ctrl.Tag = "Card" Then
            If DisplayDeck(3, j%) = Ctrl.Name Then
                Ctrl.Top = 400
                Ctrl.Left = 250 * j%
                Ctrl.ZOrder
                If frmStackView.HighlightSelectionsCheck = 1 And _
                    DisplayDeck(4, j%) = "Selected" Then
                    If DisplayDeck(6, j%) = True Then
                        BackSelected(j% - 1).Top = 400
                        BackSelected(j% - 1).Left = 250 * j%
                        BackSelected(j% - 1).ZOrder
                    Else
                        For Each NewCtrl In Controls
                            If NewCtrl.Tag = "CardSelected" Then
                                If DisplayDeck(3, j%) = Mid(NewCtrl.Name, 4) Then
                                    NewCtrl.Top = 400
                                    NewCtrl.Left = 250 * j%
                                    NewCtrl.ZOrder
                                End If
                            End If
                        Next NewCtrl
                    End If
                Else
                    If DisplayDeck(6, j%) = True Then
                        Back(j% - 1).Top = 400
                        Back(j% - 1).Left = 250 * j%
                        Back(j% - 1).ZOrder
                    End If
                End If
            End If
        End If
    Next Ctrl
    For Each Ctrl In Controls
        If Ctrl.Tag = "Position" Then
            If DisplayDeck(5, j%) = Ctrl.Name Then
                Ctrl.Top = 160
                Ctrl.Left = 250 * j% + 50
                Ctrl.ZOrder
                If frmStackView.HighlightSelectionsCheck = 1 Then
                    If DisplayDeck(4, j%) = "Selected" Then
                        Ctrl.Font.Bold = True
                        Ctrl.Font.Size = 12
                        Ctrl.Top = 60
                        Ctrl.Left = 250 * j% - 10 + 50
                    Else
                        Ctrl.Font.Bold = False
                        Ctrl.Font.Size = 8
                    End If
                Else
                    Ctrl.Font.Bold = False
                    Ctrl.Font.Size = 8
                End If
            End If
        End If
    Next Ctrl
    If frmStackView.ViewDeckBeneath Then
        If frmStackView.CountFromBack = True Then
            Ordinal(j%).Top = 2325
            Ordinal(j%).Left = 250 * j% + 50
        ElseIf frmStackView.CountFromFace = True Then
            Ordinal(53 - j%).Top = 2325
            Ordinal(53 - j%).Left = 250 * j% + 50
        End If
    ElseIf frmStackView.ViewDeckAbove Then
        If frmStackView.CountFromBack = True Then
            Ordinal(27 - j%).Top = 2325
            Ordinal(27 - j%).Left = 250 * j% + 50
        ElseIf frmStackView.CountFromFace = True Then
            Ordinal(j% + 26).Top = 2325
            Ordinal(j% + 26).Left = 250 * j% + 50
        End If
    End If
    If frmStackView.ShowPositionValues = 1 Then
        Ordinal(j%).Visible = True
    Else
        Ordinal(j%).Visible = False
    End If
Next j%
'NEW SECTION
For k% = 27 To 52
    For Each Ctrl In Controls
        If Ctrl.Tag = "Card" Then
            If DisplayDeck(3, k%) = Ctrl.Name Then
                Ctrl.Top = 3200
                Ctrl.Left = 250 * (k% - 26)
                Ctrl.ZOrder
                If frmStackView.HighlightSelectionsCheck = 1 And _
                    DisplayDeck(4, k%) = "Selected" Then
                    If DisplayDeck(6, k%) = True Then
                        BackSelected(k% - 1).Top = 3200
                        BackSelected(k% - 1).Left = 250 * (k% - 26)
                        BackSelected(k% - 1).ZOrder
                    Else
                        For Each NewCtrl In Controls
                            If NewCtrl.Tag = "CardSelected" Then
                                If DisplayDeck(3, k%) = Mid(NewCtrl.Name, 4) Then
                                    NewCtrl.Top = 3200
                                    NewCtrl.Left = 250 * (k% - 26)
                                    NewCtrl.ZOrder
                                End If
                            End If
                        Next NewCtrl
                    End If
                Else
                    If DisplayDeck(6, k%) = True Then
                        Back(k% - 1).Top = 3200
                        Back(k% - 1).Left = 250 * (k% - 26)
                        Back(k% - 1).ZOrder
                    End If
                End If
            End If
        End If
    Next Ctrl
    For Each Ctrl In Controls
        If Ctrl.Tag = "Position" Then
            If DisplayDeck(5, k%) = Ctrl.Name Then
                Ctrl.Top = 2960
                Ctrl.Left = 250 * (k% - 26) + 50
                Ctrl.ZOrder
                If frmStackView.HighlightSelectionsCheck = 1 Then
                    If DisplayDeck(4, k%) = "Selected" Then
                        Ctrl.Font.Bold = True
                        Ctrl.Font.Size = 12
                        Ctrl.Top = 2860
                        Ctrl.Left = 250 * (k% - 26) - 10 + 50
                    Else
                        Ctrl.Font.Bold = False
                        Ctrl.Font.Size = 8
                    End If
                Else
                    Ctrl.Font.Bold = False
                    Ctrl.Font.Size = 8
                End If
            End If
        End If
    Next Ctrl
    If frmStackView.ViewDeckBeneath Then
        If frmStackView.CountFromBack = True Then
            Ordinal(k%).Top = 5125
            Ordinal(k%).Left = 250 * (k% - 26) + 50
        ElseIf frmStackView.CountFromFace = True Then
            Ordinal(53 - k%).Top = 5125
            Ordinal(53 - k%).Left = 250 * (k% - 26) + 50
        End If
    ElseIf frmStackView.ViewDeckAbove Then
        If frmStackView.CountFromBack = True Then
            Ordinal(52 - k% + 27).Top = 5125
            Ordinal(52 - k% + 27).Left = 250 * (k% - 26) + 50
        ElseIf frmStackView.CountFromFace = True Then
            Ordinal(k% - 26).Top = 5125
            Ordinal(k% - 26).Left = 250 * (k% - 26) + 50
        End If
    End If
    If frmStackView.ShowPositionValues = 1 Then
        Ordinal(k%).Visible = True
    Else
        Ordinal(k%).Visible = False
    End If
Next k%
'Show the appropriate cards only
'NEW SECTION
For j% = 1 To 52
    For Each Ctrl In Controls
        If Ctrl.Tag = "Card" Then
            If DisplayDeck(3, j%) = Ctrl.Name Then
                If frmStackView.HighlightSelectionsCheck = 1 And _
                    DisplayDeck(4, j%) = "Selected" Then
                    If DisplayDeck(6, j%) = True Then
                        Ctrl.Visible = False
                        BackSelected(j% - 1).Visible = True
                    Else
                        BackSelected(j% - 1).Visible = False
                        Ctrl.Visible = False
                        For Each NewCtrl In Controls
                            If NewCtrl.Tag = "CardSelected" Then
                                If DisplayDeck(3, j%) = Mid(NewCtrl.Name, 4) Then
                                    NewCtrl.Visible = True
                                End If
                            End If
                        Next NewCtrl
                    End If
                Else
                    If DisplayDeck(6, j%) = True Then
                        Ctrl.Visible = False
                        Back(j% - 1).Visible = True
                    Else
                        Back(j% - 1).Visible = False
                        Ctrl.Visible = True
                    End If
                End If
            End If
        End If
    Next Ctrl
Next j%
For Each Ctrl In Controls
    If Ctrl.Tag = "Position" Then
        If frmStackView.ShowIndexValues = 1 Then
            Ctrl.Visible = True
        Else
            Ctrl.Visible = False
        End If
    End If
Next Ctrl
If frmMain.mnuShuffleMeter.Checked = True Then
    frmShuffleMeter.SetShuffleMeterParameters
End If
End Sub

Public Sub DisplayDeal()
PokerCardsDealt = 1
PilesShown = 0
GilbreathShown = False
GilbreathActive = False
For i% = 1 To 8
    PileLabel(i%).Visible = False
Next i%
Me.Width = 13750
If Hands < 6 Then
    Me.Height = 5700
Else
    Me.Height = 8250
End If
For Each Ctrl In Controls
    If Ctrl.Tag = "Card" Or Ctrl.Tag = "Position" Or Ctrl.Tag = "Back" _
        Or Ctrl.Tag = "BackSelected" Or Ctrl.Tag = "CardSelected" Then
        Ctrl.Visible = False
    End If
Next Ctrl
For i% = 1 To 52
    Deck(3, i%) = "Card" & Deck(2, i%)
    Deck(5, i%) = "Position" & Deck(1, i%)
    Back(i% - 1).ToolTipText = Deck(2, i%)
    BackSelected(i% - 1).ToolTipText = Deck(2, i%)
Next i%
If Hands < 6 Then
    For i% = 1 To Hands
        For j% = 1 To 5
            For Each Ctrl In Controls
                If Ctrl.Tag = "Card" Then
                    If Deck(3, (i% - 1) * 5 + j%) = Ctrl.Name Then
                        Ctrl.Top = 400
                        Ctrl.Left = (i% - 1) * 2625 + 250 * j%
                        Ctrl.ZOrder
                        If frmStackView.HighlightSelectionsCheck = 1 And _
                            Deck(4, (i% - 1) * 5 + j%) = "Selected" Then
                            If Deck(6, (i% - 1) * 5 + j%) = True Then
                                BackSelected((i% - 1) * 5 + j% - 1).Top = 400
                                BackSelected((i% - 1) * 5 + j% - 1).Left = (i% - 1) * 2625 + 250 * j%
                                BackSelected((i% - 1) * 5 + j% - 1).ZOrder
                            Else
                                For Each NewCtrl In Controls
                                    If NewCtrl.Tag = "CardSelected" Then
                                        If Deck(3, (i% - 1) * 5 + j%) = Mid(NewCtrl.Name, 4) Then
                                            NewCtrl.Top = 400
                                            NewCtrl.Left = (i% - 1) * 2625 + 250 * j%
                                            NewCtrl.ZOrder
                                        End If
                                    End If
                                Next NewCtrl
                            End If
                        Else
                            If Deck(6, (i% - 1) * 5 + j%) = True Then
                                Back((i% - 1) * 5 + j% - 1).Top = 400
                                Back((i% - 1) * 5 + j% - 1).Left = (i% - 1) * 2625 + 250 * j%
                                Back((i% - 1) * 5 + j% - 1).ZOrder
                            End If
                        End If
                    End If
                End If
            Next Ctrl
            For Each Ctrl In Controls
                If Ctrl.Tag = "Position" Then
                    If Deck(5, (i% - 1) * 5 + j%) = Ctrl.Name Then
                        Ctrl.Top = 160
                        Ctrl.Left = (i% - 1) * 2625 + 250 * j% + 50
                        If frmStackView.HighlightSelectionsCheck = 1 Then
                            If Deck(4, (i% - 1) * 5 + j%) = "Selected" Then
                                Ctrl.Font.Bold = True
                                Ctrl.Font.Size = 12
                                Ctrl.Top = 60
                                Ctrl.Left = (i% - 1) * 2625 + 250 * j% - 10 + 50
                            Else
                                Ctrl.Font.Bold = False
                                Ctrl.Font.Size = 8
                            End If
                        Else
                            Ctrl.Font.Bold = False
                            Ctrl.Font.Size = 8
                        End If
                    End If
                End If
            Next Ctrl
            If frmStackView.CountFromBack = True Then
                Ordinal((i% - 1) * 5 + j%).Top = 400 + 1925
                Ordinal((i% - 1) * 5 + j%).Left = (i% - 1) * 2625 + 250 * j% + 50
            ElseIf frmStackView.CountFromFace = True Then
                Ordinal(53 - ((i% - 1) * 5 + j%)).Top = 400 + 1925
                Ordinal(53 - ((i% - 1) * 5 + j%)).Left = (i% - 1) * 2625 + 250 * j% + 50
            End If
            If frmStackView.ShowPositionValues = 1 Then
                Ordinal((i% - 1) * 5 + j%).Visible = True
            Else
                Ordinal((i% - 1) * 5 + j%).Visible = False
            End If
        Next j%
    Next i%
    For k% = Hands * 5 + 1 To DeckCount
        For Each Ctrl In Controls
            If Deck(3, k%) = Ctrl.Name Then
                Ctrl.Top = 2900
                Ctrl.Left = 250 + (k% - Hands * 5 - 1) * 250
                Ctrl.ZOrder
                        If frmStackView.HighlightSelectionsCheck = 1 And _
                            Deck(4, k%) = "Selected" Then
                            If Deck(6, k%) = True Then
                                BackSelected(k% - 1).Top = 2900
                                BackSelected(k% - 1).Left = 250 + (k% - Hands * 5 - 1) * 250
                                BackSelected(k% - 1).ZOrder
                            Else
                                For Each NewCtrl In Controls
                                    If NewCtrl.Tag = "CardSelected" Then
                                        If Deck(3, k%) = Mid(NewCtrl.Name, 4) Then
                                            NewCtrl.Top = 2900
                                            NewCtrl.Left = 250 + (k% - Hands * 5 - 1) * 250
                                            NewCtrl.ZOrder
                                        End If
                                    End If
                                Next NewCtrl
                            End If
                        Else
                            If Deck(6, k%) = True Then
                                Back(k% - 1).Top = 2900
                                Back(k% - 1).Left = 250 + (k% - Hands * 5 - 1) * 250
                                Back(k% - 1).ZOrder
                            End If
                        End If
            End If
        Next Ctrl
        For Each Ctrl In Controls
            If Ctrl.Tag = "Position" Then
                If Deck(5, k%) = Ctrl.Name Then
                    Ctrl.Top = 2660
                    Ctrl.Left = 250 + (k% - Hands * 5 - 1) * 250 + 50
                    If frmStackView.HighlightSelectionsCheck = 1 Then
                        If Deck(4, k%) = "Selected" Then
                            Ctrl.Font.Bold = True
                            Ctrl.Font.Size = 12
                            Ctrl.Top = 2560
                            Ctrl.Left = 250 + (k% - Hands * 5 - 1) * 250 - 10 + 50
                        Else
                            Ctrl.Font.Bold = False
                            Ctrl.Font.Size = 8
                        End If
                    Else
                        Ctrl.Font.Bold = False
                        Ctrl.Font.Size = 8
                    End If
                End If
            End If
        Next Ctrl
        If frmStackView.CountFromBack = True Then
            Ordinal(k%).Top = 2900 + 1925
            Ordinal(k%).Left = 250 + (k% - Hands * 5 - 1) * 250 + 50
        ElseIf frmStackView.CountFromFace = True Then
            Ordinal(53 - k%).Top = 2900 + 1925
            Ordinal(53 - k%).Left = 250 + (k% - Hands * 5 - 1) * 250 + 50
        End If
        If frmStackView.ShowPositionValues = 1 Then
            Ordinal(k%).Visible = True
        Else
            Ordinal(k%).Visible = False
        End If
    Next k%
'Show the appropriate cards only
'NEW SECTION
For j% = 1 To 52
    For Each Ctrl In Controls
        If Ctrl.Tag = "Card" Then
            If Deck(3, j%) = Ctrl.Name Then
                If frmStackView.HighlightSelectionsCheck = 1 And _
                    Deck(4, j%) = "Selected" Then
                    If Deck(6, j%) = True Then
                        Ctrl.Visible = False
                        BackSelected(j% - 1).Visible = True
                    Else
                        BackSelected(j% - 1).Visible = False
                        Ctrl.Visible = False
                        For Each NewCtrl In Controls
                            If NewCtrl.Tag = "CardSelected" Then
                                If Deck(3, j%) = Mid(NewCtrl.Name, 4) Then
                                    NewCtrl.Visible = True
                                End If
                            End If
                        Next NewCtrl
                    End If
                Else
                    If Deck(6, j%) = True Then
                        Ctrl.Visible = False
                        Back(j% - 1).Visible = True
                    Else
                        Back(j% - 1).Visible = False
                        Ctrl.Visible = True
                    End If
                End If
            End If
        End If
    Next Ctrl
Next j%
For Each Ctrl In Controls
    If Ctrl.Tag = "Position" Then
        If frmStackView.ShowIndexValues = 1 Then
            Ctrl.Visible = True
        Else
            Ctrl.Visible = False
        End If
    End If
Next Ctrl
Else
    For i% = 1 To 5
        For j% = 1 To 5
            For Each Ctrl In Controls
                If Ctrl.Tag = "Card" Then
                    If Deck(3, (i% - 1) * 5 + j%) = Ctrl.Name Then
                        Ctrl.Top = 400
                        Ctrl.Left = (i% - 1) * 2625 + 250 * j%
                        Ctrl.ZOrder
                        If frmStackView.HighlightSelectionsCheck = 1 And _
                            Deck(4, (i% - 1) * 5 + j%) = "Selected" Then
                            If Deck(6, (i% - 1) * 5 + j%) = True Then
                                BackSelected((i% - 1) * 5 + j% - 1).Top = 400
                                BackSelected((i% - 1) * 5 + j% - 1).Left = (i% - 1) * 2625 + 250 * j%
                                BackSelected((i% - 1) * 5 + j% - 1).ZOrder
                            Else
                                For Each NewCtrl In Controls
                                    If NewCtrl.Tag = "CardSelected" Then
                                        If Deck(3, (i% - 1) * 5 + j%) = Mid(NewCtrl.Name, 4) Then
                                            NewCtrl.Top = 400
                                            NewCtrl.Left = (i% - 1) * 2625 + 250 * j%
                                            NewCtrl.ZOrder
                                        End If
                                    End If
                                Next NewCtrl
                            End If
                        Else
                            If Deck(6, (i% - 1) * 5 + j%) = True Then
                                Back((i% - 1) * 5 + j% - 1).Top = 400
                                Back((i% - 1) * 5 + j% - 1).Left = (i% - 1) * 2625 + 250 * j%
                                Back((i% - 1) * 5 + j% - 1).ZOrder
                            End If
                        End If
                    End If
                End If
            Next Ctrl
            For Each Ctrl In Controls
                If Ctrl.Tag = "Position" Then
                    If Deck(5, (i% - 1) * 5 + j%) = Ctrl.Name Then
                        Ctrl.Top = 160
                        Ctrl.Left = (i% - 1) * 2625 + 250 * j% + 50
                        If frmStackView.HighlightSelectionsCheck = 1 Then
                            If Deck(4, (i% - 1) * 5 + j%) = "Selected" Then
                                Ctrl.Font.Bold = True
                                Ctrl.Font.Size = 12
                                Ctrl.Top = 60
                                Ctrl.Left = (i% - 1) * 2625 + 250 * j% - 10 + 50
                            Else
                                Ctrl.Font.Bold = False
                                Ctrl.Font.Size = 8
                            End If
                        Else
                            Ctrl.Font.Bold = False
                            Ctrl.Font.Size = 8
                        End If
                    End If
                End If
            Next Ctrl
            If frmStackView.CountFromBack = True Then
                Ordinal((i% - 1) * 5 + j%).Top = 400 + 1925
                Ordinal((i% - 1) * 5 + j%).Left = (i% - 1) * 2625 + 250 * j% + 50
            ElseIf frmStackView.CountFromFace = True Then
                Ordinal(53 - ((i% - 1) * 5 + j%)).Top = 400 + 1925
                Ordinal(53 - ((i% - 1) * 5 + j%)).Left = (i% - 1) * 2625 + 250 * j% + 50
            End If
            If frmStackView.ShowPositionValues = 1 Then
                Ordinal((i% - 1) * 5 + j%).Visible = True
            Else
                Ordinal((i% - 1) * 5 + j%).Visible = False
            End If
        Next j%
    Next i%
    For i% = 1 To Hands - 5
        For j% = 1 To 5
            For Each Ctrl In Controls
                If Ctrl.Tag = "Card" Then
                    If Deck(3, (i% - 1) * 5 + j% + 25) = Ctrl.Name Then
                        'the "25" in the formula above addresses the fact that
                        'there are five poker hands in the top row already
                        Ctrl.Top = 2900
                        Ctrl.Left = (i% - 1) * 2625 + 250 * j%
                        Ctrl.ZOrder
                        If frmStackView.HighlightSelectionsCheck = 1 And _
                            Deck(4, (i% - 1) * 5 + j% + 25) = "Selected" Then
                            If Deck(6, (i% - 1) * 5 + j% + 25) = True Then
                                BackSelected((i% - 1) * 5 + j% + 25 - 1).Top = 2900
                                BackSelected((i% - 1) * 5 + j% + 25 - 1).Left = (i% - 1) * 2625 + 250 * j%
                                BackSelected((i% - 1) * 5 + j% + 25 - 1).ZOrder
                            Else
                                For Each NewCtrl In Controls
                                    If NewCtrl.Tag = "CardSelected" Then
                                        If Deck(3, (i% - 1) * 5 + j% + 25) = Mid(NewCtrl.Name, 4) Then
                                            NewCtrl.Top = 2900
                                            NewCtrl.Left = (i% - 1) * 2625 + 250 * j%
                                            NewCtrl.ZOrder
                                        End If
                                    End If
                                Next NewCtrl
                            End If
                        Else
                            If Deck(6, (i% - 1) * 5 + j% + 25) = True Then
                                Back((i% - 1) * 5 + j% + 25 - 1).Top = 2900
                                Back((i% - 1) * 5 + j% + 25 - 1).Left = (i% - 1) * 2625 + 250 * j%
                                Back((i% - 1) * 5 + j% + 25 - 1).ZOrder
                            End If
                        End If
                    End If
                End If
            Next Ctrl
            For Each Ctrl In Controls
                If Ctrl.Tag = "Position" Then
                    If Deck(5, (i% - 1) * 5 + j% + 25) = Ctrl.Name Then
                        'the "25" in the formula above addresses the fact that
                        'there are five poker hands in the top row already
                        Ctrl.Top = 2660
                        Ctrl.Left = (i% - 1) * 2625 + 250 * j% + 50
                        If frmStackView.HighlightSelectionsCheck = 1 Then
                            If Deck(4, (i% - 1) * 5 + j% + 25) = "Selected" Then
                        'the "25" in the formula above addresses the fact that
                        'there are five poker hands in the top row already
                                Ctrl.Font.Bold = True
                                Ctrl.Font.Size = 12
                                Ctrl.Top = 2560
                                Ctrl.Left = (i% - 1) * 2625 + 250 * j% - 10 + 50
                            Else
                                Ctrl.Font.Bold = False
                                Ctrl.Font.Size = 8
                            End If
                        Else
                            Ctrl.Font.Bold = False
                            Ctrl.Font.Size = 8
                        End If
                    End If
                End If
            Next Ctrl
            If frmStackView.CountFromBack = True Then
                Ordinal((i% - 1) * 5 + j% + 25).Top = 2900 + 1925
                Ordinal((i% - 1) * 5 + j% + 25).Left = (i% - 1) * 2625 + 250 * j% + 50
            ElseIf frmStackView.CountFromFace = True Then
                Ordinal(53 - ((i% - 1) * 5 + j% + 25)).Top = 2900 + 1925
                Ordinal(53 - ((i% - 1) * 5 + j% + 25)).Left = (i% - 1) * 2625 + 250 * j% + 50
            End If
            If frmStackView.ShowPositionValues = 1 Then
                Ordinal((i% - 1) * 5 + j% + 25).Visible = True
            Else
                Ordinal((i% - 1) * 5 + j% + 25).Visible = False
            End If
        Next j%
    Next i%
    For k% = Hands * 5 + 1 To DeckCount
        For Each Ctrl In Controls
            If Deck(3, k%) = Ctrl.Name Then
                Ctrl.Top = 5400
                Ctrl.Left = 250 + (k% - Hands * 5 - 1) * 250
                'Ctrl.Left = 10750 + (k% - (Hands * 5 + 1)) * 25
                Ctrl.ZOrder
                        If frmStackView.HighlightSelectionsCheck = 1 And _
                            Deck(4, k%) = "Selected" Then
                            If Deck(6, k%) = True Then
                                BackSelected(k% - 1).Top = 5400
                                BackSelected(k% - 1).Left = 250 + (k% - Hands * 5 - 1) * 250
                                BackSelected(k% - 1).ZOrder
                            Else
                                For Each NewCtrl In Controls
                                    If NewCtrl.Tag = "CardSelected" Then
                                        If Deck(3, k%) = Mid(NewCtrl.Name, 4) Then
                                            NewCtrl.Top = 5400
                                            NewCtrl.Left = 250 + (k% - Hands * 5 - 1) * 250
                                            NewCtrl.ZOrder
                                        End If
                                    End If
                                Next NewCtrl
                            End If
                        Else
                            If Deck(6, k%) = True Then
                                Back(k% - 1).Top = 5400
                                Back(k% - 1).Left = 250 + (k% - Hands * 5 - 1) * 250
                                Back(k% - 1).ZOrder
                            End If
                        End If
            End If
        Next Ctrl
        For Each Ctrl In Controls
            If Ctrl.Tag = "Position" Then
                If Deck(5, k%) = Ctrl.Name Then
                    Ctrl.Top = 5160
                    Ctrl.Left = 250 + (k% - Hands * 5 - 1) * 250 + 50
                    If frmStackView.HighlightSelectionsCheck = 1 Then
                        If Deck(4, k%) = "Selected" Then
                            Ctrl.Font.Bold = True
                            Ctrl.Font.Size = 12
                            Ctrl.Top = 5060
                            Ctrl.Left = 250 + (k% - Hands * 5 - 1) * 250 - 10 + 50
                        Else
                            Ctrl.Font.Bold = False
                            Ctrl.Font.Size = 8
                        End If
                    Else
                        Ctrl.Font.Bold = False
                        Ctrl.Font.Size = 8
                    End If
                End If
            End If
        Next Ctrl
        If frmStackView.CountFromBack = True Then
            Ordinal(k%).Top = 5400 + 1925
            Ordinal(k%).Left = 250 + (k% - Hands * 5 - 1) * 250 + 50
        ElseIf frmStackView.CountFromFace = True Then
            Ordinal(53 - k%).Top = 5400 + 1925
            Ordinal(53 - k%).Left = 250 + (k% - Hands * 5 - 1) * 250 + 50
        End If
        If frmStackView.ShowPositionValues = 1 Then
            Ordinal(k%).Visible = True
        Else
            Ordinal(k%).Visible = False
        End If
    Next k%
'Show the appropriate cards only
'NEW SECTION
For j% = 1 To 52
    For Each Ctrl In Controls
        If Ctrl.Tag = "Card" Then
            If Deck(3, j%) = Ctrl.Name Then
                If frmStackView.HighlightSelectionsCheck = 1 And _
                    Deck(4, j%) = "Selected" Then
                    If Deck(6, j%) = True Then
                        Ctrl.Visible = False
                        BackSelected(j% - 1).Visible = True
                    Else
                        BackSelected(j% - 1).Visible = False
                        Ctrl.Visible = False
                        For Each NewCtrl In Controls
                            If NewCtrl.Tag = "CardSelected" Then
                                If Deck(3, j%) = Mid(NewCtrl.Name, 4) Then
                                    NewCtrl.Visible = True
                                End If
                            End If
                        Next NewCtrl
                    End If
                Else
                    If Deck(6, j%) = True Then
                        Ctrl.Visible = False
                        Back(j% - 1).Visible = True
                    Else
                        Back(j% - 1).Visible = False
                        Ctrl.Visible = True
                    End If
                End If
            End If
        End If
    Next Ctrl
Next j%
For Each Ctrl In Controls
    If Ctrl.Tag = "Position" Then
        If frmStackView.ShowIndexValues = 1 Then
            Ctrl.Visible = True
        Else
            Ctrl.Visible = False
        End If
    End If
Next Ctrl
'        For k% = Hands * 5 + 1 To DeckCount
'        For Each Ctrl In Controls
'            If Deck(5, k%) = Ctrl.Name Then
'                Ctrl.Visible = False
'            End If
'        Next Ctrl
'    Next k%
End If
If frmMain.mnuShuffleMeter.Checked = True Then
    frmShuffleMeter.SetShuffleMeterParameters
End If
End Sub

Private Sub Back_DblClick(Index As Integer)
If PokerCardsDealt = 1 Then
    Dim tempTag As String
    tempTag = Back(Index).ToolTipText
    PokerDiscard (tempTag)
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(" & tempTag & ")"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = Back(Index).ToolTipText Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(" & Back(Index).ToolTipText & ")"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub

Private Sub BackSelected_DblClick(Index As Integer)
If PokerCardsDealt = 1 Then
    Dim tempTag As String
    tempTag = BackSelected(Index).ToolTipText
    PokerDiscard (tempTag)
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(" & tempTag & ")"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = BackSelected(Index).ToolTipText Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(" & BackSelected(Index).ToolTipText & ")"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub

Private Sub CardAC_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("AC")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(AC)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "AC" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(AC)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub

Private Sub Card2C_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("2C")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(2C)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "2C" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(2C)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub
Private Sub Card3C_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("3C")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(3C)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "3C" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(3C)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub
Private Sub Card4C_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("4C")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(4C)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "4C" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(4C)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub
Private Sub Card5C_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("5C")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(5C)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "5C" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(5C)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub
Private Sub Card6C_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("6C")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(6C)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "6C" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(6C)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub
Private Sub Card7C_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("7C")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(7C)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "7C" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(7C)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub
Private Sub Card8C_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("8C")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(8C)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "8C" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(8C)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub
Private Sub Card9C_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("9C")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(9C)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "9C" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(9C)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub
Private Sub Card10C_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("10C")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(10C)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "10C" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(10C)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub
Private Sub CardJC_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("JC")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(JC)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "JC" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(JC)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub
Private Sub CardQC_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("QC")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(QC)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "QC" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(QC)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub
Private Sub CardKC_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("KC")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(KC)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "KC" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(KC)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub
Private Sub CardAH_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("AH")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(AH)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "AH" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(AH)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub

Private Sub Card2H_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("2H")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(2H)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "2H" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(2H)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub
Private Sub Card3H_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("3H")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(3H)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "3H" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(3H)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub
Private Sub Card4H_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("4H")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(4H)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "4H" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(4H)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub
Private Sub Card5H_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("5H")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(5H)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "5H" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(5H)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub
Private Sub Card6H_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("6H")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(6H)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "6H" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(6H)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub
Private Sub Card7H_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("7H")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(7H)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "7H" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(7H)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub
Private Sub Card8H_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("8H")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(8H)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "8H" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(8H)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub
Private Sub Card9H_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("9H")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(9H)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "9H" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(9H)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub
Private Sub Card10H_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("10H")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(10H)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "10H" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(10H)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub
Private Sub CardJH_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("JH")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(JH)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "JH" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(JH)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub
Private Sub CardQH_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("QH")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(QH)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "QH" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(QH)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub
Private Sub CardKH_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("KH")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(KH)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "KH" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(KH)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub
Private Sub CardAS_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("AS")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(AS)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "AS" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(AS)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub

Private Sub Card2S_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("2S")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(2S)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "2S" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(2S)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub
Private Sub Card3S_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("3S")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(3S)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "3S" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(3S)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub
Private Sub Card4S_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("4S")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(4S)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "4S" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(4S)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub
Private Sub Card5S_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("5S")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(5S)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "5S" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(5S)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub
Private Sub Card6S_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("6S")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(6S)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "6S" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(6S)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub
Private Sub Card7S_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("7S")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(7S)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "7S" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(7S)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub
Private Sub Card8S_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("8S")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(8S)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "8S" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(8S)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub
Private Sub Card9S_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("9S")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(9S)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "9S" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(9S)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub
Private Sub Card10S_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("10S")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(10S)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "10S" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(10S)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub
Private Sub CardJS_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("JS")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(JS)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "JS" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(JS)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub
Private Sub CardQS_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("QS")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(QS)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "QS" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(QS)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub
Private Sub CardKS_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("KS")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(KS)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "KS" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(KS)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub
Private Sub CardAD_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("AD")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(AD)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "AD" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(AD)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub

Private Sub Card2D_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("2D")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(2D)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "2D" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(2D)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub
Private Sub Card3D_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("3D")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(3D)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "3D" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(3D)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub
Private Sub Card4D_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("4D")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(4D)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "4D" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(4D)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub
Private Sub Card5D_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("5D")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(5D)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "5D" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(5D)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub
Private Sub Card6D_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("6D")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(6D)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "6D" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(6D)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub
Private Sub Card7D_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("7D")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(7D)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "7D" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(7D)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub
Private Sub Card8D_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("8D")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(8D)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "8D" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(8D)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub
Private Sub Card9D_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("9D")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(9D)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "9D" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(9D)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub
Private Sub Card10D_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("10D")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(10D)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "10D" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(10D)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub
Private Sub CardJD_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("JD")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(JD)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "JD" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(JD)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub
Private Sub CardQD_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("QD")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(QD)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "QD" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(QD)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub
Private Sub CardKD_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("KD")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(KD)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "KD" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(KD)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub
'this next section seems to partially work to keep a hidden Deck window hidden
'it still flashes the wondow quickly at the start of some calculations
'also, the Piles view does not reactivate correctly
Private Sub Form_Activate()
If frmMain.mnuDeck.Checked = False Then
    Me.Visible = False
Else
    Me.Visible = True
End If
End Sub
'
'Private Sub Form_Load()
'If frmMain.mnuDeck.Checked = False Then
'    Me.Visible = False
'Else
'    Me.Visible = True
'End If
'End Sub

Private Sub Form_Load()
If frmMain.mnuDeck.Checked = False Then
    Me.Visible = False
Else
    Me.Visible = True
End If
'show the cards appropriately
If PilesShown = 1 Then
    If GilbreathActive Then
        If GilbreathShown Then
            frmDeck.DisplayPilesGilbreath
        Else
            frmDeck.DisplayPilesKeepGilbreathActive
        End If
    Else
        frmPiles.ShowPiles
    End If
Else
    frmStackView.ShowCards
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMain.mnuDeck.Checked = False
End Sub

'Private Sub Ordinal_Click(Index As Integer)
'ReverseCard (Index)
'SessionRecord
'If SessionRecordMode Then
'    SessionCommand = "ReverseCard(" & Index & ")"
'    frmStackView.SessionListBox.AddItem SessionCommand
'    frmStackView.SessionStatusUpdate (0)
'End If
'MsgBox ("single click")
'End Sub

Private Sub Ordinal_DblClick(Index As Integer)
If PokerCardsDealt = 1 Then
    PokerDiscard (Index)
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(" & Index & ")"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    ReverseCard (Index)
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(" & Index & ")"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub

Private Sub SelCardAC_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("AC")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(AC)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "AC" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(AC)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub

Private Sub SelCard2C_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("2C")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(2C)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "2C" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(2C)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub
Private Sub SelCard3C_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("3C")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(3C)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "3C" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(3C)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub
Private Sub SelCard4C_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("4C")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(4C)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "4C" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(4C)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub
Private Sub SelCard5C_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("5C")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(5C)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "5C" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(5C)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub
Private Sub SelCard6C_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("6C")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(6C)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "6C" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(6C)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub
Private Sub SelCard7C_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("7C")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(7C)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "7C" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(7C)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub
Private Sub SelCard8C_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("8C")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(8C)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "8C" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(8C)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub
Private Sub SelCard9C_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("9C")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(9C)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "9C" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(9C)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub
Private Sub SelCard10C_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("10C")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(10C)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "10C" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(10C)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub
Private Sub SelCardJC_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("JC")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(JC)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "JC" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(JC)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub
Private Sub SelCardQC_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("QC")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(QC)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "QC" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(QC)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub
Private Sub SelCardKC_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("KC")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(KC)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "KC" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(KC)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub
Private Sub SelCardAH_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("AH")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(AH)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "AH" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(AH)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub

Private Sub SelCard2H_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("2H")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(2H)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "2H" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(2H)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub
Private Sub SelCard3H_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("3H")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(3H)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "3H" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(3H)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub
Private Sub SelCard4H_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("4H")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(4H)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "4H" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(4H)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub
Private Sub SelCard5H_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("5H")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(5H)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "5H" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(5H)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub
Private Sub SelCard6H_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("6H")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(6H)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "6H" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(6H)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub
Private Sub SelCard7H_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("7H")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(7H)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "7H" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(7H)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub
Private Sub SelCard8H_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("8H")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(8H)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "8H" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(8H)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub
Private Sub SelCard9H_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("9H")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(9H)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "9H" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(9H)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub
Private Sub SelCard10H_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("10H")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(10H)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "10H" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(10H)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub
Private Sub SelCardJH_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("JH")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(JH)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "JH" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(JH)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub
Private Sub SelCardQH_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("QH")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(QH)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "QH" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(QH)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub
Private Sub SelCardKH_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("KH")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(KH)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "KH" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(KH)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub
Private Sub SelCardAS_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("AS")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(AS)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "AS" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(AS)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub

Private Sub SelCard2S_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("2S")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(2S)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "2S" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(2S)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub
Private Sub SelCard3S_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("3S")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(3S)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "3S" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(3S)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub
Private Sub SelCard4S_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("4S")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(4S)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "4S" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(4S)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub
Private Sub SelCard5S_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("5S")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(5S)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "5S" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(5S)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub
Private Sub SelCard6S_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("6S")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(6S)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "6S" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(6S)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub
Private Sub SelCard7S_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("7S")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(7S)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "7S" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(7S)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub
Private Sub SelCard8S_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("8S")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(8S)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "8S" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(8S)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub
Private Sub SelCard9S_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("9S")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(9S)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "9S" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(9S)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub
Private Sub SelCard10S_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("10S")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(10S)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "10S" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(10S)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub
Private Sub SelCardJS_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("JS")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(JS)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "JS" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(JS)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub
Private Sub SelCardQS_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("QS")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(QS)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "QS" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(QS)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub
Private Sub SelCardKS_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("KS")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(KS)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "KS" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(KS)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub
Private Sub SelCardAD_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("AD")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(AD)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "AD" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(AD)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub

Private Sub SelCard2D_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("2D")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(2D)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "2D" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(2D)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub
Private Sub SelCard3D_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("3D")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(3D)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "3D" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(3D)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub
Private Sub SelCard4D_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("4D")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(4D)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "4D" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(4D)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub
Private Sub SelCard5D_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("5D")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(5D)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "5D" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(5D)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub
Private Sub SelCard6D_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("6D")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(6D)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "6D" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(6D)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub
Private Sub SelCard7D_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("7D")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(7D)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "7D" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(7D)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub
Private Sub SelCard8D_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("8D")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(8D)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "8D" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(8D)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub
Private Sub SelCard9D_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("9D")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(9D)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "9D" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(9D)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub
Private Sub SelCard10D_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("10D")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(10D)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "10D" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(10D)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub
Private Sub SelCardJD_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("JD")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(JD)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "JD" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(JD)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub
Private Sub SelCardQD_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("QD")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(QD)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "QD" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(QD)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub
Private Sub SelCardKD_DblClick()
If PokerCardsDealt = 1 Then
    PokerDiscard ("KD")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "PokerDiscard(KD)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
Else
    For i% = 1 To DeckCount
        If Deck(2, i%) = "KD" Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
    If PilesShown = 1 Then
        If GilbreathActive Then
            If GilbreathShown Then
                frmDeck.DisplayPilesGilbreath
            Else
                frmDeck.DisplayPilesKeepGilbreathActive
            End If
        Else
            frmPiles.ShowPiles
        End If
    Else
        frmStackView.ShowCards
    End If
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReverseCard(KD)"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End If
End Sub

Public Sub ReverseCard(param)
Dim pSafeMatch As Boolean
pSafeMatch = False
If IsNumeric(param) Then
    For i% = 1 To 52
        If i% = Val(param) Then
            pSafeMatch = True
        End If
    Next i%
    If Not pSafeMatch Then
        MsgBox ("ReverseCard Event Error:" & Chr(13) & _
        "The single parameter can only be an integer from 1 to 52," & Chr(13) & _
        "or a card value such as AC, 2C, 3C, etc.")
        Exit Sub
    End If
    'code to turn over the card at specified position
    If frmStackView.CountFromBack Then
        Deck(6, Val(param)) = Not Deck(6, Val(param))
    ElseIf frmStackView.CountFromFace Then
        Deck(6, 53 - Val(param)) = Not Deck(6, 53 - Val(param))
    End If
Else
    For i% = 1 To 52
        If Deck(2, i%) = param Then
            pSafeMatch = True
        End If
    Next i%
    If Not pSafeMatch Then
        MsgBox ("ReverseCard Event Error:" & Chr(13) & _
        "The single parameter can only be an integer from 1 to 52," & Chr(13) & _
        "or a card value such as AC, 2C, 3C, etc.")
        Exit Sub
    End If
    'code to turn over the card
    For i% = 1 To DeckCount
        If Deck(2, i%) = param Then
            Deck(6, i%) = Not Deck(6, i%)
        End If
    Next i%
End If
'show the cards appropriately
If PilesShown = 1 Then
    If GilbreathActive Then
        If GilbreathShown Then
            frmDeck.DisplayPilesGilbreath
        Else
            frmDeck.DisplayPilesKeepGilbreathActive
        End If
    Else
        frmPiles.ShowPiles
    End If
Else
    frmStackView.ShowCards
End If
End Sub

Public Sub PokerDiscard(param)
Dim pSafeMatch As Boolean
Dim pDiscardComplete As Boolean
Dim pDiscardCounter As Integer
pSafeMatch = False
If IsNumeric(param) Then
    For i% = 1 To 50
        If i% = Val(param) Then
            pSafeMatch = True
        End If
    Next i%
    If Not pSafeMatch Then
        MsgBox ("PokerDiscard Event Error:" & Chr(13) & _
        "The single parameter can only be an integer from 1 to 50," & Chr(13) & _
        "or a card value such as AC, 2C, 3C, etc.")
        Exit Sub
    End If
    'code to discard the card
    pDiscardComplete = False
    pDiscardCounter = 1
    While Not pDiscardComplete
        If pDiscardCounter = param Then
            If pDiscardCounter > Hands * 5 Then
                DisplayDeal
                MsgBox ("You may only discard cards from active hands.")
                Exit Sub
            End If
            For j% = 1 To DeckProperties
                TempCard(j%, 1) = Deck(j%, pDiscardCounter)
                Deck(j%, pDiscardCounter) = Deck(j%, Hands * 5 + 1)
            Next j%
            For k% = Hands * 5 + 1 To 51
                For p% = 1 To DeckProperties
                    Deck(p%, k%) = Deck(p%, k% + 1)
                Next p%
            Next k%
            For j% = 1 To DeckProperties
                Deck(j%, 52) = TempCard(j%, 1)
            Next j%
            pDiscardComplete = True
        End If
        pDiscardCounter = pDiscardCounter + 1
        If pDiscardCounter > 53 Then
            MsgBox ("Poker Discard Error:" & Chr(13) & _
            "pDiscardCounter > 53" & Chr(13) & _
            "Please send email to nick@stackview.com")
        End If
    Wend
Else
    For i% = 1 To 52
        If Deck(2, i%) = param Then
            pSafeMatch = True
        End If
    Next i%
    If Not pSafeMatch Then
        MsgBox ("Poker Discard Event Error:" & Chr(13) & _
        "The single parameter can only be an integer from 1 to 50," & Chr(13) & _
        "or a card value such as AC, 2C, 3C, etc.")
        Exit Sub
    End If
    'code to discard the card
    pDiscardComplete = False
    pDiscardCounter = 1
    While Not pDiscardComplete
        If Deck(2, pDiscardCounter) = param Then
            If pDiscardCounter > Hands * 5 Then
                DisplayDeal
                MsgBox ("You may only discard cards from active hands.")
                Exit Sub
            End If
            For j% = 1 To DeckProperties
                TempCard(j%, 1) = Deck(j%, pDiscardCounter)
                Deck(j%, pDiscardCounter) = Deck(j%, Hands * 5 + 1)
            Next j%
            For k% = Hands * 5 + 1 To 51
                For p% = 1 To DeckProperties
                    Deck(p%, k%) = Deck(p%, k% + 1)
                Next p%
            Next k%
            For j% = 1 To DeckProperties
                Deck(j%, 52) = TempCard(j%, 1)
            Next j%
            pDiscardComplete = True
        End If
        pDiscardCounter = pDiscardCounter + 1
        If pDiscardCounter > 53 Then
            MsgBox ("Poker Discard Error:" & Chr(13) & _
            "pDiscardCounter > 53" & Chr(13) & _
            "Please send email to nick@stackview.com")
        End If
    Wend
End If
'show the cards
DisplayDeal
End Sub
