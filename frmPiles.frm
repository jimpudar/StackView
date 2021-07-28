VERSION 5.00
Begin VB.Form frmPiles 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Piles Control"
   ClientHeight    =   7785
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13500
   BeginProperty Font 
      Name            =   "Arial Narrow"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPiles.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7785
   ScaleWidth      =   13500
   Begin VB.Frame SwapPilesFrame 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1935
      Left            =   5640
      TabIndex        =   298
      Top             =   5760
      Width           =   6015
      Begin VB.CheckBox SwapReverseSecondPile 
         Caption         =   "Reverse Second Pile"
         Height          =   255
         Left            =   105
         TabIndex        =   313
         Top             =   1350
         Width           =   1710
      End
      Begin VB.CheckBox SwapReverseFirstPile 
         Caption         =   "Reverse First Pile"
         Height          =   255
         Left            =   105
         TabIndex        =   312
         Top             =   1080
         Width           =   1710
      End
      Begin VB.Frame SwapSecondPileFrame 
         Height          =   1455
         Left            =   3960
         TabIndex        =   307
         Top             =   360
         Width           =   1905
         Begin VB.OptionButton SwapSecondRandom 
            Caption         =   "Random"
            Height          =   270
            Left            =   75
            TabIndex        =   311
            Top             =   390
            Width           =   1380
         End
         Begin VB.OptionButton SwapSecondSecondary 
            Caption         =   "Secondary"
            Height          =   270
            Left            =   75
            TabIndex        =   310
            Top             =   165
            Value           =   -1  'True
            Width           =   1020
         End
         Begin VB.OptionButton SwapSecondSelected 
            Caption         =   "Includes Selected Card"
            Height          =   270
            Left            =   75
            TabIndex        =   309
            Top             =   615
            Width           =   1800
         End
         Begin VB.OptionButton SwapSecondNoSelected 
            Caption         =   "Random with no Selected Card"
            Height          =   495
            Left            =   75
            TabIndex        =   308
            Top             =   840
            Width           =   1470
         End
      End
      Begin VB.Frame SwapFirstPileFrame 
         Height          =   1455
         Left            =   1920
         TabIndex        =   302
         Top             =   360
         Width           =   1905
         Begin VB.OptionButton SwapFirstNoSelected 
            Caption         =   "Random with no Selected Card"
            Height          =   495
            Left            =   75
            TabIndex        =   306
            Top             =   840
            Width           =   1470
         End
         Begin VB.OptionButton SwapFirstSelected 
            Caption         =   "Includes Selected Card"
            Height          =   270
            Left            =   75
            TabIndex        =   305
            Top             =   615
            Width           =   1800
         End
         Begin VB.OptionButton SwapFirstPrimary 
            Caption         =   "Primary"
            Height          =   270
            Left            =   75
            TabIndex        =   304
            Top             =   165
            Value           =   -1  'True
            Width           =   1020
         End
         Begin VB.OptionButton SwapFirstRandom 
            Caption         =   "Random"
            Height          =   270
            Left            =   75
            TabIndex        =   303
            Top             =   390
            Width           =   1380
         End
      End
      Begin VB.CommandButton SwapPilesButton 
         Caption         =   "Swap Piles"
         Height          =   420
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   299
         TabStop         =   0   'False
         Top             =   480
         Width           =   960
      End
      Begin VB.Label Label30 
         Caption         =   "Second Pile"
         Height          =   270
         Left            =   4080
         TabIndex        =   301
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label Label16 
         Caption         =   "First Pile"
         Height          =   255
         Left            =   2040
         TabIndex        =   300
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.CommandButton RefreshDeckPiles 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Refresh Deck"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   296
      Top             =   6465
      Width           =   960
   End
   Begin VB.Frame CountFrame 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1260
      Left            =   3120
      TabIndex        =   293
      Top             =   5775
      Width           =   2460
      Begin VB.CheckBox SpecialInverseCheck 
         Caption         =   "Inverse"
         Height          =   270
         Left            =   75
         TabIndex        =   169
         Top             =   675
         Width           =   765
      End
      Begin VB.Frame Frame14 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1080
         Left            =   930
         TabIndex        =   295
         Top             =   135
         Width           =   1455
         Begin VB.OptionButton SpecialReverseOrder 
            Caption         =   "Turn Over Pile"
            Height          =   390
            Left            =   60
            TabIndex        =   172
            Top             =   690
            Width           =   1425
         End
         Begin VB.OptionButton SpecialJordan 
            Caption         =   "Jordan Count"
            Height          =   390
            Left            =   60
            TabIndex        =   171
            Top             =   390
            Width           =   1425
         End
         Begin VB.OptionButton SpecialElmsley 
            Caption         =   "Elmsley Count"
            Height          =   390
            Left            =   60
            TabIndex        =   170
            Top             =   90
            Value           =   -1  'True
            Width           =   1455
         End
      End
      Begin VB.CommandButton SpecialButton 
         Caption         =   "Special"
         Height          =   360
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   294
         TabStop         =   0   'False
         Top             =   240
         Width           =   720
      End
   End
   Begin VB.Frame CutFrame 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   2655
      Left            =   3120
      TabIndex        =   283
      Top             =   3120
      Width           =   4785
      Begin VB.CheckBox ReverseCutPortion 
         Caption         =   "Reverse Cut Portion"
         Height          =   270
         Left            =   2250
         TabIndex        =   150
         Top             =   2340
         Width           =   1815
      End
      Begin VB.Frame CutPileFrame 
         Height          =   690
         Left            =   180
         TabIndex        =   292
         Top             =   1005
         Width           =   1860
         Begin VB.OptionButton CutPrimaryPile 
            Caption         =   "Primary Pile"
            Height          =   270
            Left            =   75
            TabIndex        =   137
            Top             =   165
            Value           =   -1  'True
            Width           =   1470
         End
         Begin VB.OptionButton CutRandomPile 
            Caption         =   "Random Pile"
            Height          =   270
            Left            =   75
            TabIndex        =   138
            Top             =   390
            Width           =   1605
         End
      End
      Begin VB.Frame PlacePortionFrame 
         Height          =   2010
         Left            =   2145
         TabIndex        =   287
         Top             =   315
         Width           =   2490
         Begin VB.OptionButton TopSame 
            Caption         =   "Top of Same"
            Height          =   270
            Left            =   90
            TabIndex        =   145
            Top             =   720
            Width           =   1650
         End
         Begin VB.TextBox PlaceNewPileSpecifiedText 
            Alignment       =   2  'Center
            Height          =   330
            Left            =   1635
            TabIndex        =   149
            Top             =   1620
            Width           =   450
         End
         Begin VB.OptionButton PlaceNewPileSpecified 
            Caption         =   "New Pile Specified"
            Height          =   270
            Left            =   90
            TabIndex        =   148
            Top             =   1680
            Width           =   1605
         End
         Begin VB.OptionButton PlaceNewPileRandom 
            Caption         =   "New Pile Random"
            Height          =   270
            Left            =   90
            TabIndex        =   147
            Top             =   1440
            Width           =   1440
         End
         Begin VB.OptionButton TopRandomNotSame 
            Caption         =   "Top Random Not Same"
            Height          =   270
            Left            =   90
            TabIndex        =   288
            Top             =   1200
            Width           =   1785
         End
         Begin VB.OptionButton TopRandomAny 
            Caption         =   "Top Random Any"
            Height          =   270
            Left            =   90
            TabIndex        =   146
            Top             =   960
            Width           =   1440
         End
         Begin VB.OptionButton TopSecondary 
            Caption         =   "Top of Secondary"
            Height          =   270
            Left            =   90
            TabIndex        =   144
            Top             =   480
            Width           =   1650
         End
         Begin VB.OptionButton CompleteCut 
            Caption         =   "Complete Cut Primary / Same"
            Height          =   270
            Left            =   90
            TabIndex        =   143
            Top             =   240
            Value           =   -1  'True
            Width           =   2295
         End
         Begin VB.Label Label13 
            Caption         =   "pile #"
            Height          =   270
            Left            =   1710
            TabIndex        =   289
            Top             =   1410
            Width           =   405
         End
      End
      Begin VB.Frame CutPortionFrame 
         Height          =   1005
         Left            =   180
         TabIndex        =   285
         Top             =   1560
         Width           =   1860
         Begin VB.OptionButton CompletePile 
            Caption         =   "Complete Pile"
            Height          =   270
            Left            =   90
            TabIndex        =   140
            Top             =   405
            Width           =   1185
         End
         Begin VB.TextBox CutSpecifiedText 
            Alignment       =   2  'Center
            Height          =   330
            Left            =   1335
            TabIndex        =   142
            Top             =   585
            Width           =   450
         End
         Begin VB.OptionButton CutSpecified 
            Caption         =   "Cut Specified"
            Height          =   270
            Left            =   90
            TabIndex        =   141
            Top             =   660
            Width           =   1140
         End
         Begin VB.OptionButton CutRandom 
            Caption         =   "Cut Random"
            Height          =   270
            Left            =   90
            TabIndex        =   139
            Top             =   165
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.Label Label12 
            Caption         =   "# cards"
            Height          =   270
            Left            =   1320
            TabIndex        =   286
            Top             =   360
            Width           =   495
         End
      End
      Begin VB.CommandButton CutCardsButton 
         Caption         =   "Cut Piles"
         Height          =   420
         Left            =   225
         Style           =   1  'Graphical
         TabIndex        =   284
         TabStop         =   0   'False
         Top             =   285
         Width           =   960
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         Caption         =   "Place Cut Portion"
         Height          =   225
         Left            =   2505
         TabIndex        =   291
         Top             =   120
         Width           =   1665
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         Caption         =   "Cut Pile and Portion"
         Height          =   225
         Left            =   165
         TabIndex        =   290
         Top             =   810
         Width           =   1590
      End
   End
   Begin VB.Frame AustralianDealFrame 
      Height          =   2340
      Left            =   7995
      TabIndex        =   267
      Top             =   3420
      Width           =   5400
      Begin VB.CheckBox RandomCardSelectedCheck 
         Caption         =   "Select Random ""Down"" Card"
         Height          =   270
         Left            =   75
         TabIndex        =   152
         Top             =   1440
         Width           =   2100
      End
      Begin VB.CheckBox AustralianDealInverseCheck 
         Caption         =   "Inverse"
         Height          =   270
         Left            =   75
         TabIndex        =   153
         Top             =   1845
         Width           =   1815
      End
      Begin VB.CheckBox FinalCardSelectedCheck 
         Caption         =   "Select Final Card"
         Height          =   270
         Left            =   75
         TabIndex        =   151
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Frame AustralianReverseFrame 
         Height          =   1770
         Left            =   3795
         TabIndex        =   277
         Top             =   315
         Width           =   1545
         Begin VB.OptionButton AustralianReverseSelected 
            Caption         =   "Reverse Selected"
            Height          =   270
            Left            =   75
            TabIndex        =   163
            Top             =   390
            Width           =   1425
         End
         Begin VB.CheckBox ReverseUnderAllCheck 
            Caption         =   "Check1"
            Enabled         =   0   'False
            Height          =   270
            Left            =   1140
            TabIndex        =   167
            Top             =   1155
            Width           =   240
         End
         Begin VB.CheckBox ReverseDownRandomCheck 
            Caption         =   "Check1"
            Enabled         =   0   'False
            Height          =   270
            Left            =   825
            TabIndex        =   166
            Top             =   1440
            Width           =   240
         End
         Begin VB.CheckBox ReverseUnderRandomCheck 
            Caption         =   "Check1"
            Enabled         =   0   'False
            Height          =   270
            Left            =   1140
            TabIndex        =   168
            Top             =   1440
            Width           =   240
         End
         Begin VB.CheckBox ReverseDownAllCheck 
            Caption         =   "Check1"
            Enabled         =   0   'False
            Height          =   270
            Left            =   825
            TabIndex        =   165
            Top             =   1155
            Width           =   240
         End
         Begin VB.OptionButton AustralianNoReverse 
            Caption         =   "No Reverse"
            Height          =   270
            Left            =   75
            TabIndex        =   162
            Top             =   165
            Value           =   -1  'True
            Width           =   1170
         End
         Begin VB.OptionButton AustralianReverse 
            Caption         =   "Reverse cards"
            Height          =   270
            Left            =   75
            TabIndex        =   164
            Top             =   615
            Width           =   1320
         End
         Begin VB.Label ReverseUnderLabel 
            Alignment       =   2  'Center
            Caption         =   "Under"
            Enabled         =   0   'False
            Height          =   225
            Left            =   1095
            TabIndex        =   278
            Top             =   870
            Width           =   390
         End
         Begin VB.Label ReverseDownLabel 
            Alignment       =   2  'Center
            Caption         =   "Down"
            Enabled         =   0   'False
            Height          =   225
            Left            =   705
            TabIndex        =   281
            Top             =   870
            Width           =   405
         End
         Begin VB.Label ReverseRandomLabel 
            Caption         =   "Random"
            Enabled         =   0   'False
            Height          =   225
            Left            =   165
            TabIndex        =   280
            Top             =   1425
            Width           =   510
         End
         Begin VB.Label ReverseAllLabel 
            Caption         =   "All"
            Enabled         =   0   'False
            Height          =   225
            Left            =   165
            TabIndex        =   279
            Top             =   1155
            Width           =   510
         End
      End
      Begin VB.Frame AustralianNumberFrame 
         Height          =   1770
         Left            =   2220
         TabIndex        =   271
         Top             =   315
         Width           =   1545
         Begin VB.TextBox NumberUnderExactText 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            Height          =   285
            Left            =   1095
            TabIndex        =   160
            Top             =   1005
            Width           =   345
         End
         Begin VB.TextBox NumberDownRandomText 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            Height          =   285
            Left            =   735
            TabIndex        =   159
            Top             =   1290
            Width           =   345
         End
         Begin VB.TextBox NumberUnderRandomText 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            Height          =   285
            Left            =   1095
            TabIndex        =   161
            Top             =   1290
            Width           =   345
         End
         Begin VB.TextBox NumberDownExactText 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            Height          =   285
            Left            =   735
            TabIndex        =   158
            Top             =   1005
            Width           =   345
         End
         Begin VB.OptionButton AustralianNumberSpecified 
            Caption         =   "Specified"
            Height          =   270
            Left            =   75
            TabIndex        =   157
            Top             =   390
            Width           =   1170
         End
         Begin VB.OptionButton AustralianNumberStandard 
            Caption         =   "Standard"
            Height          =   270
            Left            =   75
            TabIndex        =   156
            Top             =   165
            Value           =   -1  'True
            Width           =   1170
         End
         Begin VB.Label NumberUnderLabel 
            Alignment       =   2  'Center
            Caption         =   "Under"
            Enabled         =   0   'False
            Height          =   225
            Left            =   1080
            TabIndex        =   276
            Top             =   750
            Width           =   390
         End
         Begin VB.Label NumberExactLabel 
            Caption         =   "Exact"
            Enabled         =   0   'False
            Height          =   225
            Left            =   150
            TabIndex        =   275
            Top             =   1035
            Width           =   510
         End
         Begin VB.Label NumberRandomLabel 
            Caption         =   "Random"
            Enabled         =   0   'False
            Height          =   225
            Left            =   150
            TabIndex        =   274
            Top             =   1305
            Width           =   510
         End
         Begin VB.Label NumberDownLabel 
            Alignment       =   2  'Center
            Caption         =   "Down"
            Enabled         =   0   'False
            Height          =   225
            Left            =   690
            TabIndex        =   273
            Top             =   750
            Width           =   405
         End
      End
      Begin VB.Frame AustralianStartFrame 
         Height          =   735
         Left            =   885
         TabIndex        =   269
         Top             =   315
         Width           =   1335
         Begin VB.OptionButton AustralianStartDownUnder 
            Caption         =   "Down / Under"
            Height          =   270
            Left            =   75
            TabIndex        =   154
            Top             =   165
            Value           =   -1  'True
            Width           =   1170
         End
         Begin VB.OptionButton AustralianStartUnderDown 
            Caption         =   "Under / Down"
            Height          =   270
            Left            =   75
            TabIndex        =   155
            Top             =   390
            Width           =   1170
         End
      End
      Begin VB.CommandButton AustralianDealButton 
         Caption         =   "Australian Deal"
         Height          =   615
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   268
         TabStop         =   0   'False
         Top             =   420
         Width           =   795
      End
      Begin VB.Label Label34 
         Alignment       =   2  'Center
         Caption         =   "Reverse"
         Height          =   225
         Left            =   4230
         TabIndex        =   282
         Top             =   135
         Width           =   600
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Caption         =   "Number"
         Height          =   225
         Left            =   2655
         TabIndex        =   272
         Top             =   135
         Width           =   600
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Caption         =   "Start"
         Height          =   225
         Left            =   1155
         TabIndex        =   270
         Top             =   135
         Width           =   600
      End
   End
   Begin VB.Frame RiffleShuffleFrame 
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1260
      Left            =   8430
      TabIndex        =   261
      Top             =   2175
      Width           =   4950
      Begin VB.CheckBox GilbreathCheck 
         Caption         =   "Gilbreath View"
         Height          =   465
         Left            =   240
         TabIndex        =   297
         Top             =   735
         Width           =   870
      End
      Begin VB.Frame Frame12 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2475
         TabIndex        =   265
         Top             =   135
         Width           =   1050
         Begin VB.OptionButton RiffleShuffleProtectPrimary 
            Caption         =   "Primary"
            Enabled         =   0   'False
            Height          =   390
            Left            =   60
            TabIndex        =   132
            Top             =   90
            Value           =   -1  'True
            Width           =   825
         End
         Begin VB.OptionButton RiffleShuffleProtectSecondary 
            Caption         =   "Secondary"
            Enabled         =   0   'False
            Height          =   390
            Left            =   60
            TabIndex        =   133
            Top             =   390
            Width           =   960
         End
      End
      Begin VB.TextBox RiffleShuffleProtectCards 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   330
         Left            =   4350
         TabIndex        =   136
         Top             =   435
         Width           =   450
      End
      Begin VB.Frame RiffleShuffleBlockDefine 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3495
         TabIndex        =   264
         Top             =   135
         Width           =   810
         Begin VB.OptionButton RiffleShuffleProtectTop 
            Caption         =   "Top"
            Enabled         =   0   'False
            Height          =   390
            Left            =   60
            TabIndex        =   134
            Top             =   90
            Value           =   -1  'True
            Width           =   585
         End
         Begin VB.OptionButton RiffleShuffleProtectBottom 
            Caption         =   "Bottom"
            Enabled         =   0   'False
            Height          =   390
            Left            =   60
            TabIndex        =   135
            Top             =   390
            Width           =   735
         End
      End
      Begin VB.Frame RiffleShuffleDefineFrame 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1305
         TabIndex        =   263
         Top             =   135
         Width           =   1200
         Begin VB.OptionButton RiffleShufflePilesProtect 
            Caption         =   "Protect Block"
            Height          =   390
            Left            =   60
            TabIndex        =   131
            Top             =   390
            Width           =   1170
         End
         Begin VB.OptionButton RiffleShufflePilesRandom 
            Caption         =   "Random"
            Height          =   390
            Left            =   60
            TabIndex        =   130
            Top             =   90
            Value           =   -1  'True
            Width           =   1170
         End
      End
      Begin VB.CommandButton RiffleShufflePileButton 
         Caption         =   "Riffle Shuffle"
         Height          =   360
         Left            =   105
         Style           =   1  'Graphical
         TabIndex        =   262
         TabStop         =   0   'False
         Top             =   330
         Width           =   1020
      End
      Begin VB.Label RiffleShuffleProtectedCardsLabel 
         Caption         =   "# cards"
         Enabled         =   0   'False
         Height          =   270
         Left            =   4335
         TabIndex        =   266
         Top             =   165
         Width           =   510
      End
   End
   Begin VB.Frame SelectReturnFrame 
      Height          =   2100
      Left            =   5640
      TabIndex        =   248
      Top             =   45
      Width           =   7740
      Begin VB.CheckBox MoveOnlyCheck 
         Caption         =   "Move only"
         Height          =   270
         Left            =   135
         TabIndex        =   107
         Top             =   1590
         Width           =   1815
      End
      Begin VB.CheckBox SelectionReverseCheck 
         Caption         =   "Reverse Selected Card"
         Height          =   270
         Left            =   135
         TabIndex        =   106
         Top             =   1305
         Width           =   1815
      End
      Begin VB.Frame ReturnPositionFrame 
         Height          =   1695
         Left            =   6090
         TabIndex        =   254
         Top             =   330
         Width           =   1575
         Begin VB.OptionButton ReturnPositionSame 
            Caption         =   "Same"
            Height          =   270
            Left            =   75
            TabIndex        =   124
            Top             =   615
            Width           =   825
         End
         Begin VB.TextBox ReturnPositionSpecifiedText 
            Alignment       =   2  'Center
            Height          =   330
            Left            =   975
            TabIndex        =   127
            Top             =   1095
            Width           =   450
         End
         Begin VB.OptionButton ReturnPositionRandom 
            Caption         =   "Random"
            Height          =   270
            Left            =   75
            TabIndex        =   125
            Top             =   840
            Width           =   825
         End
         Begin VB.OptionButton ReturnPositionSpecified 
            Caption         =   "Specified"
            Height          =   270
            Left            =   75
            TabIndex        =   126
            Top             =   1065
            Width           =   870
         End
         Begin VB.OptionButton ReturnPositionBottom 
            Caption         =   "Bottom"
            Height          =   270
            Left            =   75
            TabIndex        =   123
            Top             =   390
            Width           =   825
         End
         Begin VB.OptionButton ReturnPositionTop 
            Caption         =   "Top"
            Height          =   270
            Left            =   75
            TabIndex        =   122
            Top             =   165
            Value           =   -1  'True
            Width           =   825
         End
         Begin VB.Label Label3 
            Caption         =   "pos. #"
            Height          =   270
            Left            =   1005
            TabIndex        =   255
            Top             =   885
            Width           =   435
         End
      End
      Begin VB.Frame ReturnPileFrame 
         Height          =   1695
         Left            =   3780
         TabIndex        =   253
         Top             =   330
         Width           =   2265
         Begin VB.OptionButton ReturnPileNewPileSpecified 
            Caption         =   "New Pile Specified"
            Height          =   270
            Left            =   75
            TabIndex        =   120
            Top             =   1305
            Width           =   1500
         End
         Begin VB.TextBox ReturnPileNewPileSpecifiedText 
            Alignment       =   2  'Center
            Height          =   330
            Left            =   1710
            TabIndex        =   121
            Top             =   1290
            Width           =   450
         End
         Begin VB.OptionButton ReturnPileNewPileRandom 
            Caption         =   "New Pile Random"
            Height          =   270
            Left            =   75
            TabIndex        =   119
            Top             =   1065
            Width           =   1425
         End
         Begin VB.OptionButton ReturnPilePrimary 
            Caption         =   "Primary / Same"
            Height          =   270
            Left            =   75
            TabIndex        =   116
            Top             =   390
            Width           =   1380
         End
         Begin VB.OptionButton ReturnPileSecondary 
            Caption         =   "Secondary"
            Height          =   270
            Left            =   75
            TabIndex        =   115
            Top             =   165
            Value           =   -1  'True
            Width           =   1020
         End
         Begin VB.OptionButton ReturnPileRandomNotSame 
            Caption         =   "Random Not Same"
            Height          =   270
            Left            =   75
            TabIndex        =   117
            Top             =   615
            Width           =   1680
         End
         Begin VB.OptionButton ReturnPileRandomAny 
            Caption         =   "Random Any"
            Height          =   270
            Left            =   75
            TabIndex        =   118
            Top             =   840
            Width           =   1110
         End
         Begin VB.Label Label9 
            Caption         =   "pile #"
            Height          =   270
            Left            =   1770
            TabIndex        =   260
            Top             =   1080
            Width           =   405
         End
      End
      Begin VB.Frame SelectedCardFrame 
         Height          =   1695
         Left            =   2100
         TabIndex        =   251
         Top             =   330
         Width           =   1590
         Begin VB.OptionButton SelectedCardTop 
            Caption         =   "Top"
            Height          =   270
            Left            =   75
            TabIndex        =   110
            Top             =   165
            Value           =   -1  'True
            Width           =   825
         End
         Begin VB.OptionButton SelectedCardBottom 
            Caption         =   "Bottom"
            Height          =   270
            Left            =   75
            TabIndex        =   111
            Top             =   390
            Width           =   825
         End
         Begin VB.OptionButton SelectedCardSpecified 
            Caption         =   "Specified"
            Height          =   270
            Left            =   75
            TabIndex        =   113
            Top             =   840
            Width           =   870
         End
         Begin VB.OptionButton SelectedCardRandom 
            Caption         =   "Random"
            Height          =   270
            Left            =   75
            TabIndex        =   112
            Top             =   615
            Width           =   825
         End
         Begin VB.TextBox SelectedCardSpecifiedText 
            Alignment       =   2  'Center
            Height          =   330
            Left            =   1020
            TabIndex        =   114
            Top             =   840
            Width           =   450
         End
         Begin VB.Label Label4 
            Caption         =   "pos. #"
            Height          =   270
            Left            =   1050
            TabIndex        =   252
            Top             =   630
            Width           =   435
         End
      End
      Begin VB.Frame SelectionPileFrame 
         Height          =   915
         Left            =   825
         TabIndex        =   250
         Top             =   330
         Width           =   1245
         Begin VB.OptionButton SelectionPileRandom 
            Caption         =   "Random"
            Height          =   270
            Left            =   75
            TabIndex        =   109
            Top             =   390
            Width           =   1020
         End
         Begin VB.OptionButton SelectionPilePrimary 
            Caption         =   "Primary"
            Height          =   270
            Left            =   75
            TabIndex        =   108
            Top             =   165
            Value           =   -1  'True
            Width           =   825
         End
      End
      Begin VB.CommandButton SelectReturnButton 
         Caption         =   "Select / Return"
         Height          =   555
         Left            =   75
         Style           =   1  'Graphical
         TabIndex        =   249
         TabStop         =   0   'False
         Top             =   600
         Width           =   660
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "Return Position"
         Height          =   270
         Left            =   6255
         TabIndex        =   259
         Top             =   135
         Width           =   1215
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "Return Pile"
         Height          =   270
         Left            =   3930
         TabIndex        =   258
         Top             =   135
         Width           =   1215
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "Selected Card"
         Height          =   270
         Left            =   2280
         TabIndex        =   257
         Top             =   135
         Width           =   1215
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Selection Pile"
         Height          =   270
         Left            =   840
         TabIndex        =   256
         Top             =   135
         Width           =   1215
      End
      Begin VB.Line Line2 
         BorderColor     =   &H8000000A&
         X1              =   3720
         X2              =   3720
         Y1              =   120
         Y2              =   1725
      End
   End
   Begin VB.Frame CombineFrame 
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   930
      Left            =   5610
      TabIndex        =   233
      Top             =   2190
      Width           =   2760
      Begin VB.CommandButton CombinePilesButton 
         Caption         =   "Combine"
         Height          =   360
         Left            =   75
         Style           =   1  'Graphical
         TabIndex        =   235
         TabStop         =   0   'False
         Top             =   345
         Width           =   1020
      End
      Begin VB.Frame Frame13 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1125
         TabIndex        =   234
         Top             =   135
         Width           =   1530
         Begin VB.OptionButton CombineSecondaryTop 
            Caption         =   "Secondary on Top"
            Height          =   390
            Left            =   60
            TabIndex        =   128
            Top             =   90
            Value           =   -1  'True
            Width           =   1455
         End
         Begin VB.OptionButton CombinePrimaryTop 
            Caption         =   "Primary on Top"
            Height          =   390
            Left            =   60
            TabIndex        =   129
            Top             =   390
            Width           =   1425
         End
      End
   End
   Begin VB.Frame PileMatrixFrame 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2535
      Index           =   8
      Left            =   3120
      TabIndex        =   207
      Top             =   495
      Width           =   2310
      Begin VB.CheckBox PileMatrix 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   15
         Left            =   1440
         TabIndex        =   30
         Top             =   870
         Width           =   195
      End
      Begin VB.CheckBox PileMatrix 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   13
         Left            =   1050
         TabIndex        =   28
         Top             =   870
         Width           =   195
      End
      Begin VB.CheckBox PileMatrix 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   14
         Left            =   1245
         TabIndex        =   29
         Top             =   870
         Width           =   195
      End
      Begin VB.CheckBox PileMatrix 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   56
         Left            =   1635
         TabIndex        =   63
         Top             =   1650
         Width           =   195
      End
      Begin VB.CheckBox PileMatrix 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   66
         Left            =   1635
         TabIndex        =   71
         Top             =   1845
         Width           =   195
      End
      Begin VB.CheckBox PileMatrix 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   11
         Left            =   660
         TabIndex        =   26
         Top             =   870
         Width           =   195
      End
      Begin VB.CheckBox PileMatrix 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   25
         Left            =   1440
         TabIndex        =   38
         Top             =   1065
         Width           =   195
      End
      Begin VB.CheckBox PileMatrix 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   64
         Left            =   1245
         TabIndex        =   69
         Top             =   1845
         Width           =   195
      End
      Begin VB.CheckBox PileMatrix 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   26
         Left            =   1635
         TabIndex        =   39
         Top             =   1065
         Width           =   195
      End
      Begin VB.CheckBox PileMatrix 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   42
         Left            =   855
         TabIndex        =   51
         Top             =   1455
         Width           =   195
      End
      Begin VB.CheckBox PileMatrix 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   33
         Left            =   1050
         TabIndex        =   44
         Top             =   1260
         Width           =   195
      End
      Begin VB.CheckBox PileMatrix 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   36
         Left            =   1635
         TabIndex        =   47
         Top             =   1260
         Width           =   195
      End
      Begin VB.CheckBox PileMatrix 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   31
         Left            =   660
         TabIndex        =   42
         Top             =   1260
         Width           =   195
      End
      Begin VB.CheckBox PileMatrix 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   32
         Left            =   855
         TabIndex        =   43
         Top             =   1260
         Width           =   195
      End
      Begin VB.CheckBox PileMatrix 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   34
         Left            =   1245
         TabIndex        =   45
         Top             =   1260
         Width           =   195
      End
      Begin VB.CheckBox PileMatrix 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   22
         Left            =   855
         TabIndex        =   35
         Top             =   1065
         Width           =   195
      End
      Begin VB.CheckBox PileMatrix 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   61
         Left            =   660
         TabIndex        =   66
         Top             =   1845
         Width           =   195
      End
      Begin VB.CheckBox PileMatrix 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   62
         Left            =   855
         TabIndex        =   67
         Top             =   1845
         Width           =   195
      End
      Begin VB.CheckBox PileMatrix 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   53
         Left            =   1050
         TabIndex        =   60
         Top             =   1650
         Width           =   195
      End
      Begin VB.CheckBox PileMatrix 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   55
         Left            =   1440
         TabIndex        =   62
         Top             =   1650
         Width           =   195
      End
      Begin VB.CheckBox PileMatrix 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   16
         Left            =   1635
         TabIndex        =   31
         Top             =   870
         Width           =   195
      End
      Begin VB.CheckBox PileMatrix 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   63
         Left            =   1050
         TabIndex        =   68
         Top             =   1845
         Width           =   195
      End
      Begin VB.CheckBox PileMatrix 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   41
         Left            =   660
         TabIndex        =   50
         Top             =   1455
         Width           =   195
      End
      Begin VB.CheckBox PileMatrix 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   24
         Left            =   1245
         TabIndex        =   37
         Top             =   1065
         Width           =   195
      End
      Begin VB.CheckBox PileMatrix 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   46
         Left            =   1635
         TabIndex        =   55
         Top             =   1455
         Width           =   195
      End
      Begin VB.CheckBox PileMatrix 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   51
         Left            =   660
         TabIndex        =   58
         Top             =   1650
         Width           =   195
      End
      Begin VB.CheckBox PileMatrix 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   65
         Left            =   1440
         TabIndex        =   70
         Top             =   1845
         Width           =   195
      End
      Begin VB.CheckBox PileMatrix 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   44
         Left            =   1245
         TabIndex        =   53
         Top             =   1455
         Width           =   195
      End
      Begin VB.CheckBox PileMatrix 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   45
         Left            =   1440
         TabIndex        =   54
         Top             =   1455
         Width           =   195
      End
      Begin VB.CheckBox PileMatrix 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   35
         Left            =   1440
         TabIndex        =   46
         Top             =   1260
         Width           =   195
      End
      Begin VB.CheckBox PileMatrix 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   54
         Left            =   1245
         TabIndex        =   61
         Top             =   1650
         Width           =   195
      End
      Begin VB.CheckBox PileMatrix 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   12
         Left            =   855
         TabIndex        =   27
         Top             =   870
         Width           =   195
      End
      Begin VB.CheckBox PileMatrix 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   43
         Left            =   1050
         TabIndex        =   52
         Top             =   1455
         Width           =   195
      End
      Begin VB.CheckBox PileMatrix 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   52
         Left            =   855
         TabIndex        =   59
         Top             =   1650
         Width           =   195
      End
      Begin VB.CheckBox PileMatrix 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   23
         Left            =   1050
         TabIndex        =   36
         Top             =   1065
         Width           =   195
      End
      Begin VB.CheckBox PileMatrix 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   21
         Left            =   660
         TabIndex        =   34
         Top             =   1065
         Width           =   195
      End
      Begin VB.CheckBox PileMatrix 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   81
         Left            =   660
         TabIndex        =   82
         Top             =   2235
         Width           =   195
      End
      Begin VB.CheckBox PileMatrix 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   71
         Left            =   660
         TabIndex        =   74
         Top             =   2040
         Width           =   195
      End
      Begin VB.CheckBox PileMatrix 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   82
         Left            =   855
         TabIndex        =   83
         Top             =   2235
         Width           =   195
      End
      Begin VB.CheckBox PileMatrix 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   72
         Left            =   855
         TabIndex        =   75
         Top             =   2040
         Width           =   195
      End
      Begin VB.CheckBox PileMatrix 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   83
         Left            =   1050
         TabIndex        =   84
         Top             =   2235
         Width           =   195
      End
      Begin VB.CheckBox PileMatrix 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   73
         Left            =   1050
         TabIndex        =   76
         Top             =   2040
         Width           =   195
      End
      Begin VB.CheckBox PileMatrix 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   84
         Left            =   1245
         TabIndex        =   85
         Top             =   2235
         Width           =   195
      End
      Begin VB.CheckBox PileMatrix 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   74
         Left            =   1245
         TabIndex        =   77
         Top             =   2040
         Width           =   195
      End
      Begin VB.CheckBox PileMatrix 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   85
         Left            =   1440
         TabIndex        =   86
         Top             =   2235
         Width           =   195
      End
      Begin VB.CheckBox PileMatrix 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   75
         Left            =   1440
         TabIndex        =   78
         Top             =   2040
         Width           =   195
      End
      Begin VB.CheckBox PileMatrix 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   86
         Left            =   1635
         TabIndex        =   87
         Top             =   2235
         Width           =   195
      End
      Begin VB.CheckBox PileMatrix 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   76
         Left            =   1635
         TabIndex        =   79
         Top             =   2040
         Width           =   195
      End
      Begin VB.CheckBox PileMatrix 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   88
         Left            =   2025
         TabIndex        =   89
         Top             =   2235
         Width           =   195
      End
      Begin VB.CheckBox PileMatrix 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   78
         Left            =   2025
         TabIndex        =   81
         Top             =   2040
         Width           =   195
      End
      Begin VB.CheckBox PileMatrix 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   87
         Left            =   1830
         TabIndex        =   88
         Top             =   2235
         Width           =   195
      End
      Begin VB.CheckBox PileMatrix 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   77
         Left            =   1845
         TabIndex        =   80
         Top             =   2040
         Width           =   195
      End
      Begin VB.CheckBox PileMatrix 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   68
         Left            =   2025
         TabIndex        =   73
         Top             =   1845
         Width           =   195
      End
      Begin VB.CheckBox PileMatrix 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   58
         Left            =   2025
         TabIndex        =   65
         Top             =   1650
         Width           =   195
      End
      Begin VB.CheckBox PileMatrix 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   48
         Left            =   2025
         TabIndex        =   57
         Top             =   1455
         Width           =   195
      End
      Begin VB.CheckBox PileMatrix 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   38
         Left            =   2025
         TabIndex        =   49
         Top             =   1260
         Width           =   195
      End
      Begin VB.CheckBox PileMatrix 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   28
         Left            =   2025
         TabIndex        =   41
         Top             =   1065
         Width           =   195
      End
      Begin VB.CheckBox PileMatrix 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   18
         Left            =   2025
         TabIndex        =   33
         Top             =   870
         Width           =   195
      End
      Begin VB.CheckBox PileMatrix 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   67
         Left            =   1830
         TabIndex        =   72
         Top             =   1845
         Width           =   195
      End
      Begin VB.CheckBox PileMatrix 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   57
         Left            =   1830
         TabIndex        =   64
         Top             =   1650
         Width           =   195
      End
      Begin VB.CheckBox PileMatrix 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   47
         Left            =   1830
         TabIndex        =   56
         Top             =   1455
         Width           =   195
      End
      Begin VB.CheckBox PileMatrix 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   37
         Left            =   1830
         TabIndex        =   48
         Top             =   1260
         Width           =   195
      End
      Begin VB.CheckBox PileMatrix 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   27
         Left            =   1830
         TabIndex        =   40
         Top             =   1065
         Width           =   195
      End
      Begin VB.CheckBox PileMatrix 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   17
         Left            =   1830
         TabIndex        =   32
         Top             =   870
         Width           =   195
      End
      Begin VB.CheckBox ReverseR 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   60
         TabIndex        =   98
         Top             =   870
         Width           =   195
      End
      Begin VB.CheckBox ReverseR 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   8
         Left            =   60
         TabIndex        =   105
         Top             =   2235
         Width           =   195
      End
      Begin VB.CheckBox ReverseR 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   60
         TabIndex        =   99
         Top             =   1065
         Width           =   195
      End
      Begin VB.CheckBox ReverseR 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   60
         TabIndex        =   100
         Top             =   1260
         Width           =   195
      End
      Begin VB.CheckBox ReverseR 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   60
         TabIndex        =   101
         Top             =   1455
         Width           =   195
      End
      Begin VB.CheckBox ReverseR 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   60
         TabIndex        =   102
         Top             =   1650
         Width           =   195
      End
      Begin VB.CheckBox ReverseR 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   60
         TabIndex        =   103
         Top             =   1845
         Width           =   195
      End
      Begin VB.CheckBox ReverseR 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   7
         Left            =   60
         TabIndex        =   104
         Top             =   2040
         Width           =   195
      End
      Begin VB.CheckBox ReverseC 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   8
         Left            =   2025
         TabIndex        =   97
         Top             =   240
         Width           =   195
      End
      Begin VB.CheckBox ReverseC 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   660
         TabIndex        =   90
         Top             =   240
         Width           =   195
      End
      Begin VB.CheckBox ReverseC 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   855
         TabIndex        =   91
         Top             =   240
         Width           =   195
      End
      Begin VB.CheckBox ReverseC 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   1050
         TabIndex        =   92
         Top             =   240
         Width           =   195
      End
      Begin VB.CheckBox ReverseC 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   1245
         TabIndex        =   93
         Top             =   240
         Width           =   195
      End
      Begin VB.CheckBox ReverseC 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   1440
         TabIndex        =   94
         Top             =   240
         Width           =   195
      End
      Begin VB.CheckBox ReverseC 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   1635
         TabIndex        =   95
         Top             =   240
         Width           =   195
      End
      Begin VB.CheckBox ReverseC 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   7
         Left            =   1830
         TabIndex        =   96
         Top             =   240
         Width           =   195
      End
      Begin VB.Label Label2 
         Caption         =   "Rev"
         Height          =   240
         Left            =   45
         TabIndex        =   247
         Top             =   630
         Width           =   270
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         Caption         =   "Rev"
         Height          =   240
         Left            =   375
         TabIndex        =   246
         Top             =   195
         Width           =   270
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label39 
         Caption         =   "P"
         Height          =   210
         Left            =   330
         TabIndex        =   208
         Top             =   1185
         Width           =   120
      End
      Begin VB.Label Label40 
         Caption         =   "r"
         Height          =   210
         Left            =   345
         TabIndex        =   209
         Top             =   1305
         Width           =   120
      End
      Begin VB.Label Label41 
         Caption         =   "i"
         Height          =   210
         Left            =   345
         TabIndex        =   210
         Top             =   1455
         Width           =   120
      End
      Begin VB.Label Label42 
         Caption         =   "m"
         Height          =   210
         Left            =   315
         TabIndex        =   211
         Top             =   1590
         Width           =   120
      End
      Begin VB.Label Label43 
         Caption         =   "a"
         Height          =   210
         Left            =   330
         TabIndex        =   212
         Top             =   1710
         Width           =   120
      End
      Begin VB.Label Label44 
         Caption         =   "r"
         Height          =   210
         Left            =   345
         TabIndex        =   213
         Top             =   1845
         Width           =   120
      End
      Begin VB.Label Label45 
         Caption         =   "y"
         Height          =   210
         Left            =   330
         TabIndex        =   214
         Top             =   1965
         Width           =   120
      End
      Begin VB.Label PileMatrixLabel 
         Caption         =   "6"
         Height          =   195
         Index           =   11
         Left            =   495
         TabIndex        =   231
         Top             =   1815
         Width           =   105
      End
      Begin VB.Label PileMatrixLabel 
         Caption         =   "5"
         Height          =   195
         Index           =   10
         Left            =   495
         TabIndex        =   230
         Top             =   1620
         Width           =   105
      End
      Begin VB.Label PileMatrixLabel 
         Caption         =   "4"
         Height          =   195
         Index           =   9
         Left            =   495
         TabIndex        =   229
         Top             =   1425
         Width           =   105
      End
      Begin VB.Label PileMatrixLabel 
         Caption         =   "3"
         Height          =   195
         Index           =   8
         Left            =   495
         TabIndex        =   228
         Top             =   1230
         Width           =   105
      End
      Begin VB.Label PileMatrixLabel 
         Caption         =   "2"
         Height          =   195
         Index           =   7
         Left            =   495
         TabIndex        =   227
         Top             =   1035
         Width           =   105
      End
      Begin VB.Label PileMatrixLabel 
         Caption         =   "1"
         Height          =   195
         Index           =   6
         Left            =   495
         TabIndex        =   226
         Top             =   840
         Width           =   105
      End
      Begin VB.Label PileMatrixLabel 
         Alignment       =   2  'Center
         Caption         =   "1"
         Height          =   195
         Index           =   0
         Left            =   720
         TabIndex        =   225
         Top             =   600
         Width           =   105
      End
      Begin VB.Label PileMatrixLabel 
         Caption         =   "8"
         Height          =   195
         Index           =   14
         Left            =   495
         TabIndex        =   224
         Top             =   2205
         Width           =   105
      End
      Begin VB.Label PileMatrixLabel 
         Caption         =   "7"
         Height          =   195
         Index           =   15
         Left            =   495
         TabIndex        =   223
         Top             =   2010
         Width           =   105
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00808080&
         Height          =   1680
         Left            =   600
         Top             =   810
         Width           =   1665
      End
      Begin VB.Label PileMatrixLabel 
         Alignment       =   2  'Center
         Caption         =   "8"
         Height          =   195
         Index           =   1
         Left            =   2085
         TabIndex        =   222
         Top             =   600
         Width           =   105
      End
      Begin VB.Label PileMatrixLabel 
         Alignment       =   2  'Center
         Caption         =   "7"
         Height          =   195
         Index           =   2
         Left            =   1890
         TabIndex        =   221
         Top             =   600
         Width           =   105
      End
      Begin VB.Label PileMatrixLabel 
         Alignment       =   2  'Center
         Caption         =   "6"
         Height          =   195
         Index           =   3
         Left            =   1695
         TabIndex        =   220
         Top             =   600
         Width           =   105
      End
      Begin VB.Label PileMatrixLabel 
         Alignment       =   2  'Center
         Caption         =   "5"
         Height          =   195
         Index           =   4
         Left            =   1500
         TabIndex        =   219
         Top             =   600
         Width           =   105
      End
      Begin VB.Label PileMatrixLabel 
         Alignment       =   2  'Center
         Caption         =   "4"
         Height          =   195
         Index           =   5
         Left            =   1305
         TabIndex        =   218
         Top             =   600
         Width           =   105
      End
      Begin VB.Label PileMatrixLabel 
         Alignment       =   2  'Center
         Caption         =   "3"
         Height          =   195
         Index           =   12
         Left            =   1110
         TabIndex        =   217
         Top             =   600
         Width           =   105
      End
      Begin VB.Label PileMatrixLabel 
         Alignment       =   2  'Center
         Caption         =   "2"
         Height          =   195
         Index           =   13
         Left            =   915
         TabIndex        =   216
         Top             =   600
         Width           =   105
      End
      Begin VB.Label Label32 
         Caption         =   "Secondary"
         Height          =   240
         Left            =   1050
         TabIndex        =   215
         Top             =   405
         Width           =   750
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame PileParametersFrame 
      Caption         =   "Pile Specification"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4725
      Left            =   90
      TabIndex        =   177
      Top             =   1575
      Width           =   2910
      Begin VB.Frame DealCardsFrame 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   60
         TabIndex        =   236
         Top             =   555
         Visible         =   0   'False
         Width           =   2790
         Begin VB.Frame Frame10 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2175
            Left            =   45
            TabIndex        =   239
            Top             =   810
            Width           =   2655
            Begin VB.OptionButton CompleteRandomOption 
               Caption         =   "Complete Random Sequence"
               Height          =   270
               Left            =   75
               TabIndex        =   243
               Top             =   855
               Width           =   2490
            End
            Begin VB.OptionButton AlternatingRegularOption 
               Caption         =   "Regular Alternating Sequence"
               Height          =   270
               Left            =   75
               TabIndex        =   242
               Top             =   225
               Value           =   -1  'True
               Width           =   2490
            End
            Begin VB.TextBox NumberOfCardsToDealText 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   1305
               TabIndex        =   241
               Top             =   1305
               Width           =   450
            End
            Begin VB.OptionButton AlternatingRandomOption 
               Caption         =   "Random Alternating Sequence"
               Height          =   270
               Left            =   75
               TabIndex        =   240
               Top             =   540
               Width           =   2490
            End
            Begin VB.Label Label29 
               Caption         =   "(if blank, whole deck will be dealt)"
               Height          =   270
               Left            =   210
               TabIndex        =   245
               Top             =   1785
               Width           =   2250
            End
            Begin VB.Label Label28 
               Caption         =   "Total Number of Cards to Deal"
               Height          =   435
               Left            =   195
               TabIndex        =   244
               Top             =   1230
               Width           =   1140
            End
         End
         Begin VB.OptionButton DealAlternatingOption 
            Caption         =   "Alternating Piles"
            Height          =   270
            Left            =   60
            TabIndex        =   238
            Top             =   450
            Width           =   2565
         End
         Begin VB.OptionButton DealCompleteOption 
            Caption         =   "Complete Specified Piles"
            Height          =   270
            Left            =   60
            TabIndex        =   237
            Top             =   150
            Value           =   -1  'True
            Width           =   2280
         End
         Begin VB.Line Line4 
            BorderColor     =   &H8000000A&
            X1              =   45
            X2              =   2760
            Y1              =   765
            Y2              =   765
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   45
         TabIndex        =   204
         Top             =   195
         Width           =   2835
         Begin VB.OptionButton CutCardsOption 
            Caption         =   "Cut Cards"
            Height          =   360
            Left            =   240
            TabIndex        =   206
            Top             =   0
            Value           =   -1  'True
            Width           =   960
         End
         Begin VB.OptionButton DealCardsOption 
            Caption         =   "Deal Cards"
            Height          =   360
            Left            =   1455
            TabIndex        =   205
            Top             =   0
            Width           =   1005
         End
      End
      Begin VB.Frame CutCardsFrame 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4110
         Left            =   45
         TabIndex        =   178
         Top             =   570
         Width           =   2805
         Begin VB.Frame Frame9 
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   15
            TabIndex        =   200
            Top             =   3645
            Width           =   2745
            Begin VB.OptionButton PileRandom8 
               Caption         =   "Random"
               Height          =   270
               Left            =   285
               TabIndex        =   17
               Top             =   30
               Value           =   -1  'True
               Width           =   840
            End
            Begin VB.OptionButton PileSpecified8 
               Caption         =   "Specified"
               Height          =   270
               Left            =   1170
               TabIndex        =   25
               Top             =   30
               Width           =   840
            End
            Begin VB.TextBox SpecifiedText8 
               Alignment       =   2  'Center
               Height          =   300
               Left            =   2040
               TabIndex        =   9
               Top             =   30
               Width           =   450
            End
            Begin VB.Label Label26 
               Alignment       =   2  'Center
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
               Left            =   60
               TabIndex        =   202
               Top             =   30
               Width           =   180
            End
            Begin VB.Label PileMarker8 
               Caption         =   "A"
               Height          =   240
               Left            =   2565
               TabIndex        =   201
               Top             =   90
               Width           =   150
            End
         End
         Begin VB.Frame Frame8 
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   15
            TabIndex        =   197
            Top             =   3285
            Width           =   2745
            Begin VB.TextBox SpecifiedText7 
               Alignment       =   2  'Center
               Height          =   300
               Left            =   2040
               TabIndex        =   8
               Top             =   30
               Width           =   450
            End
            Begin VB.OptionButton PileSpecified7 
               Caption         =   "Specified"
               Height          =   270
               Left            =   1170
               TabIndex        =   24
               Top             =   30
               Width           =   840
            End
            Begin VB.OptionButton PileRandom7 
               Caption         =   "Random"
               Height          =   270
               Left            =   285
               TabIndex        =   16
               Top             =   30
               Value           =   -1  'True
               Width           =   840
            End
            Begin VB.Label Label25 
               Alignment       =   2  'Center
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
               Left            =   60
               TabIndex        =   199
               Top             =   30
               Width           =   180
            End
            Begin VB.Label PileMarker7 
               Caption         =   "A"
               Height          =   240
               Left            =   2565
               TabIndex        =   198
               Top             =   90
               Width           =   150
            End
         End
         Begin VB.Frame Frame7 
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   15
            TabIndex        =   194
            Top             =   2925
            Width           =   2745
            Begin VB.OptionButton PileRandom6 
               Caption         =   "Random"
               Height          =   270
               Left            =   285
               TabIndex        =   15
               Top             =   30
               Value           =   -1  'True
               Width           =   840
            End
            Begin VB.OptionButton PileSpecified6 
               Caption         =   "Specified"
               Height          =   270
               Left            =   1170
               TabIndex        =   23
               Top             =   30
               Width           =   840
            End
            Begin VB.TextBox SpecifiedText6 
               Alignment       =   2  'Center
               Height          =   300
               Left            =   2040
               TabIndex        =   7
               Top             =   30
               Width           =   450
            End
            Begin VB.Label Label24 
               Alignment       =   2  'Center
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
               Left            =   60
               TabIndex        =   196
               Top             =   30
               Width           =   180
            End
            Begin VB.Label PileMarker6 
               Caption         =   "A"
               Height          =   240
               Left            =   2565
               TabIndex        =   195
               Top             =   90
               Width           =   150
            End
         End
         Begin VB.Frame Frame6 
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   15
            TabIndex        =   191
            Top             =   2565
            Width           =   2745
            Begin VB.TextBox SpecifiedText5 
               Alignment       =   2  'Center
               Height          =   300
               Left            =   2040
               TabIndex        =   6
               Top             =   30
               Width           =   450
            End
            Begin VB.OptionButton PileSpecified5 
               Caption         =   "Specified"
               Height          =   270
               Left            =   1170
               TabIndex        =   22
               Top             =   30
               Width           =   840
            End
            Begin VB.OptionButton PileRandom5 
               Caption         =   "Random"
               Height          =   270
               Left            =   285
               TabIndex        =   14
               Top             =   30
               Value           =   -1  'True
               Width           =   840
            End
            Begin VB.Label Label23 
               Alignment       =   2  'Center
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
               Left            =   60
               TabIndex        =   193
               Top             =   30
               Width           =   180
            End
            Begin VB.Label PileMarker5 
               Caption         =   "A"
               Height          =   240
               Left            =   2565
               TabIndex        =   192
               Top             =   90
               Width           =   150
            End
         End
         Begin VB.Frame Frame4 
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   15
            TabIndex        =   188
            Top             =   1845
            Width           =   2745
            Begin VB.TextBox SpecifiedText3 
               Alignment       =   2  'Center
               Height          =   300
               Left            =   2040
               TabIndex        =   4
               Top             =   30
               Width           =   450
            End
            Begin VB.OptionButton PileSpecified3 
               Caption         =   "Specified"
               Height          =   270
               Left            =   1185
               TabIndex        =   20
               Top             =   30
               Width           =   840
            End
            Begin VB.OptionButton PileRandom3 
               Caption         =   "Random"
               Height          =   270
               Left            =   285
               TabIndex        =   12
               Top             =   30
               Value           =   -1  'True
               Width           =   840
            End
            Begin VB.Label Label21 
               Alignment       =   2  'Center
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
               Left            =   60
               TabIndex        =   190
               Top             =   30
               Width           =   180
            End
            Begin VB.Label PileMarker3 
               Caption         =   "A"
               Height          =   240
               Left            =   2565
               TabIndex        =   189
               Top             =   90
               Width           =   150
            End
         End
         Begin VB.Frame Frame3 
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   15
            TabIndex        =   185
            Top             =   1485
            Width           =   2745
            Begin VB.OptionButton PileRandom2 
               Caption         =   "Random"
               Height          =   270
               Left            =   285
               TabIndex        =   11
               Top             =   30
               Value           =   -1  'True
               Width           =   840
            End
            Begin VB.OptionButton PileSpecified2 
               Caption         =   "Specified"
               Height          =   270
               Left            =   1170
               TabIndex        =   19
               Top             =   30
               Width           =   840
            End
            Begin VB.TextBox SpecifiedText2 
               Alignment       =   2  'Center
               Height          =   300
               Left            =   2040
               TabIndex        =   3
               Top             =   30
               Width           =   450
            End
            Begin VB.Label Label20 
               Alignment       =   2  'Center
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
               Left            =   60
               TabIndex        =   187
               Top             =   30
               Width           =   180
            End
            Begin VB.Label PileMarker2 
               Caption         =   "A"
               Height          =   240
               Left            =   2565
               TabIndex        =   186
               Top             =   90
               Width           =   150
            End
         End
         Begin VB.Frame Frame1 
            BorderStyle     =   0  'None
            Height          =   405
            Left            =   15
            TabIndex        =   182
            Top             =   1125
            Width           =   2745
            Begin VB.TextBox SpecifiedText1 
               Alignment       =   2  'Center
               Height          =   300
               Left            =   2040
               TabIndex        =   2
               Top             =   30
               Width           =   450
            End
            Begin VB.OptionButton PileSpecified1 
               Caption         =   "Specified"
               Height          =   270
               Left            =   1170
               TabIndex        =   18
               Top             =   30
               Width           =   840
            End
            Begin VB.OptionButton PileRandom1 
               Caption         =   "Random"
               Height          =   270
               Left            =   285
               TabIndex        =   10
               Top             =   30
               Value           =   -1  'True
               Width           =   840
            End
            Begin VB.Label Label17 
               Alignment       =   2  'Center
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
               Left            =   60
               TabIndex        =   184
               Top             =   30
               Width           =   180
            End
            Begin VB.Label PileMarker1 
               Caption         =   "A"
               Height          =   240
               Left            =   2565
               TabIndex        =   183
               Top             =   90
               Width           =   150
            End
         End
         Begin VB.Frame Frame5 
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   15
            TabIndex        =   179
            Top             =   2205
            Width           =   2745
            Begin VB.OptionButton PileRandom4 
               Caption         =   "Random"
               Height          =   270
               Left            =   285
               TabIndex        =   13
               Top             =   30
               Value           =   -1  'True
               Width           =   840
            End
            Begin VB.OptionButton PileSpecified4 
               Caption         =   "Specified"
               Height          =   270
               Left            =   1170
               TabIndex        =   21
               Top             =   30
               Width           =   840
            End
            Begin VB.TextBox SpecifiedText4 
               Alignment       =   2  'Center
               Height          =   300
               Left            =   2040
               TabIndex        =   5
               Top             =   30
               Width           =   450
            End
            Begin VB.Label Label22 
               Alignment       =   2  'Center
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
               Left            =   60
               TabIndex        =   181
               Top             =   30
               Width           =   180
            End
            Begin VB.Label PileMarker4 
               Caption         =   "A"
               Height          =   240
               Left            =   2565
               TabIndex        =   180
               Top             =   90
               Width           =   150
            End
         End
         Begin VB.Label Label19 
            Caption         =   "Pile #"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   15
            TabIndex        =   203
            Top             =   870
            Width           =   735
         End
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000A&
         X1              =   0
         X2              =   2880
         Y1              =   540
         Y2              =   540
      End
   End
   Begin VB.CommandButton CreatePilesButton 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Create Piles"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   180
      Style           =   1  'Graphical
      TabIndex        =   175
      TabStop         =   0   'False
      Top             =   1125
      Width           =   1260
   End
   Begin VB.TextBox NumberOfPilesText 
      Alignment       =   2  'Center
      Height          =   330
      Left            =   1545
      TabIndex        =   1
      Top             =   1125
      Width           =   450
   End
   Begin VB.Frame PileViewpointFrame 
      Caption         =   "Pile Viewpoint"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   90
      TabIndex        =   0
      Top             =   120
      Width           =   2925
      Begin VB.OptionButton ViewPilesBeneath 
         Caption         =   "View Piles from BENEATH table (main)"
         Height          =   300
         Left            =   60
         TabIndex        =   174
         Top             =   225
         Value           =   -1  'True
         Width           =   2790
      End
      Begin VB.OptionButton ViewPilesAbove 
         Caption         =   "View Piles from ABOVE table"
         Height          =   300
         Left            =   60
         TabIndex        =   173
         Top             =   510
         Width           =   2625
      End
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H8000000F&
      BorderWidth     =   2
      Height          =   2895
      Left            =   3075
      Top             =   180
      Width           =   2400
   End
   Begin VB.Label Label27 
      Caption         =   "Pile Control Matrix"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3210
      TabIndex        =   232
      Top             =   270
      Width           =   2160
   End
   Begin VB.Label Label18 
      Caption         =   "(1 - 8 piles)"
      Height          =   270
      Left            =   2085
      TabIndex        =   176
      Top             =   1170
      Width           =   795
   End
End
Attribute VB_Name = "frmPiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub AustralianDeal(p1, p2, p3, p4, p5, p6, p7, p8, p9, p10, p11)
Dim pStartDD As Integer
Dim pFinishDD As Integer
Dim pCounterDD As Integer
Dim pRandomSelection As Integer
Dim pUpCounter As Integer
'this variable is used to count up in the case of an inverse request
pUpCounter = 1
Dim pDuck As Boolean
Dim pDeal As Boolean
Dim pileNum As Integer
pileNum = Val(Right(p1, 1))
Dim pStartingOrder As Variant
'D=down/under, U=under/down
pStartingOrder = Left(p1, 1)
Dim pInverse As Integer
pInverse = p11
pStartDD = PileTable(pileNum, 1)
pFinishDD = PileTable(pileNum, 2)
pCounterDD = pFinishDD - pStartDD + 1
'set the randomselection identifier if necessary
If p10 = "R" Then
    pRandomSelection = Int(Rnd * pCounterDD) + 1
End If
Dim pDealCounter As Integer
Dim pDuckCounter As Integer
'set the deal and duck counters
If Val(p2) > 0 Then
    pDealCounter = Val(p2)
Else
    pDealCounter = Int(Rnd * Val(p4)) + 1
End If
If Val(p3) > 0 Then
    pDuckCounter = Val(p3)
Else
    pDuckCounter = Int(Rnd * Val(p5)) + 1
End If
'control for inverse request correctly
If pInverse = 0 Then
    'set deal and duck parameters for Deal & Duck
    If pStartingOrder = "D" Then
        pDeal = True
        pDuck = False
    ElseIf pStartingOrder = "U" Then
        pDeal = False
        pDuck = True
    End If
    'do while counter>1
    'Since counter represents how many cards are in hand,
    'we do not want to do the deal and duck sequence when there is one card
    While pCounterDD > 1
        'first move the top card to the bottom of the pile
        For m% = 1 To DeckProperties
            ChangedDeck(m%, pStartDD + pCounterDD - 1) = Deck(m%, pStartDD)
        Next m%
        'next, shift the rest of the cards up one position
        For p% = pStartDD To pStartDD + pCounterDD - 2
            For m% = 1 To DeckProperties
                ChangedDeck(m%, p%) = Deck(m%, p% + 1)
            Next m%
        Next p%
        'transfer cards back to Deck()
        For p% = pStartDD To pStartDD + pCounterDD - 1
            For m% = 1 To DeckProperties
                Deck(m%, p%) = ChangedDeck(m%, p%)
            Next m%
        Next p%
        'if random selection, check on deal state and random match
        If pDeal And p10 = "R" And pRandomSelection = pCounterDD Then
            Deck(4, pStartDD + pCounterDD - 1) = "Selected"
            If frmStackView.SelectionsTextBox.Text = Empty Then
                frmStackView.SelectionsTextBox.Text = Deck(2, pStartDD + pCounterDD - 1)
            Else
                frmStackView.SelectionsTextBox.Text = frmStackView.SelectionsTextBox.Text _
                    & " " & Deck(2, pStartDD + pCounterDD - 1)
            End If
            If p6 = "S" Then
                Deck(6, pStartDD + pCounterDD - 1) = _
                    Not Deck(6, pStartDD + pCounterDD - 1)
            End If
        End If
        'check for reverse conditions
        If pDeal Then
            'when RND<0.5 then allow the random setting to reverse the card
            If p6 = 1 Or (p8 = 1 And Rnd < 0.5) Then
                Deck(6, pStartDD + pCounterDD - 1) = _
                    Not Deck(6, pStartDD + pCounterDD - 1)
            End If
        ElseIf pDuck Then
            'when RND<0.5 then allow the random setting to reverse the card
            If p7 = 1 Or (p9 = 1 And Rnd < 0.5) Then
                Deck(6, pStartDD + pCounterDD - 1) = _
                    Not Deck(6, pStartDD + pCounterDD - 1)
            End If
        End If
        'adjust counter based on current action
        If pDeal Then
            pCounterDD = pCounterDD - 1
        End If
        'adjust deal counter
        If pDeal Then
            pDealCounter = pDealCounter - 1
            If pDealCounter = 0 Then
                'switch logical settings for deal and duck
                pDuck = Not pDuck
                pDeal = Not pDeal
                'reset the deal counter
                If Val(p2) > 0 Then
                    pDealCounter = Val(p2)
                Else
                    pDealCounter = Int(Rnd * Val(p4)) + 1
                End If
            End If
        ElseIf pDuck Then
            pDuckCounter = pDuckCounter - 1
            If pDuckCounter = 0 Then
                'switch logical settings for deal and duck
                pDuck = Not pDuck
                pDeal = Not pDeal
                'reset the duck counter
                If Val(p3) > 0 Then
                    pDuckCounter = Val(p3)
                Else
                    pDuckCounter = Int(Rnd * Val(p5)) + 1
                End If
            End If
        End If
    Wend
    'special case in case there is one card in hand
    'and there are still ducks to do with reverses
    If pDuck Then
        If Val(p3) > 0 Then
            pDuckCounter = Val(p3)
        Else
            pDuckCounter = Int(Rnd * Val(p5)) + 1
        End If
        For i% = 1 To pDuckCounter
            'when RND<0.5 then allow the random setting to reverse the card
            If p7 = 1 Or (p9 = 1 And Rnd < 0.5) Then
                Deck(6, pStartDD + pCounterDD - 1) = _
                    Not Deck(6, pStartDD + pCounterDD - 1)
            End If
        Next i%
    End If
    'now set the selected card status accordingly
    'also cover the case for the last random card
    If p10 = "F" Or (p10 = "R" And pRandomSelection = 1) Then
        Deck(4, pStartDD) = "Selected"
        If frmStackView.SelectionsTextBox.Text = Empty Then
            frmStackView.SelectionsTextBox.Text = Deck(2, pStartDD)
        Else
            frmStackView.SelectionsTextBox.Text = frmStackView.SelectionsTextBox.Text _
                & " " & Deck(2, pStartDD)
        End If
        'reverse the selected card if specified
        If p6 = "S" Then
            Deck(6, pStartDD + pCounterDD - 1) = _
                Not Deck(6, pStartDD + pCounterDD - 1)
        End If
    'reverse the deal cards if specified
    ElseIf Val(p6) = 1 Or (p8 = 1 And Rnd < 0.5) Then
        Deck(6, pStartDD + pCounterDD - 1) = _
            Not Deck(6, pStartDD + pCounterDD - 1)
    End If
    ' I believe that the original code below should not have included
    'the "Or p6 = "S"" segment for the selected card, since that has already been
    'taken care of elsewhere.
    'ElseIf Val(p6) = 1 Or p6 = "S" Or (p8 = 1 And Rnd < 0.5) Then
    '    Deck(6, pStartDD + pCounterDD - 1) = _
    '        Not Deck(6, pStartDD + pCounterDD - 1)
    'End If
ElseIf pInverse = 1 Then
    'the duck counter does not need mod
    If Val(p3) > 0 Then
        pDuckCounter = Val(p3)
    'Else
    '    pDuckCounter = Int(Rnd * Val(p5)) + 1
    End If
    'adjust the deal counter based on the number of cards
    'the deal counter needs mod function unless random
    If Val(p2) > 0 Then
        pDealCounter = (pCounterDD Mod Val(p2))
        If pDealCounter = 1 Then
            pDealCounter = Val(p2)
            'this condition means that there is only one card in hand
            'for the ducks.  So, we have to reverse the card if specified,
            'and only if there are an odd number of ducks
            If Val(p7) > 0 Then
                If (pDuckCounter Mod 2) = 1 Then
                    Deck(6, pStartDD) = Not Deck(6, pStartDD)
                End If
            End If
        ElseIf pDealCounter = 0 Then
            pDealCounter = Val(p2) - 1
            'in this case, there are no preliminary ducks to handle
            'unless p2=1
            If Val(p7) > 0 Then
                If Val(p2) = 1 Then
                    If (pDuckCounter Mod 2) = 1 Then
                        Deck(6, pStartDD) = Not Deck(6, pStartDD)
                    End If
                End If
            End If
        Else
            pDealCounter = pDealCounter - 1
            'in this case, there are no preliminary ducks to handle
        End If
    End If
    'set deal and duck parameters for Deal & Duck
    'it turns out that the last thing you can do is deal,
    'so this is where we must start in all cases
    pDeal = True
    pDuck = False
    'first unselect the card if appropriate
    If p10 = "F" Then
        If Deck(4, pStartDD) = "Selected" Then
            Dim pSelectionsLength
            pSelectionsLength = Len(frmStackView.SelectionsTextBox.Text)
            Dim pSelectedCardLength
            pSelectedCardLength = Len(Deck(2, pStartDD))
            'check if there is only one card in the selections list
            If pSelectionsLength = pSelectedCardLength Then
                frmStackView.SelectionsTextBox.Text = Empty
            Else
                frmStackView.SelectionsTextBox.Text = _
                    Left(frmStackView.SelectionsTextBox.Text, pSelectionsLength - pSelectedCardLength - 1)
            End If
            Deck(4, pStartDD) = Empty
        End If
    End If
    'if the selected/final card was reversed, reverse it back
    If p6 = "S" Or Val(p6) = 1 Then
        Deck(6, pStartDD) = Not Deck(6, pStartDD)
    End If
    'do while upcounter<counterDD
    While pUpCounter < pCounterDD
        'first, shift the rest of the cards down one position
        For p% = pStartDD To pStartDD + pUpCounter - 1
            For m% = 1 To DeckProperties
                ChangedDeck(m%, p% + 1) = Deck(m%, p%)
            Next m%
        Next p%
        'next, move the bottom card to the top of the pile
        For m% = 1 To DeckProperties
            ChangedDeck(m%, pStartDD) = Deck(m%, pStartDD + pUpCounter)
        Next m%
        For p% = pStartDD To pStartDD + pUpCounter
            For m% = 1 To DeckProperties
                Deck(m%, p%) = ChangedDeck(m%, p%)
            Next m%
        Next p%
        'reverse card back if necessary
        If pDeal Then
            If p6 <> "S" And Val(p6) > 0 Then
                Deck(6, pStartDD) = Not Deck(6, pStartDD)
            End If
        ElseIf pDuck Then
            If Val(p7) > 0 Then
                Deck(6, pStartDD) = Not Deck(6, pStartDD)
            End If
        End If
        'adjust counter based on current action
        If pDeal Then
            pDealCounter = pDealCounter - 1
            If pDealCounter > 0 Then
                pUpCounter = pUpCounter + 1
            Else
                'switch logical settings for deal and duck
                pDuck = Not pDuck
                pDeal = Not pDeal
                'reset the deal counter
                If Val(p2) > 0 Then
                    pDealCounter = Val(p2)
                Else
                    pDealCounter = Int(Rnd * Val(p4)) + 1
                End If
            End If
        ElseIf pDuck Then
            pDuckCounter = pDuckCounter - 1
            If pDuckCounter = 0 Then
                pUpCounter = pUpCounter + 1
                'switch logical settings for deal and duck
                pDuck = Not pDuck
                pDeal = Not pDeal
                'reset the duck counter
                If Val(p3) > 0 Then
                    pDuckCounter = Val(p3)
                Else
                    pDuckCounter = Int(Rnd * Val(p5)) + 1
                End If
            End If
        End If
        If pUpCounter = pCounterDD - 1 And pDuck Then
            If pStartingOrder = "D" Then
                pUpCounter = pUpCounter + 1
                'this means that the original 'deal' was the starting point
                'and we are now finished with the inverse operation.
                'otherwise, do one more loop with a starting duck
            End If
        End If
    Wend
End If
PilesMatrixRefresh
CreatePiles (NumPiles)
End Sub

Private Sub AustralianDealButton_Click()
    'make sure piles have been created
    If PilesShown = 0 Then
        MsgBox ("You must create piles before you can manipulate piles.")
        Exit Sub
    End If
    ' make sure the PileMatrix has a valid selection
    Dim PileMatrixCheckSum As Integer
    PileMatrixCheckSum = 0
    For i% = 1 To 8
        For j% = 1 To 8
            If PileMatrix(10 * i% + j%).Value = 1 Then
                PileMatrixCheckSum = PileMatrixCheckSum + 1
            End If
        Next j%
    Next i%
    If PileMatrixCheckSum = 0 Then
        MsgBox ("You must check a box in the Pile Control Matrix" & Chr(13) & _
            "before you can manipulate piles.")
        Exit Sub
    End If
    Dim pTempCounter As Integer
    'identify the piles to manipulate
    Call PileMatrixQuery
    'set parameter variables
    Dim p1 As Variant
    Dim p2 As Variant
    Dim p3 As Variant
    Dim p4 As Variant
    Dim p5 As Variant
    Dim p6 As Variant
    Dim p7 As Variant
    Dim p8 As Variant
    Dim p9 As Variant
    Dim p10 As Variant
    Dim p11 As Variant
    If AustralianNumberSpecified.Value = True Then
        If (NumberDownExactText.Text <> Empty And _
            (Not IsNumeric(NumberDownExactText.Text) Or _
            Val(NumberDownExactText.Text) < 1 Or _
            Val(NumberDownExactText.Text) > 52)) Then
                NumberDownExactText.Text = Empty
                NumberDownExactText.SetFocus
                MsgBox "Please enter a valid number of cards (1 to 52)" & Chr(13) _
                    & "in the 'Down: Exact' Input Box"
                Exit Sub
        End If
        If (NumberUnderExactText.Text <> Empty And _
            (Not IsNumeric(NumberUnderExactText.Text) Or _
            Val(NumberUnderExactText.Text) < 1 Or _
            Val(NumberUnderExactText.Text) > 52)) Then
                NumberUnderExactText.Text = Empty
                NumberUnderExactText.SetFocus
                MsgBox "Please enter a valid number of cards (1 to 52)" & Chr(13) _
                    & "in the 'Under: Exact' Input Box"
                Exit Sub
        End If
        If (NumberDownRandomText.Text <> Empty And _
            (Not IsNumeric(NumberDownRandomText.Text) Or _
            Val(NumberDownRandomText.Text) < 1 Or _
            Val(NumberDownRandomText.Text) > 52)) Then
                NumberDownRandomText.Text = Empty
                NumberDownRandomText.SetFocus
                MsgBox "Please enter a valid number of cards (1 to 52)" & Chr(13) _
                    & "in the 'Down: Random' Input Box"
                Exit Sub
        End If
        If (NumberUnderRandomText.Text <> Empty And _
            (Not IsNumeric(NumberUnderRandomText.Text) Or _
            Val(NumberUnderRandomText.Text) < 1 Or _
            Val(NumberUnderRandomText.Text) > 52)) Then
                NumberUnderRandomText.Text = Empty
                NumberUnderRandomText.SetFocus
                MsgBox "Please enter a valid number of cards (1 to 52)" & Chr(13) _
                    & "in the 'Under: Random' Input Box"
                Exit Sub
        End If
        If NumberDownExactText.Text <> Empty And _
            NumberDownRandomText.Text <> Empty Then
                NumberDownExactText.Text = Empty
                NumberDownExactText.SetFocus
                NumberDownRandomText.Text = Empty
                MsgBox "You can not have numbers in both 'Exact' and 'Random'" & Chr(13) _
                    & "in the 'Down:' Input Boxes"
                Exit Sub
        End If
        If NumberUnderExactText.Text <> Empty And _
            NumberUnderRandomText.Text <> Empty Then
                NumberUnderExactText.Text = Empty
                NumberUnderExactText.SetFocus
                NumberUnderRandomText.Text = Empty
                MsgBox "You can not have numbers in both 'Exact' and 'Random'" & Chr(13) _
                    & "in the 'Under:' Input Boxes"
                Exit Sub
        End If
    End If
    If AustralianDealInverseCheck.Value = 1 Then
        If NumberDownRandomText.Text <> Empty Or _
            NumberUnderRandomText.Text <> Empty Or _
            ReverseDownRandomCheck.Value = 1 Or _
            ReverseUnderRandomCheck.Value = 1 Or _
            RandomCardSelectedCheck.Value = 1 Then
                AustralianDealInverseCheck.Value = 0
                MsgBox ("You may not select the 'Inverse' setting if you" & _
                        Chr(13) & "have any 'Random' parameters set.")
                Exit Sub
        End If
    End If
    'set first parameter
    If AustralianStartDownUnder.Value = True Then
        p1 = "D" & PileMatrixRow
    ElseIf AustralianStartUnderDown.Value = True Then
        p1 = "U" & PileMatrixRow
    End If
    'set parameters 2 through 5
    If AustralianNumberStandard.Value = True Then
        p2 = 1
        p3 = 1
        p4 = 0
        p5 = 0
    ElseIf AustralianNumberSpecified.Value = True Then
        If NumberDownExactText.Text <> Empty Then
            p2 = NumberDownExactText.Text
        Else
            p2 = 0
        End If
        If NumberUnderExactText.Text <> Empty Then
            p3 = NumberUnderExactText.Text
        Else
            p3 = 0
        End If
        If NumberDownRandomText.Text <> Empty Then
            p4 = NumberDownRandomText.Text
        Else
            p4 = 0
        End If
        If NumberUnderRandomText.Text <> Empty Then
            p5 = NumberUnderRandomText.Text
        Else
            p5 = 0
        End If
        'correct if both entries are empty
        If NumberDownExactText.Text = Empty And _
            NumberDownRandomText.Text = Empty Then
            p2 = 1
        End If
        If NumberUnderExactText.Text = Empty And _
            NumberUnderRandomText.Text = Empty Then
            p3 = 1
        End If
    End If
    'set parameters 6 through 9
    If AustralianNoReverse.Value = True Then
        p6 = 0
        p7 = 0
        p8 = 0
        p9 = 0
    ElseIf AustralianReverseSelected.Value = True Then
        'special case when reverse selected
        p6 = "S"
        p7 = 0
        p8 = 0
        p9 = 0
    ElseIf AustralianReverse.Value = True Then
        p6 = ReverseDownAllCheck.Value
        p7 = ReverseUnderAllCheck.Value
        p8 = ReverseDownRandomCheck.Value
        p9 = ReverseUnderRandomCheck.Value
    End If
    'set parameter 10
    If FinalCardSelectedCheck.Value = 1 Then
        p10 = "F"
    ElseIf RandomCardSelectedCheck.Value = 1 Then
        p10 = "R"
    Else
        p10 = 0
    End If
    'check for consistency between Reverse Selected card (P6)
    '  and Selected Card (P10)
    If p6 = "S" And p10 = 0 Then
        MsgBox ("You have indicated that the selected card is to be reversed," & Chr(13) & _
            "but you have not indicated which card will be the selected one." & Chr(13) & _
            "Please indicate " & Chr(34) & Chr(34) & "Final" & Chr(34) & Chr(34) & " or " & _
            Chr(34) & Chr(34) & "Random" & Chr(34) & Chr(34) & " by placing" & Chr(13) & _
            "a check in the appropriate box.")
        Exit Sub
    End If
    'set parameter 11
    p11 = AustralianDealInverseCheck.Value
    'identify ignoring of Reverse checks in the Pile Matrix if they are present
    pTempCounter = 0
    For i% = 1 To 8
        If ReverseR(i%).Value = 1 Then
            pTempCounter = pTempCounter + 1
        End If
        If ReverseC(i%).Value = 1 Then
            pTempCounter = pTempCounter + 1
        End If
    Next i%
    If pTempCounter > 0 Then
        MsgBox ("When Australian Deal command is performed, 'Reverse'" & _
        Chr(13) & "checkboxes in the Pile Matrix are ignored.")
    End If
    Call AustralianDeal(p1, p2, p3, p4, p5, p6, p7, p8, p9, p10, p11)
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "AustralianDeal(" & p1 _
        & ", " & p2 _
        & ", " & p3 _
        & ", " & p4 _
        & ", " & p5 _
        & ", " & p6 _
        & ", " & p7 _
        & ", " & p8 _
        & ", " & p9 _
        & ", " & p10 _
        & ", " & p11 & ")"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End Sub

Private Sub AustralianNoReverse_Click()
ReverseDownLabel.Enabled = False
ReverseUnderLabel.Enabled = False
ReverseAllLabel.Enabled = False
ReverseRandomLabel.Enabled = False
ReverseDownAllCheck.Enabled = False
ReverseDownRandomCheck.Enabled = False
ReverseDownRandomCheck.Value = 0
ReverseUnderAllCheck.Enabled = False
ReverseUnderRandomCheck.Enabled = False
ReverseUnderRandomCheck.Value = 0
End Sub

Private Sub AustralianNumberSpecified_Click()
NumberDownLabel.Enabled = True
NumberUnderLabel.Enabled = True
NumberExactLabel.Enabled = True
NumberRandomLabel.Enabled = True
NumberDownExactText.Enabled = True
NumberDownRandomText.Enabled = True
NumberUnderExactText.Enabled = True
NumberUnderRandomText.Enabled = True
End Sub

Private Sub AustralianNumberStandard_Click()
NumberDownLabel.Enabled = False
NumberUnderLabel.Enabled = False
NumberExactLabel.Enabled = False
NumberRandomLabel.Enabled = False
NumberDownExactText.Enabled = False
NumberDownRandomText.Enabled = False
NumberDownRandomText.Text = Empty
NumberUnderExactText.Enabled = False
NumberUnderRandomText.Enabled = False
NumberUnderRandomText.Text = Empty
End Sub

Private Sub AustralianReverse_Click()
ReverseDownLabel.Enabled = True
ReverseUnderLabel.Enabled = True
ReverseAllLabel.Enabled = True
ReverseRandomLabel.Enabled = True
ReverseDownAllCheck.Enabled = True
ReverseDownRandomCheck.Enabled = True
ReverseUnderAllCheck.Enabled = True
ReverseUnderRandomCheck.Enabled = True
End Sub

Private Sub AustralianReverseSelected_Click()
ReverseDownLabel.Enabled = False
ReverseUnderLabel.Enabled = False
ReverseAllLabel.Enabled = False
ReverseRandomLabel.Enabled = False
ReverseDownAllCheck.Enabled = False
ReverseDownRandomCheck.Enabled = False
ReverseDownRandomCheck.Value = 0
ReverseUnderAllCheck.Enabled = False
ReverseUnderRandomCheck.Enabled = False
ReverseUnderRandomCheck.Value = 0
End Sub

Private Sub CombinePilesButton_Click()
    'make sure piles have been created
    If PilesShown = 0 Then
        MsgBox ("You must create piles before you can manipulate piles.")
        Exit Sub
    End If
    ' make sure the PileMatrix has a valid selection
    Dim PileMatrixCheckSum As Integer
    PileMatrixCheckSum = 0
    For i% = 1 To 8
        For j% = 1 To 8
            If PileMatrix(10 * i% + j%).Value = 1 Then
                PileMatrixCheckSum = PileMatrixCheckSum + 1
            End If
        Next j%
    Next i%
    If PileMatrixCheckSum = 0 Then
        MsgBox ("You must check a box in the Pile Control Matrix" & Chr(13) & _
            "before you can manipulate piles.")
        Exit Sub
    End If
    'identify the piles to manipulate
    Dim param1 As String
    Dim param2 As String
    Call PileMatrixQuery
    If PileMatrixRow = PileMatrixColumn Then
        MsgBox ("You can not perform this operation on a single pile.")
        Exit Sub
    End If
    'set parameters to empty starting values
    param1 = ""
    param2 = ""
    'set parameters
    If CombineSecondaryTop.Value = True Then
        param2 = param2 & "T"
    ElseIf CombinePrimaryTop.Value = True Then
        param1 = param1 & "T"
    End If
    'set parameter pile numbers
    param1 = param1 & PileMatrixRow
    param2 = param2 & PileMatrixColumn
    'set parameter reverse condition
    If ReverseR(PileMatrixRow).Value = 1 Then
        param1 = param1 & "R"
    End If
    If ReverseC(PileMatrixColumn).Value = 1 Then
        param2 = param2 & "R"
    End If
    'now execute the Combine code
    Call CombinePiles(param1, param2)
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "CombinePiles(" & param1 _
        & ", " & param2 & ")"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End Sub

Public Sub CombinePiles(pileNum1Param, pileNum2Param)
    Dim tempVar As Integer
    'used as temporary variable in the new section
    'to adjust for the bottom pile being the destination pile.
    Dim maxPile As Integer
    Dim pileGap As Integer
    Dim pileAdjust As Integer
    Dim blockReverseSize As Integer
    Dim combinedPortion As Integer
    Dim firstPileSize As Integer
    Dim secondPileSize As Integer
    Dim pileNum1 As Integer
    Dim pileNum2 As Integer
    Dim pileLeft As String
    Dim pileRight As String
    'pileLeft is set to T for Top
    '   if the pileNum1 goes on top, else Empty
    'pileRight is set to T for Top
    '   if the pileNum2 goes on top, else Empty
    Dim pileLeftReverse As String
    Dim pileRightReverse As String
    'pileLeftReverse is set to R if the pile needs to
    'be reversed before the rest of the operation
    'the same is true for pileRightReverse
    '
    'decode the first paramater values
    If Len(pileNum1Param) = 1 Then
        pileLeft = Empty
        pileLeftReverse = Empty
        pileNum1 = pileNum1Param
    ElseIf Len(pileNum1Param) = 2 Then
        If IsNumeric(Left(pileNum1Param, 1)) Then
            pileNum1 = Val(Left(pileNum1Param, 1))
            pileLeftReverse = Right(pileNum1Param, 1)
            pileLeft = Empty
        Else
            pileNum1 = Val(Right(pileNum1Param, 1))
            pileLeftReverse = Empty
            pileLeft = Left(pileNum1Param, 1)
        End If
    ElseIf Len(pileNum1Param) = 3 Then
        pileLeft = Left(pileNum1Param, 1)
        pileLeftReverse = Right(pileNum1Param, 1)
        pileNum1 = Val(Mid(pileNum1Param, 2, 1))
    End If
    'decode the second paramater values
    If Len(pileNum2Param) = 1 Then
        pileRight = Empty
        pileRightReverse = Empty
        pileNum2 = pileNum2Param
    ElseIf Len(pileNum2Param) = 2 Then
        If IsNumeric(Left(pileNum2Param, 1)) Then
            pileNum2 = Val(Left(pileNum2Param, 1))
            pileRightReverse = Right(pileNum2Param, 1)
            pileRight = Empty
        Else
            pileNum2 = Val(Right(pileNum2Param, 1))
            pileRightReverse = Empty
            pileRight = Left(pileNum2Param, 1)
        End If
    ElseIf Len(pileNum2Param) = 3 Then
        pileRight = Left(pileNum2Param, 1)
        pileRightReverse = Right(pileNum2Param, 1)
        pileNum2 = Val(Mid(pileNum2Param, 2, 1))
    End If
    'set pile order info
    If pileNum1 > pileNum2 Then
        maxPile = pileNum1
    Else
        maxPile = pileNum2
    End If
    'reverse the appropriate piles if necessary
    If pileLeftReverse = "R" Then
        blockReverseSize = PileTable(pileNum1, 2) - PileTable(pileNum1, 1) + 1
        For m% = 1 To blockReverseSize
            For z% = 1 To DeckProperties
                ChangedDeck(z%, PileTable(pileNum1, 2) - m% + 1) = _
                    Deck(z%, PileTable(pileNum1, 1) + m% - 1)
            Next z%
        Next m%
        For m% = 1 To blockReverseSize
            For z% = 1 To DeckProperties
                Deck(z%, PileTable(pileNum1, 1) + m% - 1) = _
                    ChangedDeck(z%, PileTable(pileNum1, 1) + m% - 1)
            Next z%
        Next m%
        For i% = 1 To blockReverseSize
            Deck(6, PileTable(pileNum1, 1) + i% - 1) = _
                Not Deck(6, PileTable(pileNum1, 1) + i% - 1)
        Next i%
    End If
    If pileRightReverse = "R" Then
        blockReverseSize = PileTable(pileNum2, 2) - PileTable(pileNum2, 1) + 1
        For m% = 1 To blockReverseSize
            For z% = 1 To DeckProperties
                ChangedDeck(z%, PileTable(pileNum2, 2) - m% + 1) = _
                    Deck(z%, PileTable(pileNum2, 1) + m% - 1)
            Next z%
        Next m%
        For m% = 1 To blockReverseSize
            For z% = 1 To DeckProperties
                Deck(z%, PileTable(pileNum2, 1) + m% - 1) = _
                    ChangedDeck(z%, PileTable(pileNum2, 1) + m% - 1)
            Next z%
        Next m%
        For i% = 1 To blockReverseSize
            Deck(6, PileTable(pileNum2, 1) + i% - 1) = _
                Not Deck(6, PileTable(pileNum2, 1) + i% - 1)
        Next i%
    End If
    'transfer the deck to the PileDeck array
    For m% = 1 To DeckCount
        For z% = 1 To DeckProperties
            ChangedPileDeck(z%, m%) = Deck(z%, m%)
        Next z%
    Next m%
    'COMBINE CODE
    firstPileSize = PileTable(pileNum1, 2) - PileTable(pileNum1, 1) + 1
    secondPileSize = PileTable(pileNum2, 2) - PileTable(pileNum2, 1) + 1
    combinedPortion = firstPileSize + secondPileSize
    If pileLeft = "T" Then
        'here we need to handle for Primary on top
        For p% = 1 To firstPileSize
            For z% = 1 To DeckProperties
                ChangedDeck(z%, p%) = Deck(z%, PileTable(pileNum1, 1) - 1 + p%)
            Next z%
        Next p%
        For p% = 1 To secondPileSize
            For z% = 1 To DeckProperties
                ChangedDeck(z%, p% + firstPileSize) = Deck(z%, PileTable(pileNum2, 1) - 1 + p%)
            Next z%
        Next p%
    Else
        'here we need to handle for Secondary on top
        For p% = 1 To secondPileSize
            For z% = 1 To DeckProperties
                ChangedDeck(z%, p%) = Deck(z%, PileTable(pileNum2, 1) - 1 + p%)
            Next z%
        Next p%
        For p% = 1 To firstPileSize
            For z% = 1 To DeckProperties
                ChangedDeck(z%, p% + secondPileSize) = Deck(z%, PileTable(pileNum1, 1) - 1 + p%)
            Next z%
        Next p%
    End If
    'END OF COMBINE CODE
    
    'arrange the piles back into a deck
    
    'If pileNum1 = pileNum2 Then
    '    For m% = 1 To RifflePortion + ProtectedBlock
    '        For z% = 1 To DeckProperties
    '            ChangedPileDeck(z%, PileTable(pileNum1, 1) - 1 + m%) = ChangedDeck(z%, m%)
    '        Next z%
    '    Next m%
    'Else
    
    'New code to adjust for having the bottom pile define the placement
    'of the combined pile.  The older code that did not include this portion
    'used the Primary pile as the destination pile.  Logically, the Combine event
    'should mimic someone picking up the Primary pile and putting it on
    'top of the Secondary pile.
    'If pileLeft = "T" Then
    '    tempVar = pileNum2
    '    pileNum2 = pileNum1
    '    pileNum1 = tempVar
    'End If
    'this is the end of the New Code for bottom pile as destination.
    
    'I ended up deciding that the Primary pile should always define
    'the destination.
        
        pileGap = 0
        pileAdjust = 0
        For i% = 1 To NumPiles
            If i% > maxPile Then
                pileGap = 0
            End If
            If i% = pileNum2 Then
                pileAdjust = -1
                If pileNum2 < pileNum1 Then
                    pileGap = PileTable(pileNum2, 2) - PileTable(pileNum2, 1) + 1
                End If
            ElseIf i% = pileNum1 Then
                For m% = 1 To combinedPortion
                    For z% = 1 To DeckProperties
                        ChangedPileDeck(z%, PileTable(pileNum1, 1) - 1 + m% - pileGap) = _
                            ChangedDeck(z%, m%)
                    Next z%
                Next m%
                ChangedPileTable(i% + pileAdjust, 1) = PileTable(i%, 1) - pileGap
                ChangedPileTable(i% + pileAdjust, 2) = ChangedPileTable(i% + pileAdjust, 1) + _
                    combinedPortion - 1
                'adjust pileGap
                If pileNum2 < pileNum1 Then
                    pileGap = 0
                Else
                    pileGap = -(PileTable(pileNum2, 2) - PileTable(pileNum2, 1) + 1)
                End If
            Else
                For m% = 1 To PileTable(i%, 2) - PileTable(i%, 1) + 1
                    For z% = 1 To DeckProperties
                        ChangedPileDeck(z%, PileTable(i%, 1) - 1 + m% - pileGap) = _
                            Deck(z%, PileTable(i%, 1) - 1 + m%)
                    Next z%
                Next m%
                ChangedPileTable(i% + pileAdjust, 1) = PileTable(i%, 1) - pileGap
                ChangedPileTable(i% + pileAdjust, 2) = PileTable(i%, 2) - pileGap
            End If
        Next i%
        'adjust for the fact that two piles have become one
        NumPiles = NumPiles - 1
        For i% = 1 To NumPiles
            PileTable(i%, 1) = ChangedPileTable(i%, 1)
            PileTable(i%, 2) = ChangedPileTable(i%, 2)
        Next i%
    'End If
    For m% = 1 To DeckCount
        For z% = 1 To DeckProperties
            Deck(z%, m%) = ChangedPileDeck(z%, m%)
        Next z%
    Next m%
    CreatePiles (NumPiles)
    PilesMatrixRefresh
End Sub

Public Sub CreatePilesDealAlternatingRandom(paramPiles, paramCards)
Dim pCounter As Integer
Dim pSelector As Integer
Dim pStartOrder(8) As Integer
Dim pDealOrder(8) As Integer
Dim CardsPerPile As Integer
Dim PilesWithExtra As Integer
Dim pPileOrder(52) As Integer
Dim pEndOfPile As Integer
Dim pSum As Integer
CardsPerPile = Int(paramCards / paramPiles)
PilesWithExtra = paramCards Mod paramPiles
'set the RandomPiles(x,y) array to all zeros
For i% = 1 To 8
    For j% = 1 To 52
        RandomPiles(i%, j%) = 0
    Next j%
Next i%
'distribute cards
pCounter = 1
For i% = 1 To CardsPerPile
    'set pStartOrder and pDealOrder to initial values
    For z% = 1 To 8
        pStartOrder(z%) = z%
        pDealOrder(z%) = 0
    Next z%
    'establish random deal sequence
    For p% = 1 To paramPiles
        pSelector = Int(Rnd * (paramPiles + 1 - p%)) + 1
        pDealOrder(p%) = pStartOrder(pSelector)
        If pSelector <> paramPiles + 1 - p% Then
            pStartOrder(pSelector) = pStartOrder(paramPiles + 1 - p%)
        End If
    Next p%
    'place each deal round's cards in appropriate pile
    For m% = 1 To paramPiles
        RandomPiles(pDealOrder(m%), i%) = pCounter
        pCounter = pCounter + 1
    Next m%
Next i%
'handle extra cards beyond even multiples
If PilesWithExtra > 0 Then
    'set pStartOrder and pDealOrder to initial values
    For z% = 1 To 8
        pStartOrder(z%) = z%
        pDealOrder(z%) = 0
    Next z%
    'establish random deal sequence
    For p% = 1 To paramPiles
        pSelector = Int(Rnd * (paramPiles + 1 - p%)) + 1
        pDealOrder(p%) = pStartOrder(pSelector)
        If pSelector <> paramPiles + 1 - p% Then
            pStartOrder(pSelector) = pStartOrder(paramPiles + 1 - p%)
        End If
    Next p%
    'place each deal round's cards in appropriate pile
    For m% = 1 To PilesWithExtra
        RandomPiles(pDealOrder(m%), CardsPerPile + 1) = pCounter
        pCounter = pCounter + 1
    Next m%
End If
'set pPileOrder values
pCounter = 1
pSum = 0
For i% = 1 To paramPiles
    pEndOfPile = 0
    For j% = 1 To 52
        If RandomPiles(i%, 53 - j%) <> 0 Then
            If pEndOfPile = 0 Then
                PileTable(i%, 2) = 53 - j% + pSum
                pSum = pSum + 53 - j%
                pEndOfPile = 1
            End If
            If j% = 52 Then
                If i% = 1 Then
                    PileTable(1, 1) = 1
                Else
                    PileTable(i%, 1) = PileTable(i% - 1, 2) + 1
                End If
            End If
            pPileOrder(pCounter) = RandomPiles(i%, 53 - j%)
            pCounter = pCounter + 1
        End If
    Next j%
Next i%
If pCounter <= 52 Then
    PileTable(paramPiles + 1, 1) = pCounter
    PileTable(paramPiles + 1, 2) = 52
    For i% = pCounter To 52
        pPileOrder(i%) = i%
    Next i%
End If
For m% = 1 To DeckProperties
    For z% = 1 To DeckCount
        ChangedDeck(m%, z%) = Deck(m%, pPileOrder(z%))
    Next z%
Next m%
For m% = 1 To DeckProperties
    For z% = 1 To DeckCount
        Deck(m%, z%) = ChangedDeck(m%, z%)
    Next z%
Next m%
If pCounter <= 52 Then
    CreatePiles (paramPiles + 1)
Else
    CreatePiles (paramPiles)
End If
End Sub

Public Sub CreatePilesDealAlternatingRegular(paramPiles, paramCards)
Dim dCounter As Integer
Dim CardsPerPile As Integer
Dim PilesWithExtra As Integer
CardsPerPile = Int(paramCards / paramPiles)
PilesWithExtra = paramCards Mod paramPiles

'establish correct deck order
dCounter = 1
For i% = 1 To paramPiles
    If i% <= PilesWithExtra Then
        For j% = 1 To CardsPerPile + 1
            For s% = 1 To DeckProperties
                ChangedDeck(s%, dCounter) = _
                    Deck(s%, i% + (CardsPerPile + 1 - j%) * paramPiles)
            Next s%
            dCounter = dCounter + 1
        Next j%
    ElseIf i% > PilesWithExtra Then
        For j% = 1 To CardsPerPile
            For s% = 1 To DeckProperties
                ChangedDeck(s%, dCounter) = _
                    Deck(s%, i% + (CardsPerPile - j%) * paramPiles)
            Next s%
            dCounter = dCounter + 1
        Next j%
    End If
Next i%
'set balance of deck
If dCounter < 53 Then
    For k% = dCounter To 52
        For s% = 1 To DeckProperties
            ChangedDeck(s%, k%) = Deck(s%, k%)
        Next s%
    Next k%
End If
'reset original deck with new order
For m% = 1 To DeckProperties
    For z% = 1 To DeckCount
        Deck(m%, z%) = ChangedDeck(m%, z%)
    Next z%
Next m%

'establish PileTable
dCounter = 1
For i% = 1 To paramPiles
    If i% <= PilesWithExtra Then
        PileTable(i%, 1) = dCounter
        PileTable(i%, 2) = dCounter + CardsPerPile
        dCounter = dCounter + CardsPerPile + 1
    ElseIf i% > PilesWithExtra Then
        PileTable(i%, 1) = dCounter
        PileTable(i%, 2) = dCounter + CardsPerPile - 1
        dCounter = dCounter + CardsPerPile
    End If
Next i%
If dCounter <= 52 Then
    PileTable(paramPiles + 1, 1) = dCounter
    PileTable(paramPiles + 1, 2) = 52
    CreatePiles (paramPiles + 1)
Else
    CreatePiles (paramPiles)
End If
End Sub

Public Sub CreatePiles(paramPiles)
'code here to figure out where the piles go
NumPiles = Val(paramPiles)
Row1 = 0
Row2 = 0
Row3 = 0
MaxWidth = 11000
MaxHeight = 6095 'from original deck window
For i% = 1 To NumPiles
    If 250 + (PileTable(i%, 2) - PileTable(i%, 1)) * 250 + 1350 > MaxWidth Then
        MaxWidth = 250 + (PileTable(i%, 2) - PileTable(i%, 1)) * 250 + 1350 + 400
    End If
Next i%
'establish first row parameters
RowWidth = 0
For i% = 1 To NumPiles
    PileLocations(i%, 1) = 400
    PileLocations(i%, 2) = RowWidth + 250
    RowWidth = RowWidth + (PileTable(i%, 2) - PileTable(i%, 1)) * 250 + 1350 + 400
    If RowWidth <= MaxWidth Then
        Row1 = i%
    End If
Next i%
'establish second row parameters
If Row1 + 1 <= NumPiles Then
    RowWidth = 0
    For j% = Row1 + 1 To NumPiles
        PileLocations(j%, 1) = 400 + 2800
        PileLocations(j%, 2) = RowWidth + 250
        RowWidth = RowWidth + (PileTable(j%, 2) - PileTable(j%, 1)) * 250 + 1350 + 400
        If RowWidth <= MaxWidth Then
            Row2 = j%
        End If
    Next j%
    'establish third row parameters
    If Row2 + 1 <= NumPiles Then
        MaxHeight = 8900
        RowWidth = 0
        For k% = Row2 + 1 To NumPiles
            PileLocations(k%, 1) = 400 + 2800 + 2800
            PileLocations(k%, 2) = RowWidth + 250
            RowWidth = RowWidth + (PileTable(k%, 2) - PileTable(k%, 1)) * 250 + 1350 + 400
            If RowWidth <= MaxWidth Then
                Row3 = k%
            End If
        Next k%
    End If
End If
ShowPiles
End Sub

Public Sub CreatePilesGilbreath(paramPiles)
'code here to figure out where the piles go
NumPiles = Val(paramPiles)
Row1 = 0
Row2 = 0
Row3 = 0
GilbreathOffset = 0
MaxWidth = 11000
MaxHeight = 6095 'from original deck window
For i% = 1 To NumPiles
    If 250 + (PileTable(i%, 2) - PileTable(i%, 1)) * 250 + 1350 > MaxWidth Then
        MaxWidth = 250 + (PileTable(i%, 2) - PileTable(i%, 1)) * 250 + 1350 + 400
    End If
Next i%
'establish first row parameters
RowWidth = 0
For i% = 1 To NumPiles
    PileLocations(i%, 1) = 400
    PileLocations(i%, 2) = RowWidth + 250
    RowWidth = RowWidth + (PileTable(i%, 2) - PileTable(i%, 1)) * 250 + 1350 + 400
    If RowWidth <= MaxWidth Then
        Row1 = i%
    End If
Next i%
If GilbreathPileNum <= Row1 Then
    GilbreathOffset = 400
    For i% = 1 To Row1
        PileLocations(i%, 1) = PileLocations(i%, 1) + GilbreathOffset
    Next i%
End If
'establish second row parameters
If Row1 + 1 <= NumPiles Then
    RowWidth = 0
    For j% = Row1 + 1 To NumPiles
        PileLocations(j%, 1) = 400 + 2800 + GilbreathOffset
        PileLocations(j%, 2) = RowWidth + 250
        RowWidth = RowWidth + (PileTable(j%, 2) - PileTable(j%, 1)) * 250 + 1350 + 400
        If RowWidth <= MaxWidth Then
            Row2 = j%
        End If
    Next j%
    If GilbreathPileNum > Row1 And GilbreathPileNum <= Row2 Then
        GilbreathOffset = 400
        For i% = Row1 + 1 To Row2
            PileLocations(i%, 1) = PileLocations(i%, 1) + GilbreathOffset
        Next i%
    End If
    'establish third row parameters
    If Row2 + 1 <= NumPiles Then
        MaxHeight = 8900
        RowWidth = 0
        For k% = Row2 + 1 To NumPiles
            PileLocations(k%, 1) = 400 + 2800 + 2800
            PileLocations(k%, 2) = RowWidth + 250
            RowWidth = RowWidth + (PileTable(k%, 2) - PileTable(k%, 1)) * 250 + 1350 + 400
            If RowWidth <= MaxWidth Then
                Row3 = k%
            End If
        Next k%
        If GilbreathPileNum > Row2 And GilbreathPileNum <= Row3 Then
            GilbreathOffset = 400
            For i% = Row2 + 1 To NumPiles
                PileLocations(i%, 1) = PileLocations(i%, 1) + GilbreathOffset
            Next i%
        End If
    End If
End If
MaxHeight = MaxHeight + GilbreathOffset
ShowPilesGilbreath
End Sub





Private Sub DuckAndDeal(pileNum)
Dim pStartDD As Integer
Dim pFinishDD As Integer
Dim pCounterDD As Integer
Dim pDuck As Boolean
Dim pDeal As Boolean
pStartDD = PileTable(pileNum, 1)
pFinishDD = PileTable(pileNum, 2)
pCounterDD = pFinishDD - pStartDD + 1
'set deal and duck parameters for Deal & Duck
pDeal = False
pDuck = True
'do while counter>1
'Since counter represents how many cards are in hand,
'we do not want to do the deal and duck sequence when there is one card
While pCounterDD > 1
    'first move the top card to the bottom of the pile
    For m% = 1 To DeckProperties
        ChangedDeck(m%, pStartDD + pCounterDD - 1) = Deck(m%, pStartDD)
    Next m%
    'next, shift the rest of the cards up one position
    For p% = pStartDD To pStartDD + pCounterDD - 2
        For m% = 1 To DeckProperties
            ChangedDeck(m%, p%) = Deck(m%, p% + 1)
        Next m%
    Next p%
    'transfer cards back to Deck()
    For p% = pStartDD To pStartDD + pCounterDD - 1
        For m% = 1 To DeckProperties
            Deck(m%, p%) = ChangedDeck(m%, p%)
        Next m%
    Next p%
    'adjust counter based on current action
    If pDeal Then
        pCounterDD = pCounterDD - 1
    End If
    'switch logical settings for dela and duck
    pDuck = Not pDuck
    pDeal = Not pDeal
Wend
PilesMatrixRefresh
CreatePiles (NumPiles)
End Sub



Public Sub CreatePilesButton_Click()
PileError = False
DealPilesCheck
'the above procedure call is essential in case the user left 8 piles
'with alternating less than the full deck
'it also does a logical check of the inputs
If PileError Then
    Exit Sub
End If
SetNumPilesPlan
NumPiles = NumPilesPlan
'this sets NumPiles to the value in the text box
PilesMatrixRefresh
Dim pPiles As Integer
Dim pCards As Integer
Dim pCode1 As String
Dim pCode2 As String
Dim pCode3 As String
Dim pCode4 As String
Dim pCode5 As String
Dim pCode6 As String
Dim pCode7 As String
Dim pCode8 As String
For i% = 1 To 8
    For j% = 1 To 2
        PileTable(i%, j%) = 0
    Next j%
Next i%
For i% = 1 To 8
    For j% = 1 To 3
        PileLocations(i%, j%) = 0
    Next j%
Next i%
If DealCardsOption.Value = True And _
    DealAlternatingOption.Value = True And _
    AlternatingRandomOption.Value = True Then
    pPiles = Val(NumberOfPilesText.Text)
    If NumberOfCardsToDealText.Text = Empty Then
        pCards = 52
    Else
        pCards = Val(NumberOfCardsToDealText.Text)
    End If
    Call CreatePilesDealAlternatingRandom(pPiles, pCards)
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "CreatePilesDealAlternatingRandom(" & pPiles _
        & ", " & pCards & ")"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
ElseIf DealCardsOption.Value = True And _
    DealAlternatingOption.Value = True And _
    AlternatingRegularOption.Value = True Then
    pPiles = Val(NumberOfPilesText.Text)
    If NumberOfCardsToDealText.Text = Empty Then
        pCards = 52
    Else
        pCards = Val(NumberOfCardsToDealText.Text)
    End If
    Call CreatePilesDealAlternatingRegular(pPiles, pCards)
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "CreatePilesDealAlternatingRegular(" & pPiles _
        & ", " & pCards & ")"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
ElseIf DealCardsOption.Value = True And _
    DealAlternatingOption.Value = True And _
    CompleteRandomOption.Value = True Then
    pPiles = Val(NumberOfPilesText.Text)
    If NumberOfCardsToDealText.Text = Empty Then
        pCards = 52
    Else
        pCards = Val(NumberOfCardsToDealText.Text)
    End If
    Call CreatePilesDealCompleteRandom(pPiles, pCards)
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "CreatePilesDealCompleteRandom(" & pPiles _
        & ", " & pCards & ")"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
ElseIf DealCardsOption.Value = True And _
    DealCompleteOption.Value = True Then
        pPiles = Val(NumberOfPilesText.Text)
        'pile 1 code
        If PileRandom1.Value = True And _
            SpecifiedText1.Text = Empty Then
            pCode1 = "R"
        ElseIf PileRandom1.Value = True And _
            SpecifiedText1.Text <> Empty Then
            pCode1 = "R" & SpecifiedText1.Text
        End If
        If PileSpecified1.Value = True Then
            pCode1 = "S" & SpecifiedText1.Text
        End If
        'pile 2 code
        If PileRandom2.Value = True And _
            SpecifiedText2.Text = Empty Then
            pCode2 = "R"
        ElseIf PileRandom2.Value = True And _
            SpecifiedText2.Text <> Empty Then
            pCode2 = "R" & SpecifiedText2.Text
        End If
        If PileSpecified2.Value = True Then
            pCode2 = "S" & SpecifiedText2.Text
        End If
        If Val(NumberOfPilesText.Text) < 2 Then
            pCode2 = "X"
        End If
        'pile 3 code
        If PileRandom3.Value = True And _
            SpecifiedText3.Text = Empty Then
            pCode3 = "R"
        ElseIf PileRandom3.Value = True And _
            SpecifiedText3.Text <> Empty Then
            pCode3 = "R" & SpecifiedText3.Text
        End If
        If PileSpecified3.Value = True Then
            pCode3 = "S" & SpecifiedText3.Text
        End If
        If Val(NumberOfPilesText.Text) < 3 Then
            pCode3 = "X"
        End If
        'pile 4 code
        If PileRandom4.Value = True And _
            SpecifiedText4.Text = Empty Then
            pCode4 = "R"
        ElseIf PileRandom4.Value = True And _
            SpecifiedText4.Text <> Empty Then
            pCode4 = "R" & SpecifiedText4.Text
        End If
        If PileSpecified4.Value = True Then
            pCode4 = "S" & SpecifiedText4.Text
        End If
        If Val(NumberOfPilesText.Text) < 4 Then
            pCode4 = "X"
        End If
        'pile 5 code
        If PileRandom5.Value = True And _
            SpecifiedText5.Text = Empty Then
            pCode5 = "R"
        ElseIf PileRandom5.Value = True And _
            SpecifiedText5.Text <> Empty Then
            pCode5 = "R" & SpecifiedText5.Text
        End If
        If PileSpecified5.Value = True Then
            pCode5 = "S" & SpecifiedText5.Text
        End If
        If Val(NumberOfPilesText.Text) < 5 Then
            pCode5 = "X"
        End If
        'pile 6 code
        If PileRandom6.Value = True And _
            SpecifiedText6.Text = Empty Then
            pCode6 = "R"
        ElseIf PileRandom6.Value = True And _
            SpecifiedText6.Text <> Empty Then
            pCode6 = "R" & SpecifiedText6.Text
        End If
        If PileSpecified6.Value = True Then
            pCode6 = "S" & SpecifiedText6.Text
        End If
        If Val(NumberOfPilesText.Text) < 6 Then
            pCode6 = "X"
        End If
        'pile 7 code
        If PileRandom7.Value = True And _
            SpecifiedText7.Text = Empty Then
            pCode7 = "R"
        ElseIf PileRandom7.Value = True And _
            SpecifiedText7.Text <> Empty Then
            pCode7 = "R" & SpecifiedText7.Text
        End If
        If PileSpecified7.Value = True Then
            pCode7 = "S" & SpecifiedText7.Text
        End If
        If Val(NumberOfPilesText.Text) < 7 Then
            pCode7 = "X"
        End If
        'pile 8 code
        If PileRandom8.Value = True And _
            SpecifiedText8.Text = Empty Then
            pCode8 = "R"
        ElseIf PileRandom8.Value = True And _
            SpecifiedText8.Text <> Empty Then
            pCode8 = "R" & SpecifiedText8.Text
        End If
        If PileSpecified8.Value = True Then
            pCode8 = "S" & SpecifiedText8.Text
        End If
        If Val(NumberOfPilesText.Text) < 8 Then
            pCode8 = "X"
        End If
    Call CreatePilesDealComplete(pPiles, pCode1, pCode2, pCode3, _
            pCode4, pCode5, pCode6, pCode7, pCode8)
    If PileParseError Then
        PileParseError = False
        'this resets the PileParseError code, and skips the SessionRecord segment
    Else
        'SessionRecord
        If SessionRecordMode Then
            SessionCommand = "CreatePilesDealComplete(" & pPiles _
            & ", " & pCode1 _
            & ", " & pCode2 _
            & ", " & pCode3 _
            & ", " & pCode4 _
            & ", " & pCode5 _
            & ", " & pCode6 _
            & ", " & pCode7 _
            & ", " & pCode8 & ")"
            frmStackView.SessionListBox.AddItem SessionCommand
            frmStackView.SessionStatusUpdate (0)
        End If
    End If
ElseIf CutCardsOption.Value = True Then
        pPiles = Val(NumberOfPilesText.Text)
        'pile 1 code
        If PileRandom1.Value = True And _
            SpecifiedText1.Text = Empty Then
            pCode1 = "R"
        ElseIf PileRandom1.Value = True And _
            SpecifiedText1.Text <> Empty Then
            pCode1 = "R" & SpecifiedText1.Text
        End If
        If PileSpecified1.Value = True Then
            pCode1 = "S" & SpecifiedText1.Text
        End If
        'pile 2 code
        If PileRandom2.Value = True And _
            SpecifiedText2.Text = Empty Then
            pCode2 = "R"
        ElseIf PileRandom2.Value = True And _
            SpecifiedText2.Text <> Empty Then
            pCode2 = "R" & SpecifiedText2.Text
        End If
        If PileSpecified2.Value = True Then
            pCode2 = "S" & SpecifiedText2.Text
        End If
        If Val(NumberOfPilesText.Text) < 2 Then
            pCode2 = "X"
        End If
        'pile 3 code
        If PileRandom3.Value = True And _
            SpecifiedText3.Text = Empty Then
            pCode3 = "R"
        ElseIf PileRandom3.Value = True And _
            SpecifiedText3.Text <> Empty Then
            pCode3 = "R" & SpecifiedText3.Text
        End If
        If PileSpecified3.Value = True Then
            pCode3 = "S" & SpecifiedText3.Text
        End If
        If Val(NumberOfPilesText.Text) < 3 Then
            pCode3 = "X"
        End If
        'pile 4 code
        If PileRandom4.Value = True And _
            SpecifiedText4.Text = Empty Then
            pCode4 = "R"
        ElseIf PileRandom4.Value = True And _
            SpecifiedText4.Text <> Empty Then
            pCode4 = "R" & SpecifiedText4.Text
        End If
        If PileSpecified4.Value = True Then
            pCode4 = "S" & SpecifiedText4.Text
        End If
        If Val(NumberOfPilesText.Text) < 4 Then
            pCode4 = "X"
        End If
        'pile 5 code
        If PileRandom5.Value = True And _
            SpecifiedText5.Text = Empty Then
            pCode5 = "R"
        ElseIf PileRandom5.Value = True And _
            SpecifiedText5.Text <> Empty Then
            pCode5 = "R" & SpecifiedText5.Text
        End If
        If PileSpecified5.Value = True Then
            pCode5 = "S" & SpecifiedText5.Text
        End If
        If Val(NumberOfPilesText.Text) < 5 Then
            pCode5 = "X"
        End If
        'pile 6 code
        If PileRandom6.Value = True And _
            SpecifiedText6.Text = Empty Then
            pCode6 = "R"
        ElseIf PileRandom6.Value = True And _
            SpecifiedText6.Text <> Empty Then
            pCode6 = "R" & SpecifiedText6.Text
        End If
        If PileSpecified6.Value = True Then
            pCode6 = "S" & SpecifiedText6.Text
        End If
        If Val(NumberOfPilesText.Text) < 6 Then
            pCode6 = "X"
        End If
        'pile 7 code
        If PileRandom7.Value = True And _
            SpecifiedText7.Text = Empty Then
            pCode7 = "R"
        ElseIf PileRandom7.Value = True And _
            SpecifiedText7.Text <> Empty Then
            pCode7 = "R" & SpecifiedText7.Text
        End If
        If PileSpecified7.Value = True Then
            pCode7 = "S" & SpecifiedText7.Text
        End If
        If Val(NumberOfPilesText.Text) < 7 Then
            pCode7 = "X"
        End If
        'pile 8 code
        If PileRandom8.Value = True And _
            SpecifiedText8.Text = Empty Then
            pCode8 = "R"
        ElseIf PileRandom8.Value = True And _
            SpecifiedText8.Text <> Empty Then
            pCode8 = "R" & SpecifiedText8.Text
        End If
        If PileSpecified8.Value = True Then
            pCode8 = "S" & SpecifiedText8.Text
        End If
        If Val(NumberOfPilesText.Text) < 8 Then
            pCode8 = "X"
        End If
    Call CreatePilesCut(pPiles, pCode1, pCode2, pCode3, _
            pCode4, pCode5, pCode6, pCode7, pCode8)
    If PileParseError Then
        PileParseError = False
        'this resets the PileParseError code, and skips the SessionRecord segment
    Else
        'SessionRecord
        If SessionRecordMode Then
            SessionCommand = "CreatePilesCut(" & pPiles _
            & ", " & pCode1 _
            & ", " & pCode2 _
            & ", " & pCode3 _
            & ", " & pCode4 _
            & ", " & pCode5 _
            & ", " & pCode6 _
            & ", " & pCode7 _
            & ", " & pCode8 & ")"
            frmStackView.SessionListBox.AddItem SessionCommand
            frmStackView.SessionStatusUpdate (0)
        End If
    End If
End If
End Sub

Public Sub CreatePilesCut(pP, pC1, pC2, pC3, pC4, pC5, pC6, pC7, pC8)
'establish key variables
Dim PureRandomSelector As Integer
'used to adjust for card pile discrepancies of pure random amounts
Dim ApproxRandomSelector As Integer
'used to adjust for card pile discrepancies of approx random amounts
Dim PureRandomPile As Integer
'the total number of pure random cards
Dim ApproxRandomPile As Integer
'the total number of approximate random cards
Dim ExactSpecifiedPile As Integer
'the total number of exact specified cards
Dim PureRandomCount As Integer
'the number of piles that are pure random
Dim ApproxRandomCount As Integer
'the number of piles that are approximate random
Dim ExactSpecifiedCount As Integer
'the number of piles that are exact specified
Dim StrPointer As Integer
Dim PileAdjustment As Integer
'the difference between 52 and PileTotalPreAdjust
Dim PileVariation As Double
'the amount of variation to a random pile size
'set at 15% when the approximate random is specified
'set at 25% of equal breakdown when pure random
Dim PileTotalPreAdjust As Integer
'total card count with random piles before adjustment
Dim CummPileCount
'used in final segment to keep track of total pile count
Dim AllowableApproxRandomPile
'the number of cards allowed to be approx random


'initialize variables
PureRandomPile = 0
PileParseError = False
'start assuming that there is no error in pile data
ApproxRandomPile = 0
ExactSpecifiedPile = 0
PureRandomCount = 0
ApproxRandomCount = 0
ExactSpecifiedCount = 0
PileTotalPreAdjust = 0
PileCode(1) = pC1
PileCode(2) = pC2
PileCode(3) = pC3
PileCode(4) = pC4
PileCode(5) = pC5
PileCode(6) = pC6
PileCode(7) = pC7
PileCode(8) = pC8
'set X as zero values
For i% = 1 To 8
    StrPointer = InStr(PileCode(i%), "X")
    If StrPointer = 1 Then
        PileInputData(1, i%) = "X"
        PileOutputData(1, i%) = "X"
        PileOutputData(2, i%) = 0
    End If
Next i%
'identify and establish exact cards
For i% = 1 To 8
    StrPointer = InStr(PileCode(i%), "S")
    If StrPointer = 1 Then
        PileInputData(1, i%) = "S"
        PileInputData(2, i%) = Val(Right(PileCode(i%), Len(PileCode(i%)) - 1))
        PileOutputData(1, i%) = "S"
        PileOutputData(2, i%) = Val(Right(PileCode(i%), Len(PileCode(i%)) - 1))
        ExactSpecifiedPile = ExactSpecifiedPile + PileOutputData(2, i%)
        ExactSpecifiedCount = ExactSpecifiedCount + 1
    End If
Next i%
'identify pure random cards
For i% = 1 To 8
    StrPointer = InStr(PileCode(i%), "R")
    If StrPointer = 1 And Len(PileCode(i%)) = 1 Then
        PileInputData(1, i%) = "RP"
        PileInputData(2, i%) = 0
        PileOutputData(1, i%) = "RP"
        PureRandomCount = PureRandomCount + 1
    End If
Next i%
'identify and establish approx random cards
For i% = 1 To 8
    StrPointer = InStr(PileCode(i%), "R")
    If StrPointer = 1 And Len(PileCode(i%)) > 1 Then
        PileInputData(1, i%) = "RA"
        PileInputData(2, i%) = Val(Right(PileCode(i%), Len(PileCode(i%)) - 1))
        PileOutputData(1, i%) = "RA"
        PileVariation = 0.15 * PileInputData(2, i%)
        PileOutputData(2, i%) = Int(Rnd * 2 * PileVariation + _
            PileInputData(2, i%) - PileVariation) + 1
        ApproxRandomPile = ApproxRandomPile + PileOutputData(2, i%)
        ApproxRandomCount = ApproxRandomCount + 1
    End If
Next i%
'establish and fill in Pure Random values
If PureRandomCount > 0 Then
    If ExactSpecifiedPile + ApproxRandomPile >= 52 - PureRandomCount Then
        For i% = 1 To 8
            If PileInputData(1, i%) = "RP" Then
                PileOutputData(2, i%) = 1
                PureRandomPile = PureRandomPile + 1
            End If
        Next i%
    Else
        PureRandomPile = 52 - (ExactSpecifiedPile + ApproxRandomPile)
        PileVariation = 0.25 * (PureRandomPile / PureRandomCount)
        For k% = 1 To 8
            If PileInputData(1, k%) = "RP" Then
                PileOutputData(2, k%) = Int(Rnd * 2 * PileVariation + _
                    (PureRandomPile / PureRandomCount) - PileVariation) + 1
                If PileOutputData(2, k%) < 1 Then
                    PileOutputData(2, k%) = 1
                'this value must be at least 1
                End If
            End If
        Next k%
    End If
End If

'adjust Approx random Pile information
If ApproxRandomCount > 0 Then
    AllowableApproxRandomPile = 52 - PureRandomPile - ExactSpecifiedPile
    If ApproxRandomPile > AllowableApproxRandomPile Or PureRandomCount = 0 Then
    'scale down the values when they are too high,
    'but also scale up the values if there are no pure random pile to make up the difference
        For i% = 1 To 8
            If PileInputData(1, i%) = "RA" Then
                PileOutputData(2, i%) = Int(AllowableApproxRandomPile * _
                    (PileInputData(2, i%) / ApproxRandomPile))
                    If PileOutputData(2, i%) < 1 Then
                        PileOutputData(2, i%) = 1
                        'this value must be at least 1
                    End If
            End If
        Next i%
    End If
End If

'identify total card discrepancy
PileTotalPreAdjust = 0
For i% = 1 To 8
    PileTotalPreAdjust = PileTotalPreAdjust + PileOutputData(2, i%)
Next i%
PileAdjustment = 52 - PileTotalPreAdjust
'identify impossible condition for pure random to limit search of "While Loop"
If PureRandomCount > 0 Then
    If (PureRandomPile - PureRandomCount) + PileAdjustment < 0 Then
        PileAdjustment = PureRandomPile - PureRandomCount
    End If
End If

'NEW ADJUSTMENT LOGIC
'adjust for errors with Pure Random
If PureRandomCount > 0 And PileAdjustment <> 0 Then
    If PileAdjustment > 0 Then
        While PileAdjustment <> 0
            PureRandomSelector = Int(Rnd * 8) + 1
            If PileInputData(1, PureRandomSelector) = "RP" Then
                PileOutputData(2, PureRandomSelector) = _
                    PileOutputData(2, PureRandomSelector) + 1
                PileAdjustment = PileAdjustment - 1
            End If
        Wend
    ElseIf PileAdjustment < 0 Then
        While PileAdjustment <> 0
            PureRandomSelector = Int(Rnd * 8) + 1
            If PileInputData(1, PureRandomSelector) = "RP" And _
                PileOutputData(2, PureRandomSelector) > 1 Then
                PileOutputData(2, PureRandomSelector) = _
                    PileOutputData(2, PureRandomSelector) - 1
                PileAdjustment = PileAdjustment + 1
            End If
        Wend
    End If
End If

''adjust for errors with Pure Random
'If PureRandomCount > 0 Then
'    'first pass
'    For i% = 1 To 8
'        If PileInputData(1, i%) = "RP" Then
'            PileOutputData(2, i%) = Int(PileOutputData(2, i%) + _
'                PileAdjustment * (PileOutputData(2, i%) / PureRandomPile))
'            If PileOutputData(2, i%) < 1 Then
'                PileOutputData(2, i%) = 1
'            End If
'        End If
'    Next i%
'    'recalc adjustment for second pass
'    PileTotalPreAdjust = 0
'    For i% = 1 To 8
'        PileTotalPreAdjust = PileTotalPreAdjust + PileOutputData(2, i%)
'    Next i%
'    PileAdjustment = 52 - PileTotalPreAdjust
'
'    'second pass
'    For i% = 1 To 8
'        If PileAdjustment <> 0 Then
'            If PileInputData(1, i%) = "RP" Then
'                If PileOutputData(2, i%) + PileAdjustment > 0 Then
'                    PileOutputData(2, i%) = PileOutputData(2, i%) + _
'                        PileAdjustment
'                    PileAdjustment = 0
'                End If
'            End If
'        End If
'    Next i%
'End If

'identify total card discrepancy again
PileTotalPreAdjust = 0
For i% = 1 To 8
    PileTotalPreAdjust = PileTotalPreAdjust + PileOutputData(2, i%)
Next i%
PileAdjustment = 52 - PileTotalPreAdjust
'identify impossible condition for approximate random to limit search of "While Loop"
If ApproxRandomCount > 0 Then
    If (ApproxRandomPile - ApproxRandomCount) + PileAdjustment < 0 Then
        PileAdjustment = ApproxRandomPile - ApproxRandomCount
    End If
End If

'NEW ADJUSTMENT LOGIC
'adjust for errors with Approx Random
If ApproxRandomCount > 0 And PileAdjustment <> 0 Then
    If PileAdjustment > 0 Then
        While PileAdjustment <> 0
            ApproxRandomSelector = Int(Rnd * 8) + 1
            If PileInputData(1, ApproxRandomSelector) = "RA" Then
                PileOutputData(2, ApproxRandomSelector) = _
                    PileOutputData(2, ApproxRandomSelector) + 1
                PileAdjustment = PileAdjustment - 1
            End If
        Wend
    ElseIf PileAdjustment < 0 Then
        While PileAdjustment <> 0
            ApproxRandomSelector = Int(Rnd * 8) + 1
            If PileInputData(1, ApproxRandomSelector) = "RA" And _
                PileOutputData(2, ApproxRandomSelector) > 1 Then
                PileOutputData(2, ApproxRandomSelector) = _
                    PileOutputData(2, ApproxRandomSelector) - 1
                PileAdjustment = PileAdjustment + 1
            End If
        Wend
    End If
End If


''adjust for errors with Approx Random
''any remaining adjustments from last segment still active
'If ApproxRandomCount > 0 Then
'    'first pass
'    For i% = 1 To 8
'        If PileInputData(1, i%) = "RA" Then
'            PileOutputData(2, i%) = Int(PileOutputData(2, i%) + _
'                PileAdjustment * (PileOutputData(2, i%) / ApproxRandomPile))
'            If PileOutputData(2, i%) < 1 Then
'                PileOutputData(2, i%) = 1
'            End If
'        End If
'    Next i%
'
'    'recalc adjustment for second pass
'    PileTotalPreAdjust = 0
'    For i% = 1 To 8
'        PileTotalPreAdjust = PileTotalPreAdjust + PileOutputData(2, i%)
'    Next i%
'    PileAdjustment = 52 - PileTotalPreAdjust
'
'    'second pass
'    For i% = 1 To 8
'        If PileAdjustment <> 0 Then
'            If PileInputData(1, i%) = "RA" Then
'                If PileOutputData(2, i%) + PileAdjustment > 0 Then
'                    PileOutputData(2, i%) = PileOutputData(2, i%) + _
'                        PileAdjustment
'                    PileAdjustment = 0
'                End If
'            End If
'        End If
'    Next i%
'End If

'final calculation of remaining adjustments
'which can only be from an error condition
PileTotalPreAdjust = 0
For i% = 1 To 8
    PileTotalPreAdjust = PileTotalPreAdjust + PileOutputData(2, i%)
Next i%
If PileTotalPreAdjust <> 52 Then
    PileParseError = True
    MsgBox "Piles cannot be created as requested." & Chr(13) & Chr(13) _
        & "Please check your input specifications," & Chr(13) _
        & "and try again."
    Exit Sub
End If

'this next section for debugging only
'MsgBox "Piles Codes (" _
'        & PileOutputData(1, 1) _
'        & ", " & PileOutputData(1, 2) _
'        & ", " & PileOutputData(1, 3) _
'        & ", " & PileOutputData(1, 4) _
'        & ", " & PileOutputData(1, 5) _
'        & ", " & PileOutputData(1, 6) _
'        & ", " & PileOutputData(1, 7) _
'        & ", " & PileOutputData(1, 8) & ")" & Chr(13) _
'        & "Pile Data (" & PileOutputData(2, 1) _
'        & ", " & PileOutputData(2, 2) _
'        & ", " & PileOutputData(2, 3) _
'        & ", " & PileOutputData(2, 4) _
'        & ", " & PileOutputData(2, 5) _
'        & ", " & PileOutputData(2, 6) _
'        & ", " & PileOutputData(2, 7) _
'        & ", " & PileOutputData(2, 8) & ")" & Chr(13) _
'        & "Total: " & PileTotalPreAdjust
'end of debugging section
        
'translation of PileOutputData to the PileTable
PileTable(1, 1) = 1
PileTable(1, 2) = PileOutputData(2, 1)
CummPileCount = PileOutputData(2, 1)
For i% = 2 To 8
    PileTable(i%, 1) = CummPileCount + 1
    CummPileCount = CummPileCount + PileOutputData(2, i%)
    PileTable(i%, 2) = CummPileCount
Next i%

CreatePiles (pP)
End Sub

Public Sub CreatePilesDealComplete(pP, pC1, pC2, pC3, pC4, pC5, pC6, pC7, pC8)
'establish key variables
Dim PureRandomPile As Integer
'the total number of pure random cards
Dim ApproxRandomPile As Integer
'the total number of approximate random cards
Dim ExactSpecifiedPile As Integer
'the total number of exact specified cards
Dim PureRandomCount As Integer
'the number of piles that are pure random
Dim ApproxRandomCount As Integer
'the number of piles that are approximate random
Dim ExactSpecifiedCount As Integer
'the number of piles that are exact specified
Dim StrPointer As Integer
Dim PileAdjustment As Integer
'the difference between 52 and PileTotalPreAdjust
Dim PileVariation As Double
'the amount of variation to a random pile size
'set at 15% when the approximate random is specified
'set at 25% of equal breakdown when pure random
Dim PileTotalPreAdjust As Integer
'total card count with random piles before adjustment
Dim CummPileCount
'used in final segment to keep track of total pile count
Dim PileDealCounter As Integer
'used at end for setting the ChangedDeck from dealing

'initialize variables
PureRandomPile = 0
PileParseError = False
'start assuming that there is no error in pile data
ApproxRandomPile = 0
ExactSpecifiedPile = 0
PureRandomCount = 0
ApproxRandomCount = 0
ExactSpecifiedCount = 0
PileTotalPreAdjust = 0
PileCode(1) = pC1
PileCode(2) = pC2
PileCode(3) = pC3
PileCode(4) = pC4
PileCode(5) = pC5
PileCode(6) = pC6
PileCode(7) = pC7
PileCode(8) = pC8
'set X as zero values
For i% = 1 To 8
    StrPointer = InStr(PileCode(i%), "X")
    If StrPointer = 1 Then
        PileInputData(1, i%) = "X"
        PileOutputData(1, i%) = "X"
        PileOutputData(2, i%) = 0
    End If
Next i%
'identify and establish exact cards
For i% = 1 To 8
    StrPointer = InStr(PileCode(i%), "S")
    If StrPointer = 1 Then
        PileInputData(1, i%) = "S"
        PileInputData(2, i%) = Val(Right(PileCode(i%), Len(PileCode(i%)) - 1))
        PileOutputData(1, i%) = "S"
        PileOutputData(2, i%) = Val(Right(PileCode(i%), Len(PileCode(i%)) - 1))
        ExactSpecifiedPile = ExactSpecifiedPile + PileOutputData(2, i%)
        ExactSpecifiedCount = ExactSpecifiedCount + 1
    End If
Next i%
'identify pure random cards
For i% = 1 To 8
    StrPointer = InStr(PileCode(i%), "R")
    If StrPointer = 1 And Len(PileCode(i%)) = 1 Then
        PileInputData(1, i%) = "RP"
        PileInputData(2, i%) = 0
        PileOutputData(1, i%) = "RP"
        PureRandomCount = PureRandomCount + 1
    End If
Next i%
'identify and establish approx random cards
For i% = 1 To 8
    StrPointer = InStr(PileCode(i%), "R")
    If StrPointer = 1 And Len(PileCode(i%)) > 1 Then
        PileInputData(1, i%) = "RA"
        PileInputData(2, i%) = Val(Right(PileCode(i%), Len(PileCode(i%)) - 1))
        PileOutputData(1, i%) = "RA"
        PileVariation = 0.15 * PileInputData(2, i%)
        PileOutputData(2, i%) = Int(Rnd * 2 * PileVariation + _
            PileInputData(2, i%) - PileVariation) + 1
        ApproxRandomPile = ApproxRandomPile + PileOutputData(2, i%)
        ApproxRandomCount = ApproxRandomCount + 1
    End If
Next i%
'establish and fill in Pure Random values
If PureRandomCount > 0 Then
    If ExactSpecifiedPile + ApproxRandomPile > 52 - PureRandomCount Then
        For i% = 1 To 8
            If PileInputData(1, i%) = "RP" Then
                PileOutputData(2, i%) = 1
                PureRandomPile = PureRandomPile + 1
            End If
        Next i%
    Else
        PureRandomPile = 52 - (ExactSpecifiedPile + ApproxRandomPile)
        PileVariation = 0.25 * (PureRandomPile / PureRandomCount)
        For k% = 1 To 8
            If PileInputData(1, k%) = "RP" Then
                PileOutputData(2, k%) = Int(Rnd * 2 * PileVariation + _
                    (PureRandomPile / PureRandomCount) - PileVariation) + 1
                If PileOutputData(2, k%) < 1 Then
                    PileOutputData(2, k%) = 1
                'this value must be at least 1
                End If
            End If
        Next k%
    End If
End If

'identify total card discrepancy
For i% = 1 To 8
    PileTotalPreAdjust = PileTotalPreAdjust + PileOutputData(2, i%)
Next i%
PileAdjustment = 52 - PileTotalPreAdjust

'adjust for errors with Pure Random
If PureRandomCount > 0 Then
    'first pass
    For i% = 1 To 8
        If PileInputData(1, i%) = "RP" Then
            PileOutputData(2, i%) = Int(PileOutputData(2, i%) + _
                PileAdjustment * (PileOutputData(2, i%) / PureRandomPile))
            If PileOutputData(2, i%) < 1 Then
                PileOutputData(2, i%) = 1
            End If
        End If
    Next i%
    'recalc adjustment for second pass
    PileTotalPreAdjust = 0
    For i% = 1 To 8
        PileTotalPreAdjust = PileTotalPreAdjust + PileOutputData(2, i%)
    Next i%
    PileAdjustment = 52 - PileTotalPreAdjust
    
    'second pass
    For i% = 1 To 8
        If PileAdjustment <> 0 Then
            If PileInputData(1, i%) = "RP" Then
                If PileOutputData(2, i%) + PileAdjustment > 0 Then
                    PileOutputData(2, i%) = PileOutputData(2, i%) + _
                        PileAdjustment
                    PileAdjustment = 0
                End If
            End If
        End If
    Next i%
End If

'adjust for errors with Approx Random
'any remaining adjustments from last segment still active
If ApproxRandomCount > 0 Then
    'first pass
    For i% = 1 To 8
        If PileInputData(1, i%) = "RA" Then
            PileOutputData(2, i%) = Int(PileOutputData(2, i%) + _
                PileAdjustment * (PileOutputData(2, i%) / ApproxRandomPile))
            If PileOutputData(2, i%) < 1 Then
                PileOutputData(2, i%) = 1
            End If
        End If
    Next i%
    
    'recalc adjustment for second pass
    PileTotalPreAdjust = 0
    For i% = 1 To 8
        PileTotalPreAdjust = PileTotalPreAdjust + PileOutputData(2, i%)
    Next i%
    PileAdjustment = 52 - PileTotalPreAdjust
    
    'second pass
    For i% = 1 To 8
        If PileAdjustment <> 0 Then
            If PileInputData(1, i%) = "RA" Then
                If PileOutputData(2, i%) + PileAdjustment > 0 Then
                    PileOutputData(2, i%) = PileOutputData(2, i%) + _
                        PileAdjustment
                    PileAdjustment = 0
                End If
            End If
        End If
    Next i%
End If

'final calculation of remaining adjustments
'which can only be from an error condition
PileTotalPreAdjust = 0
For i% = 1 To 8
    PileTotalPreAdjust = PileTotalPreAdjust + PileOutputData(2, i%)
Next i%
If PileTotalPreAdjust <> 52 Then
    PileParseError = True
    MsgBox "Piles cannot be created as requested." & Chr(13) & Chr(13) _
        & "Please check your input specifications," & Chr(13) _
        & "and try again."
    Exit Sub
End If

'this next section for debugging only
'MsgBox "Piles Codes (" _
'        & PileOutputData(1, 1) _
'        & ", " & PileOutputData(1, 2) _
'        & ", " & PileOutputData(1, 3) _
'        & ", " & PileOutputData(1, 4) _
'        & ", " & PileOutputData(1, 5) _
'        & ", " & PileOutputData(1, 6) _
'        & ", " & PileOutputData(1, 7) _
'        & ", " & PileOutputData(1, 8) & ")" & Chr(13) _
'        & "Pile Data (" & PileOutputData(2, 1) _
'        & ", " & PileOutputData(2, 2) _
'        & ", " & PileOutputData(2, 3) _
'        & ", " & PileOutputData(2, 4) _
'        & ", " & PileOutputData(2, 5) _
'        & ", " & PileOutputData(2, 6) _
'        & ", " & PileOutputData(2, 7) _
'        & ", " & PileOutputData(2, 8) & ")" & Chr(13) _
'        & "Total: " & PileTotalPreAdjust
'end of debugging section
        
'translation of PileOutputData to the PileTable
PileTable(1, 1) = 1
PileTable(1, 2) = PileOutputData(2, 1)
CummPileCount = PileOutputData(2, 1)
For i% = 2 To 8
    PileTable(i%, 1) = CummPileCount + 1
    CummPileCount = CummPileCount + PileOutputData(2, i%)
    PileTable(i%, 2) = CummPileCount
Next i%

'rearrange the deck to simulate the dealing of piles
'before the piles are created
For i% = 1 To pP
    PileDealCounter = 1
    For j% = PileTable(i%, 1) To PileTable(i%, 2)
        For k% = 1 To DeckProperties
            ChangedDeck(k%, PileTable(i%, 1) + _
                PileOutputData(2, i%) - PileDealCounter) = Deck(k%, j%)
        Next k%
        PileDealCounter = PileDealCounter + 1
    Next j%
Next i%
For m% = 1 To DeckProperties
    For z% = 1 To DeckCount
        Deck(m%, z%) = ChangedDeck(m%, z%)
    Next z%
Next m%

CreatePiles (pP)
End Sub
Public Sub CreatePilesDealCompleteRandom(paramPiles, paramCards)
Dim pCardCounter As Integer
Dim pPileCounter(8) As Integer
Dim pSelector As Integer
Dim pStartOrder(8) As Integer
Dim pDealOrder(8) As Integer
Dim CardsPerPile As Integer
Dim PilesWithExtra As Integer
Dim pPileOrder(52) As Integer
Dim pEndOfPile As Integer
Dim pSum As Integer
'set the RandomPiles(x,y) array to all zeros
For i% = 1 To 8
    For j% = 1 To 52
        RandomPiles(i%, j%) = 0
    Next j%
Next i%
'distribute cards in first round, one card per pile
'to ensure there are the correct number of piles with at
'least one card
pCardCounter = 1
'set pStartOrder and pDealOrder to initial values
For z% = 1 To 8
    pPileCounter(z%) = 1
    pStartOrder(z%) = z%
    pDealOrder(z%) = 0
Next z%
'establish random deal sequence for first forced round of cards
For p% = 1 To paramPiles
    pSelector = Int(Rnd * (paramPiles + 1 - p%)) + 1
    pDealOrder(p%) = pStartOrder(pSelector)
    If pSelector <> paramPiles + 1 - p% Then
        pStartOrder(pSelector) = pStartOrder(paramPiles + 1 - p%)
    End If
Next p%
'place each deal round's cards in appropriate pile
For m% = 1 To paramPiles
    RandomPiles(pDealOrder(m%), pPileCounter(m%)) = pCardCounter
    pCardCounter = pCardCounter + 1
    pPileCounter(m%) = pPileCounter(m%) + 1
Next m%
'place rest of cards
If paramCards > paramPiles Then
    For i% = pCardCounter To paramCards
        pSelector = Int(Rnd * paramPiles) + 1
        RandomPiles(pSelector, pPileCounter(pSelector)) = pCardCounter
        pCardCounter = pCardCounter + 1
        pPileCounter(pSelector) = pPileCounter(pSelector) + 1
    Next i%
End If
'set pPileOrder values
pCounter = 1
pSum = 0
For i% = 1 To paramPiles
    pEndOfPile = 0
    For j% = 1 To 52
        If RandomPiles(i%, 53 - j%) <> 0 Then
            If pEndOfPile = 0 Then
                PileTable(i%, 2) = 53 - j% + pSum
                pSum = pSum + 53 - j%
                pEndOfPile = 1
            End If
            If j% = 52 Then
                If i% = 1 Then
                    PileTable(1, 1) = 1
                Else
                    PileTable(i%, 1) = PileTable(i% - 1, 2) + 1
                End If
            End If
            pPileOrder(pCounter) = RandomPiles(i%, 53 - j%)
            pCounter = pCounter + 1
        End If
    Next j%
Next i%
If pCounter <= 52 Then
    PileTable(paramPiles + 1, 1) = pCounter
    PileTable(paramPiles + 1, 2) = 52
    For i% = pCounter To 52
        pPileOrder(i%) = i%
    Next i%
End If
For m% = 1 To DeckProperties
    For z% = 1 To DeckCount
        ChangedDeck(m%, z%) = Deck(m%, pPileOrder(z%))
    Next z%
Next m%
For m% = 1 To DeckProperties
    For z% = 1 To DeckCount
        Deck(m%, z%) = ChangedDeck(m%, z%)
    Next z%
Next m%
If pCounter <= 52 Then
    CreatePiles (paramPiles + 1)
Else
    CreatePiles (paramPiles)
End If
End Sub

Public Sub CutPiles(p1, p2, p3, p4)
    'XXX-ERROR CutPiles(R,C,L,X)
    'establish working variables for this module
    Dim pSelectedPile As Integer
    'this is the number of the selected pile
    Dim pSelectedPileCards As Integer
    'this is the number of cards in the selected pile
    Dim pCutDepth As Integer
    'this is the number of cards that are cut from the pile
    Dim pCutToPile As Integer
    'this is the pile number of the return pile
    Dim pCutToCards As Integer
    'this is the number of cards in the return pile
    Dim pSuffix As String
    'st, nd, rd, th suffixes for an error message
    Dim pBackCounter As Integer
    Dim pNewPileCreated As Integer
    Dim fromCardParam As Integer
    Dim toCardParam As Integer
    Dim pCutStartCard As Integer
    Dim pCutEndCard As Integer
    Dim pCounter As Integer
    Dim cardBumpCount As Integer
    '---
    'decode first parameter
    If Left(p1, 1) = "P" Then
        pSelectedPile = Val(Right(p1, Len(p1) - 1))
        If pSelectedPile > NumPiles Then
            pSelectedPile = NumPiles
            'this last statement in case the legitimate requested pSelectedPile
            'is several numbers larger than NumPiles
        End If
    ElseIf Left(p1, 1) = "R" Then
        pSelectedPile = Int(Rnd * NumPiles) + 1
    End If
    pSelectedPileCards = PileTable(pSelectedPile, 2) - PileTable(pSelectedPile, 1) + 1
    '---
    'decode second parameter
    If Left(p2, 1) = "R" Then
        pCutDepth = Int(Rnd * pSelectedPileCards) + 1
    ElseIf Left(p2, 1) = "C" Then
        pCutDepth = pSelectedPileCards
    ElseIf Left(p2, 1) = "S" Then
        pCutDepth = Val(Right(p2, Len(p2) - 1))
        If pCutDepth > pSelectedPileCards Then
            MsgBox ("Error: The specified number of cards to cut" & Chr(13) & _
                "is greater than the number of cards in Pile " & pSelectedPile)
            Exit Sub
        End If
    End If
    '---
    'decode the fourth parameter
    If Left(p4, 1) = "R" Then
        pCutStartCard = PileTable(pSelectedPile, 1)
        pCutEndCard = PileTable(pSelectedPile, 1) + pCutDepth - 1
        For j% = pCutStartCard To pCutEndCard
            For p% = 1 To DeckProperties
                ChangedDeck(p%, j%) = Deck(p%, pCutEndCard - (j% - pCutStartCard))
            Next p%
        Next j%
        For m% = pCutStartCard To pCutEndCard
            For z% = 1 To DeckProperties
                Deck(z%, m%) = ChangedDeck(z%, m%)
            Next z%
        Next m%
        For i% = pCutStartCard To pCutEndCard
            Deck(6, i%) = Not Deck(6, i%)
        Next i%
    End If
    '---
    'decode third parameter
    If Left(p3, 1) = "P" Then
        pCutToPile = pSelectedPile
        pCutToCards = PileTable(pSelectedPile, 2) - PileTable(pSelectedPile, 1) + 1
    ElseIf Left(p3, 1) = "E" Then
        pCutToPile = pSelectedPile
        pCutToCards = PileTable(pSelectedPile, 2) - PileTable(pSelectedPile, 1) + 1
    ElseIf Left(p3, 1) = "M" Then
        pCutToPile = pSelectedPile
        pCutToCards = PileTable(pSelectedPile, 2) - PileTable(pSelectedPile, 1) + 1
    ElseIf Left(p3, 1) = "S" Then
        pCutToPile = Val(Right(p3, Len(p3) - 1))
        If pCutToPile > NumPiles Then
            pCutToPile = NumPiles
            'this last statement in case the legitimate requested pCutToPile
            'is several numbers larger than NumPiles
        End If
        pCutToCards = PileTable(pCutToPile, 2) - PileTable(pCutToPile, 1) + 1
    ElseIf Left(p3, 1) = "R" Then
        pCutToPile = Int(Rnd * NumPiles) + 1
        pCutToCards = PileTable(pCutToPile, 2) - PileTable(pCutToPile, 1) + 1
    ElseIf Left(p3, 1) = "D" Then
        'before establishing a different pile, make sure
        'there is more than 1 pile present
        If NumPiles < 2 Then
            MsgBox ("Error: When specifying 'Top Random Not Same'," & Chr(13) & _
                "there must be more than one pile present.")
            Exit Sub
        Else
            'first set the condition where the selected and return piles are the same
            pCutToPile = pSelectedPile
            'now run a While loop until the pile values are different
            While pCutToPile = pSelectedPile
                pCutToPile = Int(Rnd * NumPiles) + 1
            Wend
            pCutToCards = PileTable(pCutToPile, 2) - PileTable(pCutToPile, 1) + 1
        End If
    ElseIf Left(p3, 1) = "N" Or Left(p3, 1) = "L" Then
        If Left(p3, 1) = "N" Then
            pCutToPile = Val(Right(p3, 1))
        ElseIf Left(p3, 1) = "L" Then
            pCutToPile = Int(Rnd * (NumPiles + 1)) + 1
        End If
'        If NumPiles = 8 Then
'            MsgBox ("There can only be at most 8 piles.  You can not" & _
'                Chr(13) & "create a new pile since there are already 8 piles.")
'            Exit Sub
'        End If
        If pCutToPile > NumPiles Then
            pCutToPile = NumPiles + 1
            'this last statement in case the legitimate requested pCutToPile
            'is several numbers larger than NumPiles
            If pCutDepth = pSelectedPileCards Then
                pCutToPile = NumPiles
                'there can not be an extra pile when the whole pile is moved
            End If
        End If
    End If
    If Left(p3, 1) = "N" Or Left(p3, 1) = "L" Then
        If pSelectedPile <= pCutToPile Then
            'AAA - Below code is correct and complete
            If pCutDepth < pSelectedPileCards Then
                If NumPiles = 8 Then
                    MsgBox ("There can only be at most 8 piles.  You can not" & _
                        Chr(13) & "create a new pile since there are already 8 piles.")
                    Exit Sub
                End If
                'move selected pile
                cPileTable(pCutToPile, 1) = PileTable(pSelectedPile, 1)
                cPileTable(pCutToPile, 2) = PileTable(pSelectedPile, 1) + pCutDepth - 1
                'correct the original selected pile (necessary for next step)
                PileTable(pSelectedPile, 1) = PileTable(pSelectedPile, 1) + pCutDepth
                'calculate cardBumpCount
                cardBumpCount = 0
                For i% = pSelectedPile To pCutToPile - 1
                    cardBumpCount = cardBumpCount + PileTable(i%, 2) - PileTable(i%, 1) + 1
                Next i%
                'adjust target pile
                cPileTable(pCutToPile, 1) = cPileTable(pCutToPile, 1) + cardBumpCount
                cPileTable(pCutToPile, 2) = cPileTable(pCutToPile, 2) + cardBumpCount
                For m% = cPileTable(pCutToPile, 1) To cPileTable(pCutToPile, 2)
                    For z% = 1 To DeckProperties
                        ChangedDeck(z%, m%) = Deck(z%, m% - cardBumpCount)
                    Next z%
                Next m%
                'move "up to pSelectedPile" piles to cPileTable
                If pSelectedPile > 1 Then
                    For i% = 1 To pSelectedPile - 1
                        cPileTable(i%, 1) = PileTable(i%, 1)
                        cPileTable(i%, 2) = PileTable(i%, 2)
                        For m% = cPileTable(i%, 1) To cPileTable(i%, 2)
                            For z% = 1 To DeckProperties
                                ChangedDeck(z%, m%) = Deck(z%, m%)
                            Next z%
                        Next m%
                    Next i%
                End If
                'move "tweener" piles to cPileTable
                For i% = pSelectedPile To pCutToPile - 1
                    cPileTable(i%, 1) = PileTable(i%, 1) - pCutDepth
                    cPileTable(i%, 2) = PileTable(i%, 2) - pCutDepth
                    For m% = cPileTable(i%, 1) To cPileTable(i%, 2)
                        For z% = 1 To DeckProperties
                            ChangedDeck(z%, m%) = Deck(z%, m% + pCutDepth)
                        Next z%
                    Next m%
                Next i%
                'move "greater than pCutToPile" piles to cPileTable
                If pCutToPile <= NumPiles Then
                    For i% = pCutToPile To NumPiles
                        cPileTable(i% + 1, 1) = PileTable(i%, 1)
                        cPileTable(i% + 1, 2) = PileTable(i%, 2)
                        For m% = cPileTable(i% + 1, 1) To cPileTable(i% + 1, 2)
                            For z% = 1 To DeckProperties
                                ChangedDeck(z%, m%) = Deck(z%, m%)
                            Next z%
                        Next m%
                    Next i%
                End If
                'increment NumPiles
                NumPiles = NumPiles + 1
                'move cPileTable back to PileTable
                For i% = 1 To NumPiles
                    PileTable(i%, 1) = cPileTable(i%, 1)
                    PileTable(i%, 2) = cPileTable(i%, 2)
                Next i%
                'rebuild Deck from ChangedDeck
                For m% = 1 To DeckCount
                    For z% = 1 To DeckProperties
                        Deck(z%, m%) = ChangedDeck(z%, m%)
                    Next z%
                Next m%
            ElseIf pCutDepth = pSelectedPileCards Then
                'move selected pile
                cPileTable(pCutToPile, 1) = PileTable(pSelectedPile, 1)
                cPileTable(pCutToPile, 2) = PileTable(pSelectedPile, 2)
                'zero out the original selected pile (necessary for next step)
                PileTable(pSelectedPile, 1) = 1
                PileTable(pSelectedPile, 2) = 0
                'calculate cardBumpCount
                cardBumpCount = 0
                For i% = pSelectedPile To pCutToPile
                    cardBumpCount = cardBumpCount + PileTable(i%, 2) - PileTable(i%, 1) + 1
                Next i%
                'adjust target pile
                cPileTable(pCutToPile, 1) = cPileTable(pCutToPile, 1) + cardBumpCount
                cPileTable(pCutToPile, 2) = cPileTable(pCutToPile, 2) + cardBumpCount
                For m% = cPileTable(pCutToPile, 1) To cPileTable(pCutToPile, 2)
                    For z% = 1 To DeckProperties
                        ChangedDeck(z%, m%) = Deck(z%, m% - cardBumpCount)
                    Next z%
                Next m%
                'move "less than pSelectedPile" piles to cPileTable
                If pSelectedPile > 1 Then
                    For i% = 1 To pSelectedPile - 1
                        cPileTable(i%, 1) = PileTable(i%, 1)
                        cPileTable(i%, 2) = PileTable(i%, 2)
                        For m% = cPileTable(i%, 1) To cPileTable(i%, 2)
                            For z% = 1 To DeckProperties
                                ChangedDeck(z%, m%) = Deck(z%, m%)
                            Next z%
                        Next m%
                    Next i%
                End If
                'move "tweener" piles to cPileTable
                For i% = pSelectedPile + 1 To pCutToPile
                    cPileTable(i% - 1, 1) = PileTable(i%, 1) - pCutDepth
                    cPileTable(i% - 1, 2) = PileTable(i%, 2) - pCutDepth
                    For m% = cPileTable(i% - 1, 1) To cPileTable(i% - 1, 2)
                        For z% = 1 To DeckProperties
                            ChangedDeck(z%, m%) = Deck(z%, m% + pCutDepth)
                        Next z%
                    Next m%
                Next i%
                'move "greater than pCutToPile" piles to cPileTable
                If pCutToPile < NumPiles Then
                    For i% = pCutToPile + 1 To NumPiles
                        cPileTable(i%, 1) = PileTable(i%, 1)
                        cPileTable(i%, 2) = PileTable(i%, 2)
                        For m% = cPileTable(i%, 1) To cPileTable(i%, 2)
                            For z% = 1 To DeckProperties
                                ChangedDeck(z%, m%) = Deck(z%, m%)
                            Next z%
                        Next m%
                    Next i%
                End If
                'move cPileTable back to PileTable
                For i% = 1 To NumPiles
                    PileTable(i%, 1) = cPileTable(i%, 1)
                    PileTable(i%, 2) = cPileTable(i%, 2)
                Next i%
                'rebuild Deck from ChangedDeck
                For m% = 1 To DeckCount
                    For z% = 1 To DeckProperties
                        Deck(z%, m%) = ChangedDeck(z%, m%)
                    Next z%
                Next m%
            End If
            'AAA - Above code is correct and complete
        ElseIf pSelectedPile > pCutToPile Then
            'AAB - below code correct and complete
            If pCutDepth < pSelectedPileCards Then
                If NumPiles = 8 Then
                    MsgBox ("There can only be at most 8 piles.  You can not" & _
                        Chr(13) & "create a new pile since there are already 8 piles.")
                    Exit Sub
                End If
                'move selected pile
                cPileTable(pCutToPile, 1) = PileTable(pSelectedPile, 1)
                cPileTable(pCutToPile, 2) = PileTable(pSelectedPile, 1) + pCutDepth - 1
                'correct the original selected pile (necessary for next step)
                PileTable(pSelectedPile, 1) = PileTable(pSelectedPile, 1) + pCutDepth
                'calculate cardBumpCount
                cardBumpCount = 0
                For i% = pCutToPile To pSelectedPile - 1
                    cardBumpCount = cardBumpCount + PileTable(i%, 2) - PileTable(i%, 1) + 1
                Next i%
                'adjust target pile
                cPileTable(pCutToPile, 1) = cPileTable(pCutToPile, 1) - cardBumpCount
                cPileTable(pCutToPile, 2) = cPileTable(pCutToPile, 2) - cardBumpCount
                For m% = cPileTable(pCutToPile, 1) To cPileTable(pCutToPile, 2)
                    For z% = 1 To DeckProperties
                        ChangedDeck(z%, m%) = Deck(z%, m% + cardBumpCount)
                    Next z%
                Next m%
                'move "up to pCutToPile" piles to cPileTable
                If pCutToPile > 1 Then
                    For i% = 1 To pCutToPile - 1
                        cPileTable(i%, 1) = PileTable(i%, 1)
                        cPileTable(i%, 2) = PileTable(i%, 2)
                        For m% = cPileTable(i%, 1) To cPileTable(i%, 2)
                            For z% = 1 To DeckProperties
                                ChangedDeck(z%, m%) = Deck(z%, m%)
                            Next z%
                        Next m%
                    Next i%
                End If
                'move "tweener" piles to cPileTable
                For i% = pCutToPile To pSelectedPile - 1
                    cPileTable(i% + 1, 1) = PileTable(i%, 1) + pCutDepth
                    cPileTable(i% + 1, 2) = PileTable(i%, 2) + pCutDepth
                    For m% = cPileTable(i% + 1, 1) To cPileTable(i% + 1, 2)
                        For z% = 1 To DeckProperties
                            ChangedDeck(z%, m%) = Deck(z%, m% - pCutDepth)
                        Next z%
                    Next m%
                Next i%
                'move "greater than pCutToPile" piles to cPileTable
                'If pCutToPile <= NumPiles Then
                    For i% = pSelectedPile To NumPiles
                        cPileTable(i% + 1, 1) = PileTable(i%, 1)
                        cPileTable(i% + 1, 2) = PileTable(i%, 2)
                        For m% = cPileTable(i% + 1, 1) To cPileTable(i% + 1, 2)
                            For z% = 1 To DeckProperties
                                ChangedDeck(z%, m%) = Deck(z%, m%)
                            Next z%
                        Next m%
                    Next i%
                'End If
                'increment NumPiles
                NumPiles = NumPiles + 1
                'move cPileTable back to PileTable
                For i% = 1 To NumPiles
                    PileTable(i%, 1) = cPileTable(i%, 1)
                    PileTable(i%, 2) = cPileTable(i%, 2)
                Next i%
                'rebuild Deck from ChangedDeck
                For m% = 1 To DeckCount
                    For z% = 1 To DeckProperties
                        Deck(z%, m%) = ChangedDeck(z%, m%)
                    Next z%
                Next m%
            ElseIf pCutDepth = pSelectedPileCards Then
                'move selected pile
                cPileTable(pCutToPile, 1) = PileTable(pSelectedPile, 1)
                cPileTable(pCutToPile, 2) = PileTable(pSelectedPile, 2)
                ''zero out the original selected pile (necessary for next step)
                'PileTable(pSelectedPile, 1) = 1
                'PileTable(pSelectedPile, 2) = 0
                'calculate cardBumpCount
                cardBumpCount = 0
                For i% = pCutToPile To pSelectedPile - 1
                    cardBumpCount = cardBumpCount + PileTable(i%, 2) - PileTable(i%, 1) + 1
                Next i%
                'adjust target pile
                cPileTable(pCutToPile, 1) = cPileTable(pCutToPile, 1) - cardBumpCount
                cPileTable(pCutToPile, 2) = cPileTable(pCutToPile, 2) - cardBumpCount
                For m% = cPileTable(pCutToPile, 1) To cPileTable(pCutToPile, 2)
                    For z% = 1 To DeckProperties
                        ChangedDeck(z%, m%) = Deck(z%, m% + cardBumpCount)
                    Next z%
                Next m%
                'move "less than pCutToPile" piles to cPileTable
                If pCutToPile > 1 Then
                    For i% = 1 To pCutToPile - 1
                        cPileTable(i%, 1) = PileTable(i%, 1)
                        cPileTable(i%, 2) = PileTable(i%, 2)
                        For m% = cPileTable(i%, 1) To cPileTable(i%, 2)
                            For z% = 1 To DeckProperties
                                ChangedDeck(z%, m%) = Deck(z%, m%)
                            Next z%
                        Next m%
                    Next i%
                End If
                'move "tweener" piles to cPileTable
                For i% = pCutToPile To pSelectedPile - 1
                    cPileTable(i% + 1, 1) = PileTable(i%, 1) + pCutDepth
                    cPileTable(i% + 1, 2) = PileTable(i%, 2) + pCutDepth
                    For m% = cPileTable(i% + 1, 1) To cPileTable(i% + 1, 2)
                        For z% = 1 To DeckProperties
                            ChangedDeck(z%, m%) = Deck(z%, m% - pCutDepth)
                        Next z%
                    Next m%
                Next i%
                'move "greater than pSelectedPile" piles to cPileTable
                If pSelectedPile < NumPiles Then
                    For i% = pSelectedPile + 1 To NumPiles
                        cPileTable(i%, 1) = PileTable(i%, 1)
                        cPileTable(i%, 2) = PileTable(i%, 2)
                        For m% = cPileTable(i%, 1) To cPileTable(i%, 2)
                            For z% = 1 To DeckProperties
                                ChangedDeck(z%, m%) = Deck(z%, m%)
                            Next z%
                        Next m%
                    Next i%
                End If
                'move cPileTable back to PileTable
                For i% = 1 To NumPiles
                    PileTable(i%, 1) = cPileTable(i%, 1)
                    PileTable(i%, 2) = cPileTable(i%, 2)
                Next i%
                'rebuild Deck from ChangedDeck
                For m% = 1 To DeckCount
                    For z% = 1 To DeckProperties
                        Deck(z%, m%) = ChangedDeck(z%, m%)
                    Next z%
                Next m%
            End If
                'AAB - above code correct and complete
        End If
    '-----------------------------------------------------------
    'next major section handles when there is not a new pile
    Else
        If pSelectedPile < pCutToPile Then
            'AAC - Below code is correct and complete
            If pCutDepth < pSelectedPileCards Then
                'move selected pile
                cPileTable(pCutToPile, 1) = PileTable(pCutToPile, 1) - pCutDepth
                cPileTable(pCutToPile, 2) = PileTable(pCutToPile, 2)
                'correct the original selected pile (necessary for next step)
                PileTable(pSelectedPile, 1) = PileTable(pSelectedPile, 1) + pCutDepth
                'calculate cardBumpCount
                cardBumpCount = 0
                For i% = pSelectedPile To pCutToPile - 1
                    cardBumpCount = cardBumpCount + PileTable(i%, 2) - PileTable(i%, 1) + 1
                Next i%
                ''adjust target pile
                'cPileTable(pCutToPile, 1) = cPileTable(pCutToPile, 1) + cardBumpCount
                'cPileTable(pCutToPile, 2) = cPileTable(pCutToPile, 2) + cardBumpCount
                For m% = cPileTable(pCutToPile, 1) To cPileTable(pCutToPile, 1) + pCutDepth - 1
                    For z% = 1 To DeckProperties
                        ChangedDeck(z%, m%) = Deck(z%, m% - cardBumpCount)
                    Next z%
                Next m%
                For m% = cPileTable(pCutToPile, 1) + pCutDepth To cPileTable(pCutToPile, 2)
                    For z% = 1 To DeckProperties
                        ChangedDeck(z%, m%) = Deck(z%, m%)
                    Next z%
                Next m%
                'move "up to pSelectedPile" piles to cPileTable
                If pSelectedPile > 1 Then
                    For i% = 1 To pSelectedPile - 1
                        cPileTable(i%, 1) = PileTable(i%, 1)
                        cPileTable(i%, 2) = PileTable(i%, 2)
                        For m% = cPileTable(i%, 1) To cPileTable(i%, 2)
                            For z% = 1 To DeckProperties
                                ChangedDeck(z%, m%) = Deck(z%, m%)
                            Next z%
                        Next m%
                    Next i%
                End If
                'move "tweener" piles to cPileTable
                For i% = pSelectedPile To pCutToPile - 1
                    cPileTable(i%, 1) = PileTable(i%, 1) - pCutDepth
                    cPileTable(i%, 2) = PileTable(i%, 2) - pCutDepth
                    For m% = cPileTable(i%, 1) To cPileTable(i%, 2)
                        For z% = 1 To DeckProperties
                            ChangedDeck(z%, m%) = Deck(z%, m% + pCutDepth)
                        Next z%
                    Next m%
                Next i%
                'move "greater than pCutToPile" piles to cPileTable
                If pCutToPile <= NumPiles Then
                    For i% = pCutToPile + 1 To NumPiles
                        cPileTable(i%, 1) = PileTable(i%, 1)
                        cPileTable(i%, 2) = PileTable(i%, 2)
                        For m% = cPileTable(i%, 1) To cPileTable(i%, 2)
                            For z% = 1 To DeckProperties
                                ChangedDeck(z%, m%) = Deck(z%, m%)
                            Next z%
                        Next m%
                    Next i%
                End If
                'move cPileTable back to PileTable
                For i% = 1 To NumPiles
                    PileTable(i%, 1) = cPileTable(i%, 1)
                    PileTable(i%, 2) = cPileTable(i%, 2)
                Next i%
                'rebuild Deck from ChangedDeck
                For m% = 1 To DeckCount
                    For z% = 1 To DeckProperties
                        Deck(z%, m%) = ChangedDeck(z%, m%)
                    Next z%
                Next m%
            ElseIf pCutDepth = pSelectedPileCards Then
                'move selected pile
                cPileTable(pCutToPile, 1) = PileTable(pCutToPile, 1) - pCutDepth
                cPileTable(pCutToPile, 2) = PileTable(pCutToPile, 2)
                'zero out the original selected pile (necessary for next step)
                PileTable(pSelectedPile, 1) = 1
                PileTable(pSelectedPile, 2) = 0
                'calculate cardBumpCount
                cardBumpCount = 0
                For i% = pSelectedPile To pCutToPile - 1
                    cardBumpCount = cardBumpCount + PileTable(i%, 2) - PileTable(i%, 1) + 1
                Next i%
                ''adjust target pile
                'cPileTable(pCutToPile, 1) = cPileTable(pCutToPile, 1) + cardBumpCount
                'cPileTable(pCutToPile, 2) = cPileTable(pCutToPile, 2) + cardBumpCount
                For m% = cPileTable(pCutToPile, 1) To cPileTable(pCutToPile, 1) + pCutDepth - 1
                    For z% = 1 To DeckProperties
                        ChangedDeck(z%, m%) = Deck(z%, m% - cardBumpCount)
                    Next z%
                Next m%
                For m% = cPileTable(pCutToPile, 1) + pCutDepth To cPileTable(pCutToPile, 2)
                    For z% = 1 To DeckProperties
                        ChangedDeck(z%, m%) = Deck(z%, m%)
                    Next z%
                Next m%
                'move "less than pSelectedPile" piles to cPileTable
                If pSelectedPile > 1 Then
                    For i% = 1 To pSelectedPile - 1
                        cPileTable(i%, 1) = PileTable(i%, 1)
                        cPileTable(i%, 2) = PileTable(i%, 2)
                        For m% = cPileTable(i%, 1) To cPileTable(i%, 2)
                            For z% = 1 To DeckProperties
                                ChangedDeck(z%, m%) = Deck(z%, m%)
                            Next z%
                        Next m%
                    Next i%
                End If
                'move "tweener" piles to cPileTable
                If pCutToPile - pSelectedPile > 1 Then
                    For i% = pSelectedPile To pCutToPile - 2
                        cPileTable(i%, 1) = PileTable(i% + 1, 1) - pCutDepth
                        cPileTable(i%, 2) = PileTable(i% + 1, 2) - pCutDepth
                        For m% = cPileTable(i%, 1) To cPileTable(i%, 2)
                            For z% = 1 To DeckProperties
                                ChangedDeck(z%, m%) = Deck(z%, m% + pCutDepth)
                            Next z%
                        Next m%
                    Next i%
                End If
                'adjust pCutToPile table index
                cPileTable(pCutToPile - 1, 1) = cPileTable(pCutToPile, 1)
                cPileTable(pCutToPile - 1, 2) = cPileTable(pCutToPile, 2)
                'move "greater than pCutToPile" piles to cPileTable
                If pCutToPile < NumPiles Then
                    For i% = pCutToPile + 1 To NumPiles
                        cPileTable(i% - 1, 1) = PileTable(i%, 1)
                        cPileTable(i% - 1, 2) = PileTable(i%, 2)
                        For m% = cPileTable(i% - 1, 1) To cPileTable(i% - 1, 2)
                            For z% = 1 To DeckProperties
                                ChangedDeck(z%, m%) = Deck(z%, m%)
                            Next z%
                        Next m%
                    Next i%
                End If
                'adjust NumPiles
                NumPiles = NumPiles - 1
                'move cPileTable back to PileTable
                For i% = 1 To NumPiles
                    PileTable(i%, 1) = cPileTable(i%, 1)
                    PileTable(i%, 2) = cPileTable(i%, 2)
                Next i%
                'rebuild Deck from ChangedDeck
                For m% = 1 To DeckCount
                    For z% = 1 To DeckProperties
                        Deck(z%, m%) = ChangedDeck(z%, m%)
                    Next z%
                Next m%
            End If
            'AAC - Above code is correct and complete
        ElseIf pSelectedPile > pCutToPile Then
            'AAD - below code
            If pCutDepth < pSelectedPileCards Then
                'move selected pile
                cPileTable(pCutToPile, 1) = PileTable(pCutToPile, 1)
                cPileTable(pCutToPile, 2) = PileTable(pCutToPile, 2) + pCutDepth
                'correct the original selected pile (necessary for next step)
                PileTable(pSelectedPile, 1) = PileTable(pSelectedPile, 1) + pCutDepth
                'calculate cardBumpCount
                cardBumpCount = 0
                For i% = pCutToPile To pSelectedPile - 1
                    cardBumpCount = cardBumpCount + PileTable(i%, 2) - PileTable(i%, 1) + 1
                Next i%
                ''adjust target pile
                'cPileTable(pCutToPile, 1) = cPileTable(pCutToPile, 1) + cardBumpCount
                'cPileTable(pCutToPile, 2) = cPileTable(pCutToPile, 2) + cardBumpCount
                For m% = cPileTable(pCutToPile, 1) To cPileTable(pCutToPile, 1) + pCutDepth - 1
                    For z% = 1 To DeckProperties
                        ChangedDeck(z%, m%) = Deck(z%, m% + cardBumpCount)
                    Next z%
                Next m%
                For m% = cPileTable(pCutToPile, 1) + pCutDepth To cPileTable(pCutToPile, 2)
                    For z% = 1 To DeckProperties
                        ChangedDeck(z%, m%) = Deck(z%, m% - pCutDepth)
                    Next z%
                Next m%
                'move "up to pcuttoPile" piles to cPileTable
                If pCutToPile > 1 Then
                    For i% = 1 To pCutToPile - 1
                        cPileTable(i%, 1) = PileTable(i%, 1)
                        cPileTable(i%, 2) = PileTable(i%, 2)
                        For m% = cPileTable(i%, 1) To cPileTable(i%, 2)
                            For z% = 1 To DeckProperties
                                ChangedDeck(z%, m%) = Deck(z%, m%)
                            Next z%
                        Next m%
                    Next i%
                End If
                'move "tweener" piles to cPileTable
                If pSelectedPile - pCutToPile > 1 Then
                    For i% = pCutToPile + 1 To pSelectedPile - 1
                        cPileTable(i%, 1) = PileTable(i%, 1) + pCutDepth
                        cPileTable(i%, 2) = PileTable(i%, 2) + pCutDepth
                        For m% = cPileTable(i%, 1) To cPileTable(i%, 2)
                            For z% = 1 To DeckProperties
                                ChangedDeck(z%, m%) = Deck(z%, m% - pCutDepth)
                            Next z%
                        Next m%
                    Next i%
                End If
                'move pSelectedPile data and cards over to cPileTable
                cPileTable(pSelectedPile, 1) = PileTable(pSelectedPile, 1)
                cPileTable(pSelectedPile, 2) = PileTable(pSelectedPile, 2)
                For m% = PileTable(pSelectedPile, 1) To PileTable(pSelectedPile, 2)
                    For z% = 1 To DeckProperties
                        ChangedDeck(z%, m%) = Deck(z%, m%)
                    Next z%
                Next m%
                'move "greater than pSelectedPile" piles to cPileTable
                If pSelectedPile <= NumPiles Then
                    For i% = pSelectedPile + 1 To NumPiles
                        cPileTable(i%, 1) = PileTable(i%, 1)
                        cPileTable(i%, 2) = PileTable(i%, 2)
                        For m% = cPileTable(i%, 1) To cPileTable(i%, 2)
                            For z% = 1 To DeckProperties
                                ChangedDeck(z%, m%) = Deck(z%, m%)
                            Next z%
                        Next m%
                    Next i%
                End If
                'move cPileTable back to PileTable
                For i% = 1 To NumPiles
                    PileTable(i%, 1) = cPileTable(i%, 1)
                    PileTable(i%, 2) = cPileTable(i%, 2)
                Next i%
                'rebuild Deck from ChangedDeck
                For m% = 1 To DeckCount
                    For z% = 1 To DeckProperties
                        Deck(z%, m%) = ChangedDeck(z%, m%)
                    Next z%
                Next m%
            ElseIf pCutDepth = pSelectedPileCards Then
                'move selected pile
                cPileTable(pCutToPile, 1) = PileTable(pCutToPile, 1)
                cPileTable(pCutToPile, 2) = PileTable(pCutToPile, 2) + pCutDepth
                'calculate cardBumpCount
                cardBumpCount = 0
                For i% = pCutToPile To pSelectedPile - 1
                    cardBumpCount = cardBumpCount + PileTable(i%, 2) - PileTable(i%, 1) + 1
                Next i%
                ''adjust target pile
                'cPileTable(pCutToPile, 1) = cPileTable(pCutToPile, 1) + cardBumpCount
                'cPileTable(pCutToPile, 2) = cPileTable(pCutToPile, 2) + cardBumpCount
                For m% = cPileTable(pCutToPile, 1) To cPileTable(pCutToPile, 1) + pCutDepth - 1
                    For z% = 1 To DeckProperties
                        ChangedDeck(z%, m%) = Deck(z%, m% + cardBumpCount)
                    Next z%
                Next m%
                For m% = cPileTable(pCutToPile, 1) + pCutDepth To cPileTable(pCutToPile, 2)
                    For z% = 1 To DeckProperties
                        ChangedDeck(z%, m%) = Deck(z%, m% - pCutDepth)
                    Next z%
                Next m%
                'move "less than pCutToPile" piles to cPileTable
                If pCutToPile > 1 Then
                    For i% = 1 To pCutToPile - 1
                        cPileTable(i%, 1) = PileTable(i%, 1)
                        cPileTable(i%, 2) = PileTable(i%, 2)
                        For m% = cPileTable(i%, 1) To cPileTable(i%, 2)
                            For z% = 1 To DeckProperties
                                ChangedDeck(z%, m%) = Deck(z%, m%)
                            Next z%
                        Next m%
                    Next i%
                End If
                'move "tweener" piles to cPileTable
                If pSelectedPile - pCutToPile > 1 Then
                    For i% = pCutToPile + 1 To pSelectedPile - 1
                        cPileTable(i%, 1) = PileTable(i%, 1) + pCutDepth
                        cPileTable(i%, 2) = PileTable(i%, 2) + pCutDepth
                        For m% = cPileTable(i%, 1) To cPileTable(i%, 2)
                            For z% = 1 To DeckProperties
                                ChangedDeck(z%, m%) = Deck(z%, m% - pCutDepth)
                            Next z%
                        Next m%
                    Next i%
                End If
                'move "greater than pCutToPile" piles to cPileTable
                If pSelectedPile < NumPiles Then
                    For i% = pSelectedPile + 1 To NumPiles
                        cPileTable(i% - 1, 1) = PileTable(i%, 1)
                        cPileTable(i% - 1, 2) = PileTable(i%, 2)
                        For m% = cPileTable(i% - 1, 1) To cPileTable(i% - 1, 2)
                            For z% = 1 To DeckProperties
                                ChangedDeck(z%, m%) = Deck(z%, m%)
                            Next z%
                        Next m%
                    Next i%
                End If
                'adjust NumPiles
                NumPiles = NumPiles - 1
                'move cPileTable back to PileTable
                For i% = 1 To NumPiles
                    PileTable(i%, 1) = cPileTable(i%, 1)
                    PileTable(i%, 2) = cPileTable(i%, 2)
                Next i%
                'rebuild Deck from ChangedDeck
                For m% = 1 To DeckCount
                    For z% = 1 To DeckProperties
                        Deck(z%, m%) = ChangedDeck(z%, m%)
                    Next z%
                Next m%
            End If
                'AAD - above code
        ElseIf pSelectedPile = pCutToPile Then
            If Left(p3, 1) = "P" Or Left(p3, 1) = "E" Then
                If pCutDepth < pCutToCards Then
                    'first segment
                    For m% = PileTable(pSelectedPile, 1) To PileTable(pSelectedPile, 1) + _
                        (pCutToCards - pCutDepth) - 1
                        For z% = 1 To DeckProperties
                            ChangedDeck(z%, m%) = Deck(z%, m% + pCutDepth)
                        Next z%
                    Next m%
                    'second segment
                    For m% = PileTable(pSelectedPile, 1) + (pCutToCards - pCutDepth) To _
                        PileTable(pSelectedPile, 2)
                        For z% = 1 To DeckProperties
                            ChangedDeck(z%, m%) = Deck(z%, m% - (pCutToCards - pCutDepth))
                        Next z%
                    Next m%
                    'rebuild deck
                    For m% = PileTable(pSelectedPile, 1) To PileTable(pSelectedPile, 2)
                        For z% = 1 To DeckProperties
                            Deck(z%, m%) = ChangedDeck(z%, m%)
                        Next z%
                    Next m%
                Else
                    'do nothing for this case
                End If
            End If
        End If
    End If
    'visual pile refresh
    PilesMatrixRefresh
    CreatePiles (NumPiles)
End Sub


Private Sub CutCardsButton_Click()
    'make sure piles have been created
    If PilesShown = 0 Then
        MsgBox ("You must create piles before you can manipulate piles.")
        Exit Sub
    End If
    'set parameter variables
    Dim p1 As Variant
    Dim p2 As Variant
    Dim p3 As Variant
    Dim p4 As Variant
    Dim pCutPile As Integer
    Dim pPlacePile As Integer
    Dim pTempCounter As Integer
    ' make sure the PileMatrix has a valid selection
    ' but only if not random selections chosen
    If Not (CutRandomPile.Value = True And _
        (TopRandomAny.Value = True Or _
        TopRandomNotSame.Value = True Or _
        TopSame.Value = True Or _
        CompleteCut.Value = True Or _
        PlaceNewPileRandom.Value = True Or _
        PlaceNewPileSpecified.Value = True)) Then
            Dim PileMatrixCheckSum As Integer
            PileMatrixCheckSum = 0
            For i% = 1 To 8
                For j% = 1 To 8
                    If PileMatrix(10 * i% + j%).Value = 1 Then
                        PileMatrixCheckSum = PileMatrixCheckSum + 1
                    End If
                Next j%
            Next i%
            If PileMatrixCheckSum = 0 Then
                MsgBox ("You must check a box in the Pile Control Matrix before" & Chr(13) & _
                    "you can manipulate piles with non-random settings.")
                Exit Sub
            End If
    End If
    'initialize Pile codes
    p1 = Empty
    p2 = Empty
    p3 = Empty
    p4 = Empty
    'establish parameter codes
    'get Pile Matrix info if non random settings
    If Not (CutRandomPile.Value = True And _
        (TopRandomAny.Value = True Or _
        TopRandomNotSame.Value = True Or _
        CompleteCut.Value = True Or _
        PlaceNewPileRandom.Value = True Or _
        PlaceNewPileSpecified.Value = True)) Then
            PileMatrixQuery
    End If
    'set Cut Pile code (p1)
    If CutPrimaryPile.Value = True Then
        pCutPile = PileMatrixRow
        p1 = "P" & pCutPile
    ElseIf CutRandomPile.Value = True Then
        p1 = "R"
    End If
    'set CutPortion code (p2)
    If CutRandom.Value = True Then
        p2 = "R"
    ElseIf CompletePile.Value = True Then
        p2 = "C"
    ElseIf CutSpecified.Value = True Then
        If Not IsNumeric(CutSpecifiedText.Text) Then
            CutSpecifiedText.Text = Empty
            CutSpecifiedText.SetFocus
            MsgBox ("You can only enter numbers in the 'Specified' text box.")
            Exit Sub
        End If
        If Val(CutSpecifiedText.Text) < 1 Then
            CutSpecifiedText.Text = Empty
            CutSpecifiedText.SetFocus
            MsgBox ("You can only enter numbers greater than 0" & _
                Chr(13) & "in the 'Specified' text box.")
            Exit Sub
        End If
        p2 = "S" & CutSpecifiedText.Text
    End If
    'set Place Pile code (p3)
    If CompleteCut.Value = True Then
        If CutPrimaryPile.Value = True Then
            p3 = "P"
        ElseIf CutRandomPile.Value = True Then
            p3 = "E"
            '"E" is equivalent to Same for when the cut pile was random
            'the Call code will need to determine the random values
        End If
    ElseIf TopSecondary.Value = True Then
        pPlacePile = PileMatrixColumn
        p3 = "S" & pPlacePile
    ElseIf TopRandomAny.Value = True Then
        p3 = "R"
    ElseIf TopSame.Value = True Then
        p3 = "M"
    ElseIf TopRandomNotSame.Value = True Then
        p3 = "D"
        'means Different (not same)
    ElseIf PlaceNewPileSpecified.Value = True Then
        If Not IsNumeric(PlaceNewPileSpecifiedText.Text) Then
            PlaceNewPileSpecifiedText.Text = Empty
            PlaceNewPileSpecifiedText.SetFocus
            MsgBox ("You can only enter numbers in the 'New Pile' text box.")
            Exit Sub
        End If
        If Val(PlaceNewPileSpecifiedText.Text) < 1 Or _
            Val(PlaceNewPileSpecifiedText.Text) > 8 Then
            PlaceNewPileSpecifiedText.Text = Empty
            PlaceNewPileSpecifiedText.SetFocus
            MsgBox ("You can only enter numbers from 1 to 8" & _
                Chr(13) & "in the 'New Pile' text box.")
            Exit Sub
        End If
        p3 = "N" & PlaceNewPileSpecifiedText.Text
    ElseIf PlaceNewPileRandom.Value = True Then
        p3 = "L"
    End If
    'set Reverse code (p4)
    If ReverseCutPortion.Value = 1 Then
        p4 = "R"
    Else
        p4 = "X"
    End If
    'identify ignoring of Reverse checks in the Pile Matrix if they are present
    pTempCounter = 0
    For i% = 1 To 8
        If ReverseR(i%).Value = 1 Then
            pTempCounter = pTempCounter + 1
        End If
        If ReverseC(i%).Value = 1 Then
            pTempCounter = pTempCounter + 1
        End If
    Next i%
    If pTempCounter > 0 Then
        MsgBox ("When Cut Cards command is performed, 'Reverse'" & _
        Chr(13) & "checkboxes in the Pile Matrix are ignored.")
    End If
    Call CutPiles(p1, p2, p3, p4)
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "CutPiles(" & p1 _
        & ", " & p2 _
        & ", " & p3 _
        & ", " & p4 _
        & ")"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End Sub

Private Sub CutCardsOption_Click()
DealCardsFrame.Visible = False
CutCardsFrame.Visible = True
End Sub

Private Sub CutSpecifiedText_Change()
CutSpecified.Value = True
End Sub

Private Sub DealAlternatingOption_Click()
DealCardsFrame.Height = 3045
CutCardsFrame.Visible = False
End Sub

Private Sub DealCardsOption_Click()
DealCardsFrame.Visible = True
If DealCompleteOption.Value = True Then
    CutCardsFrame.Visible = True
Else
    CutCardsFrame.Visible = False
End If
End Sub


Private Sub DealCompleteOption_Click()
DealCardsFrame.Height = 810
CutCardsFrame.Visible = True
End Sub

Private Sub DealPilesCheck()
    If (NumberOfPilesText.Text <> Empty And _
        (Not IsNumeric(NumberOfPilesText.Text) Or _
        Val(NumberOfPilesText.Text) < 1 Or _
        Val(NumberOfPilesText.Text) > 8)) Then
'        Or _
'        NumberOfPilesText.Text = Empty Then
            NumberOfPilesText.Text = Empty
            NumberOfPilesText.SetFocus
            MsgBox "Please enter a valid number of piles (1 to 8)" & Chr(13) _
                & "in the 'Create Piles' Input Box"
            PileError = True
            Exit Sub
    End If
    If (NumberOfCardsToDealText.Text <> Empty And _
        (Not IsNumeric(NumberOfCardsToDealText.Text) Or _
        Val(NumberOfCardsToDealText.Text) < 1 Or _
        Val(NumberOfCardsToDealText.Text) > 52)) Then
            NumberOfCardsToDealText.Text = Empty
            NumberOfCardsToDealText.SetFocus
            MsgBox "Please enter a valid number of cards (1 to 52)" & Chr(13) _
                & "in the 'Cards to Deal' Input Box"
            PileError = True
            Exit Sub
    End If
    If DealCardsOption.Value = True And _
        NumberOfCardsToDealText.Text <> Empty And _
        Val(NumberOfCardsToDealText.Text) < 52 And _
        Val(NumberOfPilesText.Text) = 8 Then
        NumberOfPilesText.Text = Empty
        MsgBox "When dealing less than 52 cards," & Chr(13) _
                & "the maximum number of piles is 7." & Chr(13) _
                & "(the 8th pile is used for the balance of the deck)"
        NumberOfPilesText.SetFocus
        PileError = True
        Exit Sub
    End If
    If DealCardsOption.Value = True And _
        NumberOfCardsToDealText.Text <> Empty And _
        Val(NumberOfCardsToDealText.Text) < Val(NumberOfPilesText.Text) Then
        NumberOfCardsToDealText.Text = Empty
        MsgBox "You must deal at least as many" & Chr(13) _
                & "cards as there are piles."
        NumberOfCardsToDealText.SetFocus
        PileError = True
        Exit Sub
    End If
    
'check for error entries when the Specified option is selected
'entries must be between 1 and 52
    If PileSpecified1.Value = True And _
        (SpecifiedText1.Text = Empty Or _
         Not IsNumeric(SpecifiedText1.Text) Or _
         Val(SpecifiedText1.Text) < 1 Or _
         Val(SpecifiedText1.Text) > 52) Then
        MsgBox "If the 'Specified' option is selected," & Chr(13) _
                & "you must enter a value between 1 and 52."
        SpecifiedText1.SetFocus
        PileError = True
        Exit Sub
    End If
    If PileSpecified2.Value = True And _
        Val(NumberOfPilesText.Text) >= 2 And _
        (SpecifiedText2.Text = Empty Or _
         Not IsNumeric(SpecifiedText2.Text) Or _
         Val(SpecifiedText2.Text) < 1 Or _
         Val(SpecifiedText2.Text) > 52) Then
        MsgBox "If the 'Specified' option is selected," & Chr(13) _
                & "you must enter a value between 1 and 52."
        SpecifiedText2.SetFocus
        PileError = True
        Exit Sub
    End If
    If PileSpecified3.Value = True And _
        Val(NumberOfPilesText.Text) >= 3 And _
        (SpecifiedText3.Text = Empty Or _
         Not IsNumeric(SpecifiedText3.Text) Or _
         Val(SpecifiedText3.Text) < 1 Or _
         Val(SpecifiedText3.Text) > 52) Then
        MsgBox "If the 'Specified' option is selected," & Chr(13) _
                & "you must enter a value between 1 and 52."
        SpecifiedText3.SetFocus
        PileError = True
        Exit Sub
    End If
    If PileSpecified4.Value = True And _
        Val(NumberOfPilesText.Text) >= 4 And _
        (SpecifiedText4.Text = Empty Or _
         Not IsNumeric(SpecifiedText4.Text) Or _
         Val(SpecifiedText4.Text) < 1 Or _
         Val(SpecifiedText4.Text) > 52) Then
        MsgBox "If the 'Specified' option is selected," & Chr(13) _
                & "you must enter a value between 1 and 52."
        SpecifiedText4.SetFocus
        PileError = True
        Exit Sub
    End If
    If PileSpecified5.Value = True And _
        Val(NumberOfPilesText.Text) >= 5 And _
        (SpecifiedText5.Text = Empty Or _
         Not IsNumeric(SpecifiedText5.Text) Or _
         Val(SpecifiedText5.Text) < 1 Or _
         Val(SpecifiedText5.Text) > 52) Then
        MsgBox "If the 'Specified' option is selected," & Chr(13) _
                & "you must enter a value between 1 and 52."
        SpecifiedText5.SetFocus
        PileError = True
        Exit Sub
    End If
    If PileSpecified6.Value = True And _
        Val(NumberOfPilesText.Text) >= 6 And _
        (SpecifiedText6.Text = Empty Or _
         Not IsNumeric(SpecifiedText6.Text) Or _
         Val(SpecifiedText6.Text) < 1 Or _
         Val(SpecifiedText6.Text) > 52) Then
        MsgBox "If the 'Specified' option is selected," & Chr(13) _
                & "you must enter a value between 1 and 52."
        SpecifiedText6.SetFocus
        PileError = True
        Exit Sub
    End If
    If PileSpecified7.Value = True And _
        Val(NumberOfPilesText.Text) >= 7 And _
        (SpecifiedText7.Text = Empty Or _
         Not IsNumeric(SpecifiedText7.Text) Or _
         Val(SpecifiedText7.Text) < 1 Or _
         Val(SpecifiedText7.Text) > 52) Then
        MsgBox "If the 'Specified' option is selected," & Chr(13) _
                & "you must enter a value between 1 and 52."
        SpecifiedText7.SetFocus
        PileError = True
        Exit Sub
    End If
    If PileSpecified8.Value = True And _
        Val(NumberOfPilesText.Text) >= 8 And _
        (SpecifiedText8.Text = Empty Or _
         Not IsNumeric(SpecifiedText8.Text) Or _
         Val(SpecifiedText8.Text) < 1 Or _
         Val(SpecifiedText8.Text) > 52) Then
        MsgBox "If the 'Specified' option is selected," & Chr(13) _
                & "you must enter a value between 1 and 52."
        SpecifiedText8.SetFocus
        PileError = True
        Exit Sub
    End If
    
'check for error entries when the Random option is selected
'entries must be a numeric entry greater than 0
    If PileRandom1.Value = True And _
        SpecifiedText1.Text <> Empty And _
        Val(NumberOfPilesText.Text) >= 1 And _
        (Not IsNumeric(SpecifiedText1.Text) Or _
        Val(SpecifiedText1.Text) < 1) Then
        MsgBox "If the 'Random' option is selected," & Chr(13) _
                & "you must enter a value greater than 0."
        SpecifiedText1.SetFocus
        PileError = True
        Exit Sub
    End If
    If PileRandom2.Value = True And _
        SpecifiedText2.Text <> Empty And _
        Val(NumberOfPilesText.Text) >= 2 And _
        (Not IsNumeric(SpecifiedText2.Text) Or _
        Val(SpecifiedText2.Text) < 1) Then
        MsgBox "If the 'Random' option is selected," & Chr(13) _
                & "you must enter a value greater than 0."
        SpecifiedText2.SetFocus
        PileError = True
        Exit Sub
    End If
    If PileRandom3.Value = True And _
        SpecifiedText3.Text <> Empty And _
        Val(NumberOfPilesText.Text) >= 3 And _
        (Not IsNumeric(SpecifiedText3.Text) Or _
        Val(SpecifiedText3.Text) < 1) Then
        MsgBox "If the 'Random' option is selected," & Chr(13) _
                & "you must enter a value greater than 0."
        SpecifiedText3.SetFocus
        PileError = True
        Exit Sub
    End If
    If PileRandom4.Value = True And _
        SpecifiedText4.Text <> Empty And _
        Val(NumberOfPilesText.Text) >= 4 And _
        (Not IsNumeric(SpecifiedText4.Text) Or _
        Val(SpecifiedText4.Text) < 1) Then
        MsgBox "If the 'Random' option is selected," & Chr(13) _
                & "you must enter a value greater than 0."
        SpecifiedText4.SetFocus
        PileError = True
        Exit Sub
    End If
    If PileRandom5.Value = True And _
        SpecifiedText5.Text <> Empty And _
        Val(NumberOfPilesText.Text) >= 5 And _
        (Not IsNumeric(SpecifiedText5.Text) Or _
        Val(SpecifiedText5.Text) < 1) Then
        MsgBox "If the 'Random' option is selected," & Chr(13) _
                & "you must enter a value greater than 0."
        SpecifiedText5.SetFocus
        PileError = True
        Exit Sub
    End If
    If PileRandom6.Value = True And _
        SpecifiedText6.Text <> Empty And _
        Val(NumberOfPilesText.Text) >= 6 And _
        (Not IsNumeric(SpecifiedText6.Text) Or _
        Val(SpecifiedText6.Text) < 1) Then
        MsgBox "If the 'Random' option is selected," & Chr(13) _
                & "you must enter a value greater than 0."
        SpecifiedText6.SetFocus
        PileError = True
        Exit Sub
    End If
    If PileRandom7.Value = True And _
        SpecifiedText7.Text <> Empty And _
        Val(NumberOfPilesText.Text) >= 7 And _
        (Not IsNumeric(SpecifiedText7.Text) Or _
        Val(SpecifiedText7.Text) < 1) Then
        MsgBox "If the 'Random' option is selected," & Chr(13) _
                & "you must enter a value greater than 0."
        SpecifiedText7.SetFocus
        PileError = True
        Exit Sub
    End If
    If PileRandom8.Value = True And _
        SpecifiedText8.Text <> Empty And _
        Val(NumberOfPilesText.Text) >= 8 And _
        (Not IsNumeric(SpecifiedText8.Text) Or _
        Val(SpecifiedText8.Text) < 1) Then
        MsgBox "If the 'Random' option is selected," & Chr(13) _
                & "you must enter a value greater than 0."
        SpecifiedText8.SetFocus
        PileError = True
        Exit Sub
    End If
End Sub



Private Sub FinalCardSelectedCheck_Click()
If FinalCardSelectedCheck.Value = 1 Then
    RandomCardSelectedCheck.Value = 0
End If
End Sub

Private Sub GilbreathCheck_Click()
If PilesShown Then
    If GilbreathCheck.Value = 1 Then
        If GilbreathActive Then
            frmDeck.DisplayPilesGilbreath
        End If
    ElseIf GilbreathCheck.Value = 0 Then
        If GilbreathActive Then
            frmDeck.DisplayPilesKeepGilbreathActive
        ElseIf Not GilbreathActive Then
            frmDeck.DisplayPiles
        End If
    End If
End If
End Sub

Private Sub PlaceNewPileSpecifiedText_Change()
PlaceNewPileSpecified.Value = True
End Sub

Private Sub RandomCardSelectedCheck_Click()
If RandomCardSelectedCheck.Value = 1 Then
    FinalCardSelectedCheck.Value = 0
End If
End Sub

Private Sub Form_Activate()
If SessionRecordMode Then
    frmPiles.CreatePilesButton.BackColor = &HFF&
    frmPiles.SpecialButton.BackColor = &HFF&
    frmPiles.SwapPilesButton.BackColor = &HFF&
    frmPiles.AustralianDealButton.BackColor = &HFF&
    frmPiles.CombinePilesButton.BackColor = &HFF&
    frmPiles.RiffleShufflePileButton.BackColor = &HFF&
    frmPiles.SelectReturnButton.BackColor = &HFF&
    frmPiles.CutCardsButton.BackColor = &HFF&
Else
    frmPiles.CreatePilesButton.BackColor = &HFFC0C0
    frmPiles.SpecialButton.BackColor = &H8000000F
    frmPiles.SwapPilesButton.BackColor = &H8000000F
    frmPiles.AustralianDealButton.BackColor = &H8000000F
    frmPiles.CombinePilesButton.BackColor = &H8000000F
    frmPiles.RiffleShufflePileButton.BackColor = &H8000000F
    frmPiles.SelectReturnButton.BackColor = &H8000000F
    frmPiles.CutCardsButton.BackColor = &H8000000F
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMain.mnuPiles.Checked = False
End Sub


Private Sub NumberOfCardsToDealText_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    DealPilesCheck
    KeyAscii = 0
End If
End Sub

Private Sub NumberOfCardsToDealText_Validate(Cancel As Boolean)
DealPilesCheck
End Sub

Private Sub NumberOfPilesText_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    PileError = False
    DealPilesCheck
    If PileError = True Then
        Exit Sub
    End If
    SetNumPilesPlan
    PilesMatrixRefresh
    KeyAscii = 0
End If
End Sub

'Private Sub NumberOfPilesText_Change()
'    PileError = False
'    DealPilesCheck
'    If PileError = True Then
'        Exit Sub
'    End If
'    SetNumPilesPlan
'    PilesMatrixRefresh
'End Sub

Private Sub NumberOfPilesText_Validate(Cancel As Boolean)
If Not PilesShown Then
    PileError = False
    DealPilesCheck
    If PileError = True Then
        Exit Sub
    End If
    SetNumPilesPlan
    PilesMatrixRefresh
End If
End Sub

Private Sub PileMatrix_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
tmp = Index
For i% = 1 To 8
    For j% = 1 To 8
        If (10 * i%) + j% <> tmp Then
            PileMatrix((10 * i%) + j%).Value = 0
        End If
    Next j%
Next i%
For i% = 1 To 8
    If Int(tmp / 10) <> i% Then
        ReverseR(i%).Value = 0
    End If
    If tmp Mod 10 <> i% Then
        ReverseC(i%).Value = 0
    End If
Next i%
End Sub

Private Sub PileMatrixQuery()
For i% = 1 To 8
    For j% = 1 To 8
        If PileMatrix((10 * i%) + j%).Value = 1 Then
            PileMatrixRow = i%
            PileMatrixColumn = j%
        End If
    Next j%
Next i%
End Sub

Private Sub PilesMatrixRefresh()
Dim CorrectNumPiles As Integer
If PilesShown Then
    CorrectNumPiles = NumPiles
Else
    CorrectNumPiles = NumPilesPlan
End If
For i% = 1 To 8
    For j% = 1 To 8
        If i% <= CorrectNumPiles Then
            If j% <= CorrectNumPiles Then
                PileMatrix(10 * i% + j%).Enabled = True
            Else
                PileMatrix(10 * i% + j%).Enabled = False
                PileMatrix(10 * i% + j%).Value = 0
            End If
        Else
            PileMatrix(10 * i% + j%).Enabled = False
            PileMatrix(10 * i% + j%).Value = 0
        End If
    Next j%
Next i%
For i% = 1 To 8
    If i% <= CorrectNumPiles Then
        ReverseR(i%).Enabled = True
        ReverseC(i%).Enabled = True
    Else
        ReverseR(i%).Enabled = False
        ReverseC(i%).Enabled = False
        ReverseR(i%).Value = 0
        ReverseC(i%).Value = 0
    End If
Next i%
'next code for CutCardsFrame
Select Case CorrectNumPiles
    Case 1
        PileRandom1.Enabled = True
        PileSpecified1.Enabled = True
        SpecifiedText1.Enabled = True
        PileMarker1.Enabled = True
        PileRandom2.Enabled = False
        PileSpecified2.Enabled = False
        SpecifiedText2.Enabled = False
        PileMarker2.Enabled = False
        PileRandom3.Enabled = False
        PileSpecified3.Enabled = False
        SpecifiedText3.Enabled = False
        PileMarker3.Enabled = False
        PileRandom4.Enabled = False
        PileSpecified4.Enabled = False
        SpecifiedText4.Enabled = False
        PileMarker4.Enabled = False
        PileRandom5.Enabled = False
        PileSpecified5.Enabled = False
        SpecifiedText5.Enabled = False
        PileMarker5.Enabled = False
        PileRandom6.Enabled = False
        PileSpecified6.Enabled = False
        SpecifiedText6.Enabled = False
        PileMarker6.Enabled = False
        PileRandom7.Enabled = False
        PileSpecified7.Enabled = False
        SpecifiedText7.Enabled = False
        PileMarker7.Enabled = False
        PileRandom8.Enabled = False
        PileSpecified8.Enabled = False
        SpecifiedText8.Enabled = False
        PileMarker8.Enabled = False
    Case 2
        PileRandom1.Enabled = True
        PileSpecified1.Enabled = True
        SpecifiedText1.Enabled = True
        PileMarker1.Enabled = True
        PileRandom2.Enabled = True
        PileSpecified2.Enabled = True
        SpecifiedText2.Enabled = True
        PileMarker2.Enabled = True
        PileRandom3.Enabled = False
        PileSpecified3.Enabled = False
        SpecifiedText3.Enabled = False
        PileMarker3.Enabled = False
        PileRandom4.Enabled = False
        PileSpecified4.Enabled = False
        SpecifiedText4.Enabled = False
        PileMarker4.Enabled = False
        PileRandom5.Enabled = False
        PileSpecified5.Enabled = False
        SpecifiedText5.Enabled = False
        PileMarker5.Enabled = False
        PileRandom6.Enabled = False
        PileSpecified6.Enabled = False
        SpecifiedText6.Enabled = False
        PileMarker6.Enabled = False
        PileRandom7.Enabled = False
        PileSpecified7.Enabled = False
        SpecifiedText7.Enabled = False
        PileMarker7.Enabled = False
        PileRandom8.Enabled = False
        PileSpecified8.Enabled = False
        SpecifiedText8.Enabled = False
        PileMarker8.Enabled = False
    Case 3
        PileRandom1.Enabled = True
        PileSpecified1.Enabled = True
        SpecifiedText1.Enabled = True
        PileMarker1.Enabled = True
        PileRandom2.Enabled = True
        PileSpecified2.Enabled = True
        SpecifiedText2.Enabled = True
        PileMarker2.Enabled = True
        PileRandom3.Enabled = True
        PileSpecified3.Enabled = True
        SpecifiedText3.Enabled = True
        PileMarker3.Enabled = True
        PileRandom4.Enabled = False
        PileSpecified4.Enabled = False
        SpecifiedText4.Enabled = False
        PileMarker4.Enabled = False
        PileRandom5.Enabled = False
        PileSpecified5.Enabled = False
        SpecifiedText5.Enabled = False
        PileMarker5.Enabled = False
        PileRandom6.Enabled = False
        PileSpecified6.Enabled = False
        SpecifiedText6.Enabled = False
        PileMarker6.Enabled = False
        PileRandom7.Enabled = False
        PileSpecified7.Enabled = False
        SpecifiedText7.Enabled = False
        PileMarker7.Enabled = False
        PileRandom8.Enabled = False
        PileSpecified8.Enabled = False
        SpecifiedText8.Enabled = False
        PileMarker8.Enabled = False
    Case 4
        PileRandom1.Enabled = True
        PileSpecified1.Enabled = True
        SpecifiedText1.Enabled = True
        PileMarker1.Enabled = True
        PileRandom2.Enabled = True
        PileSpecified2.Enabled = True
        SpecifiedText2.Enabled = True
        PileMarker2.Enabled = True
        PileRandom3.Enabled = True
        PileSpecified3.Enabled = True
        SpecifiedText3.Enabled = True
        PileMarker3.Enabled = True
        PileRandom4.Enabled = True
        PileSpecified4.Enabled = True
        SpecifiedText4.Enabled = True
        PileMarker4.Enabled = True
        PileRandom5.Enabled = False
        PileSpecified5.Enabled = False
        SpecifiedText5.Enabled = False
        PileMarker5.Enabled = False
        PileRandom6.Enabled = False
        PileSpecified6.Enabled = False
        SpecifiedText6.Enabled = False
        PileMarker6.Enabled = False
        PileRandom7.Enabled = False
        PileSpecified7.Enabled = False
        SpecifiedText7.Enabled = False
        PileMarker7.Enabled = False
        PileRandom8.Enabled = False
        PileSpecified8.Enabled = False
        SpecifiedText8.Enabled = False
        PileMarker8.Enabled = False
    Case 5
        PileRandom1.Enabled = True
        PileSpecified1.Enabled = True
        SpecifiedText1.Enabled = True
        PileMarker1.Enabled = True
        PileRandom2.Enabled = True
        PileSpecified2.Enabled = True
        SpecifiedText2.Enabled = True
        PileMarker2.Enabled = True
        PileRandom3.Enabled = True
        PileSpecified3.Enabled = True
        SpecifiedText3.Enabled = True
        PileMarker3.Enabled = True
        PileRandom4.Enabled = True
        PileSpecified4.Enabled = True
        SpecifiedText4.Enabled = True
        PileMarker4.Enabled = True
        PileRandom5.Enabled = True
        PileSpecified5.Enabled = True
        SpecifiedText5.Enabled = True
        PileMarker5.Enabled = True
        PileRandom6.Enabled = False
        PileSpecified6.Enabled = False
        SpecifiedText6.Enabled = False
        PileMarker6.Enabled = False
        PileRandom7.Enabled = False
        PileSpecified7.Enabled = False
        SpecifiedText7.Enabled = False
        PileMarker7.Enabled = False
        PileRandom8.Enabled = False
        PileSpecified8.Enabled = False
        SpecifiedText8.Enabled = False
        PileMarker8.Enabled = False
    Case 6
        PileRandom1.Enabled = True
        PileSpecified1.Enabled = True
        SpecifiedText1.Enabled = True
        PileMarker1.Enabled = True
        PileRandom2.Enabled = True
        PileSpecified2.Enabled = True
        SpecifiedText2.Enabled = True
        PileMarker2.Enabled = True
        PileRandom3.Enabled = True
        PileSpecified3.Enabled = True
        SpecifiedText3.Enabled = True
        PileMarker3.Enabled = True
        PileRandom4.Enabled = True
        PileSpecified4.Enabled = True
        SpecifiedText4.Enabled = True
        PileMarker4.Enabled = True
        PileRandom5.Enabled = True
        PileSpecified5.Enabled = True
        SpecifiedText5.Enabled = True
        PileMarker5.Enabled = True
        PileRandom6.Enabled = True
        PileSpecified6.Enabled = True
        SpecifiedText6.Enabled = True
        PileMarker6.Enabled = True
        PileRandom7.Enabled = False
        PileSpecified7.Enabled = False
        SpecifiedText7.Enabled = False
        PileMarker7.Enabled = False
        PileRandom8.Enabled = False
        PileSpecified8.Enabled = False
        SpecifiedText8.Enabled = False
        PileMarker8.Enabled = False
    Case 7
        PileRandom1.Enabled = True
        PileSpecified1.Enabled = True
        SpecifiedText1.Enabled = True
        PileMarker1.Enabled = True
        PileRandom2.Enabled = True
        PileSpecified2.Enabled = True
        SpecifiedText2.Enabled = True
        PileMarker2.Enabled = True
        PileRandom3.Enabled = True
        PileSpecified3.Enabled = True
        SpecifiedText3.Enabled = True
        PileMarker3.Enabled = True
        PileRandom4.Enabled = True
        PileSpecified4.Enabled = True
        SpecifiedText4.Enabled = True
        PileMarker4.Enabled = True
        PileRandom5.Enabled = True
        PileSpecified5.Enabled = True
        SpecifiedText5.Enabled = True
        PileMarker5.Enabled = True
        PileRandom6.Enabled = True
        PileSpecified6.Enabled = True
        SpecifiedText6.Enabled = True
        PileMarker6.Enabled = True
        PileRandom7.Enabled = True
        PileSpecified7.Enabled = True
        SpecifiedText7.Enabled = True
        PileMarker7.Enabled = True
        PileRandom8.Enabled = False
        PileSpecified8.Enabled = False
        SpecifiedText8.Enabled = False
        PileMarker8.Enabled = False
    Case 8
        PileRandom1.Enabled = True
        PileSpecified1.Enabled = True
        SpecifiedText1.Enabled = True
        PileMarker1.Enabled = True
        PileRandom2.Enabled = True
        PileSpecified2.Enabled = True
        SpecifiedText2.Enabled = True
        PileMarker2.Enabled = True
        PileRandom3.Enabled = True
        PileSpecified3.Enabled = True
        SpecifiedText3.Enabled = True
        PileMarker3.Enabled = True
        PileRandom4.Enabled = True
        PileSpecified4.Enabled = True
        SpecifiedText4.Enabled = True
        PileMarker4.Enabled = True
        PileRandom5.Enabled = True
        PileSpecified5.Enabled = True
        SpecifiedText5.Enabled = True
        PileMarker5.Enabled = True
        PileRandom6.Enabled = True
        PileSpecified6.Enabled = True
        SpecifiedText6.Enabled = True
        PileMarker6.Enabled = True
        PileRandom7.Enabled = True
        PileSpecified7.Enabled = True
        SpecifiedText7.Enabled = True
        PileMarker7.Enabled = True
        PileRandom8.Enabled = True
        PileSpecified8.Enabled = True
        SpecifiedText8.Enabled = True
        PileMarker8.Enabled = True
End Select
End Sub


Private Sub PileRandom1_Click()
PileMarker1.Caption = "A"
End Sub
Private Sub PileRandom2_Click()
PileMarker2.Caption = "A"
End Sub
Private Sub PileRandom3_Click()
PileMarker3.Caption = "A"
End Sub
Private Sub PileRandom4_Click()
PileMarker4.Caption = "A"
End Sub
Private Sub PileRandom5_Click()
PileMarker5.Caption = "A"
End Sub
Private Sub PileRandom6_Click()
PileMarker6.Caption = "A"
End Sub
Private Sub PileRandom7_Click()
PileMarker7.Caption = "A"
End Sub
Private Sub PileRandom8_Click()
PileMarker8.Caption = "A"
End Sub


Private Sub PileSpecified1_Click()
PileMarker1.Caption = "E"
End Sub
Private Sub PileSpecified2_Click()
PileMarker2.Caption = "E"
End Sub
Private Sub PileSpecified3_Click()
PileMarker3.Caption = "E"
End Sub
Private Sub PileSpecified4_Click()
PileMarker4.Caption = "E"
End Sub
Private Sub PileSpecified5_Click()
PileMarker5.Caption = "E"
End Sub
Private Sub PileSpecified6_Click()
PileMarker6.Caption = "E"
End Sub
Private Sub PileSpecified7_Click()
PileMarker7.Caption = "E"
End Sub
Private Sub PileSpecified8_Click()
PileMarker8.Caption = "E"
End Sub

Private Sub RefreshDeckPiles_Click()
NumPilesPlan = 8
Call frmStackView.ShowCards
PilesMatrixRefresh
End Sub

Private Sub ReturnPileNewPileSpecifiedText_Change()
ReturnPileNewPileSpecified.Value = True
End Sub

Private Sub ReturnPositionSpecifiedText_Change()
ReturnPositionSpecified.Value = True
End Sub

Private Sub ReverseC_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
tmp = Index
Dim tmpCount As Integer
tmpCount = 0
'first turn off other checks
For i% = 1 To 8
    If i% <> tmp Then
        ReverseC(i%).Value = 0
    End If
Next i%
'second, verify that the column is legitimate by
'checking the pile matrix
For i% = 1 To 8
    If PileMatrix((10 * i%) + tmp).Value = 1 Then
        tmpCount = tmpCount + 1
    End If
Next i%
If tmpCount = 0 Then
    ReverseC(tmp).Value = 0
End If
End Sub

Private Sub ReverseDownAllCheck_Click()
If ReverseDownAllCheck.Value = 1 Then
    ReverseDownRandomCheck.Value = 0
End If
End Sub


Private Sub ReverseDownRandomCheck_Click()
If ReverseDownRandomCheck.Value = 1 Then
    ReverseDownAllCheck.Value = 0
End If
End Sub

Private Sub ReverseR_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
tmp = Index
Dim tmpCount As Integer
tmpCount = 0
'first turn off other checks
For i% = 1 To 8
    If i% <> tmp Then
        ReverseR(i%).Value = 0
    End If
Next i%
'second, verify that the row is legitimate by
'checking the pile matrix
For i% = 1 To 8
    If PileMatrix((10 * tmp) + i%).Value = 1 Then
        tmpCount = tmpCount + 1
    End If
Next i%
If tmpCount = 0 Then
    ReverseR(tmp).Value = 0
End If
End Sub

Public Sub RiffleShufflePile(pileNum1Param, pileNum2Param, riffleshufflenumber, pileNum4Param)
    Dim maxPile As Integer
    Dim pileGap As Integer
    Dim pileAdjust As Integer
    Dim blockReverseSize As Integer
    Dim pileNum1 As Integer
    Dim pileNum2 As Integer
    Dim pileLeft As String
    Dim pileRight As String
    'pileLeft is set to T for Top or B for Bottom
    '   if the pileNum1 has a protected block, else Empty
    'pileRight is set to T for Top or B for Bottom
    '   if the pileNum2 has a protected block, else Empty
    Dim pileLeftReverse As String
    Dim pileRightReverse As String
    'pileLeftReverse is set to R if the block needs to
    'be reversed before the rest of the operation
    'the same is true for pileRightReverse
    '
    'set the Gilbreath parameters
    For i% = 1 To 52
        GilbreathDeck(i%) = False
        GilbreathStatus(i%) = False
    Next i%
    'decode the first paramater values
    If Len(pileNum1Param) = 1 Then
        pileLeft = Empty
        pileLeftReverse = Empty
        pileNum1 = pileNum1Param
    ElseIf Len(pileNum1Param) = 2 Then
        If IsNumeric(Left(pileNum1Param, 1)) Then
            pileNum1 = Val(Left(pileNum1Param, 1))
            pileLeftReverse = Right(pileNum1Param, 1)
            pileLeft = Empty
        Else
            pileNum1 = Val(Right(pileNum1Param, 1))
            pileLeftReverse = Empty
            pileLeft = Left(pileNum1Param, 1)
        End If
    ElseIf Len(pileNum1Param) = 3 Then
        pileLeft = Left(pileNum1Param, 1)
        pileLeftReverse = Right(pileNum1Param, 1)
        pileNum1 = Val(Mid(pileNum1Param, 2, 1))
    End If
    'decode the second paramater values
    If Len(pileNum2Param) = 1 Then
        pileRight = Empty
        pileRightReverse = Empty
        pileNum2 = pileNum2Param
    ElseIf Len(pileNum2Param) = 2 Then
        If IsNumeric(Left(pileNum2Param, 1)) Then
            pileNum2 = Val(Left(pileNum2Param, 1))
            pileRightReverse = Right(pileNum2Param, 1)
            pileRight = Empty
        Else
            pileNum2 = Val(Right(pileNum2Param, 1))
            pileRightReverse = Empty
            pileRight = Left(pileNum2Param, 1)
        End If
    ElseIf Len(pileNum2Param) = 3 Then
        pileRight = Left(pileNum2Param, 1)
        pileRightReverse = Right(pileNum2Param, 1)
        pileNum2 = Val(Mid(pileNum2Param, 2, 1))
    End If
    'set pile order info
    If pileNum1 > pileNum2 Then
        maxPile = pileNum1
    Else
        maxPile = pileNum2
    End If
    'reverse the appropriate piles if necessary
    If pileLeftReverse = "R" Then
        blockReverseSize = PileTable(pileNum1, 2) - PileTable(pileNum1, 1) + 1
        For m% = 1 To blockReverseSize
            For z% = 1 To DeckProperties
                ChangedDeck(z%, PileTable(pileNum1, 2) - m% + 1) = _
                    Deck(z%, PileTable(pileNum1, 1) + m% - 1)
            Next z%
        Next m%
        For m% = 1 To blockReverseSize
            For z% = 1 To DeckProperties
                Deck(z%, PileTable(pileNum1, 1) + m% - 1) = _
                    ChangedDeck(z%, PileTable(pileNum1, 1) + m% - 1)
            Next z%
        Next m%
        For i% = 1 To blockReverseSize
            Deck(6, PileTable(pileNum1, 1) + i% - 1) = _
                Not Deck(6, PileTable(pileNum1, 1) + i% - 1)
        Next i%
    End If
    If pileRightReverse = "R" Then
        blockReverseSize = PileTable(pileNum2, 2) - PileTable(pileNum2, 1) + 1
        For m% = 1 To blockReverseSize
            For z% = 1 To DeckProperties
                ChangedDeck(z%, PileTable(pileNum2, 2) - m% + 1) = _
                    Deck(z%, PileTable(pileNum2, 1) + m% - 1)
            Next z%
        Next m%
        For m% = 1 To blockReverseSize
            For z% = 1 To DeckProperties
                Deck(z%, PileTable(pileNum2, 1) + m% - 1) = _
                    ChangedDeck(z%, PileTable(pileNum2, 1) + m% - 1)
            Next z%
        Next m%
        For i% = 1 To blockReverseSize
            Deck(6, PileTable(pileNum2, 1) + i% - 1) = _
                Not Deck(6, PileTable(pileNum2, 1) + i% - 1)
        Next i%
    End If
    'transfer the deck to the PileDeck array
    For m% = 1 To DeckCount
        For z% = 1 To DeckProperties
            ChangedPileDeck(z%, m%) = Deck(z%, m%)
        Next z%
    Next m%
    ProtectedBlock = riffleshufflenumber
If pileLeft = "T" Or pileRight = "T" Or _
        (pileLeft = Empty And pileRight = Empty) Then
    If pileNum1 = pileNum2 Then
        'when pileNum1=pileNum2, "T" will always be assigned to pileNum1
        'since it really doesn't matter when there is only one pile being shuffled
        RifflePortion = PileTable(pileNum1, 2) - PileTable(pileNum1, 1) + 1 - ProtectedBlock
        CutError = Int(0.2 * RifflePortion) + 1
        CutDepth = Int(Rnd * CutError) + Int((RifflePortion - CutError) / 2)
          'approximately half the unprotected deck for a riffle shuffle
        If CutDepth < 0 Then
            CutDepth = 0
                ' need to avoid an error
        End If
        RemainingCut = RifflePortion - CutDepth
        ReDim TopCut(DeckProperties, CutDepth)
        ReDim BottomCut(DeckProperties, RemainingCut)
        For i% = 1 To CutDepth
            For z% = 1 To DeckProperties
                TopCut(z%, i%) = Deck(z%, PileTable(pileNum1, 1) + ProtectedBlock - 1 + i%)
            Next z%
        Next i%
        For j% = 1 To RemainingCut
            For z% = 1 To DeckProperties
                BottomCut(z%, j%) = Deck(z%, PileTable(pileNum1, 1) + _
                    ProtectedBlock - 1 + CutDepth + j%)
            Next z%
        Next j%
        If ProtectedBlock > 0 Then
            For p% = 1 To ProtectedBlock
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, p%) = Deck(z%, PileTable(pileNum1, 1) - 1 + p%)
                Next z%
                GilbreathStatus(p%) = False
            Next p%
        End If
    Else
        'here we need to handle which pile has the protected block
        RifflePortion = PileTable(pileNum1, 2) - PileTable(pileNum1, 1) + 1 + _
            PileTable(pileNum2, 2) - PileTable(pileNum2, 1) + 1 - ProtectedBlock
        If pileLeft = "T" Then
            CutDepth = PileTable(pileNum1, 2) - PileTable(pileNum1, 1) + 1 - ProtectedBlock
            RemainingCut = RifflePortion - CutDepth
        ElseIf pileRight = "T" Then
            CutDepth = PileTable(pileNum1, 2) - PileTable(pileNum1, 1) + 1
            RemainingCut = RifflePortion - CutDepth
        ElseIf ProtectedBlock = 0 Then
            CutDepth = PileTable(pileNum1, 2) - PileTable(pileNum1, 1) + 1
            RemainingCut = RifflePortion - CutDepth
        End If
        ReDim TopCut(DeckProperties, CutDepth)
        ReDim BottomCut(DeckProperties, RemainingCut)
        For i% = 1 To CutDepth
            For z% = 1 To DeckProperties
                If pileLeft = "T" Then
                    TopCut(z%, i%) = Deck(z%, PileTable(pileNum1, 1) + ProtectedBlock - 1 + i%)
                Else
                    TopCut(z%, i%) = Deck(z%, PileTable(pileNum1, 1) - 1 + i%)
                End If
            Next z%
        Next i%
        For j% = 1 To RemainingCut
            For z% = 1 To DeckProperties
                If pileRight = "T" Then
                    BottomCut(z%, j%) = Deck(z%, PileTable(pileNum2, 1) + ProtectedBlock - 1 + j%)
                Else
                    BottomCut(z%, j%) = Deck(z%, PileTable(pileNum2, 1) - 1 + j%)
                End If
            Next z%
        Next j%
        If ProtectedBlock > 0 Then
            For p% = 1 To ProtectedBlock
                For z% = 1 To DeckProperties
                    If pileLeft = "T" Then
                        ChangedDeck(z%, p%) = Deck(z%, PileTable(pileNum1, 1) - 1 + p%)
                    Else
                        ChangedDeck(z%, p%) = Deck(z%, PileTable(pileNum2, 1) - 1 + p%)
                    End If
                Next z%
                If pileLeft = "T" Then
                    GilbreathStatus(p%) = False
                Else
                    GilbreathStatus(p%) = True
                End If
            Next p%
        End If
    End If
    TopIndex = 1
    BottomIndex = 1
    For k% = 1 To RifflePortion
        side = Rnd
        'when low, shuffle from TopCut
        'when high, shuffle from BottomCut
        If side < CutDepth / RifflePortion Then
          'compare Rnd with ratio of cuts to change odds
          'to a more even mixing of the talons
            If TopIndex <= CutDepth Then
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, ProtectedBlock + k%) = TopCut(z%, TopIndex)
                Next z%
                GilbreathStatus(ProtectedBlock + k%) = False
                TopIndex = TopIndex + 1
            Else
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, ProtectedBlock + k%) = BottomCut(z%, BottomIndex)
                Next z%
                GilbreathStatus(ProtectedBlock + k%) = True
                BottomIndex = BottomIndex + 1
            End If
        Else
            If BottomIndex <= RemainingCut Then
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, ProtectedBlock + k%) = BottomCut(z%, BottomIndex)
                Next z%
                GilbreathStatus(ProtectedBlock + k%) = True
                BottomIndex = BottomIndex + 1
            Else
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, ProtectedBlock + k%) = TopCut(z%, TopIndex)
                Next z%
                GilbreathStatus(ProtectedBlock + k%) = False
                TopIndex = TopIndex + 1
            End If
        End If
    Next k%
    'arrange the piles back into a deck
    If pileNum1 = pileNum2 Then
        For m% = 1 To RifflePortion + ProtectedBlock
            For z% = 1 To DeckProperties
                ChangedPileDeck(z%, PileTable(pileNum1, 1) - 1 + m%) = ChangedDeck(z%, m%)
            Next z%
            GilbreathDeck(PileTable(pileNum1, 1) - 1 + m%) = GilbreathStatus(m%)
        Next m%
        GilbreathPileNum = pileNum1
    Else
        pileGap = 0
        pileAdjust = 0
        For i% = 1 To NumPiles
            If i% > maxPile Then
                pileGap = 0
            End If
            If i% = pileNum2 Then
                pileAdjust = -1
                If pileNum2 < pileNum1 Then
                    pileGap = PileTable(pileNum2, 2) - PileTable(pileNum2, 1) + 1
                End If
            ElseIf i% = pileNum1 Then
                For m% = 1 To RifflePortion + ProtectedBlock
                    For z% = 1 To DeckProperties
                        ChangedPileDeck(z%, PileTable(pileNum1, 1) - 1 + m% - pileGap) = _
                            ChangedDeck(z%, m%)
                    Next z%
                    GilbreathDeck(PileTable(pileNum1, 1) - 1 + m% - pileGap) = _
                        GilbreathStatus(m%)
                Next m%
                ChangedPileTable(i% + pileAdjust, 1) = PileTable(i%, 1) - pileGap
                ChangedPileTable(i% + pileAdjust, 2) = ChangedPileTable(i% + pileAdjust, 1) + _
                    RifflePortion + ProtectedBlock - 1
                'adjust pileGap
                If pileNum2 < pileNum1 Then
                    pileGap = 0
                Else
                    pileGap = -(PileTable(pileNum2, 2) - PileTable(pileNum2, 1) + 1)
                End If
                GilbreathPileNum = pileNum1 + pileAdjust
            Else
                For m% = 1 To PileTable(i%, 2) - PileTable(i%, 1) + 1
                    For z% = 1 To DeckProperties
                        ChangedPileDeck(z%, PileTable(i%, 1) - 1 + m% - pileGap) = _
                            Deck(z%, PileTable(i%, 1) - 1 + m%)
                    Next z%
                Next m%
                ChangedPileTable(i% + pileAdjust, 1) = PileTable(i%, 1) - pileGap
                ChangedPileTable(i% + pileAdjust, 2) = PileTable(i%, 2) - pileGap
            End If
        Next i%
        'adjust for the fact that two piles have become one
        NumPiles = NumPiles - 1
        For i% = 1 To NumPiles
            PileTable(i%, 1) = ChangedPileTable(i%, 1)
            PileTable(i%, 2) = ChangedPileTable(i%, 2)
        Next i%
    End If
    For m% = 1 To DeckCount
        For z% = 1 To DeckProperties
            Deck(z%, m%) = ChangedPileDeck(z%, m%)
        Next z%
    Next m%
    If pileNum4Param = "G" Then
        CreatePilesGilbreath (NumPiles)
    Else
        CreatePiles (NumPiles)
    End If
    PilesMatrixRefresh
ElseIf pileLeft = "B" Or pileRight = "B" Then
    If pileNum1 = pileNum2 Then
        'when pileNum1=pileNum2, "B" will always be assigned to pileNum1
        'since it really doesn't matter when there is only one pile being shuffled
        RifflePortion = PileTable(pileNum1, 2) - PileTable(pileNum1, 1) + 1 - ProtectedBlock
        CutError = Int(0.2 * RifflePortion) + 1
        CutDepth = Int(Rnd * CutError) + Int((RifflePortion - CutError) / 2)
          'approximately half the unprotected deck for a riffle shuffle
        If CutDepth < 0 Then
            CutDepth = 0
                ' need to avoid an error
        End If
        RemainingCut = RifflePortion - CutDepth
        ReDim TopCut(DeckProperties, CutDepth)
        ReDim BottomCut(DeckProperties, RemainingCut)
        For i% = 1 To CutDepth
            For z% = 1 To DeckProperties
                TopCut(z%, i%) = Deck(z%, PileTable(pileNum1, 1) - 1 + i%)
            Next z%
        Next i%
        For j% = 1 To RemainingCut
            For z% = 1 To DeckProperties
                BottomCut(z%, j%) = Deck(z%, PileTable(pileNum1, 1) _
                    - 1 + CutDepth + j%)
            Next z%
        Next j%
        If ProtectedBlock > 0 Then
            For p% = 1 To ProtectedBlock
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, RifflePortion + p%) = _
                        Deck(z%, PileTable(pileNum1, 1) + RifflePortion - 1 + p%)
                Next z%
                GilbreathStatus(RifflePortion + p%) = True
            Next p%
        End If
    Else
        'here we need to handle which pile has the protected block
        RifflePortion = PileTable(pileNum1, 2) - PileTable(pileNum1, 1) + 1 + _
            PileTable(pileNum2, 2) - PileTable(pileNum2, 1) + 1 - ProtectedBlock
        If pileLeft = "B" Then
            CutDepth = PileTable(pileNum1, 2) - PileTable(pileNum1, 1) + 1 - ProtectedBlock
            RemainingCut = RifflePortion - CutDepth
        ElseIf pileRight = "B" Then
            CutDepth = PileTable(pileNum1, 2) - PileTable(pileNum1, 1) + 1
            RemainingCut = RifflePortion - CutDepth
        ElseIf ProtectedBlock = 0 Then
            CutDepth = PileTable(pileNum1, 2) - PileTable(pileNum1, 1) + 1
            RemainingCut = RifflePortion - CutDepth
        End If
        ReDim TopCut(DeckProperties, CutDepth)
        ReDim BottomCut(DeckProperties, RemainingCut)
        For i% = 1 To CutDepth
            For z% = 1 To DeckProperties
                TopCut(z%, i%) = Deck(z%, PileTable(pileNum1, 1) - 1 + i%)
            Next z%
        Next i%
        For j% = 1 To RemainingCut
            For z% = 1 To DeckProperties
                BottomCut(z%, j%) = Deck(z%, PileTable(pileNum2, 1) - 1 + j%)
            Next z%
        Next j%
        If ProtectedBlock > 0 Then
            For p% = 1 To ProtectedBlock
                For z% = 1 To DeckProperties
                    If pileLeft = "B" Then
                        ChangedDeck(z%, RifflePortion + p%) = _
                            Deck(z%, PileTable(pileNum1, 1) + CutDepth - 1 + p%)
                    Else
                        ChangedDeck(z%, RifflePortion + p%) = _
                            Deck(z%, PileTable(pileNum2, 1) + RemainingCut - 1 + p%)
                    End If
                Next z%
                If pileLeft = "B" Then
                    GilbreathStatus(RifflePortion + p%) = False
                Else
                    GilbreathStatus(RifflePortion + p%) = True
                End If
            Next p%
        End If
    End If
    TopIndex = 1
    BottomIndex = 1
    For k% = 1 To RifflePortion
        side = Rnd
        'when low, shuffle from TopCut
        'when high, shuffle from BottomCut
        If side < CutDepth / RifflePortion Then
          'compare Rnd with ratio of cuts to change odds
          'to a more even mixing of the talons
            If TopIndex <= CutDepth Then
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, k%) = TopCut(z%, TopIndex)
                Next z%
                GilbreathStatus(k%) = False
                TopIndex = TopIndex + 1
            Else
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, k%) = BottomCut(z%, BottomIndex)
                Next z%
                GilbreathStatus(k%) = True
                BottomIndex = BottomIndex + 1
            End If
        Else
            If BottomIndex <= RemainingCut Then
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, k%) = BottomCut(z%, BottomIndex)
                Next z%
                GilbreathStatus(k%) = True
                BottomIndex = BottomIndex + 1
            Else
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, k%) = TopCut(z%, TopIndex)
                Next z%
                GilbreathStatus(k%) = False
                TopIndex = TopIndex + 1
            End If
        End If
    Next k%
    'arrange the piles back into a deck
    If pileNum1 = pileNum2 Then
        For m% = 1 To RifflePortion + ProtectedBlock
            For z% = 1 To DeckProperties
                ChangedPileDeck(z%, PileTable(pileNum1, 1) - 1 + m%) = ChangedDeck(z%, m%)
            Next z%
            GilbreathDeck(PileTable(pileNum1, 1) - 1 + m%) = GilbreathStatus(m%)
        Next m%
        GilbreathPileNum = pileNum1
    Else
        pileGap = 0
        pileAdjust = 0
        For i% = 1 To NumPiles
            If i% > maxPile Then
                pileGap = 0
            End If
            If i% = pileNum2 Then
                pileAdjust = -1
                If pileNum2 < pileNum1 Then
                    pileGap = PileTable(pileNum2, 2) - PileTable(pileNum2, 1) + 1
                End If
            ElseIf i% = pileNum1 Then
                For m% = 1 To RifflePortion + ProtectedBlock
                    For z% = 1 To DeckProperties
                        ChangedPileDeck(z%, PileTable(pileNum1, 1) - 1 + m% - pileGap) = _
                            ChangedDeck(z%, m%)
                    Next z%
                    GilbreathDeck(PileTable(pileNum1, 1) - 1 + m% - pileGap) = _
                        GilbreathStatus(m%)
                Next m%
                ChangedPileTable(i% + pileAdjust, 1) = PileTable(i%, 1) - pileGap
                ChangedPileTable(i% + pileAdjust, 2) = ChangedPileTable(i% + pileAdjust, 1) + _
                    RifflePortion + ProtectedBlock - 1
                'adjust pileGap
                If pileNum2 < pileNum1 Then
                    pileGap = 0
                Else
                    pileGap = -(PileTable(pileNum2, 2) - PileTable(pileNum2, 1) + 1)
                End If
                GilbreathPileNum = pileNum1 + pileAdjust
            Else
                For m% = 1 To PileTable(i%, 2) - PileTable(i%, 1) + 1
                    For z% = 1 To DeckProperties
                        ChangedPileDeck(z%, PileTable(i%, 1) - 1 + m% - pileGap) = _
                            Deck(z%, PileTable(i%, 1) - 1 + m%)
                    Next z%
                Next m%
                ChangedPileTable(i% + pileAdjust, 1) = PileTable(i%, 1) - pileGap
                ChangedPileTable(i% + pileAdjust, 2) = PileTable(i%, 2) - pileGap
            End If
        Next i%
        'adjust for the fact that two piles have become one
        NumPiles = NumPiles - 1
        For i% = 1 To NumPiles
            PileTable(i%, 1) = ChangedPileTable(i%, 1)
            PileTable(i%, 2) = ChangedPileTable(i%, 2)
        Next i%
    End If
    For m% = 1 To DeckCount
        For z% = 1 To DeckProperties
            Deck(z%, m%) = ChangedPileDeck(z%, m%)
        Next z%
    Next m%
    If pileNum4Param = "G" Then
        CreatePilesGilbreath (NumPiles)
    Else
        CreatePiles (NumPiles)
    End If
    PilesMatrixRefresh
End If
End Sub


Private Sub ReverseUnderAllCheck_Click()
If ReverseUnderAllCheck.Value = 1 Then
    ReverseUnderRandomCheck.Value = 0
End If
End Sub

Private Sub ReverseUnderRandomCheck_Click()
If ReverseUnderRandomCheck.Value = 1 Then
    ReverseUnderAllCheck.Value = 0
End If
End Sub

Private Sub RiffleShufflePileButton_Click()
    'make sure piles have been created
    If PilesShown = 0 Then
        MsgBox ("You must create piles before you can manipulate piles.")
        Exit Sub
    End If
    ' make sure the PileMatrix has a valid selection
    Dim PileMatrixCheckSum As Integer
    PileMatrixCheckSum = 0
    For i% = 1 To 8
        For j% = 1 To 8
            If PileMatrix(10 * i% + j%).Value = 1 Then
                PileMatrixCheckSum = PileMatrixCheckSum + 1
            End If
        Next j%
    Next i%
    If PileMatrixCheckSum = 0 Then
        MsgBox ("You must check a box in the Pile Control Matrix" & Chr(13) & _
            "before you can manipulate piles.")
        Exit Sub
    End If
    'identify the piles to manipulate
    Dim param1 As String
    Dim param2 As String
    Dim param4 As String
    Call PileMatrixQuery
    'check for logical protected block size
    If RiffleShufflePilesProtect.Value = True Then
        'check for numeric data in text box
        If RiffleShuffleProtectCards.Text = Empty Then
            RiffleShuffleProtectCards.Text = "0"
        End If
        If Not IsNumeric(RiffleShuffleProtectCards.Text) Then
            RiffleShuffleProtectCards.Text = Empty
            RiffleShuffleProtectCards.SetFocus
            MsgBox ("You can only enter numbers in the Protect Cards text box.")
            Exit Sub
        End If
        If Val(RiffleShuffleProtectCards.Text) < 1 Then
            RiffleShuffleProtectCards.Text = Empty
            RiffleShuffleProtectCards.SetFocus
            MsgBox ("You can only enter numbers greater than 0" & _
                Chr(13) & "in the Protect Cards text box.")
            Exit Sub
        End If
        If RiffleShuffleProtectPrimary.Value = True Then
            If Val(RiffleShuffleProtectCards) > _
                PileTable(PileMatrixRow, 2) - PileTable(PileMatrixRow, 1) + 1 Then
                MsgBox ("You can not protect " & RiffleShuffleProtectCards.Text & " cards" & _
                    Chr(13) & "since there are only " & _
                    PileTable(PileMatrixRow, 2) - PileTable(PileMatrixRow, 1) + 1 & " cards" & _
                    Chr(13) & "in Pile " & PileMatrixRow)
                Exit Sub
            End If
        End If
        If RiffleShuffleProtectSecondary.Value = True Then
            If Val(RiffleShuffleProtectCards) > _
                PileTable(PileMatrixColumn, 2) - PileTable(PileMatrixColumn, 1) + 1 Then
                MsgBox ("You can not protect " & RiffleShuffleProtectCards.Text & " cards" & _
                    Chr(13) & "since there are only " & _
                    PileTable(PileMatrixColumn, 2) - PileTable(PileMatrixColumn, 1) + 1 & " cards" & _
                    Chr(13) & "in Pile " & PileMatrixColumn)
                Exit Sub
            End If
        End If
    ElseIf RiffleShufflePilesRandom.Value = True Then
        RiffleShuffleProtectCards.Text = "0"
    End If
    'set parameters to empty starting values
    param1 = ""
    param2 = ""
    'set protected block parameters
    If RiffleShufflePilesProtect.Value = True Then
        If RiffleShuffleProtectPrimary.Value = True Then
            If RiffleShuffleProtectTop.Value = True Then
                param1 = param1 & "T"
            ElseIf RiffleShuffleProtectBottom.Value = True Then
                param1 = param1 & "B"
            End If
        ElseIf RiffleShuffleProtectSecondary.Value = True Then
            If RiffleShuffleProtectTop.Value = True Then
                param2 = param2 & "T"
            ElseIf RiffleShuffleProtectBottom.Value = True Then
                param2 = param2 & "B"
            End If
        End If
    End If
    'set parameter pile numbers
    param1 = param1 & PileMatrixRow
    param2 = param2 & PileMatrixColumn
    'set parameter reverse condition
    If ReverseR(PileMatrixRow).Value = 1 Then
        param1 = param1 & "R"
    End If
    If ReverseC(PileMatrixColumn).Value = 1 Then
        param2 = param2 & "R"
    End If
    'set the Gilbreath View parameter
    If GilbreathCheck.Value = 1 Then
        param4 = "G"
    Else
        param4 = "X"
    End If
    'now execute the riffle shuffle
    Call RiffleShufflePile(param1, param2, Val(RiffleShuffleProtectCards), param4)
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "RiffleShufflePile(" & param1 _
        & ", " & param2 _
        & ", " & Val(RiffleShuffleProtectCards) _
        & ", " & param4 & ")"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End Sub


Private Sub RiffleShufflePilesProtect_Click()
    RiffleShuffleProtectTop.Enabled = True
    RiffleShuffleProtectBottom.Enabled = True
    RiffleShuffleProtectCards.Enabled = True
    RiffleShuffleProtectPrimary.Enabled = True
    RiffleShuffleProtectSecondary.Enabled = True
    RiffleShuffleProtectedCardsLabel.Enabled = True
End Sub

Private Sub RiffleShufflePilesRandom_Click()
    RiffleShuffleProtectTop.Enabled = False
    RiffleShuffleProtectBottom.Enabled = False
    RiffleShuffleProtectCards.Enabled = False
    RiffleShuffleProtectPrimary.Enabled = False
    RiffleShuffleProtectSecondary.Enabled = False
    RiffleShuffleProtectedCardsLabel.Enabled = False
End Sub

Public Sub SelectReturn(param1, param2, param3, param4, param5)
    'establish working variables for this module
    Dim pSelectedPile As Integer
    'this is the number of the selected pile
    Dim pSelectedPileCards As Integer
    'this is the number of cards in the selected pile
    Dim pSelectedPosition As Integer
    'this is the position from the top of the selected pile that
    'the selected card is located
    Dim pReturnPile As Integer
    'this is the pile number of the return pile
    Dim pReturnPileCards As Integer
    'this is the number of cards in the return pile
    Dim pReturnPosition As Integer
    'this is the position in the return pile that the selected
    'card will be placed
    Dim pSuffix As String
    'st, nd, rd, th suffixes for an error message
    Dim pBackCounter As Integer
    Dim pNewPileCreated As Integer
    Dim fromCardParam As Integer
    Dim toCardParam As Integer
    '---
    'decode first parameter
    If Left(param1, 1) = "P" Then
        pSelectedPile = Val(Right(param1, Len(param1) - 1))
        If pSelectedPile > NumPiles Then
            pSelectedPile = NumPiles
        End If
    ElseIf Left(param1, 1) = "R" Then
        pSelectedPile = Int(Rnd * NumPiles) + 1
    End If
    pSelectedPileCards = PileTable(pSelectedPile, 2) - PileTable(pSelectedPile, 1) + 1
    'decode second parameter
    If Left(param2, 1) = "T" Then
        pSelectedPosition = 1
    ElseIf Left(param2, 1) = "B" Then
        pSelectedPosition = pSelectedPileCards
    ElseIf Left(param2, 1) = "R" Then
        pSelectedPosition = Int(Rnd * pSelectedPileCards) + 1
    ElseIf Left(param2, 1) = "S" Then
        If Right(param2, 1) = "R" Then
            pSelectedPosition = Val(Mid(param2, 2, Len(param2) - 2))
        Else
            pSelectedPosition = Val(Mid(param2, 2, Len(param2) - 1))
        End If
        If pSelectedPosition > pSelectedPileCards Then
            If pSelectedPosition = 1 Then
                pSuffix = "st"
            ElseIf pSelectedPosition = 2 Then
                pSuffix = "nd"
            ElseIf pSelectedPosition = 3 Then
                pSuffix = "rd"
            ElseIf pSelectedPosition > 3 Then
                pSuffix = "th"
            End If
            If pSelectedPileCards = 1 Then
                MsgBox ("You can not specify the " & pSelectedPosition & _
                    pSuffix & " card" & _
                    Chr(13) & "since there is only " & _
                    pSelectedPileCards & " card" & _
                    Chr(13) & "in Pile " & pSelectedPile)
            Else
                MsgBox ("You can not specify the " & pSelectedPosition & _
                    pSuffix & " card" & _
                    Chr(13) & "since there are only " & _
                    pSelectedPileCards & " cards" & _
                    Chr(13) & "in Pile " & pSelectedPile)
            End If
            Exit Sub
        End If
    End If
    If Len(param2) > 1 And Right(param2, 1) = "R" Then
        Deck(6, PileTable(pSelectedPile, 1) + pSelectedPosition - 1) = _
            Not Deck(6, PileTable(pSelectedPile, 1) + pSelectedPosition - 1)
    End If
    'decode third parameter
    If Left(param3, 1) = "S" Then
        pReturnPile = Val(Right(param3, Len(param3) - 1))
        If pReturnPile > NumPiles Then
            pReturnPile = NumPiles
        End If
    ElseIf Left(param3, 1) = "P" Then
        pReturnPile = pSelectedPile
    ElseIf Left(param3, 1) = "E" Then
        pReturnPile = pSelectedPile
    ElseIf Left(param3, 1) = "R" Then
        pReturnPile = Int(Rnd * NumPiles) + 1
    ElseIf Left(param3, 1) = "D" Then
        'before establishing a different pile, make sure
        'there is more than 1 pile present
        If NumPiles < 2 Then
            MsgBox ("Error: When specifying 'Random Not Same'," & Chr(13) & _
                "there must be more than one pile present.")
            Exit Sub
        Else
            'first set the condition where the selected and return piles are the same
            pReturnPile = pSelectedPile
            'now run a While loop until the pile values are different
            While pReturnPile = pSelectedPile
                pReturnPile = Int(Rnd * NumPiles) + 1
            Wend
        End If
    ElseIf Left(param3, 1) = "N" Or Left(param3, 1) = "L" Then
        If Left(param3, 1) = "N" Then
            pReturnPile = Val(Right(param3, Len(param3) - 1))
        ElseIf Left(param3, 1) = "L" Then
            pReturnPile = Int(Rnd * (NumPiles + 1)) + 1
        End If
        If NumPiles = 8 Then
            MsgBox ("There can only be at most 8 piles.  You can not" & _
                Chr(13) & "create a new pile since there are already 8 piles.")
            Exit Sub
        End If
        If pReturnPile > NumPiles Then
            pReturnPile = NumPiles + 1
            'this last statement in case the legitimate requested pReturnPile
            'is several numbers larger than NumPiles
        End If
        'adjust PileTable for a new pile
        pBackCounter = NumPiles + 1
        pNewPileCreated = 0
        While pNewPileCreated = 0
            If pBackCounter = pReturnPile Then
                If pBackCounter = 1 Then
                    PileTable(pBackCounter, 1) = 1
                    PileTable(pBackCounter, 2) = 0
                Else
                    PileTable(pBackCounter, 1) = PileTable(pBackCounter - 1, 2) + 1
                    PileTable(pBackCounter, 2) = PileTable(pBackCounter, 1) - 1
                End If
                pNewPileCreated = 1
            Else
                PileTable(pBackCounter, 1) = PileTable(pBackCounter - 1, 1)
                PileTable(pBackCounter, 2) = PileTable(pBackCounter - 1, 2)
            End If
            pBackCounter = pBackCounter - 1
        Wend
        NumPiles = NumPiles + 1
        If pReturnPile <= pSelectedPile Then
            pSelectedPile = pSelectedPile + 1
        End If
    End If
    'this calculation determines how many cards are in the
    'return pile before the return card arrives
    pReturnPileCards = PileTable(pReturnPile, 2) - PileTable(pReturnPile, 1) + 1
    'decode fourth parameter
    If Left(param4, 1) = "T" Then
        pReturnPosition = 1
    ElseIf Left(param4, 1) = "B" Then
        If pReturnPile = pSelectedPile Then
            pReturnPosition = pReturnPileCards
        Else
            pReturnPosition = pReturnPileCards + 1
        End If
    ElseIf Left(param4, 1) = "E" Then
        pReturnPosition = pSelectedPosition
    ElseIf Left(param4, 1) = "R" Then
        If pReturnPile = pSelectedPile Then
            pReturnPosition = Int(Rnd * pReturnPileCards) + 1
        Else
            pReturnPosition = Int(Rnd * (pReturnPileCards + 1)) + 1
        End If
    ElseIf Left(param4, 1) = "S" Then
        pReturnPosition = Val(Mid(param4, 2, Len(param4) - 1))
    ElseIf Left(param4, 1) = "N" Then
        pReturnPosition = 1
        pReturnPileCards = 0
    End If
    If Not Left(param4, 1) = "N" Then
        If pSelectedPile = pReturnPile Then
            If pReturnPosition > pReturnPileCards Then
                If pSelectedPosition = 1 Then
                    pSuffix = "st"
                ElseIf pSelectedPosition = 2 Then
                    pSuffix = "nd"
                ElseIf pSelectedPosition = 3 Then
                    pSuffix = "rd"
                ElseIf pSelectedPosition > 3 Then
                    pSuffix = "th"
                End If
                If pReturnPileCards = 1 Then
                    MsgBox ("The " & pSelectedPosition & pSuffix & " card was selected " & _
                        "from pile " & pSelectedPile & "." & Chr(13) & Chr(13) & _
                        "You can not specify a return position of " & pReturnPosition & _
                        Chr(13) & "since there is only " & _
                        pReturnPileCards & " card" & _
                        Chr(13) & "in Return Pile " & pReturnPile)
                Else
                    MsgBox ("The " & pSelectedPosition & pSuffix & " card was selected " & _
                        "from pile " & pSelectedPile & "." & Chr(13) & Chr(13) & _
                        "You can not specify a return position of " & pReturnPosition & _
                        Chr(13) & "since there are only " & _
                        pReturnPileCards & " cards" & _
                        Chr(13) & "in Return Pile " & pReturnPile)
                End If
                Exit Sub
            End If
        Else
            If pReturnPosition > pReturnPileCards + 1 Then
                If pSelectedPosition = 1 Then
                    pSuffix = "st"
                ElseIf pSelectedPosition = 2 Then
                    pSuffix = "nd"
                ElseIf pSelectedPosition = 3 Then
                    pSuffix = "rd"
                ElseIf pSelectedPosition > 3 Then
                    pSuffix = "th"
                End If
                If pReturnPileCards + 1 = 1 Then
                    MsgBox ("The " & pSelectedPosition & pSuffix & " card was selected " & _
                        "from pile " & pSelectedPile & "." & Chr(13) & Chr(13) & _
                        "You can not specify a return position of " & pReturnPosition & _
                        Chr(13) & "since there can only be " & _
                        pReturnPileCards + 1 & " card" & _
                        Chr(13) & "in Return Pile " & pReturnPile)
                Else
                    MsgBox ("The " & pSelectedPosition & pSuffix & " card was selected " & _
                        "from pile " & pSelectedPile & "." & Chr(13) & Chr(13) & _
                        "You can not specify a return position of " & pReturnPosition & _
                        Chr(13) & "since there can only be " & _
                        pReturnPileCards + 1 & " cards" & _
                        Chr(13) & "in Return Pile " & pReturnPile)
                End If
                Exit Sub
            End If
        End If
    End If
    'decode the fifth parameter
    'identify the selected card if correct condition
    If param5 = "S" Then
        Deck(4, PileTable(pSelectedPile, 1) + pSelectedPosition - 1) = "Selected"
        If frmStackView.SelectionsTextBox.Text = Empty Then
            frmStackView.SelectionsTextBox.Text = _
                Deck(2, PileTable(pSelectedPile, 1) + pSelectedPosition - 1)
        Else
            frmStackView.SelectionsTextBox.Text = frmStackView.SelectionsTextBox.Text _
                & " " & Deck(2, PileTable(pSelectedPile, 1) + pSelectedPosition - 1)
        End If
    End If
    'before the PileTable is adjusted, record the positions of the card to move
    fromCardParam = PileTable(pSelectedPile, 1) + pSelectedPosition - 1
    'adjust Pile parameters if necessary
    Dim pPileMatch As Integer
    Dim pZeroPile As Integer
    If pSelectedPile <> pReturnPile Then
        'reduce appropriate positions by 1
        pPileMatch = 0
        pZeroPile = 0
        For i% = 1 To NumPiles
            If pPileMatch = 1 Then
                PileTable(i%, 1) = PileTable(i%, 1) - 1
                PileTable(i%, 2) = PileTable(i%, 2) - 1
            End If
            If pSelectedPile = i% Then
                PileTable(i%, 2) = PileTable(i%, 2) - 1
                pPileMatch = 1
                If PileTable(i%, 1) > PileTable(i%, 2) Then
                    pZeroPile = 1
                End If
            End If
        Next i%
        'increase appropriate positions by 1
        pPileMatch = 0
        For i% = 1 To NumPiles
            If pPileMatch = 1 Then
                PileTable(i%, 1) = PileTable(i%, 1) + 1
                PileTable(i%, 2) = PileTable(i%, 2) + 1
            End If
            If pReturnPile = i% Then
                PileTable(i%, 2) = PileTable(i%, 2) + 1
                pPileMatch = 1
            End If
        Next i%
        'correct for ZeroPile condition (a pile had one card, and it was moved)
        If pZeroPile = 1 Then
            If pSelectedPile < NumPiles Then
                For i% = 1 To NumPiles - 1
                    If i% >= pSelectedPile Then
                        PileTable(i%, 1) = PileTable(i% + 1, 1)
                        PileTable(i%, 2) = PileTable(i% + 1, 2)
                    End If
                Next i%
            End If
            NumPiles = NumPiles - 1
        End If
    End If
    'correct for ZeroPile condition (a pile had one card, and it was moved)
    If pZeroPile = 1 Then
        If pReturnPile > pSelectedPile Then
            pReturnPile = pReturnPile - 1
        End If
    End If
    'after the PileTable adjustments have been made, record the position
    'of the card to move to
    toCardParam = PileTable(pReturnPile, 1) + pReturnPosition - 1
    'move the selected card within the Deck
    For z% = 1 To DeckProperties
        ChangedDeck(z%, toCardParam) = Deck(z%, fromCardParam)
    Next z%
    If toCardParam < fromCardParam Then
        For j% = 1 To toCardParam - 1
            For z% = 1 To DeckProperties
                ChangedDeck(z%, j%) = Deck(z%, j%)
            Next z%
        Next j%
        For k% = 1 To fromCardParam - toCardParam
            For z% = 1 To DeckProperties
                ChangedDeck(z%, toCardParam + k%) = _
                Deck(z%, toCardParam - 1 + k%)
            Next z%
        Next k%
        For n% = 1 To DeckCount - fromCardParam
            For z% = 1 To DeckProperties
                ChangedDeck(z%, fromCardParam + n%) = _
                Deck(z%, fromCardParam + n%)
            Next z%
        Next n%
        For m% = 1 To DeckCount
            For z% = 1 To DeckProperties
                Deck(z%, m%) = ChangedDeck(z%, m%)
            Next z%
        Next m%
    ElseIf toCardParam > fromCardParam Then
        For k% = 1 To fromCardParam - 1
            For z% = 1 To DeckProperties
                ChangedDeck(z%, k%) = _
                Deck(z%, k%)
            Next z%
        Next k%
        For j% = 1 To toCardParam - fromCardParam
            For z% = 1 To DeckProperties
                ChangedDeck(z%, fromCardParam - 1 + j%) = _
                Deck(z%, fromCardParam + j%)
            Next z%
        Next j%
        For n% = 1 To DeckCount - toCardParam
            For z% = 1 To DeckProperties
                ChangedDeck(z%, toCardParam + n%) = _
                Deck(z%, toCardParam + n%)
            Next z%
        Next n%
        For m% = 1 To DeckCount
            For z% = 1 To DeckProperties
                Deck(z%, m%) = ChangedDeck(z%, m%)
            Next z%
        Next m%
    End If
    PilesMatrixRefresh
    CreatePiles (NumPiles)
End Sub

Private Sub SelectedCardSpecifiedText_Change()
SelectedCardSpecified.Value = True
End Sub


Private Sub SelectReturnButton_Click()
    'establish working variables for this module
    Dim paramCode1 As Variant
        'selected pile
    Dim paramCode2 As Variant
        'selected card
    Dim paramCode3 As Variant
        'return pile
    Dim paramCode4 As Variant
        'return position
    Dim paramCode5 As Variant
        'move only
    Dim pSelectedPile As Integer
    'this is the number of the selected pile
    Dim pSelectedPileCards As Integer
    'this is the number of cards in the selected pile
    Dim pSelectedPosition As Integer
    'this is the position from the top of the selected pile that
    'the selected card is located
    Dim pReturnPile As Integer
    'this is the pile number of the return pile
    Dim pReturnPileCards As Integer
    'this is the number of cards in the return pile
    Dim pReturnPosition As Integer
    'this is the position in the return pile that the selected
    'card will be placed
    Dim pSuffix As String
    'st, nd, rd, th suffixes for an error message
    Dim pTempCounter
    'used for instances of Reverse checks in the Pile Matrix
    '---
    'make sure piles have been created
    If PilesShown = 0 Then
        MsgBox ("You must create piles before you can manipulate piles.")
        Exit Sub
    End If
    ' make sure the PileMatrix has a valid selection
    ' but only if not random selections chosen
    If Not (SelectionPileRandom.Value = True And _
        (ReturnPileRandomAny.Value = True Or _
        ReturnPileRandomNotSame.Value = True Or _
        ReturnPilePrimary.Value = True Or _
        ReturnPileNewPileRandom.Value = True)) Then
            Dim PileMatrixCheckSum As Integer
            PileMatrixCheckSum = 0
            For i% = 1 To 8
                For j% = 1 To 8
                    If PileMatrix(10 * i% + j%).Value = 1 Then
                        PileMatrixCheckSum = PileMatrixCheckSum + 1
                    End If
                Next j%
            Next i%
            If PileMatrixCheckSum = 0 Then
                MsgBox ("You must check a box in the Pile Control Matrix before" & Chr(13) & _
                    "you can manipulate piles with non-random settings.")
                Exit Sub
            End If
    End If
    'initialize Pile codes
    paramCode1 = Empty
    paramCode2 = Empty
    paramCode3 = Empty
    paramCode4 = Empty
    paramCode5 = Empty
    'establish Pile codes
    'get Pile Matrix info if non random settings
    If Not (SelectionPileRandom.Value = True And _
        (ReturnPileRandomAny.Value = True Or _
        ReturnPileRandomNotSame.Value = True Or _
        ReturnPilePrimary.Value = True)) Then
            PileMatrixQuery
    End If
    'set Selection Pile code (paramCode1)
    If SelectionPilePrimary.Value = True Then
        pSelectedPile = PileMatrixRow
        paramCode1 = "P" & pSelectedPile
    ElseIf SelectionPileRandom.Value = True Then
        'pSelectedPile = Int(Rnd * NumPiles) + 1
        paramCode1 = "R"
    End If
    'set Selected Card code (paramCode2)
    'pSelectedPileCards = PileTable(pSelectedPile, 2) - PileTable(pSelectedPile, 1) + 1
    If SelectedCardTop.Value = True Then
        'pSelectedPosition = 1
        paramCode2 = "T"
    ElseIf SelectedCardBottom.Value = True Then
        'pSelectedPosition = pSelectedPileCards
        paramCode2 = "B"
    ElseIf SelectedCardRandom.Value = True Then
        'pSelectedPosition = Int(Rnd * pSelectedPileCards) + 1
        paramCode2 = "R"
    ElseIf SelectedCardSpecified.Value = True Then
        If Not IsNumeric(SelectedCardSpecifiedText.Text) Then
            SelectedCardSpecifiedText.Text = Empty
            SelectedCardSpecifiedText.SetFocus
            MsgBox ("You can only enter numbers in the 'Specified' text box.")
            Exit Sub
        End If
        If Val(SelectedCardSpecifiedText.Text) < 1 Or _
            Val(SelectedCardSpecifiedText.Text) > 52 Then
            SelectedCardSpecifiedText.Text = Empty
            SelectedCardSpecifiedText.SetFocus
            MsgBox ("You can only enter numbers from 1 to 52" & _
                Chr(13) & "in the 'Specified' text box" & _
                Chr(13) & "(up to the number of cards in the pile).")
            Exit Sub
        End If
        paramCode2 = "S" & SelectedCardSpecifiedText.Text
    End If
    'tag paramCode2 with an "R" if the reverse check was selected
    If SelectionReverseCheck.Value = 1 Then
        paramCode2 = paramCode2 & "R"
    End If
    'set Return Pile code (paramCode3)
    If ReturnPileSecondary.Value = True Then
        pReturnPile = PileMatrixColumn
        paramCode3 = "S" & pReturnPile
    ElseIf ReturnPilePrimary.Value = True Then
        If SelectionPilePrimary.Value = True Then
            paramCode3 = "P" & pSelectedPile
        ElseIf SelectionPileRandom.Value = True Then
            paramCode3 = "E"
            '"E" is equivalent to Same for when the selected pile was random
            'the Call code will need to determine the random values
        End If
    ElseIf ReturnPileRandomAny.Value = True Then
        paramCode3 = "R"
    ElseIf ReturnPileRandomNotSame.Value = True Then
        paramCode3 = "D"
        'means Different (not same)
    ElseIf ReturnPileNewPileSpecified.Value = True Then
        If Not IsNumeric(ReturnPileNewPileSpecifiedText.Text) Then
            ReturnPileNewPileSpecifiedText.Text = Empty
            ReturnPileNewPileSpecifiedText.SetFocus
            MsgBox ("You can only enter numbers in the 'New Pile' text box.")
            Exit Sub
        End If
        If Val(ReturnPileNewPileSpecifiedText.Text) < 1 Or _
            Val(ReturnPileNewPileSpecifiedText.Text) > 8 Then
            ReturnPileNewPileSpecifiedText.Text = Empty
            ReturnPileNewPileSpecifiedText.SetFocus
            MsgBox ("You can only enter numbers from 1 to 8" & _
                Chr(13) & "in the 'New Pile' text box.")
            Exit Sub
        End If
        paramCode3 = "N" & ReturnPileNewPileSpecifiedText.Text
    ElseIf ReturnPileNewPileRandom.Value = True Then
        paramCode3 = "L"
    End If
    'set Return Position parameter code (paramCode4)
    If ReturnPileNewPileRandom.Value = True Or _
        ReturnPileNewPileSpecified = True Then
        paramCode4 = "N"
    Else
        If ReturnPositionTop.Value = True Then
            paramCode4 = "T"
        ElseIf ReturnPositionBottom.Value = True Then
            paramCode4 = "B"
        ElseIf ReturnPositionSame.Value = True Then
            paramCode4 = "E"
            '"E" is equivalent to Same
        ElseIf ReturnPositionRandom.Value = True Then
            paramCode4 = "R"
        ElseIf ReturnPositionSpecified.Value = True Then
            If Not IsNumeric(ReturnPositionSpecifiedText.Text) Then
                ReturnPositionSpecifiedText.Text = Empty
                ReturnPositionSpecifiedText.SetFocus
                MsgBox ("You can only enter numbers in the 'Specified' text box.")
                Exit Sub
            End If
            If Val(ReturnPositionSpecifiedText.Text) < 1 Then
                ReturnPositionSpecifiedText.Text = Empty
                ReturnPositionSpecifiedText.SetFocus
                MsgBox ("You can only enter numbers greater than 0" & _
                    Chr(13) & "in the 'Specified' text box.")
                Exit Sub
            End If
            paramCode4 = "S" & ReturnPositionSpecifiedText.Text
        End If
    End If
    'set Move Only parameter
    If MoveOnlyCheck.Value = 1 Then
        paramCode5 = "M"
    ElseIf MoveOnlyCheck.Value = 0 Then
        paramCode5 = "S"
    End If
    
    'identify ignoring of Reverse checks in the Pile Matrix if they are present
    pTempCounter = 0
    For i% = 1 To 8
        If ReverseR(i%).Value = 1 Then
            pTempCounter = pTempCounter + 1
        End If
        If ReverseC(i%).Value = 1 Then
            pTempCounter = pTempCounter + 1
        End If
    Next i%
    If pTempCounter > 0 Then
        MsgBox ("When Select/Return command is performed, 'Reverse'" & _
        Chr(13) & "checkboxes in the Pile Matrix are ignored.")
    End If
    Call SelectReturn(paramCode1, paramCode2, paramCode3, paramCode4, paramCode5)
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "SelectReturn(" & paramCode1 _
        & ", " & paramCode2 _
        & ", " & paramCode3 _
        & ", " & paramCode4 _
        & ", " & paramCode5 _
        & ")"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End Sub


Private Sub SetNumPilesPlan()
If DealCardsOption.Value = True And _
    DealAlternatingOption.Value = True And _
    NumberOfCardsToDealText.Text <> Empty And _
    Val(NumberOfCardsToDealText.Text) < 52 Then
    NumPilesPlan = Val(NumberOfPilesText.Text) + 1
Else
    NumPilesPlan = Val(NumberOfPilesText.Text)
End If
End Sub


Public Sub ShowPiles()
PokerCardsDealt = 0
PilesShown = 1
Call frmDeck.DisplayPiles
End Sub

Public Sub ShowPilesGilbreath()
PokerCardsDealt = 0
PilesShown = 1
Call frmDeck.DisplayPilesGilbreath
End Sub


Public Sub ElmsleyCount(p1)
Dim pPileSize As Integer
Dim pParam As Variant
pParam = "P" & p1
'check that the pile has at least 4 cards
pPileSize = PileTable(p1, 2) - PileTable(p1, 1) + 1
If pPileSize < 4 Then
    MsgBox ("You must have at least 4 cards in the " & Chr(13) & _
        "specified pile (Pile " & p1 & ") to do an Elmsley Count.")
    Exit Sub
End If
Call SelectReturn(pParam, "B", pParam, "S2", "M")
End Sub

Public Sub InverseElmsleyCount(p1)
Dim pPileSize As Integer
Dim pParam As Variant
pParam = "P" & p1
'check that the pile has at least 4 cards
pPileSize = PileTable(p1, 2) - PileTable(p1, 1) + 1
If pPileSize < 4 Then
    MsgBox ("You must have at least 4 cards in the " & Chr(13) & _
        "specified pile (Pile " & p1 & ") to do an Inverse Elmsley Count.")
    Exit Sub
End If
Call SelectReturn(pParam, "S2", pParam, "B", "M")
End Sub

Public Sub JordanCount(p1)
Dim pPileSize As Integer
Dim pParam As Variant
pParam = "P" & p1
'check that the pile has at least 4 cards
pPileSize = PileTable(p1, 2) - PileTable(p1, 1) + 1
If pPileSize < 4 Then
    MsgBox ("You must have at least 4 cards in the " & Chr(13) & _
        "specified pile (Pile " & p1 & ") to do a Jordan Count.")
    Exit Sub
End If
Call SelectReturn(pParam, "S2", pParam, "B", "M")
End Sub

Public Sub InverseJordanCount(p1)
Dim pPileSize As Integer
Dim pParam As Variant
pParam = "P" & p1
'check that the pile has at least 4 cards
pPileSize = PileTable(p1, 2) - PileTable(p1, 1) + 1
If pPileSize < 4 Then
    MsgBox ("You must have at least 4 cards in the " & Chr(13) & _
        "specified pile (Pile " & p1 & ") to do an Inverse Jordan Count.")
    Exit Sub
End If
Call SelectReturn(pParam, "B", pParam, "S2", "M")
End Sub

Public Sub TurnOver(p1)
Dim pParam As Variant
pParam = "P" & p1
Call CutPiles(pParam, "C", "P", "R")
End Sub

Private Sub SpecialButton_Click()
    'make sure piles have been created
    If PilesShown = 0 Then
        MsgBox ("You must create piles before you can manipulate piles.")
        Exit Sub
    End If
    ' make sure the PileMatrix has a valid selection
    Dim PileMatrixCheckSum As Integer
    PileMatrixCheckSum = 0
    For i% = 1 To 8
        For j% = 1 To 8
            If PileMatrix(10 * i% + j%).Value = 1 Then
                PileMatrixCheckSum = PileMatrixCheckSum + 1
            End If
        Next j%
    Next i%
    If PileMatrixCheckSum = 0 Then
        MsgBox ("You must check a box in the Pile Control Matrix" & Chr(13) & _
            "before you can manipulate piles.")
        Exit Sub
    End If
    Dim pTempCounter As Integer
    'identify the piles to manipulate
    Call PileMatrixQuery
    'set parameter variables
    Dim p1 As Variant
    'set first parameter
    p1 = PileMatrixRow
    If SpecialElmsley.Value = True Then
        If SpecialInverseCheck.Value = 0 Then
            Call ElmsleyCount(p1)
            'SessionRecord
            If SessionRecordMode Then
                SessionCommand = "ElmsleyCount(" & p1 & ")"
                frmStackView.SessionListBox.AddItem SessionCommand
                frmStackView.SessionStatusUpdate (0)
            End If
        ElseIf SpecialInverseCheck.Value = 1 Then
            Call InverseElmsleyCount(p1)
            'SessionRecord
            If SessionRecordMode Then
                SessionCommand = "InverseElmsleyCount(" & p1 & ")"
                frmStackView.SessionListBox.AddItem SessionCommand
                frmStackView.SessionStatusUpdate (0)
            End If
        End If
    ElseIf SpecialJordan.Value = True Then
        If SpecialInverseCheck.Value = 0 Then
            Call JordanCount(p1)
            'SessionRecord
            If SessionRecordMode Then
                SessionCommand = "JordanCount(" & p1 & ")"
                frmStackView.SessionListBox.AddItem SessionCommand
                frmStackView.SessionStatusUpdate (0)
            End If
        ElseIf SpecialInverseCheck.Value = 1 Then
            Call InverseJordanCount(p1)
            'SessionRecord
            If SessionRecordMode Then
                SessionCommand = "InverseJordanCount(" & p1 & ")"
                frmStackView.SessionListBox.AddItem SessionCommand
                frmStackView.SessionStatusUpdate (0)
            End If
        End If
    ElseIf SpecialReverseOrder.Value = True Then
            Call TurnOver(p1)
            'SessionRecord
            If SessionRecordMode Then
                SessionCommand = "TurnOver(" & p1 & ")"
                frmStackView.SessionListBox.AddItem SessionCommand
                frmStackView.SessionStatusUpdate (0)
            End If
    End If
End Sub

Private Sub SwapPilesButton_Click()
    'establish working variables for this module
    Dim paramCode1 As Variant
        'first pile parameter
    Dim paramCode2 As Variant
        'second pile parameter
    Dim pFirstPile As Integer
    'this is the code of the first pile
    Dim pSecondPile As Integer
    'this is the code of the second pile
    Dim pSuffix As String
    'st, nd, rd, th suffixes for an error message
    Dim pTempCounter
    'used for instances of Reverse checks in the Pile Matrix
    '---
    'make sure piles have been created
    If PilesShown = 0 Then
        MsgBox ("You must create piles before you can manipulate piles.")
        Exit Sub
    End If
    ' make sure the PileMatrix has a valid selection
    ' but only if not random selections chosen
    If SwapFirstPrimary.Value = True Or _
        SwapSecondSecondary.Value = True Then
            Dim PileMatrixCheckSum As Integer
            PileMatrixCheckSum = 0
            For i% = 1 To 8
                For j% = 1 To 8
                    If PileMatrix(10 * i% + j%).Value = 1 Then
                        PileMatrixCheckSum = PileMatrixCheckSum + 1
                    End If
                Next j%
            Next i%
            If PileMatrixCheckSum = 0 Then
                MsgBox ("You must check a box in the Pile Control Matrix before" & Chr(13) & _
                    "you can manipulate piles with non-random settings.")
                Exit Sub
            End If
    End If
    'initialize Pile codes
    paramCode1 = Empty
    paramCode2 = Empty
    'establish Pile codes
    'get Pile Matrix info if non random settings
    If SwapFirstPrimary.Value = True Or _
        SwapSecondSecondary.Value = True Then
            PileMatrixQuery
    End If
    'set First Pile code (paramCode1)
    If SwapFirstPrimary.Value = True Then
        pSelectedPile = PileMatrixRow
        paramCode1 = "P" & pSelectedPile
    ElseIf SwapFirstRandom.Value = True Then
        paramCode1 = "R"
    ElseIf SwapFirstSelected.Value = True Then
        paramCode1 = "S"
    ElseIf SwapFirstNoSelected.Value = True Then
        paramCode1 = "N"
    End If
    'set Second Pile code (paramCode2)
    If SwapSecondSecondary.Value = True Then
        pSelectedPile = PileMatrixColumn
        paramCode2 = "P" & pSelectedPile
    ElseIf SwapSecondRandom.Value = True Then
        paramCode2 = "R"
    ElseIf SwapSecondSelected.Value = True Then
        paramCode2 = "S"
    ElseIf SwapSecondNoSelected.Value = True Then
        paramCode2 = "N"
    End If
    
    'identify ignoring of Reverse checks in the Pile Matrix if they are present
    pTempCounter = 0
    For i% = 1 To 8
        If ReverseR(i%).Value = 1 Then
            pTempCounter = pTempCounter + 1
        End If
        If ReverseC(i%).Value = 1 Then
            pTempCounter = pTempCounter + 1
        End If
    Next i%
    If pTempCounter > 0 Then
        MsgBox ("When Swap Piles command is performed, 'Reverse'" & _
        Chr(13) & "checkboxes in the Pile Matrix are ignored.")
    End If
    If SwapReverseFirstPile.Value = 1 Then
        paramCode1 = paramCode1 & "R"
    End If
    If SwapReverseSecondPile.Value = 1 Then
        paramCode2 = paramCode2 & "R"
    End If
    
    'run the procedure
    Call SwapPiles(paramCode1, paramCode2)
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "SwapPiles(" & paramCode1 _
        & ", " & paramCode2 _
        & ")"
        frmStackView.SessionListBox.AddItem SessionCommand
        frmStackView.SessionStatusUpdate (0)
    End If
End Sub

Public Sub SwapPiles(param1, param2)
'establish procedure variables
Dim pTemp As Integer
Dim pLeftPile As Integer
Dim pRightPile As Integer
Dim pLeftReverse As Boolean
Dim pRightReverse As Boolean
Dim pSwapPilesError As Boolean
Dim pFoundSelectedCard As Boolean
Dim pSelectedCardTextbox As String
Dim pNumberSelectedCards As Integer
Dim pLeftSelectedCard As String
Dim pRightSelectedCard As String
Dim pLeftSelectedCardPosition As Integer
Dim pRightSelectedCardPosition As Integer
Dim pStringPointer As Integer
Dim pSelectedPile1 As Integer
Dim pSelectedPile2 As Integer
Dim pNonSelectedPileFound As Boolean
Dim pValidSecondPileFound As Boolean
Dim pLeftContainsSelected As Boolean
Dim pRandomPile As Integer
Dim pPileHasSelectedCard(8) As Boolean
Dim pNumberPilesWithSelected As Integer
Dim pNumberPilesWithoutSelected As Integer
'initialize error condition to false
pSwapPilesError = False
'initialize Reverse parameters
pLeftReverse = False
pRightReverse = False
'initialize parameters
pFoundSelectedCard = False
pNonSelectedPileFound = False
pValidSecondPileFound = False
pLeftContainsSelected = False
pSelectedCardTextbox = frmStackView.SelectionsTextBox.Text
pLeftSelectedCard = Empty
pRightSelectedCard = Empty
pNumberSelectedCards = 0
pLeftSelectedCardPosition = 0
pRightSelectedCardPosition = 0
pNumberPilesWithSelected = 0
pNumberPilesWithoutSelected = 0
pStringPointer = 0
For i% = 1 To NumPiles
    pPileHasSelectedCard(i%) = False
Next i%
'ensure there are at least two piles
If NumPiles < 2 Then
    MsgBox ("There must be at least two piles to swap piles.")
End If
'new section to establish piles with/without selected cards
For i% = 1 To NumPiles
    For p% = PileTable(i%, 1) To PileTable(i%, 2)
        If Deck(4, p%) = "Selected" Then
            pPileHasSelectedCard(i%) = True
            pNumberSelectedCards = pNumberSelectedCards + 1
        End If
    Next p%
    If pPileHasSelectedCard(i%) Then
        pNumberPilesWithSelected = pNumberPilesWithSelected + 1
    Else
        pNumberPilesWithoutSelected = pNumberPilesWithoutSelected + 1
    End If
Next i%
'end of new section to establish piles with/without selected cards

'establish any selected cards
'first establish the first two selected cards
'If pSelectedCardTextbox = Empty Then
'    pLeftSelectedCardPosition = 0
'    pRightSelectedCardPosition = 0
'    pNumberSelectedCards = 0
'    pSelectedPile1 = 0
'    pSelectedPile2 = 0
'Else
'    pStringPointer = InStr(pSelectedCardTextbox, " ")
'    If pStringPointer = 0 Then
'        'in this case there is only one selected card
'        pLeftSelectedCard = pSelectedCardTextbox
'        pNumberSelectedCards = 1
'        'need to identify which pile the selection is in
'        'first identify the current deck position
'        pLeftSelectedCardPosition = 0
'        pRightSelectedCardPosition = 0
'        For i% = 1 To DeckCount
'            If Deck(4, i%) = "Selected" And _
'                Deck(2, i%) = pLeftSelectedCard Then
'                    pLeftSelectedCardPosition = i%
'            End If
'        Next i%
'        'next, identify which pile the selected card is in
'        For n% = 1 To NumPiles
'            If pLeftSelectedCardPosition >= PileTable(n%, 1) And _
'                pLeftSelectedCardPosition <= PileTable(n%, 2) Then
'                    pSelectedPile1 = n%
'            End If
'        Next n%
'        pSelectedPile2 = 0
'    Else
'        'in this case there is more than one selected card
'        pNumberSelectedCards = 2
'        'first need to identify the first selected card in the list
'        pLeftSelectedCard = Left(pSelectedCardTextbox, pStringPointer - 1)
'        'next need to strip off the first selected card from the list
'        'this allows the next parameter to also use a selected card
'        pSelectedCardTextbox = Right(pSelectedCardTextbox, _
'            Len(pSelectedCardTextbox) - pStringPointer)
'        'need to identify which pile the selection is in
'        'first identify the current deck position
'        pLeftSelectedCardPosition = 0
'        pRightSelectedCardPosition = 0
'        For i% = 1 To DeckCount
'            If Deck(4, i%) = "Selected" And _
'                Deck(2, i%) = pLeftSelectedCard Then
'                    pLeftSelectedCardPosition = i%
'            End If
'        Next i%
'        'next, identify which pile the selected card is in
'        For n% = 1 To NumPiles
'            If pLeftSelectedCardPosition >= PileTable(n%, 1) And _
'                pLeftSelectedCardPosition <= PileTable(n%, 2) Then
'                    pSelectedPile1 = n%
'            End If
'        Next n%
'        'repeat above step as a nested run for second card
'        pStringPointer = InStr(pSelectedCardTextbox, " ")
'        If pStringPointer = 0 Then
'            'in this case there is only one selected card
'            pRightSelectedCard = pSelectedCardTextbox
'            'since there is only one remaining card, clear the textbox
'            'this prevents the next parameter to use the same selected card
'            pSelectedCardTextbox = Empty
'            'establish position and pile
'            pRightSelectedCardPosition = 0
'            For i% = 1 To DeckCount
'                If Deck(4, i%) = "Selected" And _
'                    Deck(2, i%) = pRightSelectedCard Then
'                        pRightSelectedCardPosition = i%
'                End If
'            Next i%
'            'next, identify which pile the selected card is in
'            For n% = 1 To NumPiles
'                If pRightSelectedCardPosition >= PileTable(n%, 1) And _
'                    pRightSelectedCardPosition <= PileTable(n%, 2) Then
'                        pSelectedPile2 = n%
'                End If
'            Next n%
'        Else
'            'in this case there is more than one additional selected card
'            'first need to identify the first selected card in the list
'            pRightSelectedCard = Left(pSelectedCardTextbox, pStringPointer - 1)
'            'need to identify which pile the selection is in
'            'first identify the current deck position
'            pRightSelectedCardPosition = 0
'            For i% = 1 To DeckCount
'                If Deck(4, i%) = "Selected" And _
'                    Deck(2, i%) = pRightSelectedCard Then
'                        pRightSelectedCardPosition = i%
'                End If
'            Next i%
'            'next, identify which pile the selected card is in
'            For n% = 1 To NumPiles
'                If pRightSelectedCardPosition >= PileTable(n%, 1) And _
'                    pRightSelectedCardPosition <= PileTable(n%, 2) Then
'                        pSelectedPile2 = n%
'                End If
'            Next n%
'        End If
'    End If
'End If

'decode first (left) parameter
'first section checks if there are three characters which can only be "PxR"
If Len(param1) = 3 Then
    If Left(param1, 1) = "P" Then
        If IsNumeric(Mid(param1, 2, 1)) Then
            pLeftPile = Val(Mid(param1, 2, 1))
            If pLeftPile > NumPiles Then
                MsgBox ("Error: Swap Piles" & Chr(13) & _
                    "The first parameter (" & Chr(34) & param1 & Chr(34) & _
                    ") is referencing a pile number" & Chr(13) & _
                    "that is greater than the current number of piles.")
                Exit Sub
            End If
            If pLeftPile < 1 Then
                MsgBox ("Error: Swap Piles" & Chr(13) & _
                    "The first parameter (" & Chr(34) & param1 & Chr(34) & _
                    ") is referencing a pile number" & Chr(13) & _
                    "that is less than 1.  You must reference an actual pile number.")
                Exit Sub
            End If
        Else
            MsgBox ("Error: Swap Piles" & Chr(13) & _
                "The first parameter (" & Chr(34) & param1 & Chr(34) & _
                ") can only be:" & Chr(13) & _
                Chr(34) & "Px" & Chr(34) & " (specified Pile with number)" & Chr(13) & _
                Chr(34) & "R" & Chr(34) & " (Random pile)" & Chr(13) & _
                Chr(34) & "S" & Chr(34) & " (Pile with a selected card)" & Chr(13) & _
                Chr(34) & "N" & Chr(34) & " (pile with No selected card)" & Chr(13) & _
                Chr(13) & "The above parameters may have an " & Chr(34) & "R" & Chr(34) & Chr(13) & _
                "suffix which indicates the pile should be Reversed.")
            Exit Sub
        End If
    Else
        MsgBox ("Error: Swap Piles" & Chr(13) & _
            "The first parameter (" & Chr(34) & param1 & Chr(34) & _
            ") can only be:" & Chr(13) & _
            Chr(34) & "Px" & Chr(34) & " (specified Pile with number)" & Chr(13) & _
            Chr(34) & "R" & Chr(34) & " (Random pile)" & Chr(13) & _
            Chr(34) & "S" & Chr(34) & " (Pile with a selected card)" & Chr(13) & _
            Chr(34) & "N" & Chr(34) & " (pile with No selected card)" & Chr(13) & _
            Chr(13) & "The above parameters may have an " & Chr(34) & "R" & Chr(34) & Chr(13) & _
            "suffix which indicates the pile should be Reversed.")
        Exit Sub
    End If
    If Right(param1, 1) = "R" Then
        pLeftReverse = True
    Else
        MsgBox ("Error: Swap Piles" & Chr(13) & _
            "The first parameter (" & Chr(34) & param1 & Chr(34) & _
            ") can only be:" & Chr(13) & _
            Chr(34) & "Px" & Chr(34) & " (specified Pile with number)" & Chr(13) & _
            Chr(34) & "R" & Chr(34) & " (Random pile)" & Chr(13) & _
            Chr(34) & "S" & Chr(34) & " (Pile with a selected card)" & Chr(13) & _
            Chr(34) & "N" & Chr(34) & " (pile with No selected card)" & Chr(13) & _
            Chr(13) & "The above parameters may have an " & Chr(34) & "R" & Chr(34) & Chr(13) & _
            "suffix which indicates the pile should be Reversed.")
        Exit Sub
    End If
'    'identify if the first pile contains a selected card
'    If pLeftPile = pSelectedPile1 Or pLeftPile = pSelectedPile2 Then
'        pLeftContainsSelected = True
'    End If

'next check if there are two characters
ElseIf Len(param1) = 2 Then
    'if the first character is a "P" then the second must be a pile number
    If Left(param1, 1) = "P" Then
        If IsNumeric(Right(param1, 1)) Then
            pLeftPile = Val(Right(param1, 1))
            If pLeftPile > NumPiles Then
                MsgBox ("Error: Swap Piles" & Chr(13) & _
                    "The first parameter (" & Chr(34) & param1 & Chr(34) & _
                    ") is referencing a pile number" & Chr(13) & _
                    "that is greater than the current number of piles.")
                Exit Sub
            End If
            If pLeftPile < 1 Then
                MsgBox ("Error: Swap Piles" & Chr(13) & _
                    "The first parameter (" & Chr(34) & param1 & Chr(34) & _
                    ") is referencing a pile number" & Chr(13) & _
                    "that is less than 1.  You must reference an actual pile number.")
                Exit Sub
            End If
        Else
            MsgBox ("Error: Swap Piles" & Chr(13) & _
                "The first parameter (" & Chr(34) & param1 & Chr(34) & _
                ") can only be:" & Chr(13) & _
                Chr(34) & "Px" & Chr(34) & " (specified Pile with number)" & Chr(13) & _
                Chr(34) & "R" & Chr(34) & " (Random pile)" & Chr(13) & _
                Chr(34) & "S" & Chr(34) & " (Pile with a selected card)" & Chr(13) & _
                Chr(34) & "N" & Chr(34) & " (pile with No selected card)" & Chr(13) & _
                Chr(13) & "The above parameters may have an " & Chr(34) & "R" & Chr(34) & Chr(13) & _
                "suffix which indicates the pile should be Reversed.")
            Exit Sub
        End If
'        'identify if the first pile contains a selected card
'        If pLeftPile = pSelectedPile1 Or pLeftPile = pSelectedPile2 Then
'            pLeftContainsSelected = True
'        End If
    ElseIf Left(param1, 1) = "R" Then
        'set the first pile to a random number
        pLeftPile = Int(Rnd * NumPiles + 1)
        'check for a valid second parameter (can only be an "R")
        If Right(param1, 1) = "R" Then
            pLeftReverse = True
        Else
            MsgBox ("Error: Swap Piles" & Chr(13) & _
                "The first parameter (" & Chr(34) & param1 & Chr(34) & _
                ") can only be:" & Chr(13) & _
                Chr(34) & "Px" & Chr(34) & " (specified Pile with number)" & Chr(13) & _
                Chr(34) & "R" & Chr(34) & " (Random pile)" & Chr(13) & _
                Chr(34) & "S" & Chr(34) & " (Pile with a selected card)" & Chr(13) & _
                Chr(34) & "N" & Chr(34) & " (pile with No selected card)" & Chr(13) & _
                Chr(13) & "The above parameters may have an " & Chr(34) & "R" & Chr(34) & Chr(13) & _
                "suffix which indicates the pile should be Reversed.")
            Exit Sub
        End If
'        'identify if the first pile contains a selected card
'        If pLeftPile = pSelectedPile1 Or pLeftPile = pSelectedPile2 Then
'            pLeftContainsSelected = True
'        End If
    ElseIf Left(param1, 1) = "S" Then
        pLeftContainsSelected = True
        If pNumberPilesWithSelected = 0 Then
            MsgBox ("Error: Swap Piles" & Chr(13) & _
                "The first parameter (" & Chr(34) & param1 & Chr(34) & _
                ") requires there to be a selected card." & Chr(13) & _
                "However, there are no selected cards.")
            Exit Sub
        Else
            'new section for selected card piles
            pTemp = 0
            While pTemp = 0
                pRandomPile = Int(Rnd * NumPiles) + 1
                If pPileHasSelectedCard(pRandomPile) Then
                    pTemp = 99
                    'any non-zero value
                    pLeftPile = pRandomPile
                    'pNumberPilesWithSelected = pNumberPilesWithSelected - 1
                    'will be handled in later code
                End If
            Wend
            'end of new section
            
'            'establish pile based on number of selections
'            If pNumberSelectedCards = 1 Then
'                pLeftPile = pSelectedPile1
'            ElseIf pNumberSelectedCards = 2 Then
'                'establish a random selection (<0.5 then S1, else S2)
'                If Rnd < 0.5 Then
'                    pLeftPile = pSelectedPile1
'                Else
'                    pLeftPile = pSelectedPile2
'                End If
'            End If
            'check for a valid second parameter
            If Right(param1, 1) = "R" Then
                pLeftReverse = True
            Else
            MsgBox ("Error: Swap Piles" & Chr(13) & _
                "The first parameter (" & Chr(34) & param1 & Chr(34) & _
                ") can only be:" & Chr(13) & _
                Chr(34) & "Px" & Chr(34) & " (specified Pile with number)" & Chr(13) & _
                Chr(34) & "R" & Chr(34) & " (Random pile)" & Chr(13) & _
                Chr(34) & "S" & Chr(34) & " (Pile with a selected card)" & Chr(13) & _
                Chr(34) & "N" & Chr(34) & " (pile with No selected card)" & Chr(13) & _
                Chr(13) & "The above parameters may have an " & Chr(34) & "R" & Chr(34) & Chr(13) & _
                "suffix which indicates the pile should be Reversed.")
            Exit Sub
            End If
        End If
    ElseIf Left(param1, 1) = "N" Then
        'new section
        'make sure there are pile without selections
        If pNumberPilesWithoutSelected = 0 Then
            MsgBox ("Error: Swap Piles" & Chr(13) & _
                "The first parameter (" & Chr(34) & param1 & Chr(34) & _
                ") requires no selected card." & Chr(13) & _
                "However, there are no piles without a selection.")
            Exit Sub
        End If
        'find a pile without a selected card
        pTemp = 0
        While pTemp = 0
            pRandomPile = Int(Rnd * NumPiles) + 1
            If Not pPileHasSelectedCard(pRandomPile) Then
                pTemp = 99
                'any non-zero value
                pLeftPile = pRandomPile
                'pNumberPilesWithoutSelected = pNumberPilesWithoutSelected - 1
                'will be handled in later code
            End If
        Wend
        'end of new section
        
        
'        'make sure there are piles without selections before
'        'running the While...Wend loop so that it doesn't crash
'        If pNumberSelectedCards > 0 Then
'            If pNumberSelectedCards = 1 Then
'                If NumPiles < 2 Then
'                    MsgBox ("There are no piles without a valid selected card to choose.")
'                    Exit Sub
'                End If
'            ElseIf pNumberSelectedCards = 2 Then
'                'consider if both selections are in the same pile
'                If pSelectedPile1 = pSelectedPile2 Then
'                    If NumPiles < 2 Then
'                        MsgBox ("There are no piles without a valid selected card to choose.")
'                        Exit Sub
'                    End If
'                Else
'                    If NumPiles < 3 Then
'                        MsgBox ("There are no piles without a valid selected card to choose.")
'                        Exit Sub
'                    End If
'                End If
'            End If
'        End If
'        'find a pile without a valid selection
'        pNonSelectedPileFound = False
'        While Not pNonSelectedPileFound
'            pRandomPile = Int(Rnd * NumPiles) + 1
'            If pRandomPile <> pSelectedPile1 And _
'                pRandomPile <> pSelectedPile2 Then
'                pLeftPile = pRandomPile
'                pNonSelectedPileFound = True
'            End If
'        Wend
        
        'check for a valid second parameter
        If Right(param1, 1) = "R" Then
            pLeftReverse = True
        Else
            MsgBox ("Error: Swap Piles" & Chr(13) & _
                "The first parameter (" & Chr(34) & param1 & Chr(34) & _
                ") can only be:" & Chr(13) & _
                Chr(34) & "Px" & Chr(34) & " (specified Pile with number)" & Chr(13) & _
                Chr(34) & "R" & Chr(34) & " (Random pile)" & Chr(13) & _
                Chr(34) & "S" & Chr(34) & " (Pile with a selected card)" & Chr(13) & _
                Chr(34) & "N" & Chr(34) & " (pile with No selected card)" & Chr(13) & _
                Chr(13) & "The above parameters may have an " & Chr(34) & "R" & Chr(34) & Chr(13) & _
                "suffix which indicates the pile should be Reversed.")
            Exit Sub
        End If
    Else
        MsgBox ("Error: Swap Piles" & Chr(13) & _
            "The first parameter (" & Chr(34) & param1 & Chr(34) & _
            ") can only be:" & Chr(13) & _
            Chr(34) & "Px" & Chr(34) & " (specified Pile with number)" & Chr(13) & _
            Chr(34) & "R" & Chr(34) & " (Random pile)" & Chr(13) & _
            Chr(34) & "S" & Chr(34) & " (Pile with a selected card)" & Chr(13) & _
            Chr(34) & "N" & Chr(34) & " (pile with No selected card)" & Chr(13) & _
            Chr(13) & "The above parameters may have an " & Chr(34) & "R" & Chr(34) & Chr(13) & _
            "suffix which indicates the pile should be Reversed.")
        Exit Sub
    End If
ElseIf Len(param1) = 1 Then
    If param1 = "R" Then
        'set the first pile to a random number
        pLeftPile = Int(Rnd * NumPiles + 1)
        'identify if the first pile contains a selected card
        If pLeftPile = pSelectedPile1 Or pLeftPile = pSelectedPile2 Then
            pLeftContainsSelected = True
        End If
    ElseIf param1 = "S" Then
        pLeftContainsSelected = True
        If pNumberPilesWithSelected = 0 Then
            MsgBox ("Error: Swap Piles" & Chr(13) & _
                "The first parameter (" & Chr(34) & param1 & Chr(34) & _
                ") requires there to be a selected card." & Chr(13) & _
                "However, there are no selected cards.")
            Exit Sub
        Else
            'new section for selected card piles
            pTemp = 0
            While pTemp = 0
                pRandomPile = Int(Rnd * NumPiles) + 1
                If pPileHasSelectedCard(pRandomPile) Then
                    pTemp = 99
                    'any non-zero value
                    pLeftPile = pRandomPile
                    'pNumberPilesWithSelected = pNumberPilesWithSelected - 1
                    'will be handled in later code
                End If
            Wend
            'end of new section
            
'            'establish pile based on number of selections
'            If pNumberSelectedCards = 1 Then
'                pLeftPile = pSelectedPile1
'            ElseIf pNumberSelectedCards = 2 Then
'                'establish a random selection (<0.5 then S1, else S2)
'                If Rnd < 0.5 Then
'                    pLeftPile = pSelectedPile1
'                Else
'                    pLeftPile = pSelectedPile2
'                End If
'            End If
        End If
    ElseIf param1 = "N" Then
        'new section
        'make sure there are piles without selections
        If pNumberPilesWithoutSelected = 0 Then
            MsgBox ("Error: Swap Piles" & Chr(13) & _
                "The first parameter (" & Chr(34) & param1 & Chr(34) & _
                ") requires no selected card." & Chr(13) & _
                "However, there are no piles without a selection.")
            Exit Sub
        End If
        'find a pile without a selected card
        pTemp = 0
        While pTemp = 0
            pRandomPile = Int(Rnd * NumPiles) + 1
            If Not pPileHasSelectedCard(pRandomPile) Then
                pTemp = 99
                'any non-zero value
                pLeftPile = pRandomPile
                'pNumberPilesWithoutSelected = pNumberPilesWithoutSelected - 1
                'will be handled in later code
            End If
        Wend
        'end of new section
        
        
'        'make sure there are piles without selections before
'        'running the While...Wend loop so that it doesn't crash
'        If pNumberSelectedCards > 0 Then
'            If pNumberSelectedCards = 1 Then
'                If NumPiles < 2 Then
'                    MsgBox ("There are no piles without a valid selected card to choose.")
'                    Exit Sub
'                End If
'            ElseIf pNumberSelectedCards = 2 Then
'                'consider if both selections are in the same pile
'                If pSelectedPile1 = pSelectedPile2 Then
'                    If NumPiles < 2 Then
'                        MsgBox ("There are no piles without a valid selected card to choose.")
'                        Exit Sub
'                    End If
'                Else
'                    If NumPiles < 3 Then
'                        MsgBox ("There are no piles without a valid selected card to choose.")
'                        Exit Sub
'                    End If
'                End If
'            End If
'        End If
'        'find a pile without a valid selection
'        pNonSelectedPileFound = False
'        While Not pNonSelectedPileFound
'            pRandomPile = Int(Rnd * NumPiles) + 1
'            If pRandomPile <> pSelectedPile1 And _
'                pRandomPile <> pSelectedPile2 Then
'                pLeftPile = pRandomPile
'                pNonSelectedPileFound = True
'            End If
'        Wend
    Else
        MsgBox ("Error: Swap Piles" & Chr(13) & _
            "The first parameter (" & Chr(34) & param1 & Chr(34) & _
            ") can only be:" & Chr(13) & _
            Chr(34) & "Px" & Chr(34) & " (specified Pile with number)" & Chr(13) & _
            Chr(34) & "R" & Chr(34) & " (Random pile)" & Chr(13) & _
            Chr(34) & "S" & Chr(34) & " (Pile with a selected card)" & Chr(13) & _
            Chr(34) & "N" & Chr(34) & " (pile with No selected card)" & Chr(13) & _
            Chr(13) & "The above parameters may have an " & Chr(34) & "R" & Chr(34) & Chr(13) & _
            "suffix which indicates the pile should be Reversed.")
        Exit Sub
    End If
End If
'need to adjust the available piles with and without selections
If pPileHasSelectedCard(pLeftPile) Then
    pNumberPilesWithSelected = pNumberPilesWithSelected - 1
Else
    pNumberPilesWithoutSelected = pNumberPilesWithoutSelected - 1
End If

'need to identify the errors for the first parameter wherever pSwapPilesError=True
'initialize appropriate variables
pSwapPilesError = False



'decode second (right) parameter
'first section checks if there are three characters which can only be "PxR"
If Len(param2) = 3 Then
    If Left(param2, 1) = "P" Then
        If IsNumeric(Mid(param2, 2, 1)) Then
            pRightPile = Val(Mid(param2, 2, 1))
            If pRightPile > NumPiles Then
                MsgBox ("Error: Swap Piles" & Chr(13) & _
                    "The second parameter (" & Chr(34) & param2 & Chr(34) & _
                    ") is referencing a pile number" & Chr(13) & _
                    "that is greater than the current number of piles.")
                Exit Sub
            End If
            If pRightPile < 1 Then
                MsgBox ("Error: Swap Piles" & Chr(13) & _
                    "The first parameter (" & Chr(34) & param2 & Chr(34) & _
                    ") is referencing a pile number" & Chr(13) & _
                    "that is less than 1.  You must reference an actual pile number.")
                Exit Sub
            End If
            'check for same pile selection
            If pLeftPile = pRightPile Then
                MsgBox ("Error: Swap Piles" & Chr(13) & _
                    "The two piles can not be the same.  The parameter choices" & Chr(13) & _
                    "have resulted in both piles being the same.")
                Exit Sub
            End If
        Else
            MsgBox ("Error: Swap Piles" & Chr(13) & _
                "The second parameter (" & Chr(34) & param2 & Chr(34) & _
                ") can only be:" & Chr(13) & _
                Chr(34) & "Px" & Chr(34) & " (specified Pile with number)" & Chr(13) & _
                Chr(34) & "R" & Chr(34) & " (Random pile)" & Chr(13) & _
                Chr(34) & "S" & Chr(34) & " (Pile with a selected card)" & Chr(13) & _
                Chr(34) & "N" & Chr(34) & " (pile with No selected card)" & Chr(13) & _
                Chr(13) & "The above parameters may have an " & Chr(34) & "R" & Chr(34) & Chr(13) & _
                "suffix which indicates the pile should be Reversed.")
            Exit Sub
        End If
    Else
        MsgBox ("Error: Swap Piles" & Chr(13) & _
            "The second parameter (" & Chr(34) & param2 & Chr(34) & _
            ") can only be:" & Chr(13) & _
            Chr(34) & "Px" & Chr(34) & " (specified Pile with number)" & Chr(13) & _
            Chr(34) & "R" & Chr(34) & " (Random pile)" & Chr(13) & _
            Chr(34) & "S" & Chr(34) & " (Pile with a selected card)" & Chr(13) & _
            Chr(34) & "N" & Chr(34) & " (pile with No selected card)" & Chr(13) & _
            Chr(13) & "The above parameters may have an " & Chr(34) & "R" & Chr(34) & Chr(13) & _
            "suffix which indicates the pile should be Reversed.")
        Exit Sub
    End If
    If Right(param2, 1) = "R" Then
        pRightReverse = True
    Else
        MsgBox ("Error: Swap Piles" & Chr(13) & _
            "The second parameter (" & Chr(34) & param2 & Chr(34) & _
            ") can only be:" & Chr(13) & _
            Chr(34) & "Px" & Chr(34) & " (specified Pile with number)" & Chr(13) & _
            Chr(34) & "R" & Chr(34) & " (Random pile)" & Chr(13) & _
            Chr(34) & "S" & Chr(34) & " (Pile with a selected card)" & Chr(13) & _
            Chr(34) & "N" & Chr(34) & " (pile with No selected card)" & Chr(13) & _
            Chr(13) & "The above parameters may have an " & Chr(34) & "R" & Chr(34) & Chr(13) & _
            "suffix which indicates the pile should be Reversed.")
        Exit Sub
    End If
'next check if there are two characters
ElseIf Len(param2) = 2 Then
    'if the first character is a "P" then the second must be a pile number
    If Left(param2, 1) = "P" Then
        If IsNumeric(Right(param2, 1)) Then
            pRightPile = Val(Right(param2, 1))
            If pRightPile > NumPiles Then
                MsgBox ("Error: Swap Piles" & Chr(13) & _
                    "The second parameter (" & Chr(34) & param2 & Chr(34) & _
                    ") is referencing a pile number" & Chr(13) & _
                    "that is greater than the current number of piles.")
                Exit Sub
            End If
            If pRightPile < 1 Then
                MsgBox ("Error: Swap Piles" & Chr(13) & _
                    "The first parameter (" & Chr(34) & param2 & Chr(34) & _
                    ") is referencing a pile number" & Chr(13) & _
                    "that is less than 1.  You must reference an actual pile number.")
                Exit Sub
            End If
            'check for same pile selection
            If pLeftPile = pRightPile Then
                MsgBox ("Error: Swap Piles" & Chr(13) & _
                    "The two piles can not be the same.  The parameter choices" & Chr(13) & _
                    "have resulted in both piles being the same.")
                Exit Sub
            End If
        Else
            MsgBox ("Error: Swap Piles" & Chr(13) & _
                "The second parameter (" & Chr(34) & param2 & Chr(34) & _
                ") can only be:" & Chr(13) & _
                Chr(34) & "Px" & Chr(34) & " (specified Pile with number)" & Chr(13) & _
                Chr(34) & "R" & Chr(34) & " (Random pile)" & Chr(13) & _
                Chr(34) & "S" & Chr(34) & " (Pile with a selected card)" & Chr(13) & _
                Chr(34) & "N" & Chr(34) & " (pile with No selected card)" & Chr(13) & _
                Chr(13) & "The above parameters may have an " & Chr(34) & "R" & Chr(34) & Chr(13) & _
                "suffix which indicates the pile should be Reversed.")
            Exit Sub
        End If
    ElseIf Left(param2, 1) = "R" Then
        'set the second pile to a random number
        'but make sure to avoid the first pile
        pValidSecondPileFound = False
        While Not pValidSecondPileFound
            pRightPile = Int(Rnd * NumPiles + 1)
            If pLeftPile <> pRightPile Then
                pValidSecondPileFound = True
            End If
        Wend
        'check for a valid second parameter (can only be an "R")
        If Right(param2, 1) = "R" Then
            pRightReverse = True
        Else
            MsgBox ("Error: Swap Piles" & Chr(13) & _
                "The second parameter (" & Chr(34) & param2 & Chr(34) & _
                ") can only be:" & Chr(13) & _
                Chr(34) & "Px" & Chr(34) & " (specified Pile with number)" & Chr(13) & _
                Chr(34) & "R" & Chr(34) & " (Random pile)" & Chr(13) & _
                Chr(34) & "S" & Chr(34) & " (Pile with a selected card)" & Chr(13) & _
                Chr(34) & "N" & Chr(34) & " (pile with No selected card)" & Chr(13) & _
                Chr(13) & "The above parameters may have an " & Chr(34) & "R" & Chr(34) & Chr(13) & _
                "suffix which indicates the pile should be Reversed.")
            Exit Sub
        End If
    ElseIf Left(param2, 1) = "S" Then
        If pNumberPilesWithSelected = 0 Then
            If pPileHasSelectedCard(pLeftPile) Then
                MsgBox ("Error: Swap Piles" & Chr(13) & _
                    "The second parameter (" & Chr(34) & param2 & Chr(34) & _
                    ") requires there to be a selected card." & Chr(13) & _
                    "However, there are no selected cards available, since" & Chr(13) & _
                    "they are all in the Pile assigned to the first Swap Pile.")
            Else
                MsgBox ("Error: Swap Piles" & Chr(13) & _
                    "The second parameter (" & Chr(34) & param2 & Chr(34) & _
                    ") requires there to be a selected card." & Chr(13) & _
                    "However, there are no selected cards available.")
            End If
            Exit Sub
        Else
            'new section for selected card piles
            pTemp = 0
            While pTemp = 0
                pRandomPile = Int(Rnd * NumPiles) + 1
                If pPileHasSelectedCard(pRandomPile) And pRandomPile <> pLeftPile Then
                    pTemp = 99
                    'any non-zero value
                    pRightPile = pRandomPile
                    pNumberPilesWithSelected = pNumberPilesWithSelected - 1
                End If
            Wend
            'end of new section
'            'establish pile based on number of selections
'            If pNumberSelectedCards = 1 Then
'                'make sure the first pile is logically correct relative to selections
'                If Not pLeftContainsSelected Then
'                    pRightPile = pSelectedPile1
'                Else
'                    MsgBox ("Error: Swap Piles" & Chr(13) & _
'                        "The second parameter (" & Chr(34) & param2 & Chr(34) & _
'                        ") requires there to be a selected card." & Chr(13) & _
'                        "However, the first pile already contains the selected card.")
'                    Exit Sub
'                End If
'            ElseIf pNumberSelectedCards = 2 Then
'                'make sure the first pile is logically correct relative to selections
'                If Not pLeftContainsSelected Then
'                    'establish a random selection (<0.5 then S1, else S2)
'                    If Rnd < 0.5 Then
'                        pRightPile = pSelectedPile1
'                    Else
'                        pRightPile = pSelectedPile2
'                    End If
'                Else
'                    'make sure both selections are not in the first pile
'                    If pSelectedPile1 = pSelectedPile2 Then
'                        MsgBox ("Error: Swap Piles" & Chr(13) & _
'                        "The second parameter (" & Chr(34) & param2 & Chr(34) & _
'                        ") requires there to be a selected card." & Chr(13) & _
'                        "However, the first pile already contains both selected cards." & Chr(13) & _
'                        Chr(13) & _
'                        "The selected card text string is:" & Chr(13) & _
'                        frmStackView.SelectionsTextBox.Text & Chr(13) & _
'                        "The Swap Piles procedure uses only the first two selections.")
'                        Exit Sub
'                    End If
'                    'establish the pile with the non-claimed selected card
'                    If pLeftPile = pSelectedPile1 Then
'                        pRightPile = pSelectedPile2
'                    ElseIf pLeftPile = pSelectedPile2 Then
'                        pRightPile = pSelectedPile1
'                    Else
'                        MsgBox ("An illogical error condition exists." & Chr(13) & _
'                        "The first pile contains a selection, but it is not recognized.")
'                    End If
'                End If
'            End If
            
            'check for a valid second parameter
            If Right(param2, 1) = "R" Then
                pRightReverse = True
            Else
                MsgBox ("Error: Swap Piles" & Chr(13) & _
                    "The second parameter (" & Chr(34) & param2 & Chr(34) & _
                    ") can only be:" & Chr(13) & _
                    Chr(34) & "Px" & Chr(34) & " (specified Pile with number)" & Chr(13) & _
                    Chr(34) & "R" & Chr(34) & " (Random pile)" & Chr(13) & _
                    Chr(34) & "S" & Chr(34) & " (Pile with a selected card)" & Chr(13) & _
                    Chr(34) & "N" & Chr(34) & " (pile with No selected card)" & Chr(13) & _
                    Chr(13) & "The above parameters may have an " & Chr(34) & "R" & Chr(34) & Chr(13) & _
                    "suffix which indicates the pile should be Reversed.")
                Exit Sub
            End If
        End If
    ElseIf Left(param2, 1) = "N" Then
        'new section
        'make sure there are piles without selections
        If pNumberPilesWithoutSelected = 0 Then
            If Not pPileHasSelectedCard(pLeftPile) Then
                MsgBox ("Error: Swap Piles" & Chr(13) & _
                    "The second parameter (" & Chr(34) & param2 & Chr(34) & _
                    ") requires there to be no selected cards." & Chr(13) & _
                    "However, there are no Piles available, since the only" & Chr(13) & _
                    "Pile without any selections was assigned to the first Swap Pile.")
            Else
                MsgBox ("Error: Swap Piles" & Chr(13) & _
                    "The second parameter (" & Chr(34) & param2 & Chr(34) & _
                    ") requires no selected cards." & Chr(13) & _
                    "However, there are no remaining piles without any selections.")
            End If
            Exit Sub
        End If
        'find a pile without a selected card
        pTemp = 0
        While pTemp = 0
            pRandomPile = Int(Rnd * NumPiles) + 1
            If Not pPileHasSelectedCard(pRandomPile) And pRandomPile <> pLeftPile Then
                pTemp = 99
                'any non-zero value
                pRightPile = pRandomPile
                pNumberPilesWithoutSelected = pNumberPilesWithoutSelected - 1
            End If
        Wend
        'end of new section
        
'        'make sure there are piles without selections before
'        'running the While...Wend loop so that it doesn't crash
'        If pNumberSelectedCards > 0 Then
'            If pNumberSelectedCards = 1 And Not pLeftContainsSelected Then
'                If NumPiles = 2 Then
'                    MsgBox ("Error: Swap Piles" & Chr(13) & _
'                        "The first parameter (" & Chr(34) & param1 & Chr(34) & _
'                        ") resulted with no selected card." & Chr(13) & _
'                        "The second parameter (" & Chr(34) & param2 & Chr(34) & _
'                        ") requires no selected card." & Chr(13) & _
'                        "However, there are no remaining piles without a selection.")
'                    Exit Sub
'                End If
'            ElseIf pNumberSelectedCards = 2 And Not pLeftContainsSelected Then
'                'consider if both selections are in the same pile
'                If pSelectedPile1 = pSelectedPile2 Then
'                    If NumPiles = 2 Then
'                        MsgBox ("Error: Swap Piles" & Chr(13) & _
'                            "The first parameter (" & Chr(34) & param1 & Chr(34) & _
'                            ") resulted with no selected card." & Chr(13) & _
'                            "The second parameter (" & Chr(34) & param2 & Chr(34) & _
'                            ") requires no selected card." & Chr(13) & _
'                            "However, there are no remaining piles without a selection.")
'                        Exit Sub
'                    End If
'                Else
'                    'in this case, the two selections are in different piles
'                    If NumPiles < 4 Then
'                        MsgBox ("Error: Swap Piles" & Chr(13) & _
'                            "The first parameter (" & Chr(34) & param1 & Chr(34) & _
'                            ") resulted with no selected card." & Chr(13) & _
'                            "The second parameter (" & Chr(34) & param2 & Chr(34) & _
'                            ") requires no selected card." & Chr(13) & _
'                            "However, there are no remaining piles without a selection.")
'                        Exit Sub
'                    End If
'                End If
'            End If
'        End If
'        'find a pile without a valid selection
'        pNonSelectedPileFound = False
'        While Not pNonSelectedPileFound
'            pRandomPile = Int(Rnd * NumPiles) + 1
'            If pRandomPile <> pSelectedPile1 And _
'                pRandomPile <> pSelectedPile2 And _
'                pRandomPile <> pLeftPile Then
'                pRightPile = pRandomPile
'                pNonSelectedPileFound = True
'            End If
'        Wend
        
        'check for a valid second parameter
        If Right(param2, 1) = "R" Then
            pRightReverse = True
        Else
            MsgBox ("Error: Swap Piles" & Chr(13) & _
                "The second parameter (" & Chr(34) & param2 & Chr(34) & _
                ") can only be:" & Chr(13) & _
                Chr(34) & "Px" & Chr(34) & " (specified Pile with number)" & Chr(13) & _
                Chr(34) & "R" & Chr(34) & " (Random pile)" & Chr(13) & _
                Chr(34) & "S" & Chr(34) & " (Pile with a selected card)" & Chr(13) & _
                Chr(34) & "N" & Chr(34) & " (pile with No selected card)" & Chr(13) & _
                Chr(13) & "The above parameters may have an " & Chr(34) & "R" & Chr(34) & Chr(13) & _
                "suffix which indicates the pile should be Reversed.")
            Exit Sub
        End If
    Else
        MsgBox ("Error: Swap Piles" & Chr(13) & _
            "The second parameter (" & Chr(34) & param2 & Chr(34) & _
            ") can only be:" & Chr(13) & _
            Chr(34) & "Px" & Chr(34) & " (specified Pile with number)" & Chr(13) & _
            Chr(34) & "R" & Chr(34) & " (Random pile)" & Chr(13) & _
            Chr(34) & "S" & Chr(34) & " (Pile with a selected card)" & Chr(13) & _
            Chr(34) & "N" & Chr(34) & " (pile with No selected card)" & Chr(13) & _
            Chr(13) & "The above parameters may have an " & Chr(34) & "R" & Chr(34) & Chr(13) & _
            "suffix which indicates the pile should be Reversed.")
        Exit Sub
    End If
ElseIf Len(param2) = 1 Then
    If param2 = "R" Then
        'set the second pile to a random number
        'but make sure to avoid the first pile
        pValidSecondPileFound = False
        While Not pValidSecondPileFound
            pRightPile = Int(Rnd * NumPiles + 1)
            If pLeftPile <> pRightPile Then
                pValidSecondPileFound = True
            End If
        Wend
    ElseIf param2 = "S" Then
        If pNumberPilesWithSelected = 0 Then
            If pPileHasSelectedCard(pLeftPile) Then
                MsgBox ("Error: Swap Piles" & Chr(13) & _
                    "The second parameter (" & Chr(34) & param2 & Chr(34) & _
                    ") requires there to be a selected card." & Chr(13) & _
                    "However, there are no selected cards available, since" & Chr(13) & _
                    "they are all in the Pile assigned to the first Swap Pile.")
            Else
                MsgBox ("Error: Swap Piles" & Chr(13) & _
                    "The second parameter (" & Chr(34) & param2 & Chr(34) & _
                    ") requires there to be a selected card." & Chr(13) & _
                    "However, there are no selected cards available.")
            End If
            Exit Sub
        Else
            'new section for selected card piles
            pTemp = 0
            While pTemp = 0
                pRandomPile = Int(Rnd * NumPiles) + 1
                If pPileHasSelectedCard(pRandomPile) And pRandomPile <> pLeftPile Then
                    pTemp = 99
                    'any non-zero value
                    pRightPile = pRandomPile
                    pNumberPilesWithSelected = pNumberPilesWithSelected - 1
                End If
            Wend
            
'            'establish pile based on number of selections
'            If pNumberSelectedCards = 1 Then
'                'make sure the first pile is logically correct relative to selections
'                If Not pLeftContainsSelected Then
'                    pRightPile = pSelectedPile1
'                Else
'                    MsgBox ("Error: Swap Piles" & Chr(13) & _
'                        "The second parameter (" & Chr(34) & param2 & Chr(34) & _
'                        ") requires there to be a selected card." & Chr(13) & _
'                        "However, the first pile already contains the selected card.")
'                    Exit Sub
'                End If
'            ElseIf pNumberSelectedCards = 2 Then
'                'make sure the first pile is logically correct relative to selections
'                If Not pLeftContainsSelected Then
'                    'establish a random selection (<0.5 then S1, else S2)
'                    If Rnd < 0.5 Then
'                        pRightPile = pSelectedPile1
'                    Else
'                        pRightPile = pSelectedPile2
'                    End If
'                Else
'                    'make sure both selections are not in the first pile
'                    If pSelectedPile1 = pSelectedPile2 Then
'                        MsgBox ("Error: Swap Piles" & Chr(13) & _
'                        "The second parameter (" & Chr(34) & param2 & Chr(34) & _
'                        ") requires there to be a selected card." & Chr(13) & _
'                        "However, the first pile already contains both selected cards." & Chr(13) & _
'                        Chr(13) & _
'                        "The selected card text string is:" & Chr(13) & _
'                        frmStackView.SelectionsTextBox.Text & Chr(13) & _
'                        "The Swap Piles procedure uses only the first two selections.")
'                        Exit Sub
'                    End If
'                    'establish the pile with the non-claimed selected card
'                    If pLeftPile = pSelectedPile1 Then
'                        pRightPile = pSelectedPile2
'                    ElseIf pLeftPile = pSelectedPile2 Then
'                        pRightPile = pSelectedPile1
'                    Else
'                        MsgBox ("An illogical error condition exists." & Chr(13) & _
'                        "The first pile contains a selection, but it is not recognized.")
'                    End If
'                End If
'            End If
        End If
    
    ElseIf param2 = "N" Then
        'new section
        'make sure there are piles without selections
        If pNumberPilesWithoutSelected = 0 Then
            If Not pPileHasSelectedCard(pLeftPile) Then
                MsgBox ("Error: Swap Piles" & Chr(13) & _
                    "The second parameter (" & Chr(34) & param2 & Chr(34) & _
                    ") requires there to be no selected cards." & Chr(13) & _
                    "However, there are no Piles available, since the only" & Chr(13) & _
                    "Pile without any selections was assigned to the first Swap Pile.")
            Else
                MsgBox ("Error: Swap Piles" & Chr(13) & _
                    "The second parameter (" & Chr(34) & param2 & Chr(34) & _
                    ") requires no selected cards." & Chr(13) & _
                    "However, there are no remaining piles without any selections.")
            End If
            Exit Sub
        End If
        'find a pile without a selected card
        pTemp = 0
        While pTemp = 0
            pRandomPile = Int(Rnd * NumPiles) + 1
            If Not pPileHasSelectedCard(pRandomPile) And pRandomPile <> pLeftPile Then
                pTemp = 99
                'any non-zero value
                pRightPile = pRandomPile
                pNumberPilesWithoutSelected = pNumberPilesWithoutSelected - 1
            End If
        Wend
        'end of new section
        
'        'make sure there are piles without selections before
'        'running the While...Wend loop so that it doesn't crash
'        If pNumberSelectedCards > 0 Then
'            If pNumberSelectedCards = 1 And Not pLeftContainsSelected Then
'                If NumPiles = 2 Then
'                    MsgBox ("Error: Swap Piles" & Chr(13) & _
'                        "The first parameter (" & Chr(34) & param1 & Chr(34) & _
'                        ") resulted with no selected card." & Chr(13) & _
'                        "The second parameter (" & Chr(34) & param2 & Chr(34) & _
'                        ") requires no selected card." & Chr(13) & _
'                        "However, there are no remaining piles without a selection.")
'                    Exit Sub
'                End If
'            ElseIf pNumberSelectedCards = 2 And Not pLeftContainsSelected Then
'                'consider if both selections are in the same pile
'                If pSelectedPile1 = pSelectedPile2 Then
'                    If NumPiles = 2 Then
'                        MsgBox ("Error: Swap Piles" & Chr(13) & _
'                            "The first parameter (" & Chr(34) & param1 & Chr(34) & _
'                            ") resulted with no selected card." & Chr(13) & _
'                            "The second parameter (" & Chr(34) & param2 & Chr(34) & _
'                            ") requires no selected card." & Chr(13) & _
'                            "However, there are no remaining piles without a selection.")
'                        Exit Sub
'                    End If
'                Else
'                    'in this case, the two selections are in different piles
'                    If NumPiles < 4 Then
'                        MsgBox ("Error: Swap Piles" & Chr(13) & _
'                            "The first parameter (" & Chr(34) & param1 & Chr(34) & _
'                            ") resulted with no selected card." & Chr(13) & _
'                            "The second parameter (" & Chr(34) & param2 & Chr(34) & _
'                            ") requires no selected card." & Chr(13) & _
'                            "However, there are no remaining piles without a selection.")
'                        Exit Sub
'                    End If
'                End If
'            End If
'        End If
'        'find a pile without a valid selection
'        pNonSelectedPileFound = False
'        While Not pNonSelectedPileFound
'            pRandomPile = Int(Rnd * NumPiles) + 1
'            If pRandomPile <> pSelectedPile1 And _
'                pRandomPile <> pSelectedPile2 And _
'                pRandomPile <> pLeftPile Then
'                pRightPile = pRandomPile
'                pNonSelectedPileFound = True
'            End If
'        Wend
    Else
        MsgBox ("Error: Swap Piles" & Chr(13) & _
            "The second parameter (" & Chr(34) & param2 & Chr(34) & _
            ") can only be:" & Chr(13) & _
            Chr(34) & "Px" & Chr(34) & " (specified Pile with number)" & Chr(13) & _
            Chr(34) & "R" & Chr(34) & " (Random pile)" & Chr(13) & _
            Chr(34) & "S" & Chr(34) & " (Pile with a selected card)" & Chr(13) & _
            Chr(34) & "N" & Chr(34) & " (pile with No selected card)" & Chr(13) & _
            Chr(13) & "The above parameters may have an " & Chr(34) & "R" & Chr(34) & Chr(13) & _
            "suffix which indicates the pile should be Reversed.")
        Exit Sub
    End If
End If
'need to identify the errors for the second parameter wherever pSwapPilesError=True
'reverse the left pile if required
If pLeftReverse Then
    For j% = PileTable(pLeftPile, 1) To PileTable(pLeftPile, 2)
        For p% = 1 To DeckProperties
            ChangedDeck(p%, j%) = Deck(p%, PileTable(pLeftPile, 2) _
                - (j% - PileTable(pLeftPile, 1)))
        Next p%
    Next j%
    For m% = PileTable(pLeftPile, 1) To PileTable(pLeftPile, 2)
        For z% = 1 To DeckProperties
            Deck(z%, m%) = ChangedDeck(z%, m%)
        Next z%
    Next m%
    For i% = PileTable(pLeftPile, 1) To PileTable(pLeftPile, 2)
        Deck(6, i%) = Not Deck(6, i%)
    Next i%
End If
'reverse the right pile if required
If pRightReverse Then
    For j% = PileTable(pRightPile, 1) To PileTable(pRightPile, 2)
        For p% = 1 To DeckProperties
            ChangedDeck(p%, j%) = Deck(p%, PileTable(pRightPile, 2) _
                - (j% - PileTable(pRightPile, 1)))
        Next p%
    Next j%
    For m% = PileTable(pRightPile, 1) To PileTable(pRightPile, 2)
        For z% = 1 To DeckProperties
            Deck(z%, m%) = ChangedDeck(z%, m%)
        Next z%
    Next m%
    For i% = PileTable(pRightPile, 1) To PileTable(pRightPile, 2)
        Deck(6, i%) = Not Deck(6, i%)
    Next i%
End If
'ensure that pLeftPile < pRightPile
'since it doesn't matter which pile is the smaller value
'and they are being swapped
If pLeftPile > pRightPile Then
    pTemp = pLeftPile
    pLeftPile = pRightPile
    pRightPile = pTemp
End If
'now just swap the piles
Dim pLeftPileCards As Integer
Dim pRightPileCards As Integer
Dim pOffset As Integer
Dim pPosition As Integer
pLeftPileCards = PileTable(pLeftPile, 2) - PileTable(pLeftPile, 1) + 1
pRightPileCards = PileTable(pRightPile, 2) - PileTable(pRightPile, 1) + 1
pOffset = pLeftPileCards - pRightPileCards
pPosition = 1
'transfer cards up to the first pile
For i% = 1 To pLeftPile
    If i% <> pLeftPile Then
        sPileTable(i%, 1) = PileTable(i%, 1)
        sPileTable(i%, 2) = PileTable(i%, 2)
        pPosition = PileTable(i%, 2) + 1
        For k% = 1 To DeckProperties
            For m% = sPileTable(i%, 1) To sPileTable(i%, 2)
                ChangedDeck(k%, m%) = Deck(k%, m%)
            Next m%
        Next k%
    Else
        sPileTable(i%, 1) = pPosition
        sPileTable(i%, 2) = pPosition + pRightPileCards - 1
        For k% = 1 To DeckProperties
            For m% = 1 To pRightPileCards
                ChangedDeck(k%, pPosition + m% - 1) = Deck(k%, PileTable(pRightPile, 1) + m% - 1)
            Next m%
        Next k%
        pPosition = sPileTable(i%, 2) + 1
    End If
Next i%
'transfer cards up to the second pile
For i% = pLeftPile + 1 To pRightPile
    If i% <> pRightPile Then
        sPileTable(i%, 1) = PileTable(i%, 1) - pOffset
        sPileTable(i%, 2) = PileTable(i%, 2) - pOffset
        pPosition = PileTable(i%, 2) + 1 - pOffset
        For k% = 1 To DeckProperties
            For m% = sPileTable(i%, 1) To sPileTable(i%, 2)
                ChangedDeck(k%, m%) = Deck(k%, m% + pOffset)
            Next m%
        Next k%
    Else
        sPileTable(i%, 1) = pPosition
        sPileTable(i%, 2) = pPosition + pLeftPileCards - 1
        For k% = 1 To DeckProperties
            For m% = 1 To pLeftPileCards
                ChangedDeck(k%, pPosition + m% - 1) = Deck(k%, PileTable(pLeftPile, 1) + m% - 1)
            Next m%
        Next k%
        pPosition = sPileTable(i%, 2) + 1
    End If
Next i%
'finish transfer if there are any remaining piles
If pRightPile < NumPiles Then
    For i% = pRightPile + 1 To NumPiles
        sPileTable(i%, 1) = PileTable(i%, 1)
        sPileTable(i%, 2) = PileTable(i%, 2)
        pPosition = PileTable(i%, 2) + 1
        For k% = 1 To DeckProperties
            For m% = sPileTable(i%, 1) To sPileTable(i%, 2)
                ChangedDeck(k%, m%) = Deck(k%, m%)
            Next m%
        Next k%
    Next i%
End If
'reassemble PileTable
For i% = 1 To NumPiles
    PileTable(i%, 1) = sPileTable(i%, 1)
    PileTable(i%, 2) = sPileTable(i%, 2)
Next i%
'reassemble Deck
For k% = 1 To DeckProperties
    For m% = 1 To DeckCount
        Deck(k%, m%) = ChangedDeck(k%, m%)
    Next m%
Next k%
PilesMatrixRefresh
CreatePiles (NumPiles)
End Sub

Private Sub ViewPilesAbove_Click()
If PilesShown = 1 Then
    If GilbreathActive Then
        If GilbreathShown Then
            frmDeck.DisplayPilesGilbreath
        Else
            frmDeck.DisplayPilesKeepGilbreathActive
        End If
    Else
        ShowPiles
    End If
End If
End Sub

Private Sub ViewPilesBeneath_Click()
If PilesShown = 1 Then
    If GilbreathActive Then
        If GilbreathShown Then
            frmDeck.DisplayPilesGilbreath
        Else
            frmDeck.DisplayPilesKeepGilbreathActive
        End If
    Else
        ShowPiles
    End If
End If
End Sub

