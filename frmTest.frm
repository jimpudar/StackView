VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmTest 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "StackView Test"
   ClientHeight    =   6630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8745
   Icon            =   "frmTest.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   8745
   Visible         =   0   'False
   Begin VB.CheckBox MnemonicsHintCheckBox 
      Caption         =   "Auto Hint"
      Enabled         =   0   'False
      Height          =   285
      Left            =   2205
      TabIndex        =   44
      Top             =   5925
      Width           =   1020
   End
   Begin VB.CheckBox MnemonicsEnabled 
      Caption         =   "Enable Mnemonics"
      Height          =   285
      Left            =   420
      TabIndex        =   43
      Top             =   5940
      Width           =   1740
   End
   Begin VB.Frame Frame1 
      Caption         =   "Test Card"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Left            =   1890
      TabIndex        =   39
      Top             =   3045
      Width           =   1680
      Begin VB.OptionButton TestRandomCardOption 
         Caption         =   "Random Card"
         Height          =   270
         Left            =   120
         TabIndex        =   40
         Top             =   1140
         Width           =   1410
      End
      Begin VB.OptionButton TestCurrentCardOption 
         Caption         =   "Current Card"
         Height          =   270
         Left            =   120
         TabIndex        =   10
         Top             =   285
         Value           =   -1  'True
         Width           =   1200
      End
      Begin VB.OptionButton TestNextCardOption 
         Caption         =   "Next Card"
         Height          =   270
         Left            =   120
         TabIndex        =   11
         Top             =   570
         Width           =   1230
      End
      Begin VB.OptionButton TestPreviousCardOption 
         Caption         =   "Previous Card"
         Height          =   270
         Left            =   120
         TabIndex        =   12
         Top             =   855
         Width           =   1410
      End
   End
   Begin VB.Frame DeckRangeFrame 
      Caption         =   "Deck Range"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   285
      TabIndex        =   36
      Top             =   345
      Width           =   3540
      Begin VB.OptionButton DeckCurrentRangePartial 
         Caption         =   "Current Deck Position Range"
         Height          =   270
         Left            =   120
         TabIndex        =   41
         Top             =   870
         Width           =   2490
      End
      Begin VB.TextBox DeckRangePartialStartTextBox 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2715
         TabIndex        =   2
         Top             =   315
         Width           =   600
      End
      Begin VB.TextBox DeckRangePartialFinishTextBox 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2730
         TabIndex        =   3
         Top             =   840
         Width           =   600
      End
      Begin VB.OptionButton DeckRangePartial 
         Caption         =   "Stack Value Range"
         Height          =   270
         Left            =   120
         TabIndex        =   1
         Top             =   570
         Width           =   2490
      End
      Begin VB.OptionButton DeckRangeFull 
         Caption         =   "Full Deck"
         Height          =   270
         Left            =   120
         TabIndex        =   0
         Top             =   285
         Value           =   -1  'True
         Width           =   1200
      End
      Begin VB.Label Label2 
         Caption         =   "Finish"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2745
         TabIndex        =   38
         Top             =   630
         Width           =   690
      End
      Begin VB.Label Label1 
         Caption         =   "Start"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2745
         TabIndex        =   37
         Top             =   120
         Width           =   585
      End
   End
   Begin MSComctlLib.ProgressBar ProgressCardsRemaining 
      Height          =   2580
      Left            =   7875
      TabIndex        =   32
      Top             =   150
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   4551
      _Version        =   393216
      Appearance      =   1
      Max             =   52
      Orientation     =   1
      Scrolling       =   1
   End
   Begin VB.Timer TimerShow 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   8235
      Top             =   540
   End
   Begin VB.Timer TimerTest 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   8235
      Top             =   90
   End
   Begin VB.CheckBox TestTimersEnabled 
      Caption         =   "Enable Timers"
      Height          =   285
      Left            =   435
      TabIndex        =   15
      Top             =   5640
      Value           =   1  'Checked
      Width           =   1995
   End
   Begin VB.Frame Frame3 
      Caption         =   "Test Times"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1050
      Left            =   285
      TabIndex        =   23
      Top             =   4530
      Width           =   3300
      Begin VB.TextBox ShowDuration 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1275
         TabIndex        =   14
         Text            =   "2"
         Top             =   630
         Width           =   600
      End
      Begin VB.TextBox TestDuration 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1275
         TabIndex        =   13
         Text            =   "5"
         Top             =   255
         Width           =   600
      End
      Begin VB.Label Label6 
         Caption         =   "seconds"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1905
         TabIndex        =   27
         Top             =   660
         Width           =   885
      End
      Begin VB.Label Label5 
         Caption         =   "seconds"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1905
         TabIndex        =   26
         Top             =   315
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Show Duration"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   75
         TabIndex        =   25
         Top             =   660
         Width           =   1155
      End
      Begin VB.Label Label3 
         Caption         =   "Test Duration"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   75
         TabIndex        =   24
         Top             =   315
         Width           =   1050
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Test Value"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1170
      Left            =   285
      TabIndex        =   20
      Top             =   3045
      Width           =   1545
      Begin VB.OptionButton TestRandom 
         Caption         =   "Random Mix"
         Height          =   270
         Left            =   120
         TabIndex        =   9
         Top             =   855
         Value           =   -1  'True
         Width           =   1245
      End
      Begin VB.OptionButton TestPosition 
         Caption         =   "Stack Value"
         Height          =   270
         Left            =   120
         TabIndex        =   8
         Top             =   570
         Width           =   1230
      End
      Begin VB.OptionButton TestValue 
         Caption         =   "Card Value"
         Height          =   270
         Left            =   120
         TabIndex        =   7
         Top             =   285
         Width           =   1200
      End
   End
   Begin VB.Frame SequenceFrame 
      Caption         =   "Stack Sequence"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1485
      Left            =   285
      TabIndex        =   19
      Top             =   1575
      Width           =   3150
      Begin VB.OptionButton SequenceCurrent 
         Caption         =   "Current Order"
         Height          =   270
         Left            =   120
         TabIndex        =   42
         Top             =   870
         Width           =   2040
      End
      Begin VB.OptionButton SequenceForward 
         Caption         =   "Forward"
         Height          =   270
         Left            =   120
         TabIndex        =   4
         Top             =   285
         Width           =   2040
      End
      Begin VB.OptionButton SequenceBackward 
         Caption         =   "Backward"
         Height          =   270
         Left            =   120
         TabIndex        =   5
         Top             =   570
         Width           =   2040
      End
      Begin VB.OptionButton SequenceRandom 
         Caption         =   "Random"
         Height          =   270
         Left            =   120
         TabIndex        =   6
         Top             =   1155
         Value           =   -1  'True
         Width           =   2040
      End
   End
   Begin MSComctlLib.ProgressBar ProgressTest 
      Height          =   255
      Left            =   4350
      TabIndex        =   16
      Top             =   2850
      Width           =   3285
      _ExtentX        =   5794
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar ProgressShow 
      Height          =   255
      Left            =   4350
      TabIndex        =   17
      Top             =   3195
      Width           =   3285
      _ExtentX        =   5794
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Image HintButton 
      Height          =   585
      Index           =   1
      Left            =   4365
      Picture         =   "frmTest.frx":1CCA
      Top             =   3720
      Width           =   825
   End
   Begin VB.Image HintButton 
      Height          =   585
      Index           =   0
      Left            =   4365
      Picture         =   "frmTest.frx":2874
      Top             =   3720
      Width           =   825
   End
   Begin VB.Label MnemonicRight 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Mnemonic"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6225
      TabIndex        =   46
      Top             =   2310
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.Label MnemonicLeft 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Mnemonic"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4275
      TabIndex        =   45
      Top             =   2310
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.Image TestRightImage 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   5640
      Picture         =   "frmTest.frx":3583
      Top             =   1665
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Image TestLeftImage 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5655
      Picture         =   "frmTest.frx":3B53
      Top             =   1650
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image TestPreviousImage 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   5655
      Picture         =   "frmTest.frx":4107
      Top             =   1095
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Image TestNextImage 
      Appearance      =   0  'Flat
      Height          =   510
      Left            =   5640
      Picture         =   "frmTest.frx":452F
      Top             =   975
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Image TestCurrentImage 
      Appearance      =   0  'Flat
      Height          =   450
      Left            =   5640
      Picture         =   "frmTest.frx":4BA5
      Top             =   990
      Width           =   540
   End
   Begin VB.Image TestNext 
      Height          =   600
      Index           =   2
      Left            =   6525
      Picture         =   "frmTest.frx":5187
      Top             =   4410
      Visible         =   0   'False
      Width           =   1740
   End
   Begin VB.Image TestShow 
      Height          =   600
      Index           =   3
      Left            =   4530
      Picture         =   "frmTest.frx":6477
      Top             =   4410
      Visible         =   0   'False
      Width           =   1740
   End
   Begin VB.Image TestShow 
      Height          =   600
      Index           =   2
      Left            =   4530
      Picture         =   "frmTest.frx":76A3
      Top             =   4410
      Visible         =   0   'False
      Width           =   1740
   End
   Begin VB.Image TestNext 
      Height          =   600
      Index           =   5
      Left            =   6525
      Picture         =   "frmTest.frx":8A01
      Top             =   4410
      Visible         =   0   'False
      Width           =   1740
   End
   Begin VB.Image TestNext 
      Height          =   600
      Index           =   4
      Left            =   6525
      Picture         =   "frmTest.frx":9D83
      Top             =   4410
      Visible         =   0   'False
      Width           =   1740
   End
   Begin VB.Image TestNext 
      Height          =   600
      Index           =   3
      Left            =   6525
      Picture         =   "frmTest.frx":B0D0
      Top             =   4410
      Visible         =   0   'False
      Width           =   1740
   End
   Begin VB.Label TestCardsLabel2 
      Alignment       =   2  'Center
      Caption         =   "Remaining"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7680
      TabIndex        =   35
      Top             =   3225
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label TestCardsLabel1 
      Alignment       =   2  'Center
      Caption         =   "Cards"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7680
      TabIndex        =   34
      Top             =   3045
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label TestCardsRemaining 
      Alignment       =   2  'Center
      Caption         =   "52"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   7845
      TabIndex        =   33
      Top             =   2820
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label ResultTime 
      Caption         =   "104 min. 60 sec."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   345
      Left            =   5775
      TabIndex        =   31
      Top             =   5790
      Visible         =   0   'False
      Width           =   2475
   End
   Begin VB.Label ResultLabel 
      Caption         =   "Time: "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   345
      Left            =   4845
      TabIndex        =   30
      Top             =   5790
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label TestMessage 
      Caption         =   "Test uses current stack from Deck window"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   165
      TabIndex        =   29
      Top             =   6300
      Width           =   3960
   End
   Begin VB.Image TestC0 
      Height          =   1920
      Left            =   6225
      Picture         =   "frmTest.frx":C416
      Tag             =   "TCard"
      Top             =   300
      Width           =   1350
   End
   Begin VB.Label TestingStatus 
      Alignment       =   2  'Center
      Caption         =   "Not Testing"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   420
      Left            =   5190
      TabIndex        =   28
      Top             =   5280
      Width           =   2355
   End
   Begin VB.Image TestNext 
      Height          =   600
      Index           =   0
      Left            =   6525
      Picture         =   "frmTest.frx":D7E8
      Top             =   4410
      Width           =   1740
   End
   Begin VB.Image TestShow 
      Height          =   600
      Index           =   0
      Left            =   4530
      Picture         =   "frmTest.frx":E16C
      Top             =   4410
      Width           =   1740
   End
   Begin VB.Image TestToggle 
      Height          =   600
      Index           =   0
      Left            =   5460
      Picture         =   "frmTest.frx":EB81
      Top             =   3750
      Width           =   1740
   End
   Begin VB.Image TestCKD 
      Height          =   1920
      Left            =   6225
      Picture         =   "frmTest.frx":F4FC
      Tag             =   "TCard"
      Top             =   300
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image TestCQD 
      Height          =   1920
      Left            =   6225
      Picture         =   "frmTest.frx":10528
      Tag             =   "TCard"
      Top             =   300
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image TestCJD 
      Height          =   1920
      Left            =   6225
      Picture         =   "frmTest.frx":11596
      Tag             =   "TCard"
      Top             =   300
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image TestC10D 
      Height          =   1920
      Left            =   6225
      Picture         =   "frmTest.frx":125D4
      Tag             =   "TCard"
      Top             =   300
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image TestC9D 
      Height          =   1920
      Left            =   6225
      Picture         =   "frmTest.frx":13262
      Tag             =   "TCard"
      Top             =   300
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image TestC8D 
      Height          =   1920
      Left            =   6225
      Picture         =   "frmTest.frx":13E94
      Tag             =   "TCard"
      Top             =   300
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image TestC7D 
      Height          =   1920
      Left            =   6225
      Picture         =   "frmTest.frx":14A41
      Tag             =   "TCard"
      Top             =   300
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image TestC6D 
      Height          =   1920
      Left            =   6225
      Picture         =   "frmTest.frx":15535
      Tag             =   "TCard"
      Top             =   300
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image TestC5D 
      Height          =   1920
      Left            =   6225
      Picture         =   "frmTest.frx":15FA7
      Tag             =   "TCard"
      Top             =   300
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image TestC4D 
      Height          =   1920
      Left            =   6225
      Picture         =   "frmTest.frx":16989
      Tag             =   "TCard"
      Top             =   300
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image TestC3D 
      Height          =   1920
      Left            =   6225
      Picture         =   "frmTest.frx":172B9
      Tag             =   "TCard"
      Top             =   300
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image TestC2D 
      Height          =   1920
      Left            =   6225
      Picture         =   "frmTest.frx":17BA9
      Tag             =   "TCard"
      Top             =   300
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image TestCAD 
      Height          =   1920
      Left            =   6225
      Picture         =   "frmTest.frx":183EB
      Tag             =   "TCard"
      Top             =   300
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image TestCKS 
      Height          =   1920
      Left            =   6225
      Picture         =   "frmTest.frx":18BEC
      Tag             =   "TCard"
      Top             =   300
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image TestCQS 
      Height          =   1920
      Left            =   6225
      Picture         =   "frmTest.frx":19C3A
      Tag             =   "TCard"
      Top             =   300
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image TestCJS 
      Height          =   1920
      Left            =   6225
      Picture         =   "frmTest.frx":1AC39
      Tag             =   "TCard"
      Top             =   300
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image TestC10S 
      Height          =   1920
      Left            =   6225
      Picture         =   "frmTest.frx":1BBB3
      Tag             =   "TCard"
      Top             =   300
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image TestC9S 
      Height          =   1920
      Left            =   6225
      Picture         =   "frmTest.frx":1C925
      Tag             =   "TCard"
      Top             =   300
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image TestC8S 
      Height          =   1920
      Left            =   6225
      Picture         =   "frmTest.frx":1D5F5
      Tag             =   "TCard"
      Top             =   300
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image TestC7S 
      Height          =   1920
      Left            =   6225
      Picture         =   "frmTest.frx":1E244
      Tag             =   "TCard"
      Top             =   300
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image TestC6S 
      Height          =   1920
      Left            =   6225
      Picture         =   "frmTest.frx":1EDB7
      Tag             =   "TCard"
      Top             =   300
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image TestC5S 
      Height          =   1920
      Left            =   6225
      Picture         =   "frmTest.frx":1F8AB
      Tag             =   "TCard"
      Top             =   300
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image TestC4S 
      Height          =   1920
      Left            =   6225
      Picture         =   "frmTest.frx":2035A
      Tag             =   "TCard"
      Top             =   300
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image TestC3S 
      Height          =   1920
      Left            =   6225
      Picture         =   "frmTest.frx":20D48
      Tag             =   "TCard"
      Top             =   300
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image TestC2S 
      Height          =   1920
      Left            =   6225
      Picture         =   "frmTest.frx":2164D
      Tag             =   "TCard"
      Top             =   300
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image TestCAS 
      Height          =   1920
      Left            =   6225
      Picture         =   "frmTest.frx":21EE7
      Tag             =   "TCard"
      Top             =   300
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image TestCKH 
      Height          =   1920
      Left            =   6225
      Picture         =   "frmTest.frx":227F0
      Tag             =   "TCard"
      Top             =   300
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image TestCQH 
      Height          =   1920
      Left            =   6225
      Picture         =   "frmTest.frx":23845
      Tag             =   "TCard"
      Top             =   300
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image TestCJH 
      Height          =   1920
      Left            =   6225
      Picture         =   "frmTest.frx":24864
      Tag             =   "TCard"
      Top             =   300
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image TestC10H 
      Height          =   1920
      Left            =   6225
      Picture         =   "frmTest.frx":2586E
      Tag             =   "TCard"
      Top             =   300
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image TestC9H 
      Height          =   1920
      Left            =   6225
      Picture         =   "frmTest.frx":265DF
      Tag             =   "TCard"
      Top             =   300
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image TestC8H 
      Height          =   1920
      Left            =   6225
      Picture         =   "frmTest.frx":272CB
      Tag             =   "TCard"
      Top             =   300
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image TestC7H 
      Height          =   1920
      Left            =   6225
      Picture         =   "frmTest.frx":27F19
      Tag             =   "TCard"
      Top             =   300
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image TestC6H 
      Height          =   1920
      Left            =   6225
      Picture         =   "frmTest.frx":28A88
      Tag             =   "TCard"
      Top             =   300
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image TestC5H 
      Height          =   1920
      Left            =   6225
      Picture         =   "frmTest.frx":29573
      Tag             =   "TCard"
      Top             =   300
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image TestC4H 
      Height          =   1920
      Left            =   6225
      Picture         =   "frmTest.frx":29F9E
      Tag             =   "TCard"
      Top             =   300
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image TestC3H 
      Height          =   1920
      Left            =   6225
      Picture         =   "frmTest.frx":2A916
      Tag             =   "TCard"
      Top             =   300
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image TestC2H 
      Height          =   1920
      Left            =   6225
      Picture         =   "frmTest.frx":2B247
      Tag             =   "TCard"
      Top             =   300
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image TestCAH 
      Height          =   1920
      Left            =   6225
      Picture         =   "frmTest.frx":2BAD3
      Tag             =   "TCard"
      Top             =   300
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image TestCKC 
      Height          =   1920
      Left            =   6225
      Picture         =   "frmTest.frx":2C2CC
      Tag             =   "TCard"
      Top             =   300
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image TestCQC 
      Height          =   1920
      Left            =   6225
      Picture         =   "frmTest.frx":2D218
      Tag             =   "TCard"
      Top             =   300
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image TestCJC 
      Height          =   1920
      Left            =   6225
      Picture         =   "frmTest.frx":2E180
      Tag             =   "TCard"
      Top             =   300
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image TestC10C 
      Height          =   1920
      Left            =   6225
      Picture         =   "frmTest.frx":2F0C6
      Tag             =   "TCard"
      Top             =   300
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image TestC9C 
      Height          =   1920
      Left            =   6225
      Picture         =   "frmTest.frx":2FFCB
      Tag             =   "TCard"
      Top             =   300
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image TestC8C 
      Height          =   1920
      Left            =   6225
      Picture         =   "frmTest.frx":30E24
      Tag             =   "TCard"
      Top             =   300
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image TestC7C 
      Height          =   1920
      Left            =   6225
      Picture         =   "frmTest.frx":31BEA
      Tag             =   "TCard"
      Top             =   300
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image TestC6C 
      Height          =   1920
      Left            =   6225
      Picture         =   "frmTest.frx":3286C
      Tag             =   "TCard"
      Top             =   300
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image TestC5C 
      Height          =   1920
      Left            =   6225
      Picture         =   "frmTest.frx":334CE
      Tag             =   "TCard"
      Top             =   300
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image TestC4C 
      Height          =   1920
      Left            =   6225
      Picture         =   "frmTest.frx":33FDF
      Tag             =   "TCard"
      Top             =   300
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image TestC3C 
      Height          =   1920
      Left            =   6225
      Picture         =   "frmTest.frx":34A1B
      Tag             =   "TCard"
      Top             =   300
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image TestC2C 
      Height          =   1920
      Left            =   6225
      Picture         =   "frmTest.frx":35361
      Tag             =   "TCard"
      Top             =   300
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image TestCAC 
      Height          =   1920
      Left            =   6225
      Picture         =   "frmTest.frx":35C03
      Tag             =   "TCard"
      Top             =   300
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image TestV52 
      Height          =   1920
      Left            =   4260
      Picture         =   "frmTest.frx":363F8
      Tag             =   "TValue"
      Top             =   285
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image TestV51 
      Height          =   1920
      Left            =   4260
      Picture         =   "frmTest.frx":37630
      Tag             =   "TValue"
      Top             =   285
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image TestV50 
      Height          =   1920
      Left            =   4260
      Picture         =   "frmTest.frx":386DB
      Tag             =   "TValue"
      Top             =   285
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image TestV49 
      Height          =   1920
      Left            =   4260
      Picture         =   "frmTest.frx":39868
      Tag             =   "TValue"
      Top             =   285
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image TestV48 
      Height          =   1920
      Left            =   4260
      Picture         =   "frmTest.frx":3A9BC
      Tag             =   "TValue"
      Top             =   285
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image TestV47 
      Height          =   1920
      Left            =   4260
      Picture         =   "frmTest.frx":3BB25
      Tag             =   "TValue"
      Top             =   285
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image TestV46 
      Height          =   1920
      Left            =   4260
      Picture         =   "frmTest.frx":3CB94
      Tag             =   "TValue"
      Top             =   285
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image TestV45 
      Height          =   1920
      Left            =   4260
      Picture         =   "frmTest.frx":3DD37
      Tag             =   "TValue"
      Top             =   285
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image TestV44 
      Height          =   1920
      Left            =   4260
      Picture         =   "frmTest.frx":3EE3B
      Tag             =   "TValue"
      Top             =   285
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image TestV43 
      Height          =   1920
      Left            =   4260
      Picture         =   "frmTest.frx":3FE90
      Tag             =   "TValue"
      Top             =   285
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image TestV42 
      Height          =   1920
      Left            =   4260
      Picture         =   "frmTest.frx":40FB8
      Tag             =   "TValue"
      Top             =   285
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image TestV41 
      Height          =   1920
      Left            =   4260
      Picture         =   "frmTest.frx":42177
      Tag             =   "TValue"
      Top             =   285
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image TestV40 
      Height          =   1920
      Left            =   4260
      Picture         =   "frmTest.frx":43172
      Tag             =   "TValue"
      Top             =   285
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image TestV39 
      Height          =   1920
      Left            =   4260
      Picture         =   "frmTest.frx":44297
      Tag             =   "TValue"
      Top             =   285
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image TestV38 
      Height          =   1920
      Left            =   4260
      Picture         =   "frmTest.frx":454C0
      Tag             =   "TValue"
      Top             =   285
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image TestV37 
      Height          =   1920
      Left            =   4260
      Picture         =   "frmTest.frx":466F9
      Tag             =   "TValue"
      Top             =   285
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image TestV36 
      Height          =   1920
      Left            =   4260
      Picture         =   "frmTest.frx":4780C
      Tag             =   "TValue"
      Top             =   285
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image TestV35 
      Height          =   1920
      Left            =   4260
      Picture         =   "frmTest.frx":48A4A
      Tag             =   "TValue"
      Top             =   285
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image TestV34 
      Height          =   1920
      Left            =   4260
      Picture         =   "frmTest.frx":49C28
      Tag             =   "TValue"
      Top             =   285
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image TestV33 
      Height          =   1920
      Left            =   4260
      Picture         =   "frmTest.frx":4AD3E
      Tag             =   "TValue"
      Top             =   285
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image TestV32 
      Height          =   1920
      Left            =   4260
      Picture         =   "frmTest.frx":4BEEE
      Tag             =   "TValue"
      Top             =   285
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image TestV31 
      Height          =   1920
      Left            =   4260
      Picture         =   "frmTest.frx":4D10F
      Tag             =   "TValue"
      Top             =   285
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image TestV30 
      Height          =   1920
      Left            =   4260
      Picture         =   "frmTest.frx":4E178
      Tag             =   "TValue"
      Top             =   285
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image TestV29 
      Height          =   1920
      Left            =   4260
      Picture         =   "frmTest.frx":4F345
      Tag             =   "TValue"
      Top             =   285
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image TestV28 
      Height          =   1920
      Left            =   4260
      Picture         =   "frmTest.frx":505A6
      Tag             =   "TValue"
      Top             =   285
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image TestV27 
      Height          =   1920
      Left            =   4260
      Picture         =   "frmTest.frx":51861
      Tag             =   "TValue"
      Top             =   285
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image TestV26 
      Height          =   1920
      Left            =   4260
      Picture         =   "frmTest.frx":529C7
      Tag             =   "TValue"
      Top             =   285
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image TestV25 
      Height          =   1920
      Left            =   4260
      Picture         =   "frmTest.frx":53C18
      Tag             =   "TValue"
      Top             =   285
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image TestV24 
      Height          =   1920
      Left            =   4260
      Picture         =   "frmTest.frx":54E2F
      Tag             =   "TValue"
      Top             =   285
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image TestV23 
      Height          =   1920
      Left            =   4260
      Picture         =   "frmTest.frx":55FAD
      Tag             =   "TValue"
      Top             =   285
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image TestV22 
      Height          =   1920
      Left            =   4260
      Picture         =   "frmTest.frx":571A8
      Tag             =   "TValue"
      Top             =   285
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image TestV21 
      Height          =   1920
      Left            =   4260
      Picture         =   "frmTest.frx":58415
      Tag             =   "TValue"
      Top             =   285
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image TestV20 
      Height          =   1920
      Left            =   4260
      Picture         =   "frmTest.frx":5950F
      Tag             =   "TValue"
      Top             =   285
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image TestV19 
      Height          =   1920
      Left            =   4260
      Picture         =   "frmTest.frx":5A746
      Tag             =   "TValue"
      Top             =   285
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image TestV18 
      Height          =   1920
      Left            =   4260
      Picture         =   "frmTest.frx":5B81A
      Tag             =   "TValue"
      Top             =   285
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image TestV17 
      Height          =   1920
      Left            =   4260
      Picture         =   "frmTest.frx":5C91C
      Tag             =   "TValue"
      Top             =   285
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image TestV16 
      Height          =   1920
      Left            =   4260
      Picture         =   "frmTest.frx":5D8E5
      Tag             =   "TValue"
      Top             =   285
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image TestV15 
      Height          =   1920
      Left            =   4260
      Picture         =   "frmTest.frx":5E9D7
      Tag             =   "TValue"
      Top             =   285
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image TestV14 
      Height          =   1920
      Left            =   4260
      Picture         =   "frmTest.frx":5FAAF
      Tag             =   "TValue"
      Top             =   285
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image TestV13 
      Height          =   1920
      Left            =   4260
      Picture         =   "frmTest.frx":60B37
      Tag             =   "TValue"
      Top             =   285
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image TestV12 
      Height          =   1920
      Left            =   4260
      Picture         =   "frmTest.frx":61B7A
      Tag             =   "TValue"
      Top             =   285
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image TestV11 
      Height          =   1920
      Left            =   4260
      Picture         =   "frmTest.frx":62C42
      Tag             =   "TValue"
      Top             =   285
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image TestV10 
      Height          =   1920
      Left            =   4260
      Picture         =   "frmTest.frx":63B3D
      Tag             =   "TValue"
      Top             =   285
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image TestV9 
      Height          =   1920
      Left            =   4260
      Picture         =   "frmTest.frx":64BDD
      Tag             =   "TValue"
      Top             =   285
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image TestV8 
      Height          =   1920
      Left            =   4260
      Picture         =   "frmTest.frx":65B49
      Tag             =   "TValue"
      Top             =   285
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image TestV7 
      Height          =   1920
      Left            =   4260
      Picture         =   "frmTest.frx":66AB3
      Tag             =   "TValue"
      Top             =   285
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image TestV6 
      Height          =   1920
      Left            =   4260
      Picture         =   "frmTest.frx":67918
      Tag             =   "TValue"
      Top             =   285
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image TestV5 
      Height          =   1920
      Left            =   4260
      Picture         =   "frmTest.frx":688A8
      Tag             =   "TValue"
      Top             =   285
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image TestV4 
      Height          =   1920
      Left            =   4260
      Picture         =   "frmTest.frx":697B7
      Tag             =   "TValue"
      Top             =   285
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image TestV3 
      Height          =   1920
      Left            =   4260
      Picture         =   "frmTest.frx":6A63C
      Tag             =   "TValue"
      Top             =   285
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image TestV2 
      Height          =   1920
      Left            =   4260
      Picture         =   "frmTest.frx":6B55A
      Tag             =   "TValue"
      Top             =   285
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image TestV1 
      Height          =   1920
      Left            =   4260
      Picture         =   "frmTest.frx":6C506
      Tag             =   "TValue"
      Top             =   285
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Label ProgressLabelShow 
      Alignment       =   1  'Right Justify
      Caption         =   "Show"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3225
      TabIndex        =   22
      Top             =   3180
      Width           =   990
   End
   Begin VB.Label ProgressLabelTest 
      Alignment       =   1  'Right Justify
      Caption         =   "Test"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3225
      TabIndex        =   21
      Top             =   2835
      Width           =   990
   End
   Begin VB.Label Label12 
      Caption         =   "Test Design"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   270
      TabIndex        =   18
      Top             =   60
      Width           =   2520
   End
   Begin VB.Image TestV0 
      Height          =   1920
      Left            =   4260
      Picture         =   "frmTest.frx":6D2A3
      Tag             =   "TValue"
      Top             =   285
      Width           =   1350
   End
   Begin VB.Shape TestArea 
      BorderWidth     =   2
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   2625
      Left            =   4065
      Shape           =   4  'Rounded Rectangle
      Top             =   105
      Width           =   3705
   End
   Begin VB.Image TestToggle 
      Height          =   600
      Index           =   3
      Left            =   5460
      Picture         =   "frmTest.frx":6E167
      Top             =   3750
      Visible         =   0   'False
      Width           =   1740
   End
   Begin VB.Image TestToggle 
      Height          =   600
      Index           =   2
      Left            =   5460
      Picture         =   "frmTest.frx":6EA93
      Top             =   3750
      Visible         =   0   'False
      Width           =   1740
   End
   Begin VB.Image TestToggle 
      Height          =   600
      Index           =   1
      Left            =   5460
      Picture         =   "frmTest.frx":6F479
      Top             =   3750
      Visible         =   0   'False
      Width           =   1740
   End
   Begin VB.Image TestShow 
      Height          =   600
      Index           =   1
      Left            =   4530
      Picture         =   "frmTest.frx":6FD5A
      Top             =   4410
      Visible         =   0   'False
      Width           =   1740
   End
   Begin VB.Image TestNext 
      Height          =   600
      Index           =   1
      Left            =   6525
      Picture         =   "frmTest.frx":706A4
      Top             =   4410
      Visible         =   0   'False
      Width           =   1740
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   1485
      Left            =   4290
      Top             =   3645
      Width           =   4080
   End
   Begin VB.Shape Shape2 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FF8080&
      FillStyle       =   0  'Solid
      Height          =   1605
      Left            =   4215
      Top             =   3585
      Width           =   4230
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub DeckCurrentRangePartial_Click()
SequenceCurrent.Value = False
SequenceCurrent.Enabled = False
End Sub

Private Sub DeckRangeFull_Click()
DeckRangePartialStartTextBox = Empty
DeckRangePartialFinishTextBox = Empty
SequenceCurrent.Enabled = True
End Sub

Private Sub DeckRangePartial_Click()
SequenceCurrent.Value = False
SequenceCurrent.Enabled = False
End Sub

Private Sub DeckRangePartialFinishTextBox_GotFocus()
DeckRangeFull.Value = False
SequenceCurrent.Value = False
SequenceCurrent.Enabled = False
End Sub

Private Sub MnemonicIndexCalc(cardVal)
Dim mnemonicMultiplier As Integer
Dim mnemonicPosition As Integer
Select Case Right(cardVal, 1)
    Case "C"
        mnemonicMultiplier = 0
    Case "H"
        mnemonicMultiplier = 1
    Case "S"
        mnemonicMultiplier = 2
    Case "D"
        mnemonicMultiplier = 3
End Select
Select Case Left(cardVal, Len(cardVal) - 1)
    Case "A"
        mnemonicPosition = 1
    Case "2"
        mnemonicPosition = 2
    Case "3"
        mnemonicPosition = 3
    Case "4"
        mnemonicPosition = 4
    Case "5"
        mnemonicPosition = 5
    Case "6"
        mnemonicPosition = 6
    Case "7"
        mnemonicPosition = 7
    Case "8"
        mnemonicPosition = 8
    Case "9"
        mnemonicPosition = 9
    Case "10"
        mnemonicPosition = 10
    Case "J"
        mnemonicPosition = 11
    Case "Q"
        mnemonicPosition = 12
    Case "K"
        mnemonicPosition = 13
End Select
MnemonicCardIndex = mnemonicMultiplier * 13 + mnemonicPosition
End Sub

Private Sub DeckRangePartialStartTextBox_GotFocus()
DeckRangeFull.Value = False
SequenceCurrent.Value = False
SequenceCurrent.Enabled = False
End Sub


Private Sub Form_Load()
    TestCounter = 0
    HintButton(0).Visible = False
    HintButton(1).Visible = False
    MnemonicLeft.Visible = False
    MnemonicRight.Visible = False
    MnemonicLeft.Caption = Empty
    MnemonicRight.Caption = Empty
    TestCurrentCardOption.Value = True
    TestCardMode = 0
    TestCurrentImage.Visible = True
    TestNextImage.Visible = False
    TestPreviousImage.Visible = False
    TestLeftImage.Visible = False
    TestRightImage.Visible = False
    'CummTimeIntervals = 0
    TestingMode = False
    TestingStatus.Caption = "Ready"
    TestingStatus.ForeColor = &H8000&
    'set color to GREEN
    ProgressCardsRemaining.Visible = False
    TestCardsRemaining.Visible = False
    TestCardsLabel1.Visible = False
    TestCardsLabel2.Visible = False
    ProgressCardsRemaining.Value = 52
    TestCardsRemaining.Caption = 52
    DeckRangeFrame.Enabled = True
    DeckRangeFull.Enabled = True
    DeckRangePartial.Enabled = True
    DeckCurrentRangePartial.Enabled = True
    DeckRangePartialStartTextBox.Enabled = True
    DeckRangePartialFinishTextBox.Enabled = True
    SequenceFrame.Enabled = True
    SequenceForward.Enabled = True
    SequenceBackward.Enabled = True
    If DeckRangeFull.Value = True Then
        SequenceCurrent.Enabled = True
    Else
        SequenceCurrent.Enabled = False
    End If
    SequenceRandom.Enabled = True
    TestValue.Enabled = True
    TestPosition.Enabled = True
    TestRandom.Enabled = True
    TestToggle(0).Visible = True
    TestToggle(1).Visible = False
    TestToggle(2).Visible = False
    TestToggle(3).Visible = False
    TestMessage(0).Visible = True
    TestNext(0).Visible = False
    TestNext(1).Visible = True
    TestNext(2).Visible = False
    TestNext(3).Visible = False
    TestNext(4).Visible = False
    TestNext(5).Visible = False
    TestShow(0).Visible = False
    TestShow(1).Visible = True
    TestShow(2).Visible = False
    TestShow(3).Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMain.mnuTest.Checked = False
End Sub




Private Sub HintButton_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If TestingMode Then
    If TestCounter <> 0 Then
        HintButton(0).Visible = False
        HintButton(1).Visible = True
    End If
End If
End Sub

Private Sub HintButton_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If TestingMode Then
    HintButton(0).Visible = True
    HintButton(1).Visible = False
    If TestCounter <> 0 Then
        TemporaryMnemonicHint = True
        If TestingMode Or ShowingMode = 1 Then
            MnemonicHint
        End If
    End If
End If
End Sub

Private Sub MnemonicsEnabled_Click()
If MnemonicsEnabled.Value = 1 Then
    MnemonicsHintCheckBox.Enabled = True
    If TestRandomValue = 1 Or TestValue Or ShowingMode = 1 Then
        MnemonicLeft.Visible = True
    ElseIf TestRandomValue = 0 Or TestPosition Or ShowingMode = 1 Then
        MnemonicRight.Visible = True
    End If
    If MnemonicsHintCheckBox.Value = 0 Then
        HintButton(0).Visible = True
        HintButton(1).Visible = False
    Else
        HintButton(0).Visible = False
        HintButton(1).Visible = False
    End If
Else
    MnemonicsHintCheckBox.Enabled = False
    MnemonicsHintCheckBox.Value = 0
    MnemonicLeft.Visible = False
    MnemonicRight.Visible = False
    HintButton(0).Visible = False
    HintButton(1).Visible = False
End If
End Sub


Private Sub MnemonicsHintCheckBox_Click()
MnemonicLeft.Visible = False
MnemonicRight.Visible = False
If MnemonicsHintCheckBox.Value = 1 Then
    HintButton(0).Visible = False
    HintButton(1).Visible = False
Else
    HintButton(0).Visible = True
    HintButton(1).Visible = False
End If
If TestingMode Or ShowingMode = 1 Then
    MnemonicHint
    If MnemonicsHintCheckBox.Value = 0 Then
        If TestRandomValue = 1 Or TestValue Then
            MnemonicLeft.Visible = True
            MnemonicRight.Visible = False
        ElseIf TestRandomValue = 0 Or TestPosition Then
            MnemonicRight.Visible = True
            MnemonicLeft.Visible = False
        End If
    End If
End If
End Sub

Private Sub TestCurrentCardOption_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
TestCurrentImage.Visible = True
TestNextImage.Visible = False
TestPreviousImage.Visible = False
TestLeftImage.Visible = False
TestRightImage.Visible = False
End Sub


Private Sub TestNext_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If TestingMode Then
    If TestCounter = 0 Then
        TestNext(0).Visible = False
        TestNext(1).Visible = False
        TestNext(2).Visible = False
        TestNext(3).Visible = True
        TestNext(4).Visible = False
        TestNext(5).Visible = False
    ElseIf TestCounter = DeckRangeCount Then
        TestNext(0).Visible = False
        TestNext(1).Visible = False
        TestNext(2).Visible = False
        TestNext(3).Visible = False
        TestNext(4).Visible = False
        TestNext(5).Visible = True
    Else
        TestNext(0).Visible = False
        TestNext(1).Visible = True
        TestNext(2).Visible = False
        TestNext(3).Visible = False
        TestNext(4).Visible = False
        TestNext(5).Visible = False
    End If
    'cummTimeIntervals = CummTimeIntervals + TestProgressIntervals
    TimerTest.Enabled = False
    ShowingMode = 0
    TestShow(0).Visible = True
    TestShow(1).Visible = False
    TestShow(2).Visible = False
    TestShow(3).Visible = False
    'cummTimeIntervals = CummTimeIntervals + ShowProgressIntervals
    TimerShow.Enabled = False
    ProgressTest.Value = 0
    ProgressShow.Value = 0
    TestProgressIntervals = 0
    ShowProgressIntervals = 0
End If
End Sub

Private Sub TestNext_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If TestingMode Then
    If TestCounter = 0 Then
        StartTime = Timer
    End If
    If TestCounter = DeckRangeCount - 1 Then
        TestNext(0).Visible = False
        TestNext(1).Visible = False
        TestNext(2).Visible = False
        TestNext(3).Visible = False
        TestNext(4).Visible = True
        TestNext(5).Visible = False
        TestNextCard
    Else
        TestNext(0).Visible = True
        TestNext(1).Visible = False
        TestNext(2).Visible = False
        TestNext(3).Visible = False
        TestNext(4).Visible = False
        TestNext(5).Visible = False
        TestNextCard
    End If
End If
End Sub


Private Sub TestNextCardOption_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
TestCurrentImage.Visible = False
TestNextImage.Visible = True
TestPreviousImage.Visible = False
If TestValue Then
    TestLeftImage.Visible = False
    TestRightImage.Visible = True
ElseIf TestPosition Then
    TestLeftImage.Visible = True
    TestRightImage.Visible = False
End If
End Sub



Private Sub TestPosition_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'If TestNextCardOption Or TestPreviousCardOption Then
    TestLeftImage.Visible = True
    TestRightImage.Visible = False
'End If
End Sub

Private Sub TestPreviousCardOption_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
TestCurrentImage.Visible = False
TestNextImage.Visible = False
TestPreviousImage.Visible = True
If TestValue Then
    TestLeftImage.Visible = False
    TestRightImage.Visible = True
ElseIf TestPosition Then
    TestLeftImage.Visible = True
    TestRightImage.Visible = False
End If
End Sub


Private Sub TestRandom_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
TestLeftImage.Visible = False
TestRightImage.Visible = False
End Sub


Private Sub TestRandomCardOption_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
TestCurrentImage.Visible = True
TestNextImage.Visible = False
TestPreviousImage.Visible = False
If TestValue Then
    TestLeftImage.Visible = False
    TestRightImage.Visible = True
ElseIf TestPosition Then
    TestLeftImage.Visible = True
    TestRightImage.Visible = False
End If
End Sub

Private Sub TestShow_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If TestingMode Then
    If TestCounter = 0 Then
        TestShow(0).Visible = False
        TestShow(1).Visible = False
        TestShow(2).Visible = False
        TestShow(3).Visible = True
        'cummTimeIntervals = CummTimeIntervals + TestProgressIntervals
        TimerTest.Enabled = False
        ShowingMode = 0
        'cummTimeIntervals = CummTimeIntervals + ShowProgressIntervals
        TimerShow.Enabled = False
        ProgressTest.Value = 0
        ProgressShow.Value = 0
        TestProgressIntervals = 0
        ShowProgressIntervals = 0
    Else
        If ShowingMode = 1 Then
            TestShow(0).Visible = False
            TestShow(1).Visible = False
            TestShow(2).Visible = False
            TestShow(3).Visible = True
        Else
            TestShow(0).Visible = False
            TestShow(1).Visible = True
            TestShow(2).Visible = False
            TestShow(3).Visible = False
        End If
    End If
End If
End Sub

Private Sub TestShow_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If TestingMode Then
    If TestCounter = DeckRangeCount Then
        'this turns on the "Finish" button for the last card
        TestNext(0).Visible = False
        TestNext(1).Visible = False
        TestNext(2).Visible = False
        TestNext(3).Visible = False
        TestNext(4).Visible = True
        TestNext(5).Visible = False
    ElseIf TestCounter = 0 Then
        TestShow(0).Visible = True
        TestShow(1).Visible = False
        TestShow(2).Visible = False
        TestShow(3).Visible = False
        StartTime = Timer
    Else
        TestNext(0).Visible = True
        TestNext(1).Visible = False
        TestNext(2).Visible = False
        TestNext(3).Visible = False
        TestNext(4).Visible = False
        TestNext(5).Visible = False
    End If
    If TestingMode Then
        If TestCounter = 0 Then
            TestShow(0).Visible = True
            TestShow(1).Visible = False
            TestShow(2).Visible = False
            TestShow(3).Visible = False
            TestNextCard
        ElseIf ShowingMode = 1 Then
            TestShow(0).Visible = True
            TestShow(1).Visible = False
            TestShow(2).Visible = False
            TestShow(3).Visible = False
            TestNextCard
        Else
            TestShow(0).Visible = False
            TestShow(1).Visible = False
            TestShow(2).Visible = True
            TestShow(3).Visible = False
            TestShowCard
        End If
    End If
End If
End Sub

Private Sub TestTimersEnabled_Click()
If TestTimersEnabled.Value = 1 Then
    ProgressLabelTest.Visible = True
    ProgressLabelShow.Visible = True
    ProgressTest.Visible = True
    ProgressShow.Visible = True
    TestDuration.Enabled = True
    ShowDuration.Enabled = True
Else
    ProgressLabelTest.Visible = False
    ProgressLabelShow.Visible = False
    ProgressTest.Visible = False
    ProgressShow.Visible = False
    TestDuration.Enabled = False
    ShowDuration.Enabled = False
End If
End Sub

Private Sub TestToggle_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If DeckRangeFull.Value = False And _
    DeckRangePartial.Value = False And _
    DeckCurrentRangePartial.Value = False Then
    MsgBox ("You must select a Deck Range option to begin.")
    Exit Sub
End If
If SequenceForward.Value = False And _
    SequenceBackward.Value = False And _
    SequenceCurrent.Value = False And _
    SequenceRandom.Value = False Then
    MsgBox ("You must select a Stack Sequence option to begin.")
    Exit Sub
End If
'this section checks DeckRangePartial text box errors
If DeckRangePartial.Value = True And _
    (DeckRangePartialStartTextBox = Empty Or _
    DeckRangePartialFinishTextBox = Empty) Then
    MsgBox ("When you select 'Partial Stack Range', you must enter values" & Chr(13) _
        & "in both 'Start' and 'Finish' input boxes.")
    Exit Sub
End If
If DeckRangePartial.Value = True And _
    DeckRangePartialFinishTextBox.Text <> Empty And _
    (Not IsNumeric(DeckRangePartialFinishTextBox.Text) Or _
    Val(DeckRangePartialFinishTextBox.Text) < 2 Or _
    Val(DeckRangePartialFinishTextBox) > 52) Then
    DeckRangePartialFinishTextBox.Text = Empty
    MsgBox "Please enter a valid card position (2 to 52)" & Chr(13) _
        & "in the 'Finish' Input Box"
    DeckRangePartialFinishTextBox.SetFocus
    Exit Sub
End If
If DeckRangePartial.Value = True And _
    (DeckRangePartialStartTextBox.Text <> Empty And _
    DeckRangePartialFinishTextBox.Text <> Empty) And _
    Val(DeckRangePartialFinishTextBox.Text) <= _
        Val(DeckRangePartialStartTextBox) Then
    DeckRangePartialFinishTextBox.Text = Empty
    MsgBox "The 'Finish' value must be larger" & Chr(13) _
        & "then the 'Start' value."
    DeckRangePartialFinishTextBox.SetFocus
    Exit Sub
End If
If DeckRangePartial.Value = True And _
    DeckRangePartialStartTextBox.Text <> Empty And _
    (Not IsNumeric(DeckRangePartialStartTextBox.Text) Or _
    Val(DeckRangePartialStartTextBox.Text) < 1 Or _
    Val(DeckRangePartialStartTextBox) > 51) Then
    DeckRangePartialStartTextBox.Text = Empty
    MsgBox "Please enter a valid card position (1 to 51)" & Chr(13) _
        & "in the 'Start' Input Box"
    DeckRangePartialStartTextBox.SetFocus
    Exit Sub
End If
If DeckRangePartial.Value = True And _
    (DeckRangePartialStartTextBox.Text <> Empty And _
    DeckRangePartialFinishTextBox.Text <> Empty) And _
    Val(DeckRangePartialStartTextBox.Text) >= _
        Val(DeckRangePartialFinishTextBox) Then
    DeckRangePartialStartTextBox.Text = Empty
    MsgBox "The 'Start' value must be smaller" & Chr(13) _
        & "then the 'Finish' value."
    DeckRangePartialStartTextBox.SetFocus
    Exit Sub
End If
'this section checks DeckCurrentRangePartial text box errors
If DeckCurrentRangePartial.Value = True And _
    (DeckRangePartialStartTextBox = Empty Or _
    DeckRangePartialFinishTextBox = Empty) Then
    MsgBox ("When you select 'Partial Current Deck Range', you must enter values" & Chr(13) _
        & "in both 'Start' and 'Finish' input boxes.")
    Exit Sub
End If
If DeckCurrentRangePartial.Value = True And _
    DeckRangePartialFinishTextBox.Text <> Empty And _
    (Not IsNumeric(DeckRangePartialFinishTextBox.Text) Or _
    Val(DeckRangePartialFinishTextBox.Text) < 2 Or _
    Val(DeckRangePartialFinishTextBox) > 52) Then
    DeckRangePartialFinishTextBox.Text = Empty
    MsgBox "Please enter a valid card position (2 to 52)" & Chr(13) _
        & "in the 'Finish' Input Box"
    DeckRangePartialFinishTextBox.SetFocus
    Exit Sub
End If
If DeckCurrentRangePartial.Value = True And _
    (DeckRangePartialStartTextBox.Text <> Empty And _
    DeckRangePartialFinishTextBox.Text <> Empty) And _
    Val(DeckRangePartialFinishTextBox.Text) <= _
        Val(DeckRangePartialStartTextBox) Then
    DeckRangePartialFinishTextBox.Text = Empty
    MsgBox "The 'Finish' value must be larger" & Chr(13) _
        & "then the 'Start' value."
    DeckRangePartialFinishTextBox.SetFocus
    Exit Sub
End If
If DeckCurrentRangePartial.Value = True And _
    DeckRangePartialStartTextBox.Text <> Empty And _
    (Not IsNumeric(DeckRangePartialStartTextBox.Text) Or _
    Val(DeckRangePartialStartTextBox.Text) < 1 Or _
    Val(DeckRangePartialStartTextBox) > 51) Then
    DeckRangePartialStartTextBox.Text = Empty
    MsgBox "Please enter a valid card position (1 to 51)" & Chr(13) _
        & "in the 'Start' Input Box"
    DeckRangePartialStartTextBox.SetFocus
    Exit Sub
End If
If DeckCurrentRangePartial.Value = True And _
    (DeckRangePartialStartTextBox.Text <> Empty And _
    DeckRangePartialFinishTextBox.Text <> Empty) And _
    Val(DeckRangePartialStartTextBox.Text) >= _
        Val(DeckRangePartialFinishTextBox) Then
    DeckRangePartialStartTextBox.Text = Empty
    MsgBox "The 'Start' value must be smaller" & Chr(13) _
        & "then the 'Finish' value."
    DeckRangePartialStartTextBox.SetFocus
    Exit Sub
End If

If DeckRangePartial.Value = True Or _
    DeckCurrentRangePartial.Value = True Then
    DeckRangeStart = Val(DeckRangePartialStartTextBox.Text)
    DeckRangeFinish = Val(DeckRangePartialFinishTextBox.Text)
    DeckRangeCount = DeckRangeFinish - DeckRangeStart + 1
Else
    DeckRangeStart = 1
    DeckRangeFinish = 52
    DeckRangeCount = DeckRangeFinish - DeckRangeStart + 1
End If
'cummTimeIntervals = CummTimeIntervals + TestProgressIntervals
TimerTest.Enabled = False
ShowingMode = 0
'cummTimeIntervals = CummTimeIntervals + ShowProgressIntervals
TimerShow.Enabled = False
ProgressTest.Value = 0
ProgressShow.Value = 0
TestProgressIntervals = 0
ShowProgressIntervals = 0
If TestingMode Then
    TestToggle(0).Visible = False
    TestToggle(1).Visible = False
    TestToggle(2).Visible = False
    TestToggle(3).Visible = True
Else
    TestToggle(0).Visible = False
    TestToggle(1).Visible = True
    TestToggle(2).Visible = False
    TestToggle(3).Visible = False
End If
End Sub

Private Sub TestToggle_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
TestCounter = 0
ResultLabel.Visible = False
ResultTime.Visible = False
MnemonicLeft.Visible = False
MnemonicRight.Visible = False
HintButton(0).Visible = False
HintButton(1).Visible = False
MnemonicLeft.Caption = Empty
MnemonicRight.Caption = Empty
TestTimersEnabled.Enabled = False
TestDuration.Enabled = False
ShowDuration.Enabled = False
If TestTimersEnabled = 1 Then
    If Val(TestDuration.Text) >= 0.1 And _
        Val(TestDuration.Text) <= 60 And _
        Val(ShowDuration.Text) >= 0.1 And _
        Val(ShowDuration.Text) <= 60 Then
            ProgressTest.Value = 0
            ProgressTest.Max = Round(Val(TestDuration.Text), 1) * _
                (1000 / TimerTest.Interval)
            ProgressShow.Value = 0
            ProgressShow.Max = Round(Val(ShowDuration.Text), 1) * _
                (1000 / TimerShow.Interval)
    Else
        MsgBox ("Times can not be less than 0.1 seconds." & Chr(13) & _
                "Times can not be greater than 60 seconds." & Chr(13) & _
                "You may use decimals such as 2.5 seconds.")
        
        TestToggle(0).Visible = True
        TestToggle(1).Visible = False
        TestToggle(2).Visible = False
        TestToggle(3).Visible = False
        TestTimersEnabled.Enabled = True
        TestDuration.Enabled = True
        ShowDuration.Enabled = True
        Exit Sub
    End If
End If
If TestingMode Then
    ProgressCardsRemaining.Visible = False
    TestCardsRemaining.Visible = False
    TestCardsLabel1.Visible = False
    TestCardsLabel2.Visible = False
    ProgressCardsRemaining.Value = DeckRangeCount
    TestCardsRemaining.Caption = DeckRangeCount
    TestToggle(0).Visible = True
    TestToggle(1).Visible = False
    TestToggle(2).Visible = False
    TestToggle(3).Visible = False
    TestShow(0).Visible = False
    TestShow(1).Visible = True
    TestShow(2).Visible = False
    TestShow(3).Visible = False
    If TestRandom Then
        TestLeftImage.Visible = False
        TestRightImage.Visible = False
    End If
Else
    ProgressCardsRemaining.Visible = True
    TestCardsRemaining.Visible = True
    TestCardsLabel1.Visible = True
    TestCardsLabel2.Visible = True
    ProgressCardsRemaining.Value = DeckRangeCount
    TestCardsRemaining.Caption = DeckRangeCount
    TestToggle(0).Visible = False
    TestToggle(1).Visible = False
    TestToggle(2).Visible = True
    TestToggle(3).Visible = False
    'cummTimeIntervals = 0
    'cummTimeSeconds = 0
End If
TestingMode = Not TestingMode
'Toggle logical value when button pressed
'initialize deckorder
For i% = 1 To DeckRangeCount
    StartOrder(i%) = DeckRangeStart + i% - 1
Next i%
If SequenceForward Then
    For k% = 1 To DeckRangeCount
        TestOrder(k%) = StartOrder(k%)
    Next k%
ElseIf SequenceBackward Then
    For m% = 1 To DeckRangeCount
        TestOrder(m%) = StartOrder(DeckRangeCount + 1 - m%)
    Next m%
ElseIf SequenceCurrent Then
    If DeckRangeFull.Value = True Then
        'this segment is new
        For n% = 1 To DeckRangeCount
            TestOrder(n%) = StartOrder(n%)
        Next n%
        'end of new
    Else
        For n% = 1 To DeckRangeCount
            TestOrder(n%) = (Deck(1, DeckRangeStart + n% - 1))
        Next n%
    End If
ElseIf SequenceRandom Then
    For p% = 1 To DeckRangeCount
        selector = Int(Rnd * (DeckRangeCount + 1 - p%)) + 1
        TestOrder(p%) = StartOrder(selector)
        If selector <> DeckRangeCount + 1 - p% Then
            StartOrder(selector) = StartOrder(DeckRangeCount + 1 - p%)
        End If
    Next p%
End If
If TestingMode Then
    TestMessage(0).Visible = False
    TestingStatus.Caption = "Testing"
    TestingStatus.ForeColor = &HFF&
    'set color to RED
    SequenceFrame.Enabled = False
    SequenceForward.Enabled = False
    SequenceBackward.Enabled = False
    SequenceCurrent.Enabled = False
    SequenceRandom.Enabled = False
    DeckRangeFrame.Enabled = False
    DeckRangeFull.Enabled = False
    DeckRangePartial.Enabled = False
    DeckCurrentRangePartial.Enabled = False
    DeckRangePartialStartTextBox.Enabled = False
    DeckRangePartialFinishTextBox.Enabled = False
    'TestValue.Enabled = False
    'TestPosition.Enabled = False
    'TestRandom.Enabled = False
    TestNext(0).Visible = False
    TestNext(1).Visible = False
    TestNext(2).Visible = True
    TestNext(3).Visible = False
    TestNext(4).Visible = False
    TestNext(5).Visible = False
    TestShow(0).Visible = False
    TestShow(1).Visible = False
    TestShow(2).Visible = True
    TestShow(3).Visible = False
    If MnemonicsEnabled.Value = 1 Then
        If MnemonicsHintCheckBox.Value = 0 Then
            HintButton(0).Visible = True
            HintButton(1).Visible = False
        Else
            HintButton(0).Visible = False
            HintButton(1).Visible = False
        End If
    End If
Else
    TestMessage(0).Visible = True
    TestingStatus.Caption = "Ready"
    TestingStatus.ForeColor = &H8000&
    'set color to GREEN
    SequenceFrame.Enabled = True
    SequenceForward.Enabled = True
    SequenceBackward.Enabled = True
    If DeckRangeFull.Value = True Then
        SequenceCurrent.Enabled = True
    Else
        SequenceCurrent.Enabled = False
    End If
    SequenceRandom.Enabled = True
    DeckRangeFrame.Enabled = True
    DeckRangeFull.Enabled = True
    DeckRangePartial.Enabled = True
    DeckCurrentRangePartial.Enabled = True
    DeckRangePartialStartTextBox.Enabled = True
    DeckRangePartialFinishTextBox.Enabled = True
    TestValue.Enabled = True
    TestPosition.Enabled = True
    TestRandom.Enabled = True
    For Each Ctrl In Controls
        If Ctrl.Tag = "TCard" Or Ctrl.Tag = "TValue" Then
            Ctrl.Visible = False
        End If
    Next Ctrl
    TestC0.Visible = True
    TestV0.Visible = True
    TestNext(0).Visible = False
    TestNext(1).Visible = True
    TestNext(2).Visible = False
    TestNext(3).Visible = False
    TestNext(4).Visible = False
    TestNext(5).Visible = False
    TestShow(0).Visible = False
    TestShow(1).Visible = True
    TestShow(2).Visible = False
    TestShow(3).Visible = False
    TestTimersEnabled.Enabled = True
    TestDuration.Enabled = True
    ShowDuration.Enabled = True
    If MnemonicsEnabled.Value = 1 Then
        If MnemonicsHintCheckBox.Value = 0 Then
            HintButton(0).Visible = True
            HintButton(1).Visible = False
        Else
            HintButton(0).Visible = False
            HintButton(1).Visible = False
        End If
    End If
    MnemonicLeft.Caption = Empty
    MnemonicRight.Caption = Empty
End If
End Sub

Private Sub TestNextCard()
MnemonicLeft.Visible = False
MnemonicRight.Visible = False
TemporaryMnemonicHint = False
'cummTimeIntervals = CummTimeIntervals + TestProgressIntervals
TimerTest.Enabled = False
ShowingMode = 0
TestShow(0).Visible = True
TestShow(1).Visible = False
TestShow(2).Visible = False
TestShow(3).Visible = False
'cummTimeIntervals = CummTimeIntervals + ShowProgressIntervals
TimerShow.Enabled = False
ProgressTest.Value = 0
ProgressShow.Value = 0
TestProgressIntervals = 0
ShowProgressIntervals = 0
TestCounter = TestCounter + 1
TestCardsRemaining = DeckRangeCount + 1 - TestCounter
ProgressCardsRemaining.Value = Val(TestCardsRemaining)
If TestCounter = DeckRangeCount Then
    'this turns on the "Finish" button for the last card
    TestNext(0).Visible = False
    TestNext(1).Visible = False
    TestNext(2).Visible = False
    TestNext(3).Visible = False
    TestNext(4).Visible = True
    TestNext(5).Visible = False
Else
    TestNext(0).Visible = True
    TestNext(1).Visible = False
    TestNext(2).Visible = False
    TestNext(3).Visible = False
    TestNext(4).Visible = False
    TestNext(5).Visible = False
End If
If TestCounter > DeckRangeCount Then
    TestCounter = 0
    TestingMode = False
    TestingStatus.Caption = "Ready"
    TestingStatus.ForeColor = &H8000&
    'set color to GREEN
    ProgressCardsRemaining.Visible = False
    TestCardsRemaining.Visible = False
    TestCardsLabel1.Visible = False
    TestCardsLabel2.Visible = False
    ProgressCardsRemaining.Value = DeckRangeCount
    TestCardsRemaining.Caption = DeckRangeCount
    DeckRangeFrame.Enabled = True
    DeckRangeFull.Enabled = True
    DeckRangePartial.Enabled = True
    DeckCurrentRangePartial.Enabled = True
    DeckRangePartialStartTextBox.Enabled = True
    DeckRangePartialFinishTextBox.Enabled = True
    SequenceFrame.Enabled = True
    SequenceForward.Enabled = True
    SequenceBackward.Enabled = True
    If DeckRangeFull.Value = True Then
        SequenceCurrent.Enabled = True
    Else
        SequenceCurrent.Enabled = False
    End If
    SequenceRandom.Enabled = True
    TestValue.Enabled = True
    TestPosition.Enabled = True
    TestRandom.Enabled = True
    TestToggle(0).Visible = True
    TestToggle(1).Visible = False
    TestToggle(2).Visible = False
    TestToggle(3).Visible = False
    TestMessage(0).Visible = True
    For Each Ctrl In Controls
        If Ctrl.Tag = "TCard" Or Ctrl.Tag = "TValue" Then
            Ctrl.Visible = False
        End If
    Next Ctrl
    TestC0.Visible = True
    TestV0.Visible = True
    TestNext(0).Visible = False
    TestNext(1).Visible = True
    TestNext(2).Visible = False
    TestNext(3).Visible = False
    TestNext(4).Visible = False
    TestNext(5).Visible = False
    TestShow(0).Visible = False
    TestShow(1).Visible = True
    TestShow(2).Visible = False
    TestShow(3).Visible = False
    TestTimersEnabled.Enabled = True
    TestDuration.Enabled = True
    ShowDuration.Enabled = True
    'cummTimeSeconds = CummTimeIntervals * (TimerTest.Interval / 1000)
    If Timer < StartTime Then
        StartTime = StartTime - 86400
        'error condition is test crosses midnight
    End If
    ElapsedTime = Timer - StartTime
    ResultTime.Caption = Int(ElapsedTime / 60) & " min. " & _
                Round((ElapsedTime Mod 60), 1) & " sec."
    If TestTimersEnabled Then
        ResultLabel.Visible = True
        ResultTime.Visible = True
    End If
    Exit Sub
End If
For Each Ctrl In Controls
    If Ctrl.Tag = "TCard" Or Ctrl.Tag = "TValue" Then
        Ctrl.Visible = False
    End If
Next Ctrl
TestC0.Visible = True
TestV0.Visible = True
'sets the lead cards to questions
TestRandomValue = Int(Rnd * 2)
If TestRandomCardOption Then
    TestCardMode = Int(Rnd * 3)
ElseIf TestCurrentCardOption Then
    TestCardMode = 0
ElseIf TestNextCardOption Then
    TestCardMode = 1
ElseIf TestPreviousCardOption Then
    TestCardMode = 2
End If

'turn on Test Timer if checkbox is selected
If TestTimersEnabled.Value = 1 Then
    TimerTest.Enabled = True
End If
'------------------------
'set the "TestCard" symbol correctly
    TestCurrentImage.Visible = False
    TestNextImage.Visible = False
    TestPreviousImage.Visible = False
    If TestCurrentCardOption Or TestCardMode = 0 Then
        TestCurrentImage.Visible = True
    ElseIf TestNextCardOption Or TestCardMode = 1 Then
        TestNextImage.Visible = True
    ElseIf TestPreviousCardOption Or TestCardMode = 2 Then
        TestPreviousImage.Visible = True
    End If

If DeckCurrentRangePartial.Value = True Or _
    (DeckRangeFull.Value = True And _
    SequenceCurrent.Value = True) Then
    'XXXTEST
    'this section of the code is new
    'It replaces "TestOriginalDeck(1,..." with "Deck(1,..."
    For Each Ctrl In Controls
        If Ctrl.Tag = "TValue" Then
            If Deck(1, TestOrder(TestCounter)) = _
                Val(Right(Ctrl.Name, Len(Ctrl.Name) - 5)) Then
                    Ctrl.ZOrder
                    If TestRandom Then
                        If TestRandomValue = 1 Then
                            Ctrl.Visible = True
                            TestLeftImage.Visible = False
                            TestRightImage.Visible = True
                            MnemonicLeft.Caption = MnemonicPositions(Deck(1, TestOrder(TestCounter)))
                            If MnemonicsEnabled.Value = 1 Then
                                MnemonicLeft.Visible = True
                            End If
                        End If
                    Else
                        If Not TestPosition Then
                            Ctrl.Visible = True
                            MnemonicLeft.Caption = MnemonicPositions(Deck(1, TestOrder(TestCounter)))
                            If MnemonicsEnabled.Value = 1 Then
                                MnemonicLeft.Visible = True
                            End If
                        End If
                    End If
            End If
        End If
    Next Ctrl
    For Each Ctrl In Controls
        If Ctrl.Tag = "TCard" Then
            If Deck(2, TestOrder(TestCounter)) = _
                Right(Ctrl.Name, Len(Ctrl.Name) - 5) Then
                    Ctrl.ZOrder
                    If TestRandom Then
                        If TestRandomValue = 0 Then
                            Ctrl.Visible = True
                            TestLeftImage.Visible = True
                            TestRightImage.Visible = False
                            MnemonicIndexCalc (Deck(2, TestOrder(TestCounter)))
                            MnemonicRight.Caption = MnemonicCards(MnemonicCardIndex)
                            If MnemonicsEnabled.Value = 1 Then
                                MnemonicRight.Visible = True
                            End If
                        End If
                    Else
                        If Not TestValue Then
                            Ctrl.Visible = True
                            MnemonicIndexCalc (Deck(2, TestOrder(TestCounter)))
                            MnemonicRight.Caption = MnemonicCards(MnemonicCardIndex)
                            If MnemonicsEnabled.Value = 1 Then
                                MnemonicRight.Visible = True
                            End If
                        End If
                    End If
            End If
        End If
    Next Ctrl
Else
    'this section of the code is original and works
    For Each Ctrl In Controls
        If Ctrl.Tag = "TValue" Then
            If TestOriginalDeck(1, TestOrder(TestCounter)) = _
                Val(Right(Ctrl.Name, Len(Ctrl.Name) - 5)) Then
                    Ctrl.ZOrder
                    If TestRandom Then
                        If TestRandomValue = 1 Then
                            Ctrl.Visible = True
                            TestLeftImage.Visible = False
                            TestRightImage.Visible = True
                            MnemonicLeft.Caption = _
                                MnemonicPositions(TestOriginalDeck(1, TestOrder(TestCounter)))
                            If MnemonicsEnabled.Value = 1 Then
                                MnemonicLeft.Visible = True
                            End If
                        End If
                    Else
                        If Not TestPosition Then
                            Ctrl.Visible = True
                            MnemonicLeft.Caption = _
                                MnemonicPositions(TestOriginalDeck(1, TestOrder(TestCounter)))
                            If MnemonicsEnabled.Value = 1 Then
                                MnemonicLeft.Visible = True
                            End If
                        End If
                    End If
            End If
        End If
    Next Ctrl
    For Each Ctrl In Controls
        If Ctrl.Tag = "TCard" Then
            If TestOriginalDeck(2, TestOrder(TestCounter)) = _
                Right(Ctrl.Name, Len(Ctrl.Name) - 5) Then
                    Ctrl.ZOrder
                    If TestRandom Then
                        If TestRandomValue = 0 Then
                            Ctrl.Visible = True
                            TestLeftImage.Visible = True
                            TestRightImage.Visible = False
                            MnemonicIndexCalc (TestOriginalDeck(2, TestOrder(TestCounter)))
                            MnemonicRight.Caption = MnemonicCards(MnemonicCardIndex)
                            If MnemonicsEnabled.Value = 1 Then
                                MnemonicRight.Visible = True
                            End If
                        End If
                    Else
                        If Not TestValue Then
                            Ctrl.Visible = True
                            MnemonicIndexCalc (TestOriginalDeck(2, TestOrder(TestCounter)))
                            MnemonicRight.Caption = MnemonicCards(MnemonicCardIndex)
                            If MnemonicsEnabled.Value = 1 Then
                                MnemonicRight.Visible = True
                            End If
                        End If
                    End If
            End If
        End If
    Next Ctrl
End If
MnemonicHint
End Sub

Private Sub MnemonicHint()
If TestCounter = 0 Then
    Exit Sub
    'the testing has not yet begun fully, so there is no hint to show
End If
If DeckCurrentRangePartial.Value = True Or _
    (DeckRangeFull.Value = True And _
    SequenceCurrent.Value = True) Then
'this first section handles the DeckCurrentRangePartial condition
'This first IF section complete and correct
    If TestCurrentCardOption Or TestCardMode = 0 Then
        For Each Ctrl In Controls
            If Ctrl.Tag = "TValue" Then
                If Deck(1, TestOrder(TestCounter)) = _
                    Val(Right(Ctrl.Name, Len(Ctrl.Name) - 5)) Then
                    MnemonicLeft.Caption = MnemonicPositions(Deck(1, TestOrder(TestCounter)))
                    If MnemonicsHintCheckBox.Value = 1 Or TemporaryMnemonicHint Or _
                        (MnemonicsEnabled.Value = 1 And ShowingMode = 1) Then
                        MnemonicLeft.Visible = True
                    End If
                End If
            End If
        Next Ctrl
        For Each Ctrl In Controls
            If Ctrl.Tag = "TCard" Then
                If Deck(2, TestOrder(TestCounter)) = _
                    Right(Ctrl.Name, Len(Ctrl.Name) - 5) Then
                    MnemonicIndexCalc (Deck(2, TestOrder(TestCounter)))
                    MnemonicRight.Caption = MnemonicCards(MnemonicCardIndex)
                    If MnemonicsHintCheckBox.Value = 1 Or TemporaryMnemonicHint Or _
                        (MnemonicsEnabled.Value = 1 And ShowingMode = 1) Then
                        MnemonicRight.Visible = True
                    End If
                End If
            End If
        Next Ctrl
    End If
    'the above IF section is complete and correct
    
    'This second IF section complete and correct
    If TestPreviousCardOption Or TestCardMode = 2 Then
        If (TestRandom And TestRandomValue = 1) Or TestValue Then
            For Each Ctrl In Controls
                If Ctrl.Tag = "TValue" Then
                    If Deck(1, TestOrder(TestCounter)) = _
                        Val(Right(Ctrl.Name, Len(Ctrl.Name) - 5)) Then
                        MnemonicLeft.Caption = MnemonicPositions(Deck(1, TestOrder(TestCounter)))
                        If MnemonicsHintCheckBox.Value = 1 Or TemporaryMnemonicHint Or _
                            (MnemonicsEnabled.Value = 1 And ShowingMode = 1) Then
                            MnemonicLeft.Visible = True
                        End If
                    End If
                End If
            Next Ctrl
            For Each Ctrl In Controls
                If Ctrl.Tag = "TCard" Then
                    If TestOrder(TestCounter) = 1 Then
                        If Deck(2, 52) = _
                            Right(Ctrl.Name, Len(Ctrl.Name) - 5) Then
                            MnemonicIndexCalc (Deck(2, 52))
                            MnemonicRight.Caption = MnemonicCards(MnemonicCardIndex)
                            If MnemonicsHintCheckBox.Value = 1 Or TemporaryMnemonicHint Or _
                                (MnemonicsEnabled.Value = 1 And ShowingMode = 1) Then
                                MnemonicRight.Visible = True
                            End If
                        End If
                    Else
                        If Deck(2, TestOrder(TestCounter) - 1) = _
                            Right(Ctrl.Name, Len(Ctrl.Name) - 5) Then
                            MnemonicIndexCalc (Deck(2, TestOrder(TestCounter) - 1))
                            MnemonicRight.Caption = MnemonicCards(MnemonicCardIndex)
                            If MnemonicsHintCheckBox.Value = 1 Or TemporaryMnemonicHint Or _
                                (MnemonicsEnabled.Value = 1 And ShowingMode = 1) Then
                                MnemonicRight.Visible = True
                            End If
                        End If
                    End If
                End If
            Next Ctrl
        End If
        If (TestRandom And TestRandomValue = 0) Or TestPosition Then
            For Each Ctrl In Controls
                If Ctrl.Tag = "TValue" Then
                    If TestOrder(TestCounter) = 1 Then
                        If Deck(1, 52) = _
                            Val(Right(Ctrl.Name, Len(Ctrl.Name) - 5)) Then
                            MnemonicLeft.Caption = MnemonicPositions(Deck(1, 52))
                            If MnemonicsHintCheckBox.Value = 1 Or TemporaryMnemonicHint Or _
                                (MnemonicsEnabled.Value = 1 And ShowingMode = 1) Then
                                MnemonicLeft.Visible = True
                            End If
                        End If
                    Else
                        If Deck(1, TestOrder(TestCounter) - 1) = _
                            Val(Right(Ctrl.Name, Len(Ctrl.Name) - 5)) Then
                            MnemonicLeft.Caption = MnemonicPositions(Deck(1, TestOrder(TestCounter) - 1))
                            If MnemonicsHintCheckBox.Value = 1 Or TemporaryMnemonicHint Or _
                                (MnemonicsEnabled.Value = 1 And ShowingMode = 1) Then
                                MnemonicLeft.Visible = True
                            End If
                        End If
                    End If
                End If
            Next Ctrl
            For Each Ctrl In Controls
                If Ctrl.Tag = "TCard" Then
                    If Deck(2, TestOrder(TestCounter)) = _
                        Right(Ctrl.Name, Len(Ctrl.Name) - 5) Then
                        MnemonicIndexCalc (Deck(2, TestOrder(TestCounter)))
                        MnemonicRight.Caption = MnemonicCards(MnemonicCardIndex)
                        If MnemonicsHintCheckBox.Value = 1 Or TemporaryMnemonicHint Or _
                            (MnemonicsEnabled.Value = 1 And ShowingMode = 1) Then
                            MnemonicRight.Visible = True
                        End If
                    End If
                End If
            Next Ctrl
        End If
    End If
    'the above IF section is complete and correct
    
    'This third IF section complete and correct
    If TestNextCardOption Or TestCardMode = 1 Then
        If (TestRandom And TestRandomValue = 1) Or TestValue Then
            For Each Ctrl In Controls
                If Ctrl.Tag = "TValue" Then
                    If Deck(1, TestOrder(TestCounter)) = _
                        Val(Right(Ctrl.Name, Len(Ctrl.Name) - 5)) Then
                        MnemonicLeft.Caption = MnemonicPositions(Deck(1, TestOrder(TestCounter)))
                        If MnemonicsHintCheckBox.Value = 1 Or TemporaryMnemonicHint Or _
                            (MnemonicsEnabled.Value = 1 And ShowingMode = 1) Then
                            MnemonicLeft.Visible = True
                        End If
                    End If
                End If
            Next Ctrl
            For Each Ctrl In Controls
                If Ctrl.Tag = "TCard" Then
                    If TestOrder(TestCounter) = 52 Then
                        If Deck(2, 1) = _
                            Right(Ctrl.Name, Len(Ctrl.Name) - 5) Then
                            MnemonicIndexCalc (Deck(2, 1))
                            MnemonicRight.Caption = MnemonicCards(MnemonicCardIndex)
                            If MnemonicsHintCheckBox.Value = 1 Or TemporaryMnemonicHint Or _
                                (MnemonicsEnabled.Value = 1 And ShowingMode = 1) Then
                                MnemonicRight.Visible = True
                            End If
                        End If
                    Else
                        If Deck(2, TestOrder(TestCounter) + 1) = _
                            Right(Ctrl.Name, Len(Ctrl.Name) - 5) Then
                            MnemonicIndexCalc (Deck(2, TestOrder(TestCounter) + 1))
                            MnemonicRight.Caption = MnemonicCards(MnemonicCardIndex)
                            If MnemonicsHintCheckBox.Value = 1 Or TemporaryMnemonicHint Or _
                                (MnemonicsEnabled.Value = 1 And ShowingMode = 1) Then
                                MnemonicRight.Visible = True
                            End If
                        End If
                    End If
                End If
            Next Ctrl
        End If
        If (TestRandom And TestRandomValue = 0) Or TestPosition Then
            For Each Ctrl In Controls
                If Ctrl.Tag = "TValue" Then
                    If TestOrder(TestCounter) = 52 Then
                        If Deck(1, 1) = _
                            Val(Right(Ctrl.Name, Len(Ctrl.Name) - 5)) Then
                            MnemonicLeft.Caption = MnemonicPositions(Deck(1, 1))
                            If MnemonicsHintCheckBox.Value = 1 Or TemporaryMnemonicHint Or _
                                (MnemonicsEnabled.Value = 1 And ShowingMode = 1) Then
                                MnemonicLeft.Visible = True
                            End If
                        End If
                    Else
                        If Deck(1, TestOrder(TestCounter) + 1) = _
                            Val(Right(Ctrl.Name, Len(Ctrl.Name) - 5)) Then
                            MnemonicLeft.Caption = MnemonicPositions(Deck(1, TestOrder(TestCounter) + 1))
                            If MnemonicsHintCheckBox.Value = 1 Or TemporaryMnemonicHint Or _
                                (MnemonicsEnabled.Value = 1 And ShowingMode = 1) Then
                                MnemonicLeft.Visible = True
                            End If
                        End If
                    End If
                End If
            Next Ctrl
            For Each Ctrl In Controls
                If Ctrl.Tag = "TCard" Then
                    If Deck(2, TestOrder(TestCounter)) = _
                        Right(Ctrl.Name, Len(Ctrl.Name) - 5) Then
                        MnemonicIndexCalc (Deck(2, TestOrder(TestCounter)))
                        MnemonicRight.Caption = MnemonicCards(MnemonicCardIndex)
                        If MnemonicsHintCheckBox.Value = 1 Or TemporaryMnemonicHint Or _
                            (MnemonicsEnabled.Value = 1 And ShowingMode = 1) Then
                            MnemonicRight.Visible = True
                        End If
                    End If
                End If
            Next Ctrl
        End If
    End If
    'the above IF section is complete and correct
    
 Else
 'this next section is for Full Deck or DeckRangePartial
    If TestCurrentCardOption Or TestCardMode = 0 Then
        For Each Ctrl In Controls
            If Ctrl.Tag = "TValue" Then
                If TestOriginalDeck(1, TestOrder(TestCounter)) = _
                    Val(Right(Ctrl.Name, Len(Ctrl.Name) - 5)) Then
                    MnemonicLeft.Caption = MnemonicPositions(TestOriginalDeck(1, TestOrder(TestCounter)))
                    If MnemonicsHintCheckBox.Value = 1 Or TemporaryMnemonicHint Or _
                        (MnemonicsEnabled.Value = 1 And ShowingMode = 1) Then
                        MnemonicLeft.Visible = True
                    End If
                End If
            End If
        Next Ctrl
        For Each Ctrl In Controls
            If Ctrl.Tag = "TCard" Then
                If TestOriginalDeck(2, TestOrder(TestCounter)) = _
                    Right(Ctrl.Name, Len(Ctrl.Name) - 5) Then
                    MnemonicIndexCalc (TestOriginalDeck(2, TestOrder(TestCounter)))
                    MnemonicRight.Caption = MnemonicCards(MnemonicCardIndex)
                    If MnemonicsHintCheckBox.Value = 1 Or TemporaryMnemonicHint Or _
                        (MnemonicsEnabled.Value = 1 And ShowingMode = 1) Then
                        MnemonicRight.Visible = True
                    End If
                End If
            End If
        Next Ctrl
    End If
    'the above IF section is complete and correct
    
    'This second IF section complete and correct
    If TestPreviousCardOption Or TestCardMode = 2 Then
        If (TestRandom And TestRandomValue = 1) Or TestValue Then
            For Each Ctrl In Controls
                If Ctrl.Tag = "TValue" Then
                    If TestOriginalDeck(1, TestOrder(TestCounter)) = _
                        Val(Right(Ctrl.Name, Len(Ctrl.Name) - 5)) Then
                        MnemonicLeft.Caption = MnemonicPositions(TestOriginalDeck(1, TestOrder(TestCounter)))
                        If MnemonicsHintCheckBox.Value = 1 Or TemporaryMnemonicHint Or _
                            (MnemonicsEnabled.Value = 1 And ShowingMode = 1) Then
                            MnemonicLeft.Visible = True
                        End If
                    End If
                End If
            Next Ctrl
            For Each Ctrl In Controls
                If Ctrl.Tag = "TCard" Then
                    If TestOrder(TestCounter) = 1 Then
                        If TestOriginalDeck(2, 52) = _
                            Right(Ctrl.Name, Len(Ctrl.Name) - 5) Then
                            MnemonicIndexCalc (TestOriginalDeck(2, 52))
                            MnemonicRight.Caption = MnemonicCards(MnemonicCardIndex)
                            If MnemonicsHintCheckBox.Value = 1 Or TemporaryMnemonicHint Or _
                                (MnemonicsEnabled.Value = 1 And ShowingMode = 1) Then
                                MnemonicRight.Visible = True
                            End If
                        End If
                    Else
                        If TestOriginalDeck(2, TestOrder(TestCounter) - 1) = _
                            Right(Ctrl.Name, Len(Ctrl.Name) - 5) Then
                            MnemonicIndexCalc (TestOriginalDeck(2, TestOrder(TestCounter) - 1))
                            MnemonicRight.Caption = MnemonicCards(MnemonicCardIndex)
                            If MnemonicsHintCheckBox.Value = 1 Or TemporaryMnemonicHint Or _
                                (MnemonicsEnabled.Value = 1 And ShowingMode = 1) Then
                                MnemonicRight.Visible = True
                            End If
                        End If
                    End If
                End If
            Next Ctrl
        End If
        If (TestRandom And TestRandomValue = 0) Or TestPosition Then
            For Each Ctrl In Controls
                If Ctrl.Tag = "TValue" Then
                    If TestOrder(TestCounter) = 1 Then
                        If TestOriginalDeck(1, 52) = _
                            Val(Right(Ctrl.Name, Len(Ctrl.Name) - 5)) Then
                            MnemonicLeft.Caption = MnemonicPositions(TestOriginalDeck(1, 52))
                            If MnemonicsHintCheckBox.Value = 1 Or TemporaryMnemonicHint Or _
                                (MnemonicsEnabled.Value = 1 And ShowingMode = 1) Then
                                MnemonicLeft.Visible = True
                            End If
                        End If
                    Else
                        If TestOriginalDeck(1, TestOrder(TestCounter) - 1) = _
                            Val(Right(Ctrl.Name, Len(Ctrl.Name) - 5)) Then
                            MnemonicLeft.Caption = MnemonicPositions(TestOriginalDeck(1, TestOrder(TestCounter) - 1))
                            If MnemonicsHintCheckBox.Value = 1 Or TemporaryMnemonicHint Or _
                                (MnemonicsEnabled.Value = 1 And ShowingMode = 1) Then
                                MnemonicLeft.Visible = True
                            End If
                        End If
                    End If
                End If
            Next Ctrl
            For Each Ctrl In Controls
                If Ctrl.Tag = "TCard" Then
                    If TestOriginalDeck(2, TestOrder(TestCounter)) = _
                        Right(Ctrl.Name, Len(Ctrl.Name) - 5) Then
                        MnemonicIndexCalc (TestOriginalDeck(2, TestOrder(TestCounter)))
                        MnemonicRight.Caption = MnemonicCards(MnemonicCardIndex)
                        If MnemonicsHintCheckBox.Value = 1 Or TemporaryMnemonicHint Or _
                            (MnemonicsEnabled.Value = 1 And ShowingMode = 1) Then
                            MnemonicRight.Visible = True
                        End If
                    End If
                End If
            Next Ctrl
        End If
    End If
    'the above IF section is complete and correct
    
    'This third IF section complete and correct
    If TestNextCardOption Or TestCardMode = 1 Then
        If (TestRandom And TestRandomValue = 1) Or TestValue Then
            For Each Ctrl In Controls
                If Ctrl.Tag = "TValue" Then
                    If TestOriginalDeck(1, TestOrder(TestCounter)) = _
                        Val(Right(Ctrl.Name, Len(Ctrl.Name) - 5)) Then
                        MnemonicLeft.Caption = MnemonicPositions(TestOriginalDeck(1, TestOrder(TestCounter)))
                        If MnemonicsHintCheckBox.Value = 1 Or TemporaryMnemonicHint Or _
                            (MnemonicsEnabled.Value = 1 And ShowingMode = 1) Then
                            MnemonicLeft.Visible = True
                        End If
                    End If
                End If
            Next Ctrl
            For Each Ctrl In Controls
                If Ctrl.Tag = "TCard" Then
                    If TestOrder(TestCounter) = 52 Then
                        If TestOriginalDeck(2, 1) = _
                            Right(Ctrl.Name, Len(Ctrl.Name) - 5) Then
                            MnemonicIndexCalc (TestOriginalDeck(2, 1))
                            MnemonicRight.Caption = MnemonicCards(MnemonicCardIndex)
                            If MnemonicsHintCheckBox.Value = 1 Or TemporaryMnemonicHint Or _
                                (MnemonicsEnabled.Value = 1 And ShowingMode = 1) Then
                                MnemonicRight.Visible = True
                            End If
                        End If
                    Else
                        If TestOriginalDeck(2, TestOrder(TestCounter) + 1) = _
                            Right(Ctrl.Name, Len(Ctrl.Name) - 5) Then
                            MnemonicIndexCalc (TestOriginalDeck(2, TestOrder(TestCounter) + 1))
                            MnemonicRight.Caption = MnemonicCards(MnemonicCardIndex)
                            If MnemonicsHintCheckBox.Value = 1 Or TemporaryMnemonicHint Or _
                                (MnemonicsEnabled.Value = 1 And ShowingMode = 1) Then
                                MnemonicRight.Visible = True
                            End If
                        End If
                    End If
                End If
            Next Ctrl
        End If
        If (TestRandom And TestRandomValue = 0) Or TestPosition Then
            For Each Ctrl In Controls
                If Ctrl.Tag = "TValue" Then
                    If TestOrder(TestCounter) = 52 Then
                        If TestOriginalDeck(1, 1) = _
                            Val(Right(Ctrl.Name, Len(Ctrl.Name) - 5)) Then
                            MnemonicLeft.Caption = MnemonicPositions(TestOriginalDeck(1, 1))
                            If MnemonicsHintCheckBox.Value = 1 Or TemporaryMnemonicHint Or _
                                (MnemonicsEnabled.Value = 1 And ShowingMode = 1) Then
                                MnemonicLeft.Visible = True
                            End If
                        End If
                    Else
                        If TestOriginalDeck(1, TestOrder(TestCounter) + 1) = _
                            Val(Right(Ctrl.Name, Len(Ctrl.Name) - 5)) Then
                            MnemonicLeft.Caption = MnemonicPositions(TestOriginalDeck(1, TestOrder(TestCounter) + 1))
                            If MnemonicsHintCheckBox.Value = 1 Or TemporaryMnemonicHint Or _
                                (MnemonicsEnabled.Value = 1 And ShowingMode = 1) Then
                                MnemonicLeft.Visible = True
                            End If
                        End If
                    End If
                End If
            Next Ctrl
            For Each Ctrl In Controls
                If Ctrl.Tag = "TCard" Then
                    If TestOriginalDeck(2, TestOrder(TestCounter)) = _
                        Right(Ctrl.Name, Len(Ctrl.Name) - 5) Then
                        MnemonicIndexCalc (TestOriginalDeck(2, TestOrder(TestCounter)))
                        MnemonicRight.Caption = MnemonicCards(MnemonicCardIndex)
                        If MnemonicsHintCheckBox.Value = 1 Or TemporaryMnemonicHint Or _
                            (MnemonicsEnabled.Value = 1 And ShowingMode = 1) Then
                            MnemonicRight.Visible = True
                        End If
                    End If
                End If
            Next Ctrl
        End If
    End If
    'the above IF section is complete and correct
End If
End Sub

Private Sub TestShowCard()
'cummTimeIntervals = CummTimeIntervals + TestProgressIntervals
TimerTest.Enabled = False
ShowingMode = 0
TestShow(0).Visible = True
TestShow(1).Visible = False
TestShow(2).Visible = False
TestShow(3).Visible = False
'cummTimeIntervals = CummTimeIntervals + ShowProgressIntervals
TimerShow.Enabled = False
ProgressTest.Value = 0
ProgressShow.Value = 0
TestProgressIntervals = 0
ShowProgressIntervals = 0
If TestCounter > DeckRangeCount Then
    TestCounter = 0
    TestingMode = False
    TestingStatus.Caption = "Ready"
    TestingStatus.ForeColor = &H8000&
    'set color to GREEN
    SequenceFrame.Enabled = True
    SequenceForward.Enabled = True
    SequenceBackward.Enabled = True
    If DeckRangeFull.Value = True Then
        SequenceCurrent.Enabled = True
    Else
        SequenceCurrent.Enabled = False
    End If
    SequenceRandom.Enabled = True
    TestValue.Enabled = True
    TestPosition.Enabled = True
    TestRandom.Enabled = True
    TestToggle(0).Visible = True
    TestToggle(1).Visible = False
    TestToggle(2).Visible = False
    TestToggle(3).Visible = False
    TestMessage(0).Visible = True
    For Each Ctrl In Controls
        If Ctrl.Tag = "TCard" Or Ctrl.Tag = "TValue" Then
            Ctrl.Visible = False
        End If
    Next Ctrl
    TestC0.Visible = True
    TestV0.Visible = True
    Exit Sub
End If
For Each Ctrl In Controls
    If Ctrl.Tag = "TCard" Or Ctrl.Tag = "TValue" Then
        Ctrl.Visible = False
    End If
Next Ctrl
TestC0.Visible = True
TestV0.Visible = True
'sets the lead cards to questions

'turn on Test Timer if checkbox is selected
If TestTimersEnabled.Value = 1 Then
    TimerShow.Enabled = True
End If
'------------------------
ShowingMode = 1
TestShow(0).Visible = False
TestShow(1).Visible = False
TestShow(2).Visible = True
TestShow(3).Visible = False

If DeckCurrentRangePartial.Value = True Or _
    (DeckRangeFull.Value = True And _
    SequenceCurrent.Value = True) Then
'this first section handles the DeckCurrentRangePartial condition
'This first IF section complete and correct
    If TestCurrentCardOption Or TestCardMode = 0 Then
        For Each Ctrl In Controls
            If Ctrl.Tag = "TValue" Then
                If Deck(1, TestOrder(TestCounter)) = _
                    Val(Right(Ctrl.Name, Len(Ctrl.Name) - 5)) Then
                        Ctrl.ZOrder
                        Ctrl.Visible = True
                End If
            End If
        Next Ctrl
        For Each Ctrl In Controls
            If Ctrl.Tag = "TCard" Then
                If Deck(2, TestOrder(TestCounter)) = _
                    Right(Ctrl.Name, Len(Ctrl.Name) - 5) Then
                        Ctrl.ZOrder
                        Ctrl.Visible = True
                End If
            End If
        Next Ctrl
    End If
    'the above IF section is complete and correct
    
    'This second IF section complete and correct
    If TestPreviousCardOption Or TestCardMode = 2 Then
        If (TestRandom And TestRandomValue = 1) Or TestValue Then
            For Each Ctrl In Controls
                If Ctrl.Tag = "TValue" Then
                    If Deck(1, TestOrder(TestCounter)) = _
                        Val(Right(Ctrl.Name, Len(Ctrl.Name) - 5)) Then
                            Ctrl.ZOrder
                            Ctrl.Visible = True
                    End If
                End If
            Next Ctrl
            For Each Ctrl In Controls
                If Ctrl.Tag = "TCard" Then
                    If TestOrder(TestCounter) = 1 Then
                        If Deck(2, 52) = _
                            Right(Ctrl.Name, Len(Ctrl.Name) - 5) Then
                                Ctrl.ZOrder
                                Ctrl.Visible = True
                        End If
                    Else
                        If Deck(2, TestOrder(TestCounter) - 1) = _
                            Right(Ctrl.Name, Len(Ctrl.Name) - 5) Then
                                Ctrl.ZOrder
                                Ctrl.Visible = True
                        End If
                    End If
                End If
            Next Ctrl
        End If
        If (TestRandom And TestRandomValue = 0) Or TestPosition Then
            For Each Ctrl In Controls
                If Ctrl.Tag = "TValue" Then
                    If TestOrder(TestCounter) = 1 Then
                        If Deck(1, 52) = _
                            Val(Right(Ctrl.Name, Len(Ctrl.Name) - 5)) Then
                                Ctrl.ZOrder
                                Ctrl.Visible = True
                        End If
                    Else
                        If Deck(1, TestOrder(TestCounter) - 1) = _
                            Val(Right(Ctrl.Name, Len(Ctrl.Name) - 5)) Then
                                Ctrl.ZOrder
                                Ctrl.Visible = True
                        End If
                    End If
                End If
            Next Ctrl
            For Each Ctrl In Controls
                If Ctrl.Tag = "TCard" Then
                    If Deck(2, TestOrder(TestCounter)) = _
                        Right(Ctrl.Name, Len(Ctrl.Name) - 5) Then
                            Ctrl.ZOrder
                            Ctrl.Visible = True
                    End If
                End If
            Next Ctrl
        End If
    End If
    'the above IF section is complete and correct
    
    'This third IF section complete and correct
    If TestNextCardOption Or TestCardMode = 1 Then
        If (TestRandom And TestRandomValue = 1) Or TestValue Then
            For Each Ctrl In Controls
                If Ctrl.Tag = "TValue" Then
                    If Deck(1, TestOrder(TestCounter)) = _
                        Val(Right(Ctrl.Name, Len(Ctrl.Name) - 5)) Then
                            Ctrl.ZOrder
                            Ctrl.Visible = True
                    End If
                End If
            Next Ctrl
            For Each Ctrl In Controls
                If Ctrl.Tag = "TCard" Then
                    If TestOrder(TestCounter) = 52 Then
                        If Deck(2, 1) = _
                            Right(Ctrl.Name, Len(Ctrl.Name) - 5) Then
                                Ctrl.ZOrder
                                Ctrl.Visible = True
                        End If
                    Else
                        If Deck(2, TestOrder(TestCounter) + 1) = _
                            Right(Ctrl.Name, Len(Ctrl.Name) - 5) Then
                                Ctrl.ZOrder
                                Ctrl.Visible = True
                        End If
                    End If
                End If
            Next Ctrl
        End If
        If (TestRandom And TestRandomValue = 0) Or TestPosition Then
            For Each Ctrl In Controls
                If Ctrl.Tag = "TValue" Then
                    If TestOrder(TestCounter) = 52 Then
                        If Deck(1, 1) = _
                            Val(Right(Ctrl.Name, Len(Ctrl.Name) - 5)) Then
                                Ctrl.ZOrder
                                Ctrl.Visible = True
                        End If
                    Else
                        If Deck(1, TestOrder(TestCounter) + 1) = _
                            Val(Right(Ctrl.Name, Len(Ctrl.Name) - 5)) Then
                                Ctrl.ZOrder
                                Ctrl.Visible = True
                        End If
                    End If
                End If
            Next Ctrl
            For Each Ctrl In Controls
                If Ctrl.Tag = "TCard" Then
                    If Deck(2, TestOrder(TestCounter)) = _
                        Right(Ctrl.Name, Len(Ctrl.Name) - 5) Then
                            Ctrl.ZOrder
                            Ctrl.Visible = True
                    End If
                End If
            Next Ctrl
        End If
    End If
    'the above IF section is complete and correct
 Else
 'this next section is for Full Deck or DeckRangePartial
    If TestCurrentCardOption Or TestCardMode = 0 Then
        For Each Ctrl In Controls
            If Ctrl.Tag = "TValue" Then
                If TestOriginalDeck(1, TestOrder(TestCounter)) = _
                    Val(Right(Ctrl.Name, Len(Ctrl.Name) - 5)) Then
                        Ctrl.ZOrder
                        Ctrl.Visible = True
                End If
            End If
        Next Ctrl
        For Each Ctrl In Controls
            If Ctrl.Tag = "TCard" Then
                If TestOriginalDeck(2, TestOrder(TestCounter)) = _
                    Right(Ctrl.Name, Len(Ctrl.Name) - 5) Then
                        Ctrl.ZOrder
                        Ctrl.Visible = True
                End If
            End If
        Next Ctrl
    End If
    'the above IF section is complete and correct
    
    'This second IF section complete and correct
    If TestPreviousCardOption Or TestCardMode = 2 Then
        If (TestRandom And TestRandomValue = 1) Or TestValue Then
            For Each Ctrl In Controls
                If Ctrl.Tag = "TValue" Then
                    If TestOriginalDeck(1, TestOrder(TestCounter)) = _
                        Val(Right(Ctrl.Name, Len(Ctrl.Name) - 5)) Then
                            Ctrl.ZOrder
                            Ctrl.Visible = True
                    End If
                End If
            Next Ctrl
            For Each Ctrl In Controls
                If Ctrl.Tag = "TCard" Then
                    If TestOrder(TestCounter) = 1 Then
                        If TestOriginalDeck(2, 52) = _
                            Right(Ctrl.Name, Len(Ctrl.Name) - 5) Then
                                Ctrl.ZOrder
                                Ctrl.Visible = True
                        End If
                    Else
                        If TestOriginalDeck(2, TestOrder(TestCounter) - 1) = _
                            Right(Ctrl.Name, Len(Ctrl.Name) - 5) Then
                                Ctrl.ZOrder
                                Ctrl.Visible = True
                        End If
                    End If
                End If
            Next Ctrl
        End If
        If (TestRandom And TestRandomValue = 0) Or TestPosition Then
            For Each Ctrl In Controls
                If Ctrl.Tag = "TValue" Then
                    If TestOrder(TestCounter) = 1 Then
                        If TestOriginalDeck(1, 52) = _
                            Val(Right(Ctrl.Name, Len(Ctrl.Name) - 5)) Then
                                Ctrl.ZOrder
                                Ctrl.Visible = True
                        End If
                    Else
                        If TestOriginalDeck(1, TestOrder(TestCounter) - 1) = _
                            Val(Right(Ctrl.Name, Len(Ctrl.Name) - 5)) Then
                                Ctrl.ZOrder
                                Ctrl.Visible = True
                        End If
                    End If
                End If
            Next Ctrl
            For Each Ctrl In Controls
                If Ctrl.Tag = "TCard" Then
                    If TestOriginalDeck(2, TestOrder(TestCounter)) = _
                        Right(Ctrl.Name, Len(Ctrl.Name) - 5) Then
                            Ctrl.ZOrder
                            Ctrl.Visible = True
                    End If
                End If
            Next Ctrl
        End If
    End If
    'the above IF section is complete and correct
    
    'This third IF section complete and correct
    If TestNextCardOption Or TestCardMode = 1 Then
        If (TestRandom And TestRandomValue = 1) Or TestValue Then
            For Each Ctrl In Controls
                If Ctrl.Tag = "TValue" Then
                    If TestOriginalDeck(1, TestOrder(TestCounter)) = _
                        Val(Right(Ctrl.Name, Len(Ctrl.Name) - 5)) Then
                            Ctrl.ZOrder
                            Ctrl.Visible = True
                    End If
                End If
            Next Ctrl
            For Each Ctrl In Controls
                If Ctrl.Tag = "TCard" Then
                    If TestOrder(TestCounter) = 52 Then
                        If TestOriginalDeck(2, 1) = _
                            Right(Ctrl.Name, Len(Ctrl.Name) - 5) Then
                                Ctrl.ZOrder
                                Ctrl.Visible = True
                        End If
                    Else
                        If TestOriginalDeck(2, TestOrder(TestCounter) + 1) = _
                            Right(Ctrl.Name, Len(Ctrl.Name) - 5) Then
                                Ctrl.ZOrder
                                Ctrl.Visible = True
                        End If
                    End If
                End If
            Next Ctrl
        End If
        If (TestRandom And TestRandomValue = 0) Or TestPosition Then
            For Each Ctrl In Controls
                If Ctrl.Tag = "TValue" Then
                    If TestOrder(TestCounter) = 52 Then
                        If TestOriginalDeck(1, 1) = _
                            Val(Right(Ctrl.Name, Len(Ctrl.Name) - 5)) Then
                                Ctrl.ZOrder
                                Ctrl.Visible = True
                        End If
                    Else
                        If TestOriginalDeck(1, TestOrder(TestCounter) + 1) = _
                            Val(Right(Ctrl.Name, Len(Ctrl.Name) - 5)) Then
                                Ctrl.ZOrder
                                Ctrl.Visible = True
                        End If
                    End If
                End If
            Next Ctrl
            For Each Ctrl In Controls
                If Ctrl.Tag = "TCard" Then
                    If TestOriginalDeck(2, TestOrder(TestCounter)) = _
                        Right(Ctrl.Name, Len(Ctrl.Name) - 5) Then
                            Ctrl.ZOrder
                            Ctrl.Visible = True
                    End If
                End If
            Next Ctrl
        End If
    End If
    'the above IF section is complete and correct
End If
MnemonicHint
End Sub



Private Sub TestValue_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'If TestNextCardOption Or TestPreviousCardOption Then
    TestLeftImage.Visible = False
    TestRightImage.Visible = True
'End If
End Sub

Private Sub TimerShow_Timer()
ShowProgressIntervals = ShowProgressIntervals + 1
If ProgressShow.Value < ProgressShow.Max Then
    ProgressShow.Value = ShowProgressIntervals
Else
    ShowingMode = 0
    TestShow(0).Visible = True
    TestShow(1).Visible = False
    TestShow(2).Visible = False
    TestShow(3).Visible = False
    'cummTimeIntervals = CummTimeIntervals + ShowProgressIntervals
    TimerShow.Enabled = False
    ProgressShow.Value = 0
    ShowProgressIntervals = 0
    If TestCounter = DeckRangeCount - 1 Then
        'this turns on the "Finish" button for the last card
        TestNext(0).Visible = False
        TestNext(1).Visible = False
        TestNext(2).Visible = False
        TestNext(3).Visible = False
        TestNext(4).Visible = True
        TestNext(5).Visible = False
    Else
        TestNext(0).Visible = True
        TestNext(1).Visible = False
        TestNext(2).Visible = False
        TestNext(3).Visible = False
        TestNext(4).Visible = False
        TestNext(5).Visible = False
    End If
    TestNextCard
End If
End Sub

Private Sub TimerTest_Timer()
TestProgressIntervals = TestProgressIntervals + 1
If ProgressTest.Value < ProgressTest.Max Then
    ProgressTest.Value = TestProgressIntervals
Else
    'cummTimeIntervals = CummTimeIntervals + TestProgressIntervals
    TimerTest.Enabled = False
    ProgressTest.Value = 0
    TestProgressIntervals = 0
    TestShowCard
End If
End Sub
