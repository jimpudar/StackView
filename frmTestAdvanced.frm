VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmTestAdvanced 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "StackView Advanced Test"
   ClientHeight    =   6345
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   Icon            =   "frmTestAdvanced.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   11640
   Begin VB.Frame TestQuestionsFrame 
      Caption         =   "Test Questions"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1230
      Left            =   105
      TabIndex        =   44
      Top             =   2730
      Width           =   2955
      Begin VB.OptionButton NewBottomCardQuestion 
         Caption         =   "New Bottom Card Stack Value"
         Height          =   270
         Left            =   120
         TabIndex        =   47
         Top             =   855
         Width           =   2610
      End
      Begin VB.OptionButton CardsToCutQuestion 
         Caption         =   "Cards to Cut"
         Height          =   270
         Left            =   120
         TabIndex        =   46
         Top             =   285
         Value           =   -1  'True
         Width           =   2205
      End
      Begin VB.OptionButton NewTopCardQuestion 
         Caption         =   "New Top Card Stack Value"
         Height          =   270
         Left            =   120
         TabIndex        =   45
         Top             =   570
         Width           =   2370
      End
   End
   Begin VB.Frame DeckOrderFrame 
      Caption         =   "Deck Order"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   105
      TabIndex        =   38
      Top             =   555
      Width           =   2895
      Begin VB.CheckBox CurrentDeckRandomCut 
         Caption         =   "Check1"
         Height          =   195
         Left            =   2190
         TabIndex        =   43
         Top             =   645
         Width           =   225
      End
      Begin VB.CheckBox StartingStackRandomCut 
         Caption         =   "Check1"
         Height          =   195
         Left            =   2190
         TabIndex        =   42
         Top             =   375
         Width           =   225
      End
      Begin VB.OptionButton DeckOrderStartingOption 
         Caption         =   "Starting Stack Order"
         Height          =   270
         Left            =   120
         TabIndex        =   40
         Top             =   285
         Width           =   1770
      End
      Begin VB.OptionButton DeckOrderCurrentOption 
         Caption         =   "Current Deck Order"
         Height          =   270
         Left            =   120
         TabIndex        =   39
         Top             =   570
         Value           =   -1  'True
         Width           =   1905
      End
      Begin VB.Label RandomCutLabel 
         Caption         =   "Random Cut"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1920
         TabIndex        =   41
         Top             =   150
         Width           =   885
      End
   End
   Begin VB.Frame DesiredPositionFrame 
      Caption         =   "Desired Position"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2430
      TabIndex        =   34
      Top             =   1500
      Width           =   2190
      Begin VB.TextBox DesiredPositionText 
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
         Left            =   1470
         TabIndex        =   37
         Top             =   645
         Width           =   600
      End
      Begin VB.OptionButton DesiredPositionRandom 
         Caption         =   "Random Position"
         Height          =   270
         Left            =   120
         TabIndex        =   36
         Top             =   300
         Value           =   -1  'True
         Width           =   1665
      End
      Begin VB.OptionButton DesiredPositionSpecified 
         Caption         =   "Specified Position Value"
         Height          =   420
         Left            =   120
         TabIndex        =   35
         Top             =   615
         Width           =   1365
      End
   End
   Begin VB.Frame DesiredCardFrame 
      Caption         =   "Desired Card"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   105
      TabIndex        =   30
      Top             =   1500
      Width           =   2190
      Begin VB.TextBox DesiredCardText 
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
         Left            =   1470
         TabIndex        =   33
         Top             =   645
         Width           =   600
      End
      Begin VB.OptionButton DesiredCardRandom 
         Caption         =   "Random Card"
         Height          =   270
         Left            =   135
         TabIndex        =   32
         Top             =   300
         Value           =   -1  'True
         Width           =   1470
      End
      Begin VB.OptionButton DesiredCardSpecified 
         Caption         =   "Specified Stack Value"
         Height          =   420
         Left            =   120
         TabIndex        =   31
         Top             =   615
         Width           =   1230
      End
   End
   Begin VB.Frame KnownCardFrame 
      Caption         =   "Known Card"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   3150
      TabIndex        =   26
      Top             =   570
      Width           =   1455
      Begin VB.OptionButton KnownBottomCard 
         Caption         =   "Bottom Card"
         Height          =   270
         Left            =   120
         TabIndex        =   28
         Top             =   570
         Width           =   1305
      End
      Begin VB.OptionButton KnownTopCard 
         Caption         =   "Top Card"
         Height          =   270
         Left            =   120
         TabIndex        =   27
         Top             =   285
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.Frame TestTimesFrame 
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
      Left            =   105
      TabIndex        =   11
      Top             =   4035
      Width           =   2970
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
         Text            =   "30"
         Top             =   255
         Width           =   600
      End
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
         TabIndex        =   12
         Text            =   "10"
         Top             =   630
         Width           =   600
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
         TabIndex        =   17
         Top             =   315
         Width           =   1050
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
         TabIndex        =   16
         Top             =   660
         Width           =   1155
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
         TabIndex        =   15
         Top             =   315
         Width           =   855
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
         TabIndex        =   14
         Top             =   660
         Width           =   885
      End
   End
   Begin VB.CheckBox TestTimersEnabled 
      Caption         =   "Enable Timers"
      Height          =   285
      Left            =   240
      TabIndex        =   10
      Top             =   5235
      Value           =   1  'Checked
      Width           =   1995
   End
   Begin VB.Timer TimerTest 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   9825
      Top             =   1560
   End
   Begin VB.Timer TimerShow 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   9825
      Top             =   2100
   End
   Begin VB.Frame AnswerFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1530
      Left            =   3270
      TabIndex        =   0
      Top             =   2820
      Width           =   3450
      Begin VB.TextBox AnswerInput 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1320
         TabIndex        =   1
         Top             =   450
         Width           =   1215
      End
      Begin VB.Label AnswerLabel 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cards to Cut?"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   30
         TabIndex        =   2
         Top             =   90
         Width           =   3360
      End
      Begin VB.Image Guess 
         Height          =   765
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":1CCA
         Top             =   435
         Width           =   825
      End
      Begin VB.Image InCorrect 
         Height          =   765
         Index           =   52
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":2820
         Top             =   435
         Width           =   825
      End
      Begin VB.Image InCorrect 
         Height          =   765
         Index           =   51
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":3AA1
         Top             =   435
         Width           =   825
      End
      Begin VB.Image InCorrect 
         Height          =   765
         Index           =   50
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":4C9D
         Top             =   435
         Width           =   825
      End
      Begin VB.Image InCorrect 
         Height          =   765
         Index           =   49
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":5F28
         Top             =   435
         Width           =   825
      End
      Begin VB.Image InCorrect 
         Height          =   765
         Index           =   48
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":7178
         Top             =   435
         Width           =   825
      End
      Begin VB.Image InCorrect 
         Height          =   765
         Index           =   47
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":83E8
         Top             =   435
         Width           =   825
      End
      Begin VB.Image InCorrect 
         Height          =   765
         Index           =   46
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":95A1
         Top             =   435
         Width           =   825
      End
      Begin VB.Image InCorrect 
         Height          =   765
         Index           =   45
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":A807
         Top             =   435
         Width           =   825
      End
      Begin VB.Image InCorrect 
         Height          =   765
         Index           =   44
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":BA2D
         Top             =   435
         Width           =   825
      End
      Begin VB.Image InCorrect 
         Height          =   765
         Index           =   43
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":CC11
         Top             =   435
         Width           =   825
      End
      Begin VB.Image InCorrect 
         Height          =   765
         Index           =   42
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":DE71
         Top             =   435
         Width           =   825
      End
      Begin VB.Image InCorrect 
         Height          =   765
         Index           =   41
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":F08E
         Top             =   435
         Width           =   825
      End
      Begin VB.Image InCorrect 
         Height          =   765
         Index           =   40
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":10202
         Top             =   435
         Width           =   825
      End
      Begin VB.Image InCorrect 
         Height          =   765
         Index           =   39
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":11416
         Top             =   435
         Width           =   825
      End
      Begin VB.Image InCorrect 
         Height          =   765
         Index           =   38
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":12706
         Top             =   435
         Width           =   825
      End
      Begin VB.Image InCorrect 
         Height          =   765
         Index           =   37
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":13A16
         Top             =   435
         Width           =   825
      End
      Begin VB.Image InCorrect 
         Height          =   765
         Index           =   36
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":14C49
         Top             =   435
         Width           =   825
      End
      Begin VB.Image InCorrect 
         Height          =   765
         Index           =   35
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":15F36
         Top             =   435
         Width           =   825
      End
      Begin VB.Image InCorrect 
         Height          =   765
         Index           =   34
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":171E7
         Top             =   435
         Width           =   825
      End
      Begin VB.Image InCorrect 
         Height          =   765
         Index           =   33
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":1844E
         Top             =   435
         Width           =   825
      End
      Begin VB.Image InCorrect 
         Height          =   765
         Index           =   32
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":1973F
         Top             =   435
         Width           =   825
      End
      Begin VB.Image InCorrect 
         Height          =   765
         Index           =   31
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":1A9E4
         Top             =   435
         Width           =   825
      End
      Begin VB.Image InCorrect 
         Height          =   765
         Index           =   30
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":1BBCA
         Top             =   435
         Width           =   825
      End
      Begin VB.Image InCorrect 
         Height          =   765
         Index           =   29
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":1CE69
         Top             =   435
         Width           =   825
      End
      Begin VB.Image InCorrect 
         Height          =   765
         Index           =   28
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":1E11F
         Top             =   435
         Width           =   825
      End
      Begin VB.Image InCorrect 
         Height          =   765
         Index           =   27
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":1F3EB
         Top             =   435
         Width           =   825
      End
      Begin VB.Image InCorrect 
         Height          =   765
         Index           =   26
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":205E2
         Top             =   435
         Width           =   825
      End
      Begin VB.Image InCorrect 
         Height          =   765
         Index           =   25
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":2188D
         Top             =   435
         Width           =   825
      End
      Begin VB.Image InCorrect 
         Height          =   765
         Index           =   24
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":22B1A
         Top             =   435
         Width           =   825
      End
      Begin VB.Image InCorrect 
         Height          =   765
         Index           =   23
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":23D51
         Top             =   435
         Width           =   825
      End
      Begin VB.Image InCorrect 
         Height          =   765
         Index           =   22
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":25002
         Top             =   435
         Width           =   825
      End
      Begin VB.Image InCorrect 
         Height          =   765
         Index           =   21
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":2625B
         Top             =   435
         Width           =   825
      End
      Begin VB.Image InCorrect 
         Height          =   765
         Index           =   20
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":27406
         Top             =   435
         Width           =   825
      End
      Begin VB.Image InCorrect 
         Height          =   765
         Index           =   19
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":28669
         Top             =   435
         Width           =   825
      End
      Begin VB.Image InCorrect 
         Height          =   765
         Index           =   18
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":29850
         Top             =   435
         Width           =   825
      End
      Begin VB.Image InCorrect 
         Height          =   765
         Index           =   17
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":2AA5C
         Top             =   435
         Width           =   825
      End
      Begin VB.Image InCorrect 
         Height          =   765
         Index           =   16
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":2BB52
         Top             =   435
         Width           =   825
      End
      Begin VB.Image InCorrect 
         Height          =   765
         Index           =   15
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":2CD28
         Top             =   435
         Width           =   825
      End
      Begin VB.Image InCorrect 
         Height          =   765
         Index           =   14
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":2DEE9
         Top             =   435
         Width           =   825
      End
      Begin VB.Image InCorrect 
         Height          =   765
         Index           =   13
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":2F04B
         Top             =   435
         Width           =   825
      End
      Begin VB.Image InCorrect 
         Height          =   765
         Index           =   12
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":30240
         Top             =   435
         Width           =   825
      End
      Begin VB.Image InCorrect 
         Height          =   765
         Index           =   11
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":313D0
         Top             =   435
         Width           =   825
      End
      Begin VB.Image InCorrect 
         Height          =   765
         Index           =   10
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":32485
         Top             =   435
         Width           =   825
      End
      Begin VB.Image InCorrect 
         Height          =   765
         Index           =   9
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":3360E
         Top             =   435
         Width           =   825
      End
      Begin VB.Image InCorrect 
         Height          =   765
         Index           =   8
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":34674
         Top             =   435
         Width           =   825
      End
      Begin VB.Image InCorrect 
         Height          =   765
         Index           =   7
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":3571B
         Top             =   435
         Width           =   825
      End
      Begin VB.Image InCorrect 
         Height          =   765
         Index           =   6
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":366A1
         Top             =   435
         Width           =   825
      End
      Begin VB.Image InCorrect 
         Height          =   765
         Index           =   5
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":37727
         Top             =   435
         Width           =   825
      End
      Begin VB.Image InCorrect 
         Height          =   765
         Index           =   4
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":3879F
         Top             =   435
         Width           =   825
      End
      Begin VB.Image InCorrect 
         Height          =   765
         Index           =   3
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":397A6
         Top             =   435
         Width           =   825
      End
      Begin VB.Image InCorrect 
         Height          =   765
         Index           =   2
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":3A83A
         Top             =   435
         Width           =   825
      End
      Begin VB.Image InCorrect 
         Height          =   765
         Index           =   1
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":3B886
         Top             =   435
         Width           =   825
      End
      Begin VB.Image InCorrect 
         Height          =   765
         Index           =   0
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":3C715
         Top             =   435
         Width           =   825
      End
      Begin VB.Image Correct 
         Height          =   765
         Index           =   52
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":3D72C
         Top             =   435
         Width           =   825
      End
      Begin VB.Image Correct 
         Height          =   765
         Index           =   51
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":3E9BC
         Top             =   435
         Width           =   825
      End
      Begin VB.Image Correct 
         Height          =   765
         Index           =   50
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":3FBC9
         Top             =   435
         Width           =   825
      End
      Begin VB.Image Correct 
         Height          =   765
         Index           =   49
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":40E43
         Top             =   435
         Width           =   825
      End
      Begin VB.Image Correct 
         Height          =   765
         Index           =   48
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":420B3
         Top             =   435
         Width           =   825
      End
      Begin VB.Image Correct 
         Height          =   765
         Index           =   47
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":43326
         Top             =   435
         Width           =   825
      End
      Begin VB.Image Correct 
         Height          =   765
         Index           =   46
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":444DE
         Top             =   435
         Width           =   825
      End
      Begin VB.Image Correct 
         Height          =   765
         Index           =   45
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":4575D
         Top             =   435
         Width           =   825
      End
      Begin VB.Image Correct 
         Height          =   765
         Index           =   44
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":469BC
         Top             =   435
         Width           =   825
      End
      Begin VB.Image Correct 
         Height          =   765
         Index           =   43
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":47C15
         Top             =   435
         Width           =   825
      End
      Begin VB.Image Correct 
         Height          =   765
         Index           =   42
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":48E75
         Top             =   435
         Width           =   825
      End
      Begin VB.Image Correct 
         Height          =   765
         Index           =   41
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":4A0A4
         Top             =   435
         Width           =   825
      End
      Begin VB.Image Correct 
         Height          =   765
         Index           =   40
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":4B250
         Top             =   435
         Width           =   825
      End
      Begin VB.Image Correct 
         Height          =   765
         Index           =   39
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":4C47A
         Top             =   435
         Width           =   825
      End
      Begin VB.Image Correct 
         Height          =   765
         Index           =   38
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":4D76D
         Top             =   435
         Width           =   825
      End
      Begin VB.Image Correct 
         Height          =   765
         Index           =   37
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":4EA63
         Top             =   435
         Width           =   825
      End
      Begin VB.Image Correct 
         Height          =   765
         Index           =   36
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":4FCB1
         Top             =   435
         Width           =   825
      End
      Begin VB.Image Correct 
         Height          =   765
         Index           =   35
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":50FAD
         Top             =   435
         Width           =   825
      End
      Begin VB.Image Correct 
         Height          =   765
         Index           =   34
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":522A0
         Top             =   435
         Width           =   825
      End
      Begin VB.Image Correct 
         Height          =   765
         Index           =   33
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":5358D
         Top             =   435
         Width           =   825
      End
      Begin VB.Image Correct 
         Height          =   765
         Index           =   32
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":54888
         Top             =   435
         Width           =   825
      End
      Begin VB.Image Correct 
         Height          =   765
         Index           =   31
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":55B4F
         Top             =   435
         Width           =   825
      End
      Begin VB.Image Correct 
         Height          =   765
         Index           =   30
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":56D72
         Top             =   435
         Width           =   825
      End
      Begin VB.Image Correct 
         Height          =   765
         Index           =   29
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":5801B
         Top             =   435
         Width           =   825
      End
      Begin VB.Image Correct 
         Height          =   765
         Index           =   28
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":592B5
         Top             =   435
         Width           =   825
      End
      Begin VB.Image Correct 
         Height          =   765
         Index           =   27
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":5A552
         Top             =   435
         Width           =   825
      End
      Begin VB.Image Correct 
         Height          =   765
         Index           =   26
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":5B72D
         Top             =   435
         Width           =   825
      End
      Begin VB.Image Correct 
         Height          =   765
         Index           =   25
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":5C9E8
         Top             =   435
         Width           =   825
      End
      Begin VB.Image Correct 
         Height          =   765
         Index           =   24
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":5DC90
         Top             =   435
         Width           =   825
      End
      Begin VB.Image Correct 
         Height          =   765
         Index           =   23
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":5EF36
         Top             =   435
         Width           =   825
      End
      Begin VB.Image Correct 
         Height          =   765
         Index           =   22
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":601CB
         Top             =   435
         Width           =   825
      End
      Begin VB.Image Correct 
         Height          =   765
         Index           =   21
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":61426
         Top             =   435
         Width           =   825
      End
      Begin VB.Image Correct 
         Height          =   765
         Index           =   20
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":625FB
         Top             =   435
         Width           =   825
      End
      Begin VB.Image Correct 
         Height          =   765
         Index           =   19
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":63840
         Top             =   435
         Width           =   825
      End
      Begin VB.Image Correct 
         Height          =   765
         Index           =   18
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":649C4
         Top             =   435
         Width           =   825
      End
      Begin VB.Image Correct 
         Height          =   765
         Index           =   17
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":65B68
         Top             =   435
         Width           =   825
      End
      Begin VB.Image Correct 
         Height          =   765
         Index           =   16
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":66BF1
         Top             =   435
         Width           =   825
      End
      Begin VB.Image Correct 
         Height          =   765
         Index           =   15
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":67D6B
         Top             =   435
         Width           =   825
      End
      Begin VB.Image Correct 
         Height          =   765
         Index           =   14
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":68EE6
         Top             =   435
         Width           =   825
      End
      Begin VB.Image Correct 
         Height          =   765
         Index           =   13
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":6A048
         Top             =   435
         Width           =   825
      End
      Begin VB.Image Correct 
         Height          =   765
         Index           =   12
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":6B205
         Top             =   435
         Width           =   825
      End
      Begin VB.Image Correct 
         Height          =   765
         Index           =   11
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":6C38B
         Top             =   435
         Width           =   825
      End
      Begin VB.Image Correct 
         Height          =   765
         Index           =   10
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":6D42A
         Top             =   435
         Width           =   825
      End
      Begin VB.Image Correct 
         Height          =   765
         Index           =   9
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":6E580
         Top             =   435
         Width           =   825
      End
      Begin VB.Image Correct 
         Height          =   765
         Index           =   8
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":6F5BC
         Top             =   435
         Width           =   825
      End
      Begin VB.Image Correct 
         Height          =   765
         Index           =   7
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":7063C
         Top             =   435
         Width           =   825
      End
      Begin VB.Image Correct 
         Height          =   765
         Index           =   6
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":715AF
         Top             =   435
         Width           =   825
      End
      Begin VB.Image Correct 
         Height          =   765
         Index           =   5
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":72639
         Top             =   435
         Width           =   825
      End
      Begin VB.Image Correct 
         Height          =   765
         Index           =   4
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":7369D
         Top             =   435
         Width           =   825
      End
      Begin VB.Image Correct 
         Height          =   765
         Index           =   3
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":7469B
         Top             =   435
         Width           =   825
      End
      Begin VB.Image Correct 
         Height          =   765
         Index           =   2
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":75710
         Top             =   435
         Width           =   825
      End
      Begin VB.Image Correct 
         Height          =   765
         Index           =   1
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":7674E
         Top             =   435
         Width           =   825
      End
      Begin VB.Image Correct 
         Height          =   765
         Index           =   0
         Left            =   255
         Picture         =   "frmTestAdvanced.frx":7759A
         Top             =   435
         Width           =   825
      End
   End
   Begin MSComctlLib.ProgressBar ProgressTest 
      Height          =   255
      Left            =   4005
      TabIndex        =   3
      Top             =   4530
      Width           =   2850
      _ExtentX        =   5027
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar ProgressShow 
      Height          =   255
      Left            =   4020
      TabIndex        =   4
      Top             =   4800
      Width           =   2850
      _ExtentX        =   5027
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar ProgressCardsRemaining 
      Height          =   2220
      Left            =   6930
      TabIndex        =   7
      Top             =   2775
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   3916
      _Version        =   393216
      Appearance      =   1
      Max             =   30
      Orientation     =   1
      Scrolling       =   1
   End
   Begin VB.Label ExplanationLabel2 
      Alignment       =   2  'Center
      Caption         =   "A score of 100 or more is excellent"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   7980
      TabIndex        =   56
      Top             =   6060
      Visible         =   0   'False
      Width           =   2640
   End
   Begin VB.Label ExplanationLabel1 
      Alignment       =   2  'Center
      Caption         =   "Score = (# Correct * 3 / Seconds) * 100"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   7995
      TabIndex        =   55
      Top             =   5850
      Visible         =   0   'False
      Width           =   2595
   End
   Begin VB.Label ScoreLabel2 
      Caption         =   "104 min. 60 sec."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   345
      Left            =   9000
      TabIndex        =   54
      Top             =   5460
      Visible         =   0   'False
      Width           =   2475
   End
   Begin VB.Label PercentCorrectLabel2 
      Caption         =   "104 min. 60 sec."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   345
      Left            =   9000
      TabIndex        =   53
      Top             =   5160
      Visible         =   0   'False
      Width           =   2475
   End
   Begin VB.Label ScoreLabel1 
      Alignment       =   1  'Right Justify
      Caption         =   "Score:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   345
      Left            =   7785
      TabIndex        =   52
      Top             =   5460
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label PercentCorrectLabel1 
      Alignment       =   1  'Right Justify
      Caption         =   "% Correct:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   345
      Left            =   7410
      TabIndex        =   51
      Top             =   5160
      Visible         =   0   'False
      Width           =   1425
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
      Left            =   8220
      TabIndex        =   50
      Top             =   4455
      Width           =   2355
   End
   Begin VB.Label ResultLabel 
      Alignment       =   1  'Right Justify
      Caption         =   "Time: "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   345
      Left            =   8085
      TabIndex        =   49
      Top             =   4860
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label ResultTime 
      Caption         =   "104 min. 60 sec."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   345
      Left            =   9000
      TabIndex        =   48
      Top             =   4860
      Visible         =   0   'False
      Width           =   2475
   End
   Begin VB.Image TestToggle 
      Height          =   600
      Index           =   0
      Left            =   8505
      Picture         =   "frmTestAdvanced.frx":7859D
      Top             =   2925
      Width           =   1740
   End
   Begin VB.Image TestToggle 
      Height          =   600
      Index           =   3
      Left            =   8505
      Picture         =   "frmTestAdvanced.frx":78F18
      Top             =   2925
      Visible         =   0   'False
      Width           =   1740
   End
   Begin VB.Image TestToggle 
      Height          =   600
      Index           =   2
      Left            =   8505
      Picture         =   "frmTestAdvanced.frx":79844
      Top             =   2925
      Visible         =   0   'False
      Width           =   1740
   End
   Begin VB.Image TestToggle 
      Height          =   600
      Index           =   1
      Left            =   8505
      Picture         =   "frmTestAdvanced.frx":7A22A
      Top             =   2925
      Visible         =   0   'False
      Width           =   1740
   End
   Begin VB.Image TestNext 
      Height          =   600
      Index           =   1
      Left            =   9570
      Picture         =   "frmTestAdvanced.frx":7AB0B
      Top             =   3585
      Visible         =   0   'False
      Width           =   1740
   End
   Begin VB.Image TestShow 
      Height          =   600
      Index           =   1
      Left            =   7575
      Picture         =   "frmTestAdvanced.frx":7B3E5
      Top             =   3585
      Visible         =   0   'False
      Width           =   1740
   End
   Begin VB.Image TestShow 
      Height          =   600
      Index           =   0
      Left            =   7575
      Picture         =   "frmTestAdvanced.frx":7BD2F
      Top             =   3585
      Width           =   1740
   End
   Begin VB.Image TestNext 
      Height          =   600
      Index           =   0
      Left            =   9570
      Picture         =   "frmTestAdvanced.frx":7C744
      Top             =   3585
      Width           =   1740
   End
   Begin VB.Image TestNext 
      Height          =   600
      Index           =   3
      Left            =   9570
      Picture         =   "frmTestAdvanced.frx":7D0C8
      Top             =   3585
      Visible         =   0   'False
      Width           =   1740
   End
   Begin VB.Image TestNext 
      Height          =   600
      Index           =   4
      Left            =   9570
      Picture         =   "frmTestAdvanced.frx":7E40E
      Top             =   3585
      Visible         =   0   'False
      Width           =   1740
   End
   Begin VB.Image TestNext 
      Height          =   600
      Index           =   5
      Left            =   9570
      Picture         =   "frmTestAdvanced.frx":7F75B
      Top             =   3585
      Visible         =   0   'False
      Width           =   1740
   End
   Begin VB.Image TestShow 
      Height          =   600
      Index           =   2
      Left            =   7575
      Picture         =   "frmTestAdvanced.frx":80ADD
      Top             =   3585
      Visible         =   0   'False
      Width           =   1740
   End
   Begin VB.Image TestShow 
      Height          =   600
      Index           =   3
      Left            =   7575
      Picture         =   "frmTestAdvanced.frx":81E3B
      Top             =   3585
      Visible         =   0   'False
      Width           =   1740
   End
   Begin VB.Image TestNext 
      Height          =   600
      Index           =   2
      Left            =   9570
      Picture         =   "frmTestAdvanced.frx":83067
      Top             =   3585
      Visible         =   0   'False
      Width           =   1740
   End
   Begin VB.Image DesiredPosition 
      Height          =   1920
      Index           =   0
      Left            =   8280
      Picture         =   "frmTestAdvanced.frx":84357
      Top             =   720
      Width           =   1350
   End
   Begin VB.Label Label12 
      Caption         =   "Advanced Test Design"
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
      Left            =   105
      TabIndex        =   29
      Top             =   120
      Width           =   2520
   End
   Begin VB.Image DesiredCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   0
      Left            =   6765
      Picture         =   "frmTestAdvanced.frx":8521B
      Tag             =   "DesiredCardBack"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image KnownCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   0
      Left            =   4995
      Picture         =   "frmTestAdvanced.frx":865ED
      Tag             =   "KnownCardBack"
      Top             =   720
      Width           =   1350
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
      Left            =   6690
      TabIndex        =   25
      Top             =   5370
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Position"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   8475
      TabIndex        =   24
      Top             =   345
      Width           =   975
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Desired"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8475
      TabIndex        =   23
      Top             =   90
      Width           =   960
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Card"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6885
      TabIndex        =   22
      Top             =   375
      Width           =   975
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Desired"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6885
      TabIndex        =   21
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Card"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5115
      TabIndex        =   20
      Top             =   405
      Width           =   975
   End
   Begin VB.Label KnownCardLabel 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Top"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5115
      TabIndex        =   19
      Top             =   150
      Width           =   975
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
      Left            =   90
      TabIndex        =   18
      Top             =   5865
      Width           =   3960
   End
   Begin VB.Label TestCardsRemaining 
      Alignment       =   2  'Center
      Caption         =   "30"
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
      Left            =   6855
      TabIndex        =   9
      Top             =   4995
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label TestCardsLabel1 
      Alignment       =   2  'Center
      Caption         =   "Tests"
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
      Left            =   6855
      TabIndex        =   8
      Top             =   5175
      Visible         =   0   'False
      Width           =   360
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
      TabIndex        =   6
      Top             =   4515
      Width           =   660
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
      Left            =   3135
      TabIndex        =   5
      Top             =   4785
      Width           =   750
   End
   Begin VB.Image DesiredPosition 
      Height          =   1920
      Index           =   52
      Left            =   8280
      Picture         =   "frmTestAdvanced.frx":879BF
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image DesiredPosition 
      Height          =   1920
      Index           =   51
      Left            =   8280
      Picture         =   "frmTestAdvanced.frx":88BF7
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image DesiredPosition 
      Height          =   1920
      Index           =   50
      Left            =   8280
      Picture         =   "frmTestAdvanced.frx":89CA2
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image DesiredPosition 
      Height          =   1920
      Index           =   49
      Left            =   8280
      Picture         =   "frmTestAdvanced.frx":8AE2F
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image DesiredPosition 
      Height          =   1920
      Index           =   48
      Left            =   8280
      Picture         =   "frmTestAdvanced.frx":8BF83
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image DesiredPosition 
      Height          =   1920
      Index           =   47
      Left            =   8280
      Picture         =   "frmTestAdvanced.frx":8D0EC
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image DesiredPosition 
      Height          =   1920
      Index           =   46
      Left            =   8280
      Picture         =   "frmTestAdvanced.frx":8E15B
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image DesiredPosition 
      Height          =   1920
      Index           =   45
      Left            =   8280
      Picture         =   "frmTestAdvanced.frx":8F2FE
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image DesiredPosition 
      Height          =   1920
      Index           =   44
      Left            =   8280
      Picture         =   "frmTestAdvanced.frx":90402
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image DesiredPosition 
      Height          =   1920
      Index           =   43
      Left            =   8280
      Picture         =   "frmTestAdvanced.frx":91457
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image DesiredPosition 
      Height          =   1920
      Index           =   42
      Left            =   8280
      Picture         =   "frmTestAdvanced.frx":9257F
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image DesiredPosition 
      Height          =   1920
      Index           =   41
      Left            =   8280
      Picture         =   "frmTestAdvanced.frx":9373E
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image DesiredPosition 
      Height          =   1920
      Index           =   40
      Left            =   8280
      Picture         =   "frmTestAdvanced.frx":94739
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image DesiredPosition 
      Height          =   1920
      Index           =   39
      Left            =   8280
      Picture         =   "frmTestAdvanced.frx":9585E
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image DesiredPosition 
      Height          =   1920
      Index           =   38
      Left            =   8280
      Picture         =   "frmTestAdvanced.frx":96A87
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image DesiredPosition 
      Height          =   1920
      Index           =   37
      Left            =   8280
      Picture         =   "frmTestAdvanced.frx":97CC0
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image DesiredPosition 
      Height          =   1920
      Index           =   36
      Left            =   8280
      Picture         =   "frmTestAdvanced.frx":98DD3
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image DesiredPosition 
      Height          =   1920
      Index           =   35
      Left            =   8280
      Picture         =   "frmTestAdvanced.frx":9A011
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image DesiredPosition 
      Height          =   1920
      Index           =   34
      Left            =   8280
      Picture         =   "frmTestAdvanced.frx":9B1EF
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image DesiredPosition 
      Height          =   1920
      Index           =   33
      Left            =   8280
      Picture         =   "frmTestAdvanced.frx":9C305
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image DesiredPosition 
      Height          =   1920
      Index           =   32
      Left            =   8280
      Picture         =   "frmTestAdvanced.frx":9D4B5
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image DesiredPosition 
      Height          =   1920
      Index           =   31
      Left            =   8280
      Picture         =   "frmTestAdvanced.frx":9E6D6
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image DesiredPosition 
      Height          =   1920
      Index           =   30
      Left            =   8280
      Picture         =   "frmTestAdvanced.frx":9F73F
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image DesiredPosition 
      Height          =   1920
      Index           =   29
      Left            =   8280
      Picture         =   "frmTestAdvanced.frx":A090C
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image DesiredPosition 
      Height          =   1920
      Index           =   28
      Left            =   8280
      Picture         =   "frmTestAdvanced.frx":A1B6D
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image DesiredPosition 
      Height          =   1920
      Index           =   27
      Left            =   8280
      Picture         =   "frmTestAdvanced.frx":A2E28
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image DesiredPosition 
      Height          =   1920
      Index           =   26
      Left            =   8280
      Picture         =   "frmTestAdvanced.frx":A3F8E
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image DesiredPosition 
      Height          =   1920
      Index           =   25
      Left            =   8280
      Picture         =   "frmTestAdvanced.frx":A51DF
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image DesiredPosition 
      Height          =   1920
      Index           =   24
      Left            =   8280
      Picture         =   "frmTestAdvanced.frx":A63F6
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image DesiredPosition 
      Height          =   1920
      Index           =   23
      Left            =   8280
      Picture         =   "frmTestAdvanced.frx":A7574
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image DesiredPosition 
      Height          =   1920
      Index           =   22
      Left            =   8280
      Picture         =   "frmTestAdvanced.frx":A876F
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image DesiredPosition 
      Height          =   1920
      Index           =   21
      Left            =   8280
      Picture         =   "frmTestAdvanced.frx":A99DC
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image DesiredPosition 
      Height          =   1920
      Index           =   20
      Left            =   8280
      Picture         =   "frmTestAdvanced.frx":AAAD6
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image DesiredPosition 
      Height          =   1920
      Index           =   19
      Left            =   8280
      Picture         =   "frmTestAdvanced.frx":ABD0D
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image DesiredPosition 
      Height          =   1920
      Index           =   18
      Left            =   8280
      Picture         =   "frmTestAdvanced.frx":ACDE1
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image DesiredPosition 
      Height          =   1920
      Index           =   17
      Left            =   8280
      Picture         =   "frmTestAdvanced.frx":ADEE3
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image DesiredPosition 
      Height          =   1920
      Index           =   16
      Left            =   8280
      Picture         =   "frmTestAdvanced.frx":AEEAC
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image DesiredPosition 
      Height          =   1920
      Index           =   15
      Left            =   8280
      Picture         =   "frmTestAdvanced.frx":AFF9E
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image DesiredPosition 
      Height          =   1920
      Index           =   14
      Left            =   8280
      Picture         =   "frmTestAdvanced.frx":B1076
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image DesiredPosition 
      Height          =   1920
      Index           =   13
      Left            =   8280
      Picture         =   "frmTestAdvanced.frx":B20FE
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image DesiredPosition 
      Height          =   1920
      Index           =   12
      Left            =   8280
      Picture         =   "frmTestAdvanced.frx":B3141
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image DesiredPosition 
      Height          =   1920
      Index           =   11
      Left            =   8280
      Picture         =   "frmTestAdvanced.frx":B4209
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image DesiredPosition 
      Height          =   1920
      Index           =   10
      Left            =   8280
      Picture         =   "frmTestAdvanced.frx":B5104
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image DesiredPosition 
      Height          =   1920
      Index           =   9
      Left            =   8280
      Picture         =   "frmTestAdvanced.frx":B61A4
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image DesiredPosition 
      Height          =   1920
      Index           =   8
      Left            =   8280
      Picture         =   "frmTestAdvanced.frx":B7110
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image DesiredPosition 
      Height          =   1920
      Index           =   7
      Left            =   8280
      Picture         =   "frmTestAdvanced.frx":B807A
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image DesiredPosition 
      Height          =   1920
      Index           =   6
      Left            =   8280
      Picture         =   "frmTestAdvanced.frx":B8EDF
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image DesiredPosition 
      Height          =   1920
      Index           =   5
      Left            =   8280
      Picture         =   "frmTestAdvanced.frx":B9E6F
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image DesiredPosition 
      Height          =   1920
      Index           =   4
      Left            =   8280
      Picture         =   "frmTestAdvanced.frx":BAD7E
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image DesiredPosition 
      Height          =   1920
      Index           =   3
      Left            =   8280
      Picture         =   "frmTestAdvanced.frx":BBC03
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image DesiredPosition 
      Height          =   1920
      Index           =   2
      Left            =   8280
      Picture         =   "frmTestAdvanced.frx":BCB21
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image DesiredPosition 
      Height          =   1920
      Index           =   1
      Left            =   8280
      Picture         =   "frmTestAdvanced.frx":BDACD
      Top             =   720
      Width           =   1350
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FF8080&
      FillStyle       =   0  'Solid
      Height          =   1650
      Left            =   3210
      Top             =   2760
      Width           =   3585
   End
   Begin VB.Image DesiredCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   52
      Left            =   6765
      Picture         =   "frmTestAdvanced.frx":BE86A
      Tag             =   "DesiredCardAC"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image DesiredCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   1
      Left            =   6765
      Picture         =   "frmTestAdvanced.frx":BF05F
      Tag             =   "DesiredCard2C"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image DesiredCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   2
      Left            =   6765
      Picture         =   "frmTestAdvanced.frx":BF6E8
      Tag             =   "DesiredCard2D"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image DesiredCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   3
      Left            =   6765
      Picture         =   "frmTestAdvanced.frx":BFF2A
      Tag             =   "DesiredCard2H"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image DesiredCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   4
      Left            =   6765
      Picture         =   "frmTestAdvanced.frx":C07B6
      Tag             =   "DesiredCard2S"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image DesiredCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   5
      Left            =   6765
      Picture         =   "frmTestAdvanced.frx":C1050
      Tag             =   "DesiredCard3C"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image DesiredCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   6
      Left            =   6765
      Picture         =   "frmTestAdvanced.frx":C1996
      Tag             =   "DesiredCard3D"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image DesiredCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   7
      Left            =   6765
      Picture         =   "frmTestAdvanced.frx":C2286
      Tag             =   "DesiredCard3H"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image DesiredCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   8
      Left            =   6765
      Picture         =   "frmTestAdvanced.frx":C2BB7
      Tag             =   "DesiredCard3S"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image DesiredCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   9
      Left            =   6765
      Picture         =   "frmTestAdvanced.frx":C34BC
      Tag             =   "DesiredCard4C"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image DesiredCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   10
      Left            =   6765
      Picture         =   "frmTestAdvanced.frx":C3EF8
      Tag             =   "DesiredCard4D"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image DesiredCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   11
      Left            =   6765
      Picture         =   "frmTestAdvanced.frx":C4828
      Tag             =   "DesiredCard4H"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image DesiredCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   12
      Left            =   6765
      Picture         =   "frmTestAdvanced.frx":C51A0
      Tag             =   "DesiredCard4S"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image DesiredCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   13
      Left            =   6765
      Picture         =   "frmTestAdvanced.frx":C5B8E
      Tag             =   "DesiredCard5C"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image DesiredCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   14
      Left            =   6765
      Picture         =   "frmTestAdvanced.frx":C669F
      Tag             =   "DesiredCard5D"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image DesiredCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   15
      Left            =   6765
      Picture         =   "frmTestAdvanced.frx":C7081
      Tag             =   "DesiredCard5H"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image DesiredCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   16
      Left            =   6765
      Picture         =   "frmTestAdvanced.frx":C7AAC
      Tag             =   "DesiredCard5S"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image DesiredCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   17
      Left            =   6765
      Picture         =   "frmTestAdvanced.frx":C855B
      Tag             =   "DesiredCard6C"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image DesiredCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   18
      Left            =   6765
      Picture         =   "frmTestAdvanced.frx":C91BD
      Tag             =   "DesiredCard6D"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image DesiredCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   19
      Left            =   6765
      Picture         =   "frmTestAdvanced.frx":C9C2F
      Tag             =   "DesiredCard6H"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image DesiredCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   20
      Left            =   6765
      Picture         =   "frmTestAdvanced.frx":CA71A
      Tag             =   "DesiredCard6S"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image DesiredCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   21
      Left            =   6765
      Picture         =   "frmTestAdvanced.frx":CB20E
      Tag             =   "DesiredCard7C"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image DesiredCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   22
      Left            =   6765
      Picture         =   "frmTestAdvanced.frx":CBE90
      Tag             =   "DesiredCard7D"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image DesiredCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   23
      Left            =   6765
      Picture         =   "frmTestAdvanced.frx":CC984
      Tag             =   "DesiredCard7H"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image DesiredCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   24
      Left            =   6765
      Picture         =   "frmTestAdvanced.frx":CD4F3
      Tag             =   "DesiredCard7S"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image DesiredCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   25
      Left            =   6765
      Picture         =   "frmTestAdvanced.frx":CE066
      Tag             =   "DesiredCard8C"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image DesiredCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   26
      Left            =   6765
      Picture         =   "frmTestAdvanced.frx":CEE2C
      Tag             =   "DesiredCard8D"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image DesiredCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   27
      Left            =   6765
      Picture         =   "frmTestAdvanced.frx":CF9D9
      Tag             =   "DesiredCard8H"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image DesiredCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   28
      Left            =   6765
      Picture         =   "frmTestAdvanced.frx":D0627
      Tag             =   "DesiredCard8S"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image DesiredCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   29
      Left            =   6765
      Picture         =   "frmTestAdvanced.frx":D1276
      Tag             =   "DesiredCard9C"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image DesiredCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   30
      Left            =   6765
      Picture         =   "frmTestAdvanced.frx":D20CF
      Tag             =   "DesiredCard9D"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image DesiredCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   31
      Left            =   6765
      Picture         =   "frmTestAdvanced.frx":D2D01
      Tag             =   "DesiredCard9H"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image DesiredCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   32
      Left            =   6765
      Picture         =   "frmTestAdvanced.frx":D39ED
      Tag             =   "DesiredCard9S"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image DesiredCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   33
      Left            =   6765
      Picture         =   "frmTestAdvanced.frx":D46BD
      Tag             =   "DesiredCard10C"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image DesiredCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   34
      Left            =   6765
      Picture         =   "frmTestAdvanced.frx":D55C2
      Tag             =   "DesiredCard10D"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image DesiredCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   35
      Left            =   6765
      Picture         =   "frmTestAdvanced.frx":D6250
      Tag             =   "DesiredCard10H"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image DesiredCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   36
      Left            =   6765
      Picture         =   "frmTestAdvanced.frx":D6FC1
      Tag             =   "DesiredCard10S"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image DesiredCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   37
      Left            =   6765
      Picture         =   "frmTestAdvanced.frx":D7D33
      Tag             =   "DesiredCardJC"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image DesiredCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   38
      Left            =   6765
      Picture         =   "frmTestAdvanced.frx":D8C79
      Tag             =   "DesiredCardJD"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image DesiredCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   39
      Left            =   6765
      Picture         =   "frmTestAdvanced.frx":D9CB7
      Tag             =   "DesiredCardJH"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image DesiredCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   40
      Left            =   6765
      Picture         =   "frmTestAdvanced.frx":DACC1
      Tag             =   "DesiredCardJS"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image DesiredCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   41
      Left            =   6765
      Picture         =   "frmTestAdvanced.frx":DBC3B
      Tag             =   "DesiredCardQC"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image DesiredCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   42
      Left            =   6765
      Picture         =   "frmTestAdvanced.frx":DCBA3
      Tag             =   "DesiredCardQD"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image DesiredCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   43
      Left            =   6765
      Picture         =   "frmTestAdvanced.frx":DDC11
      Tag             =   "DesiredCardQH"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image DesiredCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   44
      Left            =   6765
      Picture         =   "frmTestAdvanced.frx":DEC30
      Tag             =   "DesiredCardQS"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image DesiredCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   45
      Left            =   6765
      Picture         =   "frmTestAdvanced.frx":DFC2F
      Tag             =   "DesiredCardKC"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image DesiredCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   46
      Left            =   6765
      Picture         =   "frmTestAdvanced.frx":E0B7B
      Tag             =   "DesiredCardKD"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image DesiredCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   47
      Left            =   6765
      Picture         =   "frmTestAdvanced.frx":E1BA7
      Tag             =   "DesiredCardKH"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image DesiredCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   48
      Left            =   6765
      Picture         =   "frmTestAdvanced.frx":E2BFC
      Tag             =   "DesiredCardKS"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image DesiredCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   49
      Left            =   6765
      Picture         =   "frmTestAdvanced.frx":E3C4A
      Tag             =   "DesiredCardAD"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image DesiredCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   50
      Left            =   6765
      Picture         =   "frmTestAdvanced.frx":E444B
      Tag             =   "DesiredCardAH"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image DesiredCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   51
      Left            =   6765
      Picture         =   "frmTestAdvanced.frx":E4C44
      Tag             =   "DesiredCardAS"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image KnownCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   52
      Left            =   4995
      Picture         =   "frmTestAdvanced.frx":E554D
      Tag             =   "KnownCardAC"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image KnownCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   1
      Left            =   4995
      Picture         =   "frmTestAdvanced.frx":E5D42
      Tag             =   "KnownCard2C"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image KnownCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   40
      Left            =   4995
      Picture         =   "frmTestAdvanced.frx":E63CB
      Tag             =   "KnownCard2D"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image KnownCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   14
      Left            =   4995
      Picture         =   "frmTestAdvanced.frx":E6C0D
      Tag             =   "KnownCard2H"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image KnownCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   27
      Left            =   4995
      Picture         =   "frmTestAdvanced.frx":E7499
      Tag             =   "KnownCard2S"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image KnownCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   2
      Left            =   4995
      Picture         =   "frmTestAdvanced.frx":E7D33
      Tag             =   "KnownCard3C"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image KnownCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   41
      Left            =   4995
      Picture         =   "frmTestAdvanced.frx":E8679
      Tag             =   "KnownCard3D"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image KnownCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   15
      Left            =   4995
      Picture         =   "frmTestAdvanced.frx":E8F69
      Tag             =   "KnownCard3H"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image KnownCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   28
      Left            =   4995
      Picture         =   "frmTestAdvanced.frx":E989A
      Tag             =   "KnownCard3S"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image KnownCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   3
      Left            =   4995
      Picture         =   "frmTestAdvanced.frx":EA19F
      Tag             =   "KnownCard4C"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image KnownCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   42
      Left            =   4995
      Picture         =   "frmTestAdvanced.frx":EABDB
      Tag             =   "KnownCard4D"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image KnownCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   16
      Left            =   4995
      Picture         =   "frmTestAdvanced.frx":EB50B
      Tag             =   "KnownCard4H"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image KnownCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   29
      Left            =   4995
      Picture         =   "frmTestAdvanced.frx":EBE83
      Tag             =   "KnownCard4S"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image KnownCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   4
      Left            =   4995
      Picture         =   "frmTestAdvanced.frx":EC871
      Tag             =   "KnownCard5C"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image KnownCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   43
      Left            =   4995
      Picture         =   "frmTestAdvanced.frx":ED382
      Tag             =   "KnownCard5D"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image KnownCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   17
      Left            =   4995
      Picture         =   "frmTestAdvanced.frx":EDD64
      Tag             =   "KnownCard5H"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image KnownCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   30
      Left            =   4995
      Picture         =   "frmTestAdvanced.frx":EE78F
      Tag             =   "KnownCard5S"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image KnownCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   5
      Left            =   4995
      Picture         =   "frmTestAdvanced.frx":EF23E
      Tag             =   "KnownCard6C"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image KnownCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   44
      Left            =   4995
      Picture         =   "frmTestAdvanced.frx":EFEA0
      Tag             =   "KnownCard6D"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image KnownCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   18
      Left            =   4995
      Picture         =   "frmTestAdvanced.frx":F0912
      Tag             =   "KnownCard6H"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image KnownCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   31
      Left            =   4995
      Picture         =   "frmTestAdvanced.frx":F13FD
      Tag             =   "KnownCard6S"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image KnownCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   6
      Left            =   4995
      Picture         =   "frmTestAdvanced.frx":F1EF1
      Tag             =   "KnownCard7C"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image KnownCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   45
      Left            =   4995
      Picture         =   "frmTestAdvanced.frx":F2B73
      Tag             =   "KnownCard7D"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image KnownCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   19
      Left            =   4995
      Picture         =   "frmTestAdvanced.frx":F3667
      Tag             =   "KnownCard7H"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image KnownCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   32
      Left            =   4995
      Picture         =   "frmTestAdvanced.frx":F41D6
      Tag             =   "KnownCard7S"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image KnownCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   7
      Left            =   4995
      Picture         =   "frmTestAdvanced.frx":F4D49
      Tag             =   "KnownCard8C"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image KnownCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   46
      Left            =   4995
      Picture         =   "frmTestAdvanced.frx":F5B0F
      Tag             =   "KnownCard8D"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image KnownCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   20
      Left            =   4995
      Picture         =   "frmTestAdvanced.frx":F66BC
      Tag             =   "KnownCard8H"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image KnownCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   33
      Left            =   4995
      Picture         =   "frmTestAdvanced.frx":F730A
      Tag             =   "KnownCard8S"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image KnownCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   8
      Left            =   4995
      Picture         =   "frmTestAdvanced.frx":F7F59
      Tag             =   "KnownCard9C"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image KnownCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   47
      Left            =   4995
      Picture         =   "frmTestAdvanced.frx":F8DB2
      Tag             =   "KnownCard9D"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image KnownCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   21
      Left            =   4995
      Picture         =   "frmTestAdvanced.frx":F99E4
      Tag             =   "KnownCard9H"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image KnownCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   34
      Left            =   4995
      Picture         =   "frmTestAdvanced.frx":FA6D0
      Tag             =   "KnownCard9S"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image KnownCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   9
      Left            =   4995
      Picture         =   "frmTestAdvanced.frx":FB3A0
      Tag             =   "KnownCard10C"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image KnownCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   48
      Left            =   4995
      Picture         =   "frmTestAdvanced.frx":FC2A5
      Tag             =   "KnownCard10D"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image KnownCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   22
      Left            =   4995
      Picture         =   "frmTestAdvanced.frx":FCF33
      Tag             =   "KnownCard10H"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image KnownCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   35
      Left            =   4995
      Picture         =   "frmTestAdvanced.frx":FDCA4
      Tag             =   "KnownCard10S"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image KnownCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   10
      Left            =   4995
      Picture         =   "frmTestAdvanced.frx":FEA16
      Tag             =   "KnownCardJC"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image KnownCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   49
      Left            =   4995
      Picture         =   "frmTestAdvanced.frx":FF95C
      Tag             =   "KnownCardJD"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image KnownCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   23
      Left            =   4995
      Picture         =   "frmTestAdvanced.frx":10099A
      Tag             =   "KnownCardJH"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image KnownCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   36
      Left            =   4995
      Picture         =   "frmTestAdvanced.frx":1019A4
      Tag             =   "KnownCardJS"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image KnownCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   11
      Left            =   4995
      Picture         =   "frmTestAdvanced.frx":10291E
      Tag             =   "KnownCardQC"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image KnownCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   50
      Left            =   4995
      Picture         =   "frmTestAdvanced.frx":103886
      Tag             =   "KnownCardQD"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image KnownCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   24
      Left            =   4995
      Picture         =   "frmTestAdvanced.frx":1048F4
      Tag             =   "KnownCardQH"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image KnownCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   37
      Left            =   4995
      Picture         =   "frmTestAdvanced.frx":105913
      Tag             =   "KnownCardQS"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image KnownCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   12
      Left            =   4995
      Picture         =   "frmTestAdvanced.frx":106912
      Tag             =   "KnownCardKC"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image KnownCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   51
      Left            =   4995
      Picture         =   "frmTestAdvanced.frx":10785E
      Tag             =   "KnownCardKD"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image KnownCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   25
      Left            =   4995
      Picture         =   "frmTestAdvanced.frx":10888A
      Tag             =   "KnownCardKH"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image KnownCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   38
      Left            =   4995
      Picture         =   "frmTestAdvanced.frx":1098DF
      Tag             =   "KnownCardKS"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image KnownCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   39
      Left            =   4995
      Picture         =   "frmTestAdvanced.frx":10A92D
      Tag             =   "KnownCardAD"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image KnownCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   13
      Left            =   4995
      Picture         =   "frmTestAdvanced.frx":10B12E
      Tag             =   "KnownCardAH"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image KnownCard 
      Appearance      =   0  'Flat
      Height          =   1920
      Index           =   26
      Left            =   4995
      Picture         =   "frmTestAdvanced.frx":10B927
      Tag             =   "KnownCardAS"
      Top             =   720
      Width           =   1350
   End
   Begin VB.Shape Shape2 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   1485
      Left            =   7335
      Top             =   2820
      Width           =   4080
   End
   Begin VB.Shape Shape3 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FF8080&
      FillStyle       =   0  'Solid
      Height          =   1605
      Left            =   7260
      Top             =   2760
      Width           =   4230
   End
End
Attribute VB_Name = "frmTestAdvanced"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub AnswerInput_KeyPress(KeyAscii As Integer)
If Not AdvTestingMode Then
    Exit Sub
End If
If AdvShowingMode = 1 Then
    TestingStatus.Caption = "Showing"
    If KeyAscii = 13 Then
        Guess.ZOrder
        TestNextCard
        KeyAscii = 0
        Exit Sub
    End If
ElseIf AdvShowingMode = 0 Then
    TestingStatus.Caption = "Testing"
    If KeyAscii = 13 Then
        If (AnswerInput.Text <> Empty And _
            (Not IsNumeric(AnswerInput.Text) Or _
            Val(AnswerInput.Text) < 0 Or _
            Val(AnswerInput.Text) > 52)) Or _
            AnswerInput.Text = Empty Then
            AnswerInputKeyEntry = 99
            '99 means that an error occurred, and to show the correct answer as an error
        Else
            AnswerInputKeyEntry = Val(AnswerInput.Text)
        End If
        TestShowCard
        KeyAscii = 0
        Exit Sub
    End If
End If
End Sub

Private Sub CardsToCutQuestion_Click()
AnswerLabel.Caption = "Cards to Cut?"
If AdvTestingMode Then
    AnswerInput.SetFocus
End If
If AdvShowingMode = 1 Then
    TestingStatus.Caption = "Showing"
    InCorrect(CardsToCut).ZOrder
End If
End Sub


Private Sub Form_Activate()
'initialize deck information

'initialize original deck order
For i% = 1 To DeckProperties
    For j% = 1 To DeckCount
        OriginalDeckOrder(i%, j%) = Deck(i%, j%)
    Next j%
Next i%

'set AdvDeckCurrent order
For i% = 1 To DeckProperties
    For j% = 1 To DeckCount
        AdvDeckCurrent(i%, j%) = Deck(i%, j%)
    Next j%
Next i%

'set AdvDeckOriginal order
For m% = 1 To 52
    For n% = 1 To 52
        If Val(Deck(1, n%)) = m% Then
            For p% = 1 To DeckProperties
                AdvDeckOriginal(p%, m%) = Deck(p%, n%)
            Next p%
        End If
    Next n%
Next m%
End Sub

Private Sub StartingStackRandomCut_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
DeckOrderStartingOption.Value = True
For i% = 1 To DeckProperties
    For j% = 1 To DeckCount
        Deck(i%, j%) = AdvDeckOriginal(i%, j%)
    Next j%
Next i%
Call frmDeck.DisplayCards
End Sub

Private Sub CurrentDeckRandomCut_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
DeckOrderCurrentOption.Value = True
For i% = 1 To DeckProperties
    For j% = 1 To DeckCount
        Deck(i%, j%) = AdvDeckCurrent(i%, j%)
    Next j%
Next i%
Call frmDeck.DisplayCards
End Sub

Private Sub DeckOrderCurrentOption_Click()
StartingStackRandomCut.Value = 0
For i% = 1 To DeckProperties
    For j% = 1 To DeckCount
        Deck(i%, j%) = AdvDeckCurrent(i%, j%)
    Next j%
Next i%
Call frmDeck.DisplayCards
End Sub

Private Sub DeckOrderStartingOption_Click()
CurrentDeckRandomCut.Value = 0
For i% = 1 To DeckProperties
    For j% = 1 To DeckCount
        Deck(i%, j%) = AdvDeckOriginal(i%, j%)
    Next j%
Next i%
Call frmDeck.DisplayCards
End Sub

Private Sub DesiredCardRandom_Click()
DesiredCardText = Empty
DesiredCard(0).ZOrder
End Sub


Private Sub DesiredCardText_GotFocus()
DesiredCardSpecified.Value = True
End Sub

Private Sub DesiredCardText_LostFocus()
If DesiredCardSpecified.Value = True And _
    DesiredCardText.Text <> Empty And _
    (Not IsNumeric(DesiredCardText.Text) Or _
    Val(DesiredCardText.Text) < 1 Or _
    Val(DesiredCardText.Text) > 52) Then
    DesiredCardText.Text = Empty
    MsgBox "Please enter a valid card position (1 to 52)" & Chr(13) _
        & "in the 'Specified Stack Value' Input Box"
    DesiredCardText.SetFocus
    Exit Sub
End If
If DesiredCardText.Text <> Empty Then
    For Each Ctrl In Controls
        If Left(Ctrl.Tag, 11) = "DesiredCard" Then
            If TestOriginalDeck(2, Val(DesiredCardText.Text)) = Right(Ctrl.Tag, Len(Ctrl.Tag) - 11) Then
                Ctrl.ZOrder
                Ctrl.Visible = True
            End If
        End If
    Next Ctrl
End If
End Sub

Private Sub DesiredPositionRandom_Click()
DesiredPosition(0).ZOrder
DesiredPositionText = Empty
End Sub



Private Sub DesiredPositionText_GotFocus()
DesiredPositionSpecified.Value = True
End Sub

Private Sub DesiredPositionText_LostFocus()
If DesiredPositionSpecified.Value = True And _
    DesiredPositionText.Text <> Empty And _
    (Not IsNumeric(DesiredPositionText.Text) Or _
    Val(DesiredPositionText.Text) < 1 Or _
    Val(DesiredPositionText.Text) > 52) Then
    DesiredPositionText.Text = Empty
    MsgBox "Please enter a valid card position (1 to 52)" & Chr(13) _
        & "in the 'Specified Position Value' Input Box"
    DesiredPositionText.SetFocus
    Exit Sub
End If
If DesiredPositionText.Text <> Empty Then
    DesiredPosition(Val(DesiredPositionText.Text)).ZOrder
End If
End Sub

Private Sub Form_Load()
'all settings at default values
AdvTestRange = 30
AdvTestCounter = 0
CorrectAnswerCount = 0
DeckOrderStartingOption.Value = False
DeckOrderCurrentOption.Value = True
KnownTopCard.Value = True
KnownBottomCard.Value = False
DesiredCardRandom.Value = True
DesiredPositionRandom.Value = True
CardsToCutQuestion.Value = True
AnswerInput.Text = Empty
StartingStackRandomCut.Value = 0
CurrentDeckRandomCut.Value = 0
DesiredCardText.Text = Empty
DesiredPositionText.Text = Empty
TestTimersEnabled.Value = 1

DesiredPosition(0).Visible = True
DesiredPosition(0).ZOrder
DesiredCard(0).Visible = True
DesiredCard(0).ZOrder
KnownCard(0).Visible = True
KnownCard(0).ZOrder


'set starting conditions
AdvTestingMode = False
TestingStatus.Caption = "Ready"
TestingStatus.ForeColor = &H8000&
'set color to GREEN
ProgressCardsRemaining.Visible = False
TestCardsRemaining.Visible = False
TestCardsLabel1.Visible = False
TestCardsLabel2.Visible = False
ProgressCardsRemaining.Value = AdvTestRange
TestCardsRemaining.Caption = AdvTestRange
DeckOrderFrame.Enabled = True
KnownCardFrame.Enabled = True
DesiredCardFrame.Enabled = True
DesiredPositionFrame.Enabled = True
DeckOrderStartingOption.Enabled = True
DeckOrderCurrentOption.Enabled = True
StartingStackRandomCut.Enabled = True
CurrentDeckRandomCut.Enabled = True
KnownTopCard.Enabled = True
KnownBottomCard.Enabled = True
DesiredCardRandom.Enabled = True
DesiredCardSpecified.Enabled = True
DesiredCardText.Enabled = True
DesiredPositionRandom.Enabled = True
DesiredPositionSpecified.Enabled = True
DesiredPositionText.Enabled = True
TestQuestionsFrame.Enabled = True
TestTimesFrame.Enabled = True
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

Private Sub KnownBottomCard_Click()
KnownCardLabel = "Bottom"
End Sub

Private Sub KnownTopCard_Click()
KnownCardLabel = "Top"
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmMain.mnuAdvancedTest.Checked = False
For i% = 1 To DeckProperties
    For j% = 1 To DeckCount
        Deck(i%, j%) = OriginalDeckOrder(i%, j%)
    Next j%
Next i%
Call frmDeck.DisplayCards
End Sub

Private Sub NewBottomCardQuestion_Click()
AnswerLabel.Caption = "New Bottom Card Stack Value?"
If AdvTestingMode Then
    AnswerInput.SetFocus
End If
If AdvShowingMode = 1 Then
    TestingStatus.Caption = "Showing"
    InCorrect(NewBottomCard).ZOrder
End If
End Sub

Private Sub NewTopCardQuestion_Click()
AnswerLabel.Caption = "New Top Card Stack Value?"
If AdvTestingMode Then
    AnswerInput.SetFocus
End If
If AdvShowingMode = 1 Then
    TestingStatus.Caption = "Showing"
    InCorrect(NewTopCard).ZOrder
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
If DesiredCardSpecified.Value = True And _
    DesiredCardText = Empty Then
    MsgBox ("When you select 'Specified Stack Value', you must enter a value" & Chr(13) _
        & "in the input box.")
    DesiredCardText.SetFocus
    Exit Sub
End If
If DeckOrderStartingOption.Value = False And _
    DeckOrderCurrentOption.Value = False Then
    MsgBox ("You must select a 'Deck Order'.")
    Exit Sub
End If
If KnownTopCard.Value = False And _
    KnownBottomCard.Value = False Then
    MsgBox ("You must select a 'Known Card'.")
    Exit Sub
End If
If DesiredCardSpecified.Value = True And _
    DesiredCardText.Text <> Empty And _
    (Not IsNumeric(DesiredCardText.Text) Or _
    Val(DesiredCardText.Text) < 1 Or _
    Val(DesiredCardText.Text) > 52) Then
    DesiredCardText.Text = Empty
    MsgBox "Please enter a valid card position (1 to 52)" & Chr(13) _
        & "in the 'Specified Stack Value' Input Box"
    DesiredCardText.SetFocus
    Exit Sub
End If
If DesiredPositionSpecified.Value = True And _
    DesiredPositionText = Empty Then
    MsgBox ("When you select 'Specified Position Value', you must enter a value" & Chr(13) _
        & "in the input box.")
    DesiredPositionText.SetFocus
    Exit Sub
End If
If DesiredPositionSpecified.Value = True And _
    DesiredPositionText.Text <> Empty And _
    (Not IsNumeric(DesiredPositionText.Text) Or _
    Val(DesiredPositionText.Text) < 1 Or _
    Val(DesiredPositionText.Text) > 52) Then
    DesiredPositionText.Text = Empty
    MsgBox "Please enter a valid card position (1 to 52)" & Chr(13) _
        & "in the 'Specified Position Value' Input Box"
    DesiredPositionText.SetFocus
    Exit Sub
End If
TimerTest.Enabled = False
AdvShowingMode = 0
TestingStatus.Caption = "Testing"
TimerShow.Enabled = False
ProgressTest.Value = 0
ProgressShow.Value = 0
AdvTestProgressIntervals = 0
AdvShowProgressIntervals = 0
If AdvTestingMode Then
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
Guess.ZOrder
AdvTestCounter = 0
CorrectAnswerCount = 0
ResultLabel.Visible = False
ResultTime.Visible = False
PercentCorrectLabel1.Visible = False
PercentCorrectLabel2.Visible = False
ScoreLabel1.Visible = False
ScoreLabel2.Visible = False
ExplanationLabel1.Visible = False
ExplanationLabel2.Visible = False
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
If AdvTestingMode Then
    ProgressCardsRemaining.Visible = False
    TestCardsRemaining.Visible = False
    TestCardsLabel1.Visible = False
    TestCardsLabel2.Visible = False
    ProgressCardsRemaining.Value = AdvTestRange
    TestCardsRemaining.Caption = AdvTestRange
    TestToggle(0).Visible = True
    TestToggle(1).Visible = False
    TestToggle(2).Visible = False
    TestToggle(3).Visible = False
    TestShow(0).Visible = False
    TestShow(1).Visible = True
    TestShow(2).Visible = False
    TestShow(3).Visible = False
Else
    ProgressCardsRemaining.Visible = True
    TestCardsRemaining.Visible = True
    TestCardsLabel1.Visible = True
    TestCardsLabel2.Visible = True
    ProgressCardsRemaining.Value = AdvTestRange
    TestCardsRemaining.Caption = AdvTestRange
    TestToggle(0).Visible = False
    TestToggle(1).Visible = False
    TestToggle(2).Visible = True
    TestToggle(3).Visible = False
End If
AdvTestingMode = Not AdvTestingMode
'Toggle logical value when button pressed
'initialize test sequences
For i% = 1 To AdvTestRange
    DesiredCardSequence(i%) = Int(Rnd * 52) + 1
    DesiredPositionSequence(i%) = Int(Rnd * 52) + 1
Next i%

If AdvTestingMode Then
    TestMessage(0).Visible = False
    TestingStatus.Caption = "Testing"
    TestingStatus.ForeColor = &HFF&
    'set color to RED
    DeckOrderFrame.Enabled = False
    KnownCardFrame.Enabled = False
    DesiredCardFrame.Enabled = False
    DesiredPositionFrame.Enabled = False
    DeckOrderStartingOption.Enabled = False
    DeckOrderCurrentOption.Enabled = False
    StartingStackRandomCut.Enabled = False
    CurrentDeckRandomCut.Enabled = False
    KnownTopCard.Enabled = False
    KnownBottomCard.Enabled = False
    DesiredCardRandom.Enabled = False
    DesiredCardSpecified.Enabled = False
    DesiredCardText.Enabled = False
    DesiredPositionRandom.Enabled = False
    DesiredPositionSpecified.Enabled = False
    DesiredPositionText.Enabled = False
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
Else
    TestMessage(0).Visible = True
    TestingStatus.Caption = "Ready"
    TestingStatus.ForeColor = &H8000&
    'set color to GREEN
    DeckOrderFrame.Enabled = True
    KnownCardFrame.Enabled = True
    DesiredCardFrame.Enabled = True
    DesiredPositionFrame.Enabled = True
    DeckOrderStartingOption.Enabled = True
    DeckOrderCurrentOption.Enabled = True
    StartingStackRandomCut.Enabled = True
    CurrentDeckRandomCut.Enabled = True
    KnownTopCard.Enabled = True
    KnownBottomCard.Enabled = True
    DesiredCardRandom.Enabled = True
    DesiredCardSpecified.Enabled = True
    DesiredCardText.Enabled = True
    DesiredPositionRandom.Enabled = True
    DesiredPositionSpecified.Enabled = True
    DesiredPositionText.Enabled = True
    KnownCard(0).ZOrder
    DesiredCard(0).ZOrder
    DesiredPosition(0).ZOrder
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
    'new section to return deck back after test is stopped
    For i% = 1 To DeckProperties
        For j% = 1 To DeckCount
            Deck(i%, j%) = OriginalDeckOrder(i%, j%)
        Next j%
    Next i%
    Call frmDeck.DisplayCards
    'end of new section
End If
End Sub


Private Sub TestNext_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If AdvTestingMode Then
    If AdvTestCounter = 0 Then
        TestNext(0).Visible = False
        TestNext(1).Visible = False
        TestNext(2).Visible = False
        TestNext(3).Visible = True
        TestNext(4).Visible = False
        TestNext(5).Visible = False
    ElseIf AdvTestCounter = AdvTestRange Then
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
    TimerTest.Enabled = False
    AdvShowingMode = 0
    TestingStatus.Caption = "Testing"
    TestShow(0).Visible = True
    TestShow(1).Visible = False
    TestShow(2).Visible = False
    TestShow(3).Visible = False
    TimerShow.Enabled = False
    ProgressTest.Value = 0
    ProgressShow.Value = 0
    AdvTestProgressIntervals = 0
    AdvShowProgressIntervals = 0
End If
End Sub

Private Sub TestNext_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If AdvTestingMode Then
    If AdvTestCounter = 0 Then
        AdvStartTime = Timer
    End If
    If AdvTestCounter = AdvTestRange - 1 Then
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

Private Sub TestShow_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If AdvTestingMode Then
    If AdvTestCounter = 0 Then
        TestShow(0).Visible = False
        TestShow(1).Visible = False
        TestShow(2).Visible = False
        TestShow(3).Visible = True
        TimerTest.Enabled = False
        AdvShowingMode = 0
        TestingStatus.Caption = "Testing"
        TimerShow.Enabled = False
        ProgressTest.Value = 0
        ProgressShow.Value = 0
        AdvTestProgressIntervals = 0
        AdvShowProgressIntervals = 0
    Else
        If AdvShowingMode = 1 Then
            TestingStatus.Caption = "Showing"
            TestShow(0).Visible = False
            TestShow(1).Visible = False
            TestShow(2).Visible = False
            TestShow(3).Visible = True
        Else
            TestingStatus.Caption = "Testing"
            TestShow(0).Visible = False
            TestShow(1).Visible = True
            TestShow(2).Visible = False
            TestShow(3).Visible = False
        End If
    End If
End If
End Sub

Private Sub TestShow_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If AdvTestingMode Then
    If AdvTestCounter = AdvTestRange Then
        'this turns on the "Finish" button for the last card
        TestNext(0).Visible = False
        TestNext(1).Visible = False
        TestNext(2).Visible = False
        TestNext(3).Visible = False
        TestNext(4).Visible = True
        TestNext(5).Visible = False
    ElseIf AdvTestCounter = 0 Then
        TestShow(0).Visible = True
        TestShow(1).Visible = False
        TestShow(2).Visible = False
        TestShow(3).Visible = False
        AdvStartTime = Timer
    Else
        TestNext(0).Visible = True
        TestNext(1).Visible = False
        TestNext(2).Visible = False
        TestNext(3).Visible = False
        TestNext(4).Visible = False
        TestNext(5).Visible = False
    End If
    If AdvTestingMode Then
        If AdvTestCounter = 0 Then
            TestShow(0).Visible = True
            TestShow(1).Visible = False
            TestShow(2).Visible = False
            TestShow(3).Visible = False
            TestNextCard
        ElseIf AdvShowingMode = 1 Then
            TestingStatus.Caption = "Showing"
            TestShow(0).Visible = True
            TestShow(1).Visible = False
            TestShow(2).Visible = False
            TestShow(3).Visible = False
            TestNextCard
        Else
            TestingStatus.Caption = "Testing"
            TestShow(0).Visible = False
            TestShow(1).Visible = False
            TestShow(2).Visible = True
            TestShow(3).Visible = False
            TestShowCard
        End If
    End If
End If
End Sub

Private Sub TimerShow_Timer()
AdvShowProgressIntervals = AdvShowProgressIntervals + 1
If ProgressShow.Value < ProgressShow.Max Then
    ProgressShow.Value = AdvShowProgressIntervals
Else
    AdvShowingMode = 0
    TestingStatus.Caption = "Testing"
    TestShow(0).Visible = True
    TestShow(1).Visible = False
    TestShow(2).Visible = False
    TestShow(3).Visible = False
    TimerShow.Enabled = False
    ProgressShow.Value = 0
    AdvShowProgressIntervals = 0
    If AdvTestCounter = AdvTestRange - 1 Then
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
AdvTestProgressIntervals = AdvTestProgressIntervals + 1
If ProgressTest.Value < ProgressTest.Max Then
    ProgressTest.Value = AdvTestProgressIntervals
Else
    'cummTimeIntervals = CummTimeIntervals + TestProgressIntervals
    TimerTest.Enabled = False
    ProgressTest.Value = 0
    AdvTestProgressIntervals = 0
    TestShowCard
End If
End Sub

Private Sub TestNextCard()
Guess.ZOrder
AnswerInput.Text = Empty
AnswerInput.SetFocus
TimerTest.Enabled = False
AdvShowingMode = 0
TestingStatus.Caption = "Testing"
TestShow(0).Visible = True
TestShow(1).Visible = False
TestShow(2).Visible = False
TestShow(3).Visible = False
TimerShow.Enabled = False
ProgressTest.Value = 0
ProgressShow.Value = 0
AdvTestProgressIntervals = 0
AdvShowProgressIntervals = 0
AdvTestCounter = AdvTestCounter + 1
TestCardsRemaining = AdvTestRange + 1 - AdvTestCounter
ProgressCardsRemaining.Value = Val(TestCardsRemaining)
If AdvTestCounter = AdvTestRange Then
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
If AdvTestCounter > AdvTestRange Then
    AdvTestCounter = 0
    AdvTestingMode = False
    TestingStatus.Caption = "Ready"
    TestingStatus.ForeColor = &H8000&
    'set color to GREEN
    ProgressCardsRemaining.Visible = False
    TestCardsRemaining.Visible = False
    TestCardsLabel1.Visible = False
    TestCardsLabel2.Visible = False
    ProgressCardsRemaining.Value = AdvTestRange
    TestCardsRemaining.Caption = AdvTestRange
    DeckOrderFrame.Enabled = True
    KnownCardFrame.Enabled = True
    DesiredCardFrame.Enabled = True
    DesiredPositionFrame.Enabled = True
    DeckOrderStartingOption.Enabled = True
    DeckOrderCurrentOption.Enabled = True
    StartingStackRandomCut.Enabled = True
    CurrentDeckRandomCut.Enabled = True
    KnownTopCard.Enabled = True
    KnownBottomCard.Enabled = True
    DesiredCardRandom.Enabled = True
    DesiredCardSpecified.Enabled = True
    DesiredCardText.Enabled = True
    DesiredPositionRandom.Enabled = True
    DesiredPositionSpecified.Enabled = True
    DesiredPositionText.Enabled = True
    TestToggle(0).Visible = True
    TestToggle(1).Visible = False
    TestToggle(2).Visible = False
    TestToggle(3).Visible = False
    TestMessage(0).Visible = True
    KnownCard(0).ZOrder
    DesiredCard(0).ZOrder
    DesiredPosition(0).ZOrder
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
    If Timer < AdvStartTime Then
        AdvStartTime = AdvStartTime - 86400
        'error condition is test crosses midnight
    End If
    AdvElapsedTime = Timer - AdvStartTime
    ResultTime.Caption = Int(AdvElapsedTime / 60) & " min. " & _
                Round((AdvElapsedTime Mod 60), 1) & " sec."
    PercentCorrectLabel2.Caption = Round((CorrectAnswerCount / 30) * 100, 1) & "%"
    ScoreLabel2.Caption = Round(100 * ((3 * CorrectAnswerCount) / AdvElapsedTime), 1)
    If TestTimersEnabled Then
        ResultLabel.Visible = True
        ResultTime.Visible = True
        PercentCorrectLabel1.Visible = True
        PercentCorrectLabel2.Visible = True
        ScoreLabel1.Visible = True
        ScoreLabel2.Visible = True
        ExplanationLabel1.Visible = True
        ExplanationLabel2.Visible = True
    End If
    'new section to return deck back after test is complete
    For i% = 1 To DeckProperties
        For j% = 1 To DeckCount
            Deck(i%, j%) = OriginalDeckOrder(i%, j%)
        Next j%
    Next i%
    Call frmDeck.DisplayCards
    'end of new section
    Exit Sub
End If
KnownCard(0).ZOrder
DesiredCard(0).ZOrder
DesiredPosition(0).ZOrder
'sets the lead cards to questions
'turn on Test Timer if checkbox is selected
If TestTimersEnabled.Value = 1 Then
    TimerTest.Enabled = True
End If
'------------------------
'Show correct Known Card
If DeckOrderStartingOption.Value = True Then
    If StartingStackRandomCut.Value = 1 Then
        Call frmStackView.CutDeckRandom("X")
    End If
    If KnownTopCard.Value = True Then
        For Each Ctrl In Controls
            If Ctrl.Tag = "KnownCard" & Deck(2, 1) Then
                Ctrl.ZOrder
            End If
        Next Ctrl
    ElseIf KnownBottomCard.Value = True Then
        For Each Ctrl In Controls
            If Ctrl.Tag = "KnownCard" & Deck(2, 52) Then
                Ctrl.ZOrder
            End If
        Next Ctrl
    End If
ElseIf DeckOrderCurrentOption.Value = True Then
    If CurrentDeckRandomCut.Value = 1 Then
        Call frmStackView.CutDeckRandom("X")
    End If
    If KnownTopCard.Value = True Then
        For Each Ctrl In Controls
            If Ctrl.Tag = "KnownCard" & Deck(2, 1) Then
                Ctrl.ZOrder
            End If
        Next Ctrl
    ElseIf KnownBottomCard.Value = True Then
        For Each Ctrl In Controls
            If Ctrl.Tag = "KnownCard" & Deck(2, 52) Then
                Ctrl.ZOrder
            End If
        Next Ctrl
    End If
End If

'Show correct Desired Card
If DesiredCardRandom.Value = True Then
    For i% = 1 To 52
        If Val(Deck(1, i%)) = DesiredCardSequence(AdvTestCounter) Then
            AdvDesiredCard = Deck(1, i%)
            AdvDesiredCardText = Deck(2, i%)
        End If
    Next i%
ElseIf DesiredCardSpecified.Value = True Then
    For i% = 1 To 52
        If Val(Deck(1, i%)) = DesiredCardText.Text Then
            AdvDesiredCard = Deck(1, i%)
            AdvDesiredCardText = Deck(2, i%)
        End If
    Next i%
End If
For Each Ctrl In Controls
    If Ctrl.Tag = "DesiredCard" & AdvDesiredCardText Then
        Ctrl.ZOrder
    End If
Next Ctrl

'Show correct Desired Position
If DesiredPositionRandom.Value = True Then
    AdvDesiredPosition = DesiredPositionSequence(AdvTestCounter)
    DesiredPosition(DesiredPositionSequence(AdvTestCounter)).ZOrder
ElseIf DesiredPositionSpecified.Value = True Then
    AdvDesiredPosition = DesiredPositionText.Text
    DesiredPosition(Val(DesiredPositionText.Text)).ZOrder
End If

'establish shifted desired car position for adjusted calculation
For i% = 1 To 52
    If Deck(1, i%) = AdvDesiredCard Then
        AdvDesiredCardShift = i%
    End If
Next i%
End Sub

Private Sub TestShowCard()
AnswerInput.Text = Empty
AnswerInput.SetFocus
TimerTest.Enabled = False
AdvShowingMode = 0
TestingStatus.Caption = "Testing"
TestShow(0).Visible = True
TestShow(1).Visible = False
TestShow(2).Visible = False
TestShow(3).Visible = False
TimerShow.Enabled = False
ProgressTest.Value = 0
ProgressShow.Value = 0
AdvTestProgressIntervals = 0
AdvShowProgressIntervals = 0
If AdvTestCounter > AdvTestRange Then
    AdvTestCounter = 0
    AdvTestingMode = False
    TestingStatus.Caption = "Ready"
    TestingStatus.ForeColor = &H8000&
    'set color to GREEN
    DeckOrderFrame.Enabled = True
    KnownCardFrame.Enabled = True
    DesiredCardFrame.Enabled = True
    DesiredPositionFrame.Enabled = True
    DeckOrderStartingOption.Enabled = True
    DeckOrderCurrentOption.Enabled = True
    StartingStackRandomCut.Enabled = True
    CurrentDeckRandomCut.Enabled = True
    KnownTopCard.Enabled = True
    KnownBottomCard.Enabled = True
    DesiredCardRandom.Enabled = True
    DesiredCardSpecified.Enabled = True
    DesiredCardText.Enabled = True
    DesiredPositionRandom.Enabled = True
    DesiredPositionSpecified.Enabled = True
    DesiredPositionText.Enabled = True
    TestToggle(0).Visible = True
    TestToggle(1).Visible = False
    TestToggle(2).Visible = False
    TestToggle(3).Visible = False
    TestMessage(0).Visible = True
    KnownCard(0).ZOrder
    DesiredCard(0).ZOrder
    DesiredPosition(0).ZOrder
    Exit Sub
End If

'turn on Test Timer if checkbox is selected
If TestTimersEnabled.Value = 1 Then
    TimerShow.Enabled = True
End If
'------------------------
AdvShowingMode = 1
TestingStatus.Caption = "Showing"
TestShow(0).Visible = False
TestShow(1).Visible = False
TestShow(2).Visible = True
TestShow(3).Visible = False

'calculate correct answers
CardsToCut = (52 + AdvDesiredCardShift - AdvDesiredPosition) Mod 52
NewTopCardShift = (52 + AdvDesiredCardShift - AdvDesiredPosition + 1) Mod 52
If NewTopCardShift = 0 Then
    NewTopCardShift = 52
End If
NewTopCard = Deck(1, NewTopCardShift)
NewBottomCardShift = NewTopCardShift - 1
If NewBottomCardShift = 0 Then
    NewBottomCardShift = 52
End If
NewBottomCard = Deck(1, NewBottomCardShift)

'compare correct answer to test result
If CardsToCutQuestion.Value = True Then
    If AnswerInputKeyEntry = CardsToCut Then
        Correct(CardsToCut).ZOrder
        CorrectAnswerCount = CorrectAnswerCount + 1
    Else
        InCorrect(CardsToCut).ZOrder
    End If
ElseIf NewTopCardQuestion.Value = True Then
    If AnswerInputKeyEntry = NewTopCard Then
        Correct(NewTopCard).ZOrder
        CorrectAnswerCount = CorrectAnswerCount + 1
    Else
        InCorrect(NewTopCard).ZOrder
    End If
ElseIf NewBottomCardQuestion.Value = True Then
    If AnswerInputKeyEntry = NewBottomCard Then
        Correct(NewBottomCard).ZOrder
        CorrectAnswerCount = CorrectAnswerCount + 1
    Else
        InCorrect(NewBottomCard).ZOrder
    End If
End If
End Sub


