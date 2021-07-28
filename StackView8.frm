VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmStackView 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "StackView Control"
   ClientHeight    =   7425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9795
   BeginProperty Font 
      Name            =   "Arial Narrow"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "StackView8.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   7425
   ScaleWidth      =   9795
   Visible         =   0   'False
   Begin TabDlg.SSTab StackViewDialog 
      Height          =   7095
      Left            =   120
      TabIndex        =   0
      Top             =   195
      Width           =   9570
      _ExtentX        =   16880
      _ExtentY        =   12515
      _Version        =   393216
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Control"
      TabPicture(0)   =   "StackView8.frx":57E2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "ResetDeckFrame"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "ConvShuffleFrame"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "FaroShuffleFrame"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Selection"
      TabPicture(1)   =   "StackView8.frx":57FE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "SelectionsLabel"
      Tab(1).Control(1)=   "Label4"
      Tab(1).Control(2)=   "Label18"
      Tab(1).Control(3)=   "Label20"
      Tab(1).Control(4)=   "FreeChoiceHandlingFrame"
      Tab(1).Control(5)=   "ForceFrame"
      Tab(1).Control(6)=   "FreeChoiceSpreadFrame"
      Tab(1).Control(7)=   "Frame11"
      Tab(1).Control(8)=   "HighlightSelectionsCheck"
      Tab(1).Control(9)=   "ClearSelections"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "ShowIndexValues"
      Tab(1).Control(11)=   "SelectionsTextBox"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "ShowPositionValues"
      Tab(1).Control(13)=   "SwapCardsFrame"
      Tab(1).Control(14)=   "ReverseSelections"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "ViewpointFrame"
      Tab(1).Control(16)=   "ClearReversedCards"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).ControlCount=   17
      TabCaption(2)   =   "Sessions"
      TabPicture(2)   =   "StackView8.frx":581A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "SessionInsertMacro"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "SessionRecursionList(9)"
      Tab(2).Control(2)=   "SessionRecursionList(8)"
      Tab(2).Control(3)=   "SessionRecursionList(7)"
      Tab(2).Control(4)=   "SessionRecursionList(6)"
      Tab(2).Control(5)=   "SessionRecursionList(10)"
      Tab(2).Control(6)=   "SessionRecursionList(5)"
      Tab(2).Control(7)=   "SessionRecursionList(4)"
      Tab(2).Control(8)=   "SessionRecursionList(3)"
      Tab(2).Control(9)=   "SessionRecursionList(2)"
      Tab(2).Control(10)=   "SessionRecursionList(1)"
      Tab(2).Control(11)=   "SessionRecordToggle(0)"
      Tab(2).Control(12)=   "SessionRecordToggle(3)"
      Tab(2).Control(13)=   "SessionRecordToggle(1)"
      Tab(2).Control(14)=   "SessionListBox"
      Tab(2).Control(15)=   "SessionPlayALL"
      Tab(2).Control(15).Enabled=   0   'False
      Tab(2).Control(16)=   "SessionClearALL"
      Tab(2).Control(16).Enabled=   0   'False
      Tab(2).Control(17)=   "SessionEventMoveUp"
      Tab(2).Control(17).Enabled=   0   'False
      Tab(2).Control(18)=   "SessionEventMoveDown"
      Tab(2).Control(18).Enabled=   0   'False
      Tab(2).Control(19)=   "SessionEventDelete"
      Tab(2).Control(19).Enabled=   0   'False
      Tab(2).Control(20)=   "SessionPlayEvent"
      Tab(2).Control(20).Enabled=   0   'False
      Tab(2).Control(21)=   "SessionRecordToggle(2)"
      Tab(2).Control(22)=   "Label17"
      Tab(2).Control(23)=   "SessionFileName"
      Tab(2).Control(24)=   "Label6"
      Tab(2).Control(25)=   "Label7"
      Tab(2).Control(26)=   "Label8"
      Tab(2).Control(27)=   "Label9"
      Tab(2).Control(28)=   "Label10"
      Tab(2).Control(29)=   "Label11"
      Tab(2).Control(30)=   "SessionRecordLabel"
      Tab(2).Control(31)=   "Line2"
      Tab(2).Control(32)=   "Line3"
      Tab(2).Control(33)=   "SessionRecordingStatus"
      Tab(2).ControlCount=   34
      Begin VB.CommandButton ClearReversedCards 
         Height          =   255
         Left            =   -68025
         TabIndex        =   211
         TabStop         =   0   'False
         Top             =   915
         Width           =   375
      End
      Begin VB.Frame ViewpointFrame 
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
         Left            =   -71265
         TabIndex        =   208
         Top             =   900
         Width           =   2955
         Begin VB.OptionButton ViewDeckAbove 
            Caption         =   "View Deck from ABOVE table"
            Height          =   300
            Left            =   75
            TabIndex        =   210
            Top             =   390
            Width           =   2625
         End
         Begin VB.OptionButton ViewDeckBeneath 
            Caption         =   "View Deck from BENEATH table (main)"
            Height          =   300
            Left            =   90
            TabIndex        =   209
            Top             =   150
            Value           =   -1  'True
            Width           =   2820
         End
      End
      Begin VB.CommandButton ReverseSelections 
         Height          =   255
         Left            =   -68025
         TabIndex        =   199
         TabStop         =   0   'False
         Top             =   1230
         Width           =   375
      End
      Begin VB.CommandButton SessionInsertMacro 
         DownPicture     =   "StackView8.frx":5836
         Height          =   315
         Left            =   -70110
         Picture         =   "StackView8.frx":5B1F
         Style           =   1  'Graphical
         TabIndex        =   194
         TabStop         =   0   'False
         Top             =   5205
         Width           =   345
      End
      Begin VB.ListBox SessionRecursionList 
         Height          =   510
         Index           =   9
         Left            =   -67455
         TabIndex        =   193
         Top             =   1920
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.ListBox SessionRecursionList 
         Height          =   510
         Index           =   8
         Left            =   -66165
         TabIndex        =   192
         Top             =   1320
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.ListBox SessionRecursionList 
         Height          =   510
         Index           =   7
         Left            =   -66600
         TabIndex        =   191
         Top             =   1305
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.ListBox SessionRecursionList 
         Height          =   510
         Index           =   6
         Left            =   -67020
         TabIndex        =   190
         Top             =   1350
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.ListBox SessionRecursionList 
         Columns         =   1
         Height          =   510
         Index           =   10
         Left            =   -67005
         TabIndex        =   189
         Top             =   1965
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.ListBox SessionRecursionList 
         Height          =   510
         Index           =   5
         Left            =   -67455
         TabIndex        =   188
         Top             =   1335
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.ListBox SessionRecursionList 
         Height          =   510
         Index           =   4
         Left            =   -66180
         TabIndex        =   187
         Top             =   705
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.ListBox SessionRecursionList 
         Height          =   510
         Index           =   3
         Left            =   -66600
         TabIndex        =   186
         Top             =   720
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.ListBox SessionRecursionList 
         Height          =   510
         Index           =   2
         Left            =   -67020
         TabIndex        =   185
         Top             =   720
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.ListBox SessionRecursionList 
         Height          =   510
         Index           =   1
         Left            =   -67455
         TabIndex        =   184
         Top             =   720
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Frame SwapCardsFrame 
         Caption         =   "Swap Cards"
         Height          =   5385
         Left            =   -68235
         TabIndex        =   182
         Top             =   1575
         Width           =   2550
         Begin VB.CheckBox SwapCardsNoSelection 
            Caption         =   "No Selections"
            Height          =   255
            Left            =   1335
            TabIndex        =   20
            Top             =   270
            Width           =   1155
         End
         Begin VB.CheckBox SwapSameSuitRandom 
            Caption         =   "Random"
            Height          =   255
            Left            =   465
            TabIndex        =   40
            Top             =   4530
            Width           =   1275
         End
         Begin VB.CheckBox SwapSameSuitClub 
            Caption         =   "Club"
            Height          =   255
            Left            =   465
            TabIndex        =   41
            Top             =   4740
            Width           =   630
         End
         Begin VB.CheckBox SwapSameSuitHeart 
            Caption         =   "Heart"
            Height          =   255
            Left            =   465
            TabIndex        =   42
            Top             =   4950
            Width           =   645
         End
         Begin VB.CheckBox SwapSameSuitSpade 
            Caption         =   "Spade"
            Height          =   255
            Left            =   1140
            TabIndex        =   43
            Top             =   4740
            Width           =   1275
         End
         Begin VB.CheckBox SwapSameSuitDiamond 
            Caption         =   "Diamond"
            Height          =   255
            Left            =   1140
            TabIndex        =   44
            Top             =   4950
            Width           =   1275
         End
         Begin VB.CheckBox SwapDifferentSuitsRandom 
            Caption         =   "Random"
            Height          =   255
            Left            =   465
            TabIndex        =   34
            Top             =   3555
            Width           =   1275
         End
         Begin VB.CheckBox SwapDifferentSuitsClub 
            Caption         =   "Club"
            Height          =   255
            Left            =   465
            TabIndex        =   35
            Top             =   3765
            Width           =   630
         End
         Begin VB.CheckBox SwapDifferentSuitsHeart 
            Caption         =   "Heart"
            Height          =   255
            Left            =   465
            TabIndex        =   36
            Top             =   3975
            Width           =   645
         End
         Begin VB.CheckBox SwapDifferentSuitsSpade 
            Caption         =   "Spade"
            Height          =   255
            Left            =   1140
            TabIndex        =   37
            Top             =   3765
            Width           =   1275
         End
         Begin VB.CheckBox SwapDifferentSuitsDiamond 
            Caption         =   "Diamond"
            Height          =   255
            Left            =   1140
            TabIndex        =   38
            Top             =   3975
            Width           =   1275
         End
         Begin VB.CheckBox SwapSameColorBlack 
            Caption         =   "Black"
            Height          =   255
            Left            =   465
            TabIndex        =   32
            Top             =   3030
            Width           =   1275
         End
         Begin VB.CheckBox SwapSameColorRed 
            Caption         =   "Red"
            Height          =   255
            Left            =   465
            TabIndex        =   31
            Top             =   2820
            Width           =   1275
         End
         Begin VB.CheckBox SwapSameColorRandom 
            Caption         =   "Random"
            Height          =   255
            Left            =   465
            TabIndex        =   30
            Top             =   2610
            Width           =   1275
         End
         Begin VB.OptionButton SwapSameColorOption 
            Caption         =   "Same Color"
            Height          =   255
            Left            =   165
            TabIndex        =   29
            Top             =   2400
            Width           =   1755
         End
         Begin VB.OptionButton SwapDifferentColorsOption 
            Caption         =   "Different Colors"
            Height          =   255
            Left            =   165
            TabIndex        =   28
            Top             =   2085
            Width           =   1755
         End
         Begin VB.OptionButton SwapDifferentSuitsOption 
            Caption         =   "Different Suits"
            Height          =   255
            Left            =   165
            TabIndex        =   33
            Top             =   3345
            Width           =   1755
         End
         Begin VB.OptionButton SwapSameSuitOption 
            Caption         =   "Same Suit"
            Height          =   255
            Left            =   165
            TabIndex        =   39
            Top             =   4320
            Width           =   1755
         End
         Begin VB.TextBox SwapPosition2 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1320
            TabIndex        =   27
            Top             =   1695
            Width           =   495
         End
         Begin VB.TextBox SwapPosition1 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   540
            TabIndex        =   26
            Top             =   1695
            Width           =   495
         End
         Begin VB.OptionButton SwapSpecifiedPositionsOption 
            Caption         =   "Specified Card Positions"
            Height          =   255
            Left            =   165
            TabIndex        =   25
            Top             =   1485
            Width           =   1890
         End
         Begin VB.TextBox SwapValue2 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1320
            TabIndex        =   24
            Top             =   1125
            Width           =   495
         End
         Begin VB.TextBox SwapValue1 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   540
            TabIndex        =   23
            Top             =   1125
            Width           =   495
         End
         Begin VB.OptionButton SwapSpecifiedCardsOption 
            Caption         =   "Specified Card Stack Values"
            Height          =   255
            Left            =   165
            TabIndex        =   22
            Top             =   915
            Width           =   2190
         End
         Begin VB.OptionButton SwapRandomOption 
            Caption         =   "Random"
            Height          =   255
            Left            =   165
            TabIndex        =   21
            Top             =   660
            Value           =   -1  'True
            Width           =   1755
         End
         Begin VB.CommandButton SwapCardsButton 
            Height          =   255
            Left            =   180
            Style           =   1  'Graphical
            TabIndex        =   183
            TabStop         =   0   'False
            Top             =   285
            Width           =   375
         End
      End
      Begin VB.PictureBox SessionRecordToggle 
         BorderStyle     =   0  'None
         Height          =   510
         Index           =   0
         Left            =   -70260
         Picture         =   "StackView8.frx":6296
         ScaleHeight     =   510
         ScaleWidth      =   1080
         TabIndex        =   181
         Top             =   1320
         Width           =   1080
      End
      Begin VB.PictureBox SessionRecordToggle 
         BorderStyle     =   0  'None
         Height          =   510
         Index           =   3
         Left            =   -70260
         Picture         =   "StackView8.frx":697C
         ScaleHeight     =   510
         ScaleWidth      =   1080
         TabIndex        =   180
         Top             =   1320
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.PictureBox SessionRecordToggle 
         BorderStyle     =   0  'None
         Height          =   510
         Index           =   1
         Left            =   -70260
         Picture         =   "StackView8.frx":6F86
         ScaleHeight     =   510
         ScaleWidth      =   1080
         TabIndex        =   179
         Top             =   1320
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.ListBox SessionListBox 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4380
         Left            =   -74775
         TabIndex        =   169
         Top             =   975
         Width           =   4290
      End
      Begin VB.CommandButton SessionPlayALL 
         DownPicture     =   "StackView8.frx":75B0
         Height          =   315
         Left            =   -70110
         Picture         =   "StackView8.frx":78C2
         Style           =   1  'Graphical
         TabIndex        =   168
         TabStop         =   0   'False
         Top             =   2100
         Width           =   345
      End
      Begin VB.CommandButton SessionClearALL 
         DownPicture     =   "StackView8.frx":7BE2
         Height          =   315
         Left            =   -70110
         Picture         =   "StackView8.frx":7F18
         Style           =   1  'Graphical
         TabIndex        =   167
         TabStop         =   0   'False
         Top             =   4320
         Width           =   345
      End
      Begin VB.CommandButton SessionEventMoveUp 
         DownPicture     =   "StackView8.frx":8269
         Height          =   315
         Left            =   -70110
         Picture         =   "StackView8.frx":857A
         Style           =   1  'Graphical
         TabIndex        =   166
         TabStop         =   0   'False
         Top             =   3210
         Width           =   345
      End
      Begin VB.CommandButton SessionEventMoveDown 
         DownPicture     =   "StackView8.frx":889D
         Height          =   315
         Left            =   -70110
         Picture         =   "StackView8.frx":8BA3
         Style           =   1  'Graphical
         TabIndex        =   165
         TabStop         =   0   'False
         Top             =   3660
         Width           =   345
      End
      Begin VB.CommandButton SessionEventDelete 
         DownPicture     =   "StackView8.frx":8EC1
         Height          =   315
         Left            =   -70110
         Picture         =   "StackView8.frx":91AA
         Style           =   1  'Graphical
         TabIndex        =   164
         TabStop         =   0   'False
         Top             =   4740
         Width           =   345
      End
      Begin VB.CommandButton SessionPlayEvent 
         DownPicture     =   "StackView8.frx":94C5
         Height          =   315
         Left            =   -70110
         Picture         =   "StackView8.frx":97F5
         Style           =   1  'Graphical
         TabIndex        =   163
         TabStop         =   0   'False
         Top             =   2535
         Width           =   345
      End
      Begin VB.PictureBox SessionRecordToggle 
         BorderStyle     =   0  'None
         Height          =   510
         Index           =   2
         Left            =   -70260
         Picture         =   "StackView8.frx":9B3E
         ScaleHeight     =   510
         ScaleWidth      =   1080
         TabIndex        =   162
         Top             =   1320
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.CheckBox ShowPositionValues 
         Caption         =   "Show Position Values"
         Height          =   225
         Left            =   -74610
         TabIndex        =   3
         Top             =   900
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.TextBox SelectionsTextBox 
         Height          =   330
         Left            =   -72000
         Locked          =   -1  'True
         TabIndex        =   159
         TabStop         =   0   'False
         Top             =   570
         Width           =   3855
      End
      Begin VB.CheckBox ShowIndexValues 
         Caption         =   "Show Stack Values"
         Height          =   225
         Left            =   -74610
         TabIndex        =   2
         Top             =   660
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.CommandButton ClearSelections 
         Height          =   255
         Left            =   -68025
         TabIndex        =   158
         TabStop         =   0   'False
         Top             =   600
         Width           =   375
      End
      Begin VB.CheckBox HighlightSelectionsCheck 
         Caption         =   "Highlight Selections"
         Height          =   255
         Left            =   -74610
         TabIndex        =   1
         Top             =   420
         Value           =   1  'Checked
         Width           =   1545
      End
      Begin VB.Frame Frame11 
         Height          =   510
         Left            =   -74325
         TabIndex        =   157
         Top             =   1035
         Width           =   2895
         Begin VB.OptionButton CountFromBack 
            Caption         =   "Count from back"
            Height          =   240
            Left            =   105
            TabIndex        =   4
            Top             =   180
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.OptionButton CountFromFace 
            Caption         =   "Count from face"
            Height          =   240
            Left            =   1515
            TabIndex        =   5
            Top             =   180
            Width           =   1335
         End
      End
      Begin VB.Frame FreeChoiceSpreadFrame 
         Caption         =   "Free Choice from Spread"
         Height          =   1905
         Left            =   -74715
         TabIndex        =   151
         Top             =   2940
         Width           =   6435
         Begin VB.CheckBox FreeChoiceReverseCheck 
            Caption         =   "Reverse Selected Card"
            Height          =   255
            Left            =   2865
            TabIndex        =   198
            Top             =   1575
            Width           =   2070
         End
         Begin VB.ComboBox FreeChoiceSelectSpecificCombo 
            Height          =   345
            ItemData        =   "StackView8.frx":A1EC
            Left            =   900
            List            =   "StackView8.frx":A1F9
            TabIndex        =   12
            Top             =   975
            Width           =   1815
         End
         Begin VB.CommandButton FreeChoiceSpreadButton 
            Height          =   255
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   154
            TabStop         =   0   'False
            Top             =   480
            Width           =   375
         End
         Begin VB.Frame FreeChoiceSpreadSelectFrame 
            BorderStyle     =   0  'None
            Caption         =   "Frame24"
            ClipControls    =   0   'False
            Height          =   480
            Left            =   600
            TabIndex        =   153
            Top             =   480
            Width           =   2040
            Begin VB.OptionButton FreeChoiceSpreadSelectAnyCardOption 
               Caption         =   "Any Card"
               Height          =   255
               Left            =   120
               TabIndex        =   10
               Top             =   0
               Value           =   -1  'True
               Width           =   1635
            End
            Begin VB.OptionButton FreeChoiceSpreadSelectSpecificOption 
               Caption         =   "Specific Section"
               Height          =   255
               Left            =   120
               TabIndex        =   11
               Top             =   240
               Width           =   1695
            End
         End
         Begin VB.Frame FreeChoiceSpreadReturnFrame 
            BorderStyle     =   0  'None
            Caption         =   "Frame25"
            Height          =   975
            Left            =   2610
            TabIndex        =   152
            Top             =   480
            Width           =   1935
            Begin VB.OptionButton FreeChoiceSpreadReturnOriginalOption 
               Caption         =   "Original Position"
               Height          =   255
               Left            =   240
               TabIndex        =   13
               Top             =   0
               Width           =   1575
            End
            Begin VB.OptionButton FreeChoiceSpreadReturnAnywhereOption 
               Caption         =   "Anywhere"
               Height          =   255
               Left            =   240
               TabIndex        =   14
               Top             =   240
               Value           =   -1  'True
               Width           =   1455
            End
            Begin VB.OptionButton FreeChoiceSpreadReturnPositionOption 
               Caption         =   "Controlled Position"
               Height          =   255
               Left            =   240
               TabIndex        =   15
               Top             =   480
               Width           =   1695
            End
            Begin VB.OptionButton FreeChoiceSpreadReturnSectionOption 
               Caption         =   "Specific Section"
               Height          =   255
               Left            =   240
               TabIndex        =   17
               Top             =   720
               Width           =   1575
            End
         End
         Begin VB.TextBox FreeChoiceSpreadReturnPositionTextBox 
            Alignment       =   2  'Center
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
            Height          =   285
            Left            =   4530
            TabIndex        =   16
            Top             =   930
            Width           =   495
         End
         Begin VB.ComboBox FreeChoiceSpreadReturnSectionCombo 
            Height          =   345
            ItemData        =   "StackView8.frx":A224
            Left            =   4530
            List            =   "StackView8.frx":A231
            TabIndex        =   18
            Top             =   1200
            Width           =   1815
         End
         Begin VB.Label FreeChoiceSpreadSelectLabel 
            Caption         =   "Select"
            Height          =   255
            Left            =   720
            TabIndex        =   156
            Top             =   240
            Width           =   975
         End
         Begin VB.Label FreeChoiceSpreadReturnLabel 
            Caption         =   "Return"
            Height          =   255
            Left            =   2850
            TabIndex        =   155
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame ForceFrame 
         Caption         =   "Force"
         Height          =   1365
         Left            =   -74715
         TabIndex        =   145
         Top             =   1575
         Width           =   6420
         Begin VB.CheckBox ForceReverseCheck 
            Caption         =   "Reverse Forced Card"
            Height          =   255
            Left            =   2880
            TabIndex        =   196
            Top             =   1035
            Width           =   2070
         End
         Begin VB.TextBox ForceReturnPositionTextBox 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   4635
            TabIndex        =   9
            Top             =   735
            Width           =   495
         End
         Begin VB.CommandButton ForceButton 
            Height          =   255
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   147
            TabStop         =   0   'False
            Top             =   480
            Width           =   375
         End
         Begin VB.Frame ForceReturnFrame 
            BorderStyle     =   0  'None
            Caption         =   "Frame24"
            ClipControls    =   0   'False
            Height          =   495
            Left            =   2760
            TabIndex        =   146
            Top             =   495
            Width           =   1830
            Begin VB.OptionButton ForceReturnPositionOption 
               Caption         =   "Controlled Position"
               Height          =   255
               Left            =   120
               TabIndex        =   8
               Top             =   240
               Width           =   1575
            End
            Begin VB.OptionButton ForceReturnAnywhereOption 
               Caption         =   "Anywhere"
               Height          =   255
               Left            =   120
               TabIndex        =   7
               Top             =   0
               Value           =   -1  'True
               Width           =   1455
            End
         End
         Begin VB.TextBox ForcePositionTextBox 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1440
            TabIndex        =   6
            Top             =   480
            Width           =   495
         End
         Begin VB.Label ForceReturnLabel 
            Caption         =   "Return"
            Height          =   255
            Left            =   2880
            TabIndex        =   150
            Top             =   255
            Width           =   975
         End
         Begin VB.Label ForceCardLabel 
            Caption         =   "Force Card"
            Height          =   255
            Left            =   720
            TabIndex        =   149
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label ForcePositionLabel 
            Caption         =   "Position"
            Height          =   255
            Left            =   720
            TabIndex        =   148
            Top             =   480
            Width           =   735
         End
      End
      Begin VB.Frame FreeChoiceHandlingFrame 
         Caption         =   "Free Choice with Special Handling"
         Height          =   2130
         Left            =   -74730
         TabIndex        =   133
         Top             =   4860
         Width           =   6435
         Begin VB.CheckBox SpecialReverseThirdCheck 
            Caption         =   "Third Card"
            Height          =   255
            Left            =   4905
            TabIndex        =   202
            Top             =   1185
            Width           =   1245
         End
         Begin VB.CheckBox SpecialReverseSecondCheck 
            Caption         =   "Second Card"
            Height          =   255
            Left            =   4905
            TabIndex        =   201
            Top             =   885
            Width           =   1245
         End
         Begin VB.CheckBox SpecialReverseFirstCheck 
            Caption         =   "First Card"
            Height          =   255
            Left            =   4905
            TabIndex        =   197
            Top             =   585
            Width           =   1245
         End
         Begin VB.PictureBox SpecialHandlingPicture 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1140
            Index           =   8
            Left            =   690
            Picture         =   "StackView8.frx":A25C
            ScaleHeight     =   1140
            ScaleWidth      =   4020
            TabIndex        =   134
            TabStop         =   0   'False
            Top             =   885
            Width           =   4020
         End
         Begin VB.CommandButton FreeChoiceHandlingButton 
            Height          =   255
            Left            =   165
            Style           =   1  'Graphical
            TabIndex        =   135
            TabStop         =   0   'False
            Top             =   405
            Width           =   375
         End
         Begin VB.ComboBox FreeChoiceHandlingSelectCombo 
            Height          =   345
            ItemData        =   "StackView8.frx":A882
            Left            =   1275
            List            =   "StackView8.frx":A89E
            TabIndex        =   19
            Text            =   "Select special selection handling"
            Top             =   375
            Width           =   2535
         End
         Begin VB.PictureBox SpecialHandlingPicture 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1140
            Index           =   7
            Left            =   690
            Picture         =   "StackView8.frx":A95A
            ScaleHeight     =   1140
            ScaleWidth      =   4020
            TabIndex        =   143
            Top             =   885
            Width           =   4020
         End
         Begin VB.PictureBox SpecialHandlingPicture 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1140
            Index           =   6
            Left            =   690
            Picture         =   "StackView8.frx":BC68
            ScaleHeight     =   1140
            ScaleWidth      =   4020
            TabIndex        =   142
            Top             =   885
            Width           =   4020
         End
         Begin VB.PictureBox SpecialHandlingPicture 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1140
            Index           =   0
            Left            =   690
            Picture         =   "StackView8.frx":CD31
            ScaleHeight     =   1140
            ScaleWidth      =   4020
            TabIndex        =   141
            Top             =   885
            Width           =   4020
         End
         Begin VB.PictureBox SpecialHandlingPicture 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1140
            Index           =   5
            Left            =   690
            Picture         =   "StackView8.frx":DA0A
            ScaleHeight     =   1140
            ScaleWidth      =   4020
            TabIndex        =   140
            Top             =   885
            Width           =   4020
         End
         Begin VB.PictureBox SpecialHandlingPicture 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1140
            Index           =   4
            Left            =   690
            Picture         =   "StackView8.frx":ED64
            ScaleHeight     =   1140
            ScaleWidth      =   4020
            TabIndex        =   139
            Top             =   885
            Width           =   4020
         End
         Begin VB.PictureBox SpecialHandlingPicture 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1140
            Index           =   3
            Left            =   690
            Picture         =   "StackView8.frx":FCF6
            ScaleHeight     =   1140
            ScaleWidth      =   4020
            TabIndex        =   138
            Top             =   885
            Width           =   4020
         End
         Begin VB.PictureBox SpecialHandlingPicture 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1140
            Index           =   2
            Left            =   690
            Picture         =   "StackView8.frx":10875
            ScaleHeight     =   1140
            ScaleWidth      =   4020
            TabIndex        =   137
            Top             =   885
            Width           =   4020
         End
         Begin VB.PictureBox SpecialHandlingPicture 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1140
            Index           =   1
            Left            =   690
            Picture         =   "StackView8.frx":11ABF
            ScaleHeight     =   1140
            ScaleWidth      =   4020
            TabIndex        =   136
            Top             =   885
            Width           =   4020
         End
         Begin VB.Label Label19 
            Caption         =   "Reverse Selected Card(s)"
            Height          =   225
            Left            =   4575
            TabIndex        =   203
            Top             =   300
            Width           =   1695
         End
         Begin VB.Shape Shape1 
            BorderWidth     =   3
            Height          =   1200
            Left            =   660
            Top             =   855
            Width           =   4080
         End
         Begin VB.Label FreeChoiceHandlingSelectLabel 
            Caption         =   "Select"
            Height          =   255
            Left            =   750
            TabIndex        =   144
            Top             =   420
            Width           =   975
         End
      End
      Begin VB.Frame FaroShuffleFrame 
         Caption         =   "Faro Shuffles"
         Height          =   1290
         Left            =   120
         TabIndex        =   111
         Top             =   5085
         Width           =   8985
         Begin VB.CheckBox OutFaroReverseCheck 
            Caption         =   "Reverse Top Block"
            Height          =   240
            Left            =   7260
            TabIndex        =   87
            Top             =   795
            Width           =   1530
         End
         Begin VB.CheckBox InFaroReverseCheck 
            Caption         =   "Reverse Top Block"
            Height          =   240
            Left            =   7260
            TabIndex        =   79
            Top             =   435
            Width           =   1530
         End
         Begin VB.CheckBox OutFaroInverseCheck 
            Caption         =   "Inverse"
            Height          =   240
            Left            =   6255
            TabIndex        =   86
            Top             =   795
            Width           =   885
         End
         Begin VB.CheckBox InFaroInverseCheck 
            Caption         =   "Inverse"
            Height          =   240
            Left            =   6255
            TabIndex        =   78
            Top             =   450
            Width           =   885
         End
         Begin VB.TextBox OutFaroInteriorTextBox 
            Alignment       =   2  'Center
            DataField       =   "22"
            Height          =   285
            Left            =   5535
            TabIndex        =   85
            Top             =   780
            Width           =   495
         End
         Begin VB.TextBox InFaroInteriorTextBox 
            Alignment       =   2  'Center
            DataField       =   "22"
            Height          =   285
            Left            =   5535
            TabIndex        =   77
            Top             =   420
            Width           =   495
         End
         Begin VB.CommandButton InFaroButton 
            Height          =   255
            Left            =   135
            Style           =   1  'Graphical
            TabIndex        =   98
            TabStop         =   0   'False
            Top             =   420
            Width           =   375
         End
         Begin VB.TextBox InFaroFromTopTextBox 
            Alignment       =   2  'Center
            DataField       =   "22"
            Height          =   285
            Left            =   3390
            TabIndex        =   74
            Top             =   420
            Width           =   495
         End
         Begin VB.Frame InFaroStartWeaveFrame 
            BorderStyle     =   0  'None
            Height          =   300
            Left            =   4035
            TabIndex        =   115
            Top             =   480
            Width           =   1395
            Begin VB.OptionButton InFaroStartWeaveTopOption 
               Caption         =   "Top"
               Height          =   195
               Left            =   0
               TabIndex        =   75
               Top             =   15
               Width           =   585
            End
            Begin VB.OptionButton InFaroStartWeaveBottomOption 
               Caption         =   "Bottom"
               Height          =   195
               Left            =   630
               TabIndex        =   76
               Top             =   15
               Width           =   1095
            End
         End
         Begin VB.CommandButton OutFaroButton 
            Height          =   255
            Left            =   135
            Style           =   1  'Graphical
            TabIndex        =   99
            TabStop         =   0   'False
            Top             =   780
            Width           =   375
         End
         Begin VB.TextBox OutFaroFromTopTextBox 
            Alignment       =   2  'Center
            DataField       =   "22"
            Height          =   285
            Left            =   3390
            TabIndex        =   82
            Top             =   780
            Width           =   495
         End
         Begin VB.Frame OutFaroStartWeaveFrame 
            BorderStyle     =   0  'None
            Height          =   390
            Left            =   4035
            TabIndex        =   114
            Top             =   705
            Width           =   1485
            Begin VB.OptionButton OutFaroStartWeaveBottomOption 
               Caption         =   "Bottom"
               Height          =   255
               Left            =   630
               TabIndex        =   84
               Top             =   105
               Width           =   1095
            End
            Begin VB.OptionButton OutFaroStartWeaveTopOption 
               Caption         =   "Top"
               Height          =   255
               Left            =   0
               TabIndex        =   83
               Top             =   120
               Width           =   570
            End
         End
         Begin VB.Frame InFaroTypeFrame 
            BorderStyle     =   0  'None
            Height          =   270
            Left            =   1260
            TabIndex        =   113
            Top             =   450
            Width           =   2295
            Begin VB.OptionButton InFaroStandardOption 
               Caption         =   "Standard"
               Height          =   255
               Left            =   105
               TabIndex        =   72
               Top             =   -30
               Value           =   -1  'True
               Width           =   975
            End
            Begin VB.OptionButton InFaroSpecialOption 
               Caption         =   "Special"
               Height          =   255
               Left            =   1200
               TabIndex        =   73
               Top             =   -30
               Width           =   1095
            End
         End
         Begin VB.Frame OutFaroTypeFrame 
            BorderStyle     =   0  'None
            Height          =   405
            Left            =   1260
            TabIndex        =   112
            Top             =   660
            Width           =   2295
            Begin VB.OptionButton OutFaroSpecialOption 
               Caption         =   "Special"
               Height          =   255
               Left            =   1200
               TabIndex        =   81
               Top             =   120
               Width           =   1095
            End
            Begin VB.OptionButton OutFaroStandardOption 
               Caption         =   "Standard"
               Height          =   240
               Left            =   105
               TabIndex        =   80
               Top             =   105
               Value           =   -1  'True
               Width           =   975
            End
         End
         Begin VB.Label Label2 
            Caption         =   "Interior Position"
            Height          =   255
            Left            =   5310
            TabIndex        =   127
            Top             =   135
            Width           =   1215
         End
         Begin VB.Label FaroShuffleStartWeaveLabel 
            Caption         =   "Start Weave"
            Height          =   255
            Left            =   4245
            TabIndex        =   119
            Top             =   135
            Width           =   1215
         End
         Begin VB.Label FaroShuffleFromTopLabel 
            Alignment       =   2  'Center
            Caption         =   "From Top"
            Height          =   255
            Left            =   3270
            TabIndex        =   118
            Top             =   135
            Width           =   735
         End
         Begin VB.Label InFaroLabel 
            Caption         =   "In Faro"
            Height          =   255
            Left            =   600
            TabIndex        =   117
            Top             =   420
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "Out Faro"
            Height          =   255
            Left            =   600
            TabIndex        =   116
            Top             =   780
            Width           =   1335
         End
      End
      Begin VB.Frame ConvShuffleFrame 
         Caption         =   "Conventional Shuffles"
         Height          =   3420
         Left            =   135
         TabIndex        =   102
         Top             =   1605
         Width           =   8970
         Begin VB.CheckBox RiffleReverseBottomCheck 
            Caption         =   "Reverse Bottom Block"
            Height          =   240
            Left            =   6720
            TabIndex        =   206
            Top             =   765
            Width           =   1755
         End
         Begin VB.CheckBox RiffleReverseTopCheck 
            Caption         =   "Reverse Top Block"
            Height          =   240
            Left            =   6720
            TabIndex        =   204
            Top             =   555
            Width           =   1530
         End
         Begin VB.CheckBox RunSingleCardsReverseCheck 
            Caption         =   "Reverse Run Cards"
            Height          =   240
            Left            =   4095
            TabIndex        =   63
            Top             =   1815
            Width           =   1890
         End
         Begin VB.CheckBox ShiftTopBlockReverseCheck 
            Caption         =   "Reverse Top Block"
            Height          =   240
            Left            =   4095
            TabIndex        =   67
            Top             =   2385
            Width           =   1530
         End
         Begin VB.CheckBox MoveCardInverseCheck 
            Caption         =   "Inverse"
            Height          =   240
            Left            =   3195
            TabIndex        =   70
            Top             =   3030
            Width           =   900
         End
         Begin VB.CheckBox MoveCardReverseCheck 
            Caption         =   "Reverse Card"
            Height          =   240
            Left            =   4095
            TabIndex        =   71
            Top             =   3045
            Width           =   1290
         End
         Begin VB.TextBox MoveCardToTextBox 
            Alignment       =   2  'Center
            DataField       =   "22"
            Height          =   285
            Left            =   2580
            TabIndex        =   69
            Top             =   3000
            Width           =   495
         End
         Begin VB.TextBox MoveCardFromTextBox 
            Alignment       =   2  'Center
            DataField       =   "22"
            Height          =   285
            Left            =   1905
            TabIndex        =   68
            Top             =   3000
            Width           =   495
         End
         Begin VB.CommandButton MoveCardButton 
            Height          =   255
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   97
            TabStop         =   0   'False
            Top             =   3000
            Width           =   375
         End
         Begin VB.CheckBox ShiftTopBlockInverseCheck 
            Caption         =   "Inverse"
            Height          =   240
            Left            =   3195
            TabIndex        =   66
            Top             =   2385
            Width           =   900
         End
         Begin VB.TextBox ShiftTopDepthTextBox 
            Alignment       =   2  'Center
            DataField       =   "22"
            Height          =   285
            Left            =   2580
            TabIndex        =   65
            Top             =   2370
            Width           =   495
         End
         Begin VB.TextBox ShiftTopBlockTextBox 
            Alignment       =   2  'Center
            DataField       =   "22"
            Height          =   285
            Left            =   1905
            TabIndex        =   64
            Top             =   2385
            Width           =   495
         End
         Begin VB.CommandButton ShiftTopBlockButton 
            Height          =   255
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   96
            TabStop         =   0   'False
            Top             =   2400
            Width           =   375
         End
         Begin VB.CommandButton OverhandShuffleButton 
            Height          =   255
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   93
            TabStop         =   0   'False
            Top             =   960
            Width           =   375
         End
         Begin VB.CheckBox RunSingleInverseCheck 
            Caption         =   "Inverse"
            Height          =   240
            Left            =   3195
            TabIndex        =   62
            Top             =   1815
            Width           =   915
         End
         Begin VB.TextBox RunSingleTextBox 
            Alignment       =   2  'Center
            DataField       =   "22"
            Height          =   285
            Left            =   1905
            TabIndex        =   61
            Top             =   1800
            Width           =   495
         End
         Begin VB.CommandButton RunSingleButton 
            Height          =   255
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   95
            TabStop         =   0   'False
            Top             =   1845
            Width           =   375
         End
         Begin VB.TextBox OverhandNumberTextBox 
            Alignment       =   2  'Center
            DataField       =   "22"
            Height          =   345
            Left            =   4395
            TabIndex        =   54
            Top             =   945
            Width           =   495
         End
         Begin VB.TextBox CutPreciseTextBox 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   3360
            TabIndex        =   58
            Top             =   1320
            Width           =   495
         End
         Begin VB.ComboBox CutSpecialCombo 
            Height          =   345
            ItemData        =   "StackView8.frx":12A47
            Left            =   4995
            List            =   "StackView8.frx":12A60
            TabIndex        =   60
            Top             =   1320
            Width           =   1575
         End
         Begin VB.ComboBox OverhandLocationCombo 
            Height          =   345
            ItemData        =   "StackView8.frx":12AA5
            Left            =   4980
            List            =   "StackView8.frx":12AAF
            TabIndex        =   55
            Top             =   945
            Width           =   1335
         End
         Begin VB.ComboBox RiffleLocationCombo 
            Height          =   345
            ItemData        =   "StackView8.frx":12AC0
            Left            =   4980
            List            =   "StackView8.frx":12ACA
            TabIndex        =   51
            Top             =   585
            Width           =   1335
         End
         Begin VB.Frame CutTypeFrame 
            BorderStyle     =   0  'None
            Height          =   570
            Left            =   1440
            TabIndex        =   105
            Top             =   1200
            Width           =   7380
            Begin VB.CheckBox CutReverseBottomCheck 
               Caption         =   "Reverse Bottom Block"
               Height          =   240
               Left            =   5280
               TabIndex        =   207
               Top             =   255
               Width           =   1860
            End
            Begin VB.CheckBox CutReverseTopCheck 
               Caption         =   "Reverse Top Block"
               Height          =   240
               Left            =   5280
               TabIndex        =   205
               Top             =   45
               Width           =   1530
            End
            Begin VB.OptionButton CutPreciseOption 
               Caption         =   "Precise"
               Height          =   255
               Left            =   1080
               TabIndex        =   57
               Top             =   120
               Width           =   855
            End
            Begin VB.OptionButton CutRandomOption 
               Caption         =   "Random"
               Height          =   255
               Left            =   120
               TabIndex        =   56
               Top             =   120
               Value           =   -1  'True
               Width           =   975
            End
            Begin VB.OptionButton CutSpecialOption 
               Caption         =   "Special Random"
               Height          =   375
               Left            =   2715
               TabIndex        =   59
               Top             =   165
               Width           =   1155
            End
         End
         Begin VB.Frame OverhandTypeFrame 
            BorderStyle     =   0  'None
            Height          =   375
            Left            =   1440
            TabIndex        =   104
            Top             =   840
            Width           =   2895
            Begin VB.OptionButton OverhandRandomOption 
               Caption         =   "Random"
               Height          =   255
               Left            =   120
               TabIndex        =   52
               Top             =   120
               Value           =   -1  'True
               Width           =   975
            End
            Begin VB.OptionButton OverhandProtectOption 
               Caption         =   "Protect Block"
               Height          =   255
               Left            =   1080
               TabIndex        =   53
               Top             =   120
               Width           =   1575
            End
         End
         Begin VB.Frame RiffleTypeFrame 
            BorderStyle     =   0  'None
            Height          =   375
            Left            =   1440
            TabIndex        =   103
            Top             =   480
            Width           =   2895
            Begin VB.OptionButton RiffleProtectOption 
               Caption         =   "Protect Block"
               Height          =   255
               Left            =   1080
               TabIndex        =   49
               Top             =   120
               Width           =   1575
            End
            Begin VB.OptionButton RiffleRandomOption 
               Caption         =   "Random"
               Height          =   255
               Left            =   120
               TabIndex        =   48
               Top             =   105
               Value           =   -1  'True
               Width           =   975
            End
         End
         Begin VB.CommandButton CutShuffleButton 
            Height          =   255
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   94
            TabStop         =   0   'False
            Top             =   1320
            Width           =   375
         End
         Begin VB.TextBox RiffleNumberTextBox 
            Alignment       =   2  'Center
            DataField       =   "22"
            Height          =   345
            Left            =   4395
            TabIndex        =   50
            Top             =   585
            Width           =   495
         End
         Begin VB.CommandButton RiffleShuffleButton 
            Height          =   255
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   92
            TabStop         =   0   'False
            Top             =   615
            Width           =   375
         End
         Begin VB.Label Label16 
            Alignment       =   2  'Center
            Caption         =   "To"
            Height          =   255
            Left            =   2655
            TabIndex        =   130
            Top             =   2745
            Width           =   345
         End
         Begin VB.Label Label15 
            Alignment       =   2  'Center
            Caption         =   "From"
            Height          =   255
            Left            =   1935
            TabIndex        =   129
            Top             =   2745
            Width           =   495
         End
         Begin VB.Label Label3 
            Caption         =   "Move Card"
            Height          =   255
            Left            =   720
            TabIndex        =   128
            Top             =   3000
            Width           =   765
         End
         Begin VB.Label Label14 
            Alignment       =   2  'Center
            Caption         =   "Depth"
            Height          =   255
            Left            =   2475
            TabIndex        =   126
            Top             =   2115
            Width           =   735
         End
         Begin VB.Label Label13 
            Alignment       =   2  'Center
            Caption         =   "Block"
            Height          =   255
            Left            =   1800
            TabIndex        =   125
            Top             =   2115
            Width           =   735
         End
         Begin VB.Label Label12 
            Caption         =   "Shift Top Block"
            Height          =   255
            Left            =   705
            TabIndex        =   124
            Top             =   2430
            Width           =   1035
         End
         Begin VB.Label Label5 
            Caption         =   "Run Single Cards"
            Height          =   255
            Left            =   690
            TabIndex        =   121
            Top             =   1830
            Width           =   1335
         End
         Begin VB.Label CutLabel 
            Caption         =   "Cut"
            Height          =   255
            Left            =   720
            TabIndex        =   110
            Top             =   1320
            Width           =   1335
         End
         Begin VB.Label OverhandLabel 
            Caption         =   "Overhand"
            Height          =   255
            Left            =   720
            TabIndex        =   109
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label RiffleLabel 
            Caption         =   "Riffle"
            Height          =   255
            Left            =   720
            TabIndex        =   108
            Top             =   600
            Width           =   1335
         End
         Begin VB.Label ConvShuffleNumberLabel 
            Alignment       =   2  'Center
            Caption         =   "Number"
            Height          =   255
            Left            =   4305
            TabIndex        =   107
            Top             =   285
            Width           =   735
         End
         Begin VB.Label ConvShuffleLocationLabel 
            Caption         =   "Location"
            Height          =   255
            Left            =   5145
            TabIndex        =   106
            Top             =   285
            Width           =   1215
         End
      End
      Begin VB.Frame ResetDeckFrame 
         Caption         =   "Arrange Cards"
         Height          =   1050
         Left            =   135
         TabIndex        =   100
         Top             =   480
         Width           =   8970
         Begin VB.CommandButton RefreshView 
            Height          =   255
            Left            =   7695
            Style           =   1  'Graphical
            TabIndex        =   131
            TabStop         =   0   'False
            Top             =   495
            Width           =   375
         End
         Begin VB.ComboBox AssemblePokerDealCombo 
            Height          =   345
            ItemData        =   "StackView8.frx":12ADB
            Left            =   5730
            List            =   "StackView8.frx":12AE8
            TabIndex        =   47
            Top             =   630
            Width           =   1530
         End
         Begin VB.CommandButton AssemblePokerDealButton 
            Height          =   255
            Left            =   3735
            Style           =   1  'Graphical
            TabIndex        =   91
            TabStop         =   0   'False
            Top             =   660
            Width           =   375
         End
         Begin VB.CommandButton ResetCurrentDeckButton 
            Height          =   255
            Left            =   105
            MaskColor       =   &H80000000&
            Style           =   1  'Graphical
            TabIndex        =   89
            TabStop         =   0   'False
            Top             =   660
            Width           =   375
         End
         Begin VB.ComboBox PokerDealCombo 
            Height          =   345
            ItemData        =   "StackView8.frx":12B09
            Left            =   5040
            List            =   "StackView8.frx":12B28
            TabIndex        =   46
            Top             =   240
            Width           =   2220
         End
         Begin VB.CommandButton PokerDealButton 
            Height          =   255
            Left            =   3735
            Style           =   1  'Graphical
            TabIndex        =   90
            TabStop         =   0   'False
            Top             =   285
            Width           =   375
         End
         Begin VB.CommandButton ResetDeckButton 
            Height          =   255
            Left            =   105
            MaskColor       =   &H80000000&
            Style           =   1  'Graphical
            TabIndex        =   88
            TabStop         =   0   'False
            Top             =   285
            Width           =   375
         End
         Begin VB.ComboBox SetStackCombo 
            Height          =   345
            ItemData        =   "StackView8.frx":12B7E
            Left            =   1275
            List            =   "StackView8.frx":12BAC
            TabIndex        =   45
            Top             =   255
            Width           =   2220
         End
         Begin VB.Label Label30 
            Caption         =   "Refresh Deck View"
            Height          =   465
            Left            =   8160
            TabIndex        =   132
            Top             =   390
            Width           =   750
         End
         Begin VB.Label AssemblePokerDealLabel 
            Caption         =   "Assemble Poker Deal"
            Height          =   255
            Left            =   4185
            TabIndex        =   123
            Top             =   660
            Width           =   1470
         End
         Begin VB.Label ResetCurrentDeckLabel 
            Caption         =   "Reset Current Deck Order"
            Height          =   255
            Left            =   600
            TabIndex        =   122
            Top             =   690
            Width           =   2235
         End
         Begin VB.Label PokerDealLabel 
            Caption         =   "Poker Deal"
            Height          =   255
            Left            =   4185
            TabIndex        =   120
            Top             =   270
            Width           =   1335
         End
         Begin VB.Label SetStackLabel 
            Caption         =   "Set Stack"
            Height          =   255
            Left            =   585
            TabIndex        =   101
            Top             =   270
            Width           =   1335
         End
      End
      Begin VB.Label Label20 
         Caption         =   "Clear All Reversed Cards"
         Height          =   255
         Left            =   -67575
         TabIndex        =   212
         Top             =   930
         Width           =   1920
      End
      Begin VB.Label Label18 
         Caption         =   "Reverse All Selections"
         Height          =   255
         Left            =   -67575
         TabIndex        =   200
         Top             =   1230
         Width           =   1920
      End
      Begin VB.Label Label17 
         Caption         =   "Insert Session File as Macro"
         Height          =   300
         Left            =   -69600
         TabIndex        =   195
         Top             =   5235
         Width           =   2040
      End
      Begin VB.Label SessionFileName 
         Appearance      =   0  'Flat
         Caption         =   "No current Session"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   -74760
         TabIndex        =   178
         Top             =   645
         Width           =   4305
      End
      Begin VB.Label Label6 
         Caption         =   "Play ALL current Session"
         Height          =   300
         Left            =   -69600
         TabIndex        =   177
         Top             =   2160
         Width           =   2040
      End
      Begin VB.Label Label7 
         Caption         =   "Clear ALL Session Events"
         Height          =   300
         Left            =   -69600
         TabIndex        =   176
         Top             =   4350
         Width           =   2040
      End
      Begin VB.Label Label8 
         Caption         =   "Move Event UP"
         Height          =   300
         Left            =   -69600
         TabIndex        =   175
         Top             =   3255
         Width           =   2040
      End
      Begin VB.Label Label9 
         Caption         =   "Move Event DOWN"
         Height          =   300
         Left            =   -69600
         TabIndex        =   174
         Top             =   3690
         Width           =   2040
      End
      Begin VB.Label Label10 
         Caption         =   "Delete current Event"
         Height          =   300
         Left            =   -69600
         TabIndex        =   173
         Top             =   4785
         Width           =   2040
      End
      Begin VB.Label Label11 
         Caption         =   "Play current Event"
         Height          =   300
         Left            =   -69600
         TabIndex        =   172
         Top             =   2565
         Width           =   2040
      End
      Begin VB.Label SessionRecordLabel 
         Caption         =   "Start Recording"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   300
         Left            =   -69150
         TabIndex        =   171
         Top             =   1425
         Width           =   1695
      End
      Begin VB.Line Line2 
         X1              =   -70110
         X2              =   -66015
         Y1              =   3015
         Y2              =   3015
      End
      Begin VB.Line Line3 
         X1              =   -70125
         X2              =   -66030
         Y1              =   4110
         Y2              =   4110
      End
      Begin VB.Label SessionRecordingStatus 
         Caption         =   "Not Recording"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -70215
         TabIndex        =   170
         Top             =   975
         Width           =   2655
      End
      Begin VB.Label Label4 
         Caption         =   "Clear All Selections"
         Height          =   255
         Left            =   -67575
         TabIndex        =   161
         Top             =   615
         Width           =   1920
      End
      Begin VB.Label SelectionsLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Selections"
         Height          =   255
         Left            =   -72855
         TabIndex        =   160
         Top             =   600
         Width           =   795
      End
   End
End
Attribute VB_Name = "frmStackView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub Card2C_DragDrop(Source As Control, X As Single, Y As Single)
    Source.Move (X - DragX + OriginalLeft), (Y - DragY + OriginalTop)
End Sub

Public Sub Card2C_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Card2C.Drag 1
    DragX = X
    DragY = Y
    OriginalTop = Card2C.Top
    OriginalLeft = Card2C.Left
End Sub




Private Sub ClearReversedCards_Click()
For i% = 1 To DeckCount
    Deck(6, i%) = False
Next i%
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
ElseIf PokerCardsDealt = 1 Then
    ShowDeal
Else
    ShowCards
End If
End Sub

Public Sub ClearSelections_Click()
'NumberOfSelectedCards = 0
'Erase SelectedCards
SelectionsTextBox.Text = Empty
For i% = 1 To DeckCount
    Deck(4, i%) = Empty
Next i%
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
ElseIf PokerCardsDealt = 1 Then
    ShowDeal
Else
    ShowCards
End If
End Sub
Private Sub CutDeckPrecise(cutdepthparameter, paramreverse)
    Dim pReverse As String
    Dim pCutStartCard As Integer
    Dim pCutEndCard As Integer
    pReverse = paramreverse
    CutDepth = Val(cutdepthparameter)
    If CutDepth = 0 Then
        Exit Sub
    End If
    'reverse a block if required
    If pReverse = "T" Then
        pCutStartCard = 1
        pCutEndCard = CutDepth
    ElseIf pReverse = "B" Then
        pCutStartCard = CutDepth + 1
        pCutEndCard = 52
    End If
    If Not pReverse = "X" Then
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
    'resume cut procedure
    i = 1
    For j% = CutDepth + 1 To DeckCount
        For z% = 1 To DeckProperties
            ChangedDeck(z%, i) = Deck(z%, j%)
        Next z%
        i = i + 1
    Next j%
    For k% = 1 To CutDepth
        For z% = 1 To DeckProperties
            ChangedDeck(z%, i) = Deck(z%, k%)
        Next z%
        i = i + 1
    Next k%
    For m% = 1 To DeckCount
        For z% = 1 To DeckProperties
            Deck(z%, m%) = ChangedDeck(z%, m%)
        Next z%
    Next m%
    ShowCards
End Sub

Public Sub AssemblePokerDealButton_Click()
If PokerCardsDealt = 0 Then
    MsgBox ("You may only use this command" & Chr(13) & _
        "if there are poker hands dealt.")
    Exit Sub
End If
If AssemblePokerDealCombo.ListIndex = -1 Then
    AssemblePokerDealCombo.Text = "Select order"
    MsgBox "Please select a valid response from" & Chr(13) _
        & "the 'Assemble Poker Deal' Dropdown Box"
    Exit Sub
End If
AssemblePokerDeal (AssemblePokerDealCombo.Text)
'SessionRecord
If SessionRecordMode Then
    SessionCommand = "AssemblePokerDeal(" & AssemblePokerDealCombo.Text & ")"
    SessionListBox.AddItem SessionCommand
    SessionStatusUpdate (0)
End If
'If AssemblePokerDealCombo.ListIndex = 0 Then
'    AssemblePokerDeal ("Backwards")
'    Debug.Print "List Index 0 - backwards"
'    'SessionRecord
'    If SessionRecordMode Then
'        SessionCommand = "AssemblePokerDeal(" & Chr(34) & "Backwards" & Chr(34) & ")"
'        SessionListBox.AddItem SessionCommand
'        SessionStatusUpdate (0)
'    End If
'ElseIf AssemblePokerDealCombo.ListIndex = 1 Then
'    AssemblePokerDeal ("Forwards")
'    Debug.Print "List Index 1 - forwards"
'    'SessionRecord
'    If SessionRecordMode Then
'        SessionCommand = "AssemblePokerDeal(" & Chr(34) & "Forwards" & Chr(34) & ")"
'        SessionListBox.AddItem SessionCommand
'        SessionStatusUpdate (0)
'    End If
'End If
End Sub
Public Sub AssemblePokerDeal(order)
If PokerCardsDealt = 0 Then
    MsgBox ("In a session, the " & Chr(34) & "AssemblePokerDeal" _
            & Chr(34) & " command" & Chr(13) & _
        "must come immediately after a " & Chr(34) & "PokerDeal" & Chr(34) & " command.")
    Exit Sub
End If
If order = Chr(34) & "Backwards" & Chr(34) Or _
    order = "Backwards" Then
    ShowCards
ElseIf order = Chr(34) & "Forwards" & Chr(34) Or _
    order = "Forwards" Then
    For k% = 1 To Hands
        For m% = 1 To 5
            For p% = 1 To DeckProperties
                ChangedDeck(p%, (k% - 1) * 5 + m%) = _
                        Deck(p%, (Hands - k%) * 5 + m%)
            Next p%
        Next m%
    Next k%
    For i% = (Hands * 5) + 1 To DeckCount
        For p% = 1 To DeckProperties
            ChangedDeck(p%, i%) = Deck(p%, i%)
        Next p%
    Next i%
    For m% = 1 To DeckCount
        For z% = 1 To DeckProperties
            Deck(z%, m%) = ChangedDeck(z%, m%)
        Next z%
    Next m%
    ShowCards
ElseIf order = Chr(34) & "Unwind" & Chr(34) Or _
    order = "Unwind" Then
    For m% = 1 To DeckCount
        For z% = 1 To DeckProperties
            Deck(z%, m%) = UnwindDeck(z%, m%)
        Next z%
    Next m%
    ShowCards
End If
End Sub

Private Sub ShiftTopBlock(sblock, sdepth)
ShiftBlock = Val(sblock)
ShiftDepth = Val(sdepth)
    i = 1
    For j% = ShiftBlock + 1 To ShiftBlock + ShiftDepth
        For z% = 1 To DeckProperties
            ChangedDeck(z%, i) = Deck(z%, j%)
        Next z%
        i = i + 1
    Next j%
    For k% = 1 To ShiftBlock
        For z% = 1 To DeckProperties
            ChangedDeck(z%, i) = Deck(z%, k%)
        Next z%
        i = i + 1
    Next k%
    If ShiftBlock + ShiftDepth < 52 Then
        For p% = ShiftBlock + ShiftDepth + 1 To DeckCount
            For z% = 1 To DeckProperties
                ChangedDeck(z%, i) = Deck(z%, p%)
            Next z%
            i = i + 1
        Next p%
    End If
    For m% = 1 To DeckCount
        For z% = 1 To DeckProperties
            Deck(z%, m%) = ChangedDeck(z%, m%)
        Next z%
    Next m%
    ShowCards
End Sub

Private Sub ShiftTopBlockReverse(sblock, sdepth)
ShiftBlock = Val(sblock)
ShiftDepth = Val(sdepth)
ReverseBlock (ShiftBlock)
    i = 1
    For j% = ShiftBlock + 1 To ShiftBlock + ShiftDepth
        For z% = 1 To DeckProperties
            ChangedDeck(z%, i) = Deck(z%, j%)
        Next z%
        i = i + 1
    Next j%
    For k% = 1 To ShiftBlock
        For z% = 1 To DeckProperties
            ChangedDeck(z%, i) = Deck(z%, k%)
        Next z%
        i = i + 1
    Next k%
    If ShiftBlock + ShiftDepth < 52 Then
        For p% = ShiftBlock + ShiftDepth + 1 To DeckCount
            For z% = 1 To DeckProperties
                ChangedDeck(z%, i) = Deck(z%, p%)
            Next z%
            i = i + 1
        Next p%
    End If
    For m% = 1 To DeckCount
        For z% = 1 To DeckProperties
            Deck(z%, m%) = ChangedDeck(z%, m%)
        Next z%
    Next m%
    ShowCards
End Sub


Private Sub InverseShiftTopBlock(sblock, sdepth)
'for the inverse, just inverse the first two declarations
'to the opposite assignments from the regular subroutine
ShiftBlock = Val(sdepth)
ShiftDepth = Val(sblock)
    i = 1
    For j% = ShiftBlock + 1 To ShiftBlock + ShiftDepth
        For z% = 1 To DeckProperties
            ChangedDeck(z%, i) = Deck(z%, j%)
        Next z%
        i = i + 1
    Next j%
    For k% = 1 To ShiftBlock
        For z% = 1 To DeckProperties
            ChangedDeck(z%, i) = Deck(z%, k%)
        Next z%
        i = i + 1
    Next k%
    If ShiftBlock + ShiftDepth < 52 Then
        For p% = ShiftBlock + ShiftDepth + 1 To DeckCount
            For z% = 1 To DeckProperties
                ChangedDeck(z%, i) = Deck(z%, p%)
            Next z%
            i = i + 1
        Next p%
    End If
    For m% = 1 To DeckCount
        For z% = 1 To DeckProperties
            Deck(z%, m%) = ChangedDeck(z%, m%)
        Next z%
    Next m%
    ShowCards
End Sub

Private Sub InverseShiftTopBlockReverse(sblock, sdepth)
'for the inverse, just inverse the first two declarations
'to the opposite assignments from the regular subroutine
ShiftBlock = Val(sdepth)
ShiftDepth = Val(sblock)
    i = 1
    For j% = ShiftBlock + 1 To ShiftBlock + ShiftDepth
        For z% = 1 To DeckProperties
            ChangedDeck(z%, i) = Deck(z%, j%)
        Next z%
        i = i + 1
    Next j%
    For k% = 1 To ShiftBlock
        For z% = 1 To DeckProperties
            ChangedDeck(z%, i) = Deck(z%, k%)
        Next z%
        i = i + 1
    Next k%
    If ShiftBlock + ShiftDepth < 52 Then
        For p% = ShiftBlock + ShiftDepth + 1 To DeckCount
            For z% = 1 To DeckProperties
                ChangedDeck(z%, i) = Deck(z%, p%)
            Next z%
            i = i + 1
        Next p%
    End If
    For m% = 1 To DeckCount
        For z% = 1 To DeckProperties
            Deck(z%, m%) = ChangedDeck(z%, m%)
        Next z%
    Next m%
    ReverseBlock (ShiftDepth)
    'need to use ShiftDepth due to earlier switch of parameters for the inverse action
    ShowCards
End Sub


Private Sub CountFromBack_Click()
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
ElseIf PokerCardsDealt = 1 Then
    ShowDeal
Else
    ShowCards
End If
End Sub

Private Sub CountFromFace_Click()
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
ElseIf PokerCardsDealt = 1 Then
    ShowDeal
Else
    ShowCards
End If
End Sub

Public Sub CutPreciseOption_Click()
    CutSpecialCombo.Text = Empty
End Sub

Public Sub CutPreciseTextBox_GotFocus()
    CutPreciseOption.Value = True
End Sub

Public Sub CutRandomOption_Click()
    CutPreciseTextBox.Text = Empty
    CutSpecialCombo.Text = Empty
End Sub

Public Sub CutShuffleButton_Click()
Dim pReverse As String
If CutReverseTopCheck.Value = 1 Then
    pReverse = "T"
ElseIf CutReverseBottomCheck.Value = 1 Then
    pReverse = "B"
Else
    pReverse = "X"
End If
If CutPreciseOption.Value = True And _
    (Not IsNumeric(CutPreciseTextBox.Text) Or _
    Val(CutPreciseTextBox.Text) < 1 Or _
    Val(CutPreciseTextBox.Text) > 52) Then
    CutPreciseTextBox.Text = Empty
    MsgBox "Please enter a valid card number (1 to 52)" & Chr(13) _
        & "in the 'Cut: Precise' Input Box"
    Exit Sub
End If
If CutSpecialOption = True And _
    CutSpecialCombo.ListIndex = -1 Then
    CutSpecialCombo.Text = "Select cut area"
    MsgBox "Please select a valid response from" & Chr(13) _
        & "the 'Cut: Special Random' Dropdown Box"
    Exit Sub
End If
If CutRandomOption.Value = True Then
    Call CutDeckRandom(pReverse)
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "CutDeckRandom(" & pReverse & ")"
        SessionListBox.AddItem SessionCommand
        SessionStatusUpdate (0)
    End If
ElseIf CutPreciseOption.Value = True Then
    Call CutDeckPrecise(Val(CutPreciseTextBox.Text), pReverse)
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "CutDeckPrecise(" & CutPreciseTextBox.Text & _
        ", " & pReverse & ")"
        SessionListBox.AddItem SessionCommand
        SessionStatusUpdate (0)
    End If
ElseIf CutSpecialOption.Value = True Then
    Call CutSpecialRandom(CutSpecialCombo.Text, pReverse)
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "CutSpecialRandom(" & CutSpecialCombo.Text & _
        ", " & pReverse & ")"
        SessionListBox.AddItem SessionCommand
        SessionStatusUpdate (0)
    End If
End If
End Sub

Public Sub CutSpecialCombo_GotFocus()
    CutSpecialOption.Value = True
End Sub

Public Sub CutSpecialOption_Click()
    CutPreciseTextBox.Text = Empty
End Sub

Public Sub CutSpecialRandom(csrp, paramreverse)
    If csrp = "Quarter" Then
        CutDepth = Int(Rnd * 10) + 8
    ElseIf csrp = "Third" Then
        CutDepth = Int(Rnd * 12) + 11
    ElseIf csrp = "Half" Then
        CutDepth = Int(Rnd * 20) + 16
    ElseIf csrp = "Two Thirds" Then
        CutDepth = Int(Rnd * 12) + 28
    ElseIf csrp = "Three Quarters" Then
        CutDepth = Int(Rnd * 10) + 34
    ElseIf csrp = "Shallow" Then
        CutDepth = Int(Rnd * 20) + 5
    ElseIf csrp = "Deep" Then
        CutDepth = Int(Rnd * 20) + 28
    End If
    Dim pReverse As String
    Dim pCutStartCard As Integer
    Dim pCutEndCard As Integer
    pReverse = paramreverse
    'reverse a block if required
    If pReverse = "T" Then
        pCutStartCard = 1
        pCutEndCard = CutDepth
    ElseIf pReverse = "B" Then
        pCutStartCard = CutDepth + 1
        pCutEndCard = 52
    End If
    If Not pReverse = "X" Then
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
    'resume cut procedure
    i = 1
    For j% = CutDepth + 1 To DeckCount
        For z% = 1 To DeckProperties
            ChangedDeck(z%, i) = Deck(z%, j%)
        Next z%
        i = i + 1
    Next j%
    For k% = 1 To CutDepth
        For z% = 1 To DeckProperties
            ChangedDeck(z%, i) = Deck(z%, k%)
        Next z%
        i = i + 1
    Next k%
    For m% = 1 To DeckCount
        For z% = 1 To DeckProperties
            Deck(z%, m%) = ChangedDeck(z%, m%)
        Next z%
    Next m%
    ShowCards
End Sub

Public Sub ReverseBlock(blocksizeparam)
BlockSize = Val(blocksizeparam)
For m% = 1 To BlockSize
    For z% = 1 To DeckProperties
        ChangedDeck(z%, BlockSize - m% + 1) = Deck(z%, m%)
    Next z%
Next m%
For m% = 1 To BlockSize
    For z% = 1 To DeckProperties
        Deck(z%, m%) = ChangedDeck(z%, m%)
    Next z%
Next m%
For i% = 1 To BlockSize
    Deck(6, i%) = Not Deck(6, i%)
Next i%
ShowCards
End Sub

Public Sub ForceButton_Click()
Dim paramreverse As String
If ForceReverseCheck.Value = 0 Then
    paramreverse = "X"
ElseIf ForceReverseCheck.Value = 1 Then
    paramreverse = "R"
End If
If Not IsNumeric(ForcePositionTextBox.Text) Or _
    Val(ForcePositionTextBox.Text) < 1 Or _
    Val(ForcePositionTextBox.Text) > 52 Then
    ForcePositionTextBox.Text = Empty
    MsgBox "Please enter a valid card position (1 to 52)" & Chr(13) _
        & "in the 'Force Card: Position' Input Box"
    Exit Sub
End If
If ForceReturnPositionOption = True And _
    (Not IsNumeric(ForceReturnPositionTextBox.Text) Or _
    Val(ForceReturnPositionTextBox.Text) < 1 Or _
    Val(ForceReturnPositionTextBox.Text) > 52) Then
    ForceReturnPositionTextBox.Text = Empty
    MsgBox "Please enter a valid card position (1 to 52)" & Chr(13) _
        & "in the 'Return: Controlled Position' Input Box"
    Exit Sub
End If
If ForceReturnAnywhereOption = True Then
    ForceReturnPositionTextBox.Text = Empty
    Call ForceCard(Val(ForcePositionTextBox.Text), paramreverse)
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ForceCard(" & ForcePositionTextBox.Text & _
            ", " & paramreverse & ")"
        SessionListBox.AddItem SessionCommand
        SessionStatusUpdate (0)
    End If
    ReturnCard ("Anywhere")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReturnCard(" & Chr(34) & "Anywhere" & Chr(34) & ")"
        SessionListBox.AddItem SessionCommand
        SessionStatusUpdate (0)
    End If
ElseIf ForceReturnPositionOption = True Then
    Call ForceCard(Val(ForcePositionTextBox.Text), paramreverse)
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ForceCard(" & ForcePositionTextBox.Text & _
            ", " & paramreverse & ")"
        SessionListBox.AddItem SessionCommand
        SessionStatusUpdate (0)
    End If
    ReturnCard (Val(ForceReturnPositionTextBox.Text))
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "ReturnCard(" & ForceReturnPositionTextBox.Text & ")"
        SessionListBox.AddItem SessionCommand
        SessionStatusUpdate (0)
    End If
End If
End Sub
Public Sub FreeChoiceSpreadSelect(freesc, paramR)
If freesc = Chr(34) & "Any Card" & Chr(34) Or _
    freesc = "Any Card" Then
    SelectedCard = Int(Rnd * DeckCount) + 1
ElseIf freesc = "Top Third" Then
    SelectedCard = Int(Rnd * 17) + 1
ElseIf freesc = "Middle Third" Then
    SelectedCard = Int(Rnd * 18) + 18
ElseIf freesc = "Bottom Third" Then
    SelectedCard = Int(Rnd * 17) + 36
End If
Deck(4, SelectedCard) = "Selected"
If SelectionsTextBox.Text = Empty Then
    SelectionsTextBox.Text = Deck(2, SelectedCard)
Else
    SelectionsTextBox.Text = SelectionsTextBox.Text & " " & Deck(2, SelectedCard)
End If
If paramR = "R" Then
    Deck(6, SelectedCard) = Not Deck(6, SelectedCard)
End If
End Sub
Public Sub ForceCard(fcp, paramR)
SelectedCard = Val(fcp)
Deck(4, SelectedCard) = "Selected"
If SelectionsTextBox.Text = Empty Then
    SelectionsTextBox.Text = Deck(2, SelectedCard)
Else
    SelectionsTextBox.Text = SelectionsTextBox.Text & " " & Deck(2, SelectedCard)
End If
If paramR = "R" Then
    Deck(6, SelectedCard) = Not Deck(6, SelectedCard)
End If
End Sub
Public Sub ForceReturnAnywhereOption_Click()
    ForceReturnPositionTextBox.Text = Empty
End Sub



Public Sub ForceReturnPositionTextBox_GotFocus()
    ForceReturnPositionOption.Value = True
End Sub

Private Sub Form_Activate()
SessionRecordButtons
End Sub

Public Sub Form_DragDrop(Source As Control, X As Single, Y As Single)
    Source.Move (X - DragX), (Y - DragY)
End Sub





Public Sub FreeChoiceHandlingButton_Click()
Dim paramReverse1 As String
Dim paramReverse2 As String
Dim paramReverse3 As String
If SpecialReverseFirstCheck.Value = 0 Then
    paramReverse1 = "X"
ElseIf SpecialReverseFirstCheck.Value = 1 Then
    paramReverse1 = "R"
End If
If SpecialReverseSecondCheck.Value = 0 Then
    paramReverse2 = "X"
ElseIf SpecialReverseSecondCheck.Value = 1 Then
    paramReverse2 = "R"
End If
If SpecialReverseThirdCheck.Value = 0 Then
    paramReverse3 = "X"
ElseIf SpecialReverseThirdCheck.Value = 1 Then
    paramReverse3 = "R"
End If
If FreeChoiceHandlingSelectCombo.ListIndex = -1 Then
    FreeChoiceHandlingSelectCombo.Text = "Select special selection handling"
    For i% = 0 To 7
        SpecialHandlingPicture(i%).Visible = False
    Next i%
    SpecialHandlingPicture(8).Visible = True
End If
If FreeChoiceHandlingSelectCombo.ListIndex = 0 Then
    Call SelectCardsCutSelectNext1(paramReverse1)
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "SelectCardsCutSelectNext1(" & _
            paramReverse1 & ")"
        SessionListBox.AddItem SessionCommand
        SessionStatusUpdate (0)
    End If
ElseIf FreeChoiceHandlingSelectCombo.ListIndex = 1 Then
    Call SelectCardsCutSelectNext2(paramReverse1, paramReverse2)
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "SelectCardsCutSelectNext2(" & _
            paramReverse1 & ", " & paramReverse2 & ")"
        SessionListBox.AddItem SessionCommand
        SessionStatusUpdate (0)
    End If
ElseIf FreeChoiceHandlingSelectCombo.ListIndex = 2 Then
    Call SelectCardsCutSelectNext3(paramReverse1, paramReverse2, paramReverse3)
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "SelectCardsCutSelectNext3(" & _
            paramReverse1 & ", " & paramReverse2 & ", " & paramReverse3 & ")"
        SessionListBox.AddItem SessionCommand
        SessionStatusUpdate (0)
    End If
ElseIf FreeChoiceHandlingSelectCombo.ListIndex = 3 Then
    Call SelectCardsCutSelectFace1(paramReverse1)
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "SelectCardsCutSelectFace1(" & _
            paramReverse1 & ")"
        SessionListBox.AddItem SessionCommand
        SessionStatusUpdate (0)
    End If
ElseIf FreeChoiceHandlingSelectCombo.ListIndex = 4 Then
    Call SelectCardsCutSelectFace2(paramReverse1, paramReverse2)
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "SelectCardsCutSelectFace2(" & _
            paramReverse1 & ", " & paramReverse2 & ")"
        SessionListBox.AddItem SessionCommand
        SessionStatusUpdate (0)
    End If
ElseIf FreeChoiceHandlingSelectCombo.ListIndex = 5 Then
    Call SelectCardsCutSelectFace3(paramReverse1, paramReverse2, paramReverse3)
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "SelectCardsCutSelectFace3(" & _
            paramReverse1 & ", " & paramReverse2 & ", " & paramReverse3 & ")"
        SessionListBox.AddItem SessionCommand
        SessionStatusUpdate (0)
    End If
ElseIf FreeChoiceHandlingSelectCombo.ListIndex = 6 Then
    Call SelectCardsCutSelectNextRepeat(paramReverse1, paramReverse2)
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "SelectCardsCutSelectNextRepeat(" & _
            paramReverse1 & ", " & paramReverse2 & ")"
        SessionListBox.AddItem SessionCommand
        SessionStatusUpdate (0)
    End If
ElseIf FreeChoiceHandlingSelectCombo.ListIndex = 7 Then
    Call SelectCardsCutSelectNextRepeat2(paramReverse1, paramReverse2, paramReverse3)
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "SelectCardsCutSelectNextRepeat2(" & _
            paramReverse1 & ", " & paramReverse2 & ", " & paramReverse3 & ")"
        SessionListBox.AddItem SessionCommand
        SessionStatusUpdate (0)
    End If
End If
End Sub

Public Sub FreeChoiceHandlingSelectCombo_Click()
i% = FreeChoiceHandlingSelectCombo.ListIndex
For j% = 0 To 8
    If j% = i% Then
        SpecialHandlingPicture(j%).Visible = True
    Else
        SpecialHandlingPicture(j%).Visible = False
    End If
Next j%
Select Case i%
    Case 0
        SpecialReverseFirstCheck.Enabled = True
        SpecialReverseSecondCheck.Enabled = False
        SpecialReverseThirdCheck.Enabled = False
        SpecialReverseSecondCheck.Value = 0
        SpecialReverseThirdCheck.Value = 0
    Case 1
        SpecialReverseFirstCheck.Enabled = True
        SpecialReverseSecondCheck.Enabled = True
        SpecialReverseThirdCheck.Enabled = False
        SpecialReverseThirdCheck.Value = 0
    Case 2
        SpecialReverseFirstCheck.Enabled = True
        SpecialReverseSecondCheck.Enabled = True
        SpecialReverseThirdCheck.Enabled = True
    Case 3
        SpecialReverseFirstCheck.Enabled = True
        SpecialReverseSecondCheck.Enabled = False
        SpecialReverseThirdCheck.Enabled = False
        SpecialReverseSecondCheck.Value = 0
        SpecialReverseThirdCheck.Value = 0
    Case 4
        SpecialReverseFirstCheck.Enabled = True
        SpecialReverseSecondCheck.Enabled = True
        SpecialReverseThirdCheck.Enabled = False
        SpecialReverseThirdCheck.Value = 0
    Case 5
        SpecialReverseFirstCheck.Enabled = True
        SpecialReverseSecondCheck.Enabled = True
        SpecialReverseThirdCheck.Enabled = True
    Case 6
        SpecialReverseFirstCheck.Enabled = True
        SpecialReverseSecondCheck.Enabled = True
        SpecialReverseThirdCheck.Enabled = False
        SpecialReverseThirdCheck.Value = 0
    Case 7
        SpecialReverseFirstCheck.Enabled = True
        SpecialReverseSecondCheck.Enabled = True
        SpecialReverseThirdCheck.Enabled = True
    End Select
End Sub



Public Sub FreeChoiceSelectSpecificCombo_Click()
If FreeChoiceSelectSpecificCombo.ListIndex = -1 Then
    FreeChoiceSelectSpecificCombo.Text = "Identify selection area"
End If
End Sub

Public Sub FreeChoiceSelectSpecificCombo_GotFocus()
    FreeChoiceSpreadSelectSpecificOption.Value = True
End Sub

Public Sub FreeChoiceSpreadButton_Click()
    Dim paramreverse As String
    If FreeChoiceReverseCheck.Value = 0 Then
        paramreverse = "X"
    ElseIf FreeChoiceReverseCheck.Value = 1 Then
        paramreverse = "R"
    End If
    ' this first section checks for error conditions
    ' in the input text boxes and combo lists
    If FreeChoiceSpreadSelectSpecificOption = True And _
        FreeChoiceSelectSpecificCombo.ListIndex = -1 Then
        FreeChoiceSelectSpecificCombo.Text = "Identify selection area"
        MsgBox "Please select a valid response from" & Chr(13) _
            & "the 'Select: Specific Section' Dropdown Box"
        Exit Sub
    End If
    If FreeChoiceSpreadReturnPositionOption = True And _
        (Not IsNumeric(FreeChoiceSpreadReturnPositionTextBox.Text) Or _
        Val(FreeChoiceSpreadReturnPositionTextBox.Text) < 1 Or _
        Val(FreeChoiceSpreadReturnPositionTextBox.Text) > 52) Then
        FreeChoiceSpreadReturnPositionTextBox.Text = Empty
        MsgBox "Please enter a valid card position (1 to 52)" & Chr(13) _
            & "in the 'Controlled Position' Input Box"
        Exit Sub
    End If
    If FreeChoiceSpreadReturnSectionOption = True And _
        FreeChoiceSpreadReturnSectionCombo.ListIndex = -1 Then
        FreeChoiceSpreadReturnSectionCombo.Text = "Select return area"
        MsgBox "Please select a valid response from" & Chr(13) _
            & "the 'Return: Specific Section' Dropdown Box"
        Exit Sub
    End If
    ' this second section runs the codes based on the combinations
    If FreeChoiceSpreadSelectAnyCardOption = True And _
        FreeChoiceSpreadReturnOriginalOption = True Then
        FreeChoiceSelectSpecificCombo.Text = ""
        FreeChoiceSpreadReturnPositionTextBox.Text = ""
        FreeChoiceSpreadReturnSectionCombo.Text = ""
        Call FreeChoiceSpreadSelect("Any Card", paramreverse)
        'SessionRecord
        If SessionRecordMode Then
            SessionCommand = "FreeChoiceSpreadSelect(" & Chr(34) & "Any Card" & Chr(34) & _
                ", " & paramreverse & ")"
            SessionListBox.AddItem SessionCommand
            SessionStatusUpdate (0)
        End If
        ReturnCard ("Original Position")
        'SessionRecord
        If SessionRecordMode Then
            SessionCommand = "ReturnCard(" & Chr(34) & "Original Position" & Chr(34) & ")"
            SessionListBox.AddItem SessionCommand
            SessionStatusUpdate (0)
        End If
    ElseIf FreeChoiceSpreadSelectAnyCardOption = True And _
        FreeChoiceSpreadReturnAnywhereOption = True Then
        FreeChoiceSelectSpecificCombo.Text = ""
        FreeChoiceSpreadReturnPositionTextBox.Text = ""
        FreeChoiceSpreadReturnSectionCombo.Text = ""
        Call FreeChoiceSpreadSelect("Any Card", paramreverse)
        'SessionRecord
        If SessionRecordMode Then
            SessionCommand = "FreeChoiceSpreadSelect(" & Chr(34) & "Any Card" & Chr(34) & _
                ", " & paramreverse & ")"
            SessionListBox.AddItem SessionCommand
            SessionStatusUpdate (0)
        End If
        ReturnCard ("Anywhere")
        'SessionRecord
        If SessionRecordMode Then
            SessionCommand = "ReturnCard(" & Chr(34) & "Anywhere" & Chr(34) & ")"
            SessionListBox.AddItem SessionCommand
            SessionStatusUpdate (0)
        End If
    ElseIf FreeChoiceSpreadSelectAnyCardOption = True And _
        FreeChoiceSpreadReturnPositionOption = True Then
        FreeChoiceSelectSpecificCombo.Text = ""
        FreeChoiceSpreadReturnSectionCombo.Text = ""
        Call FreeChoiceSpreadSelect("Any Card", paramreverse)
        'SessionRecord
        If SessionRecordMode Then
            SessionCommand = "FreeChoiceSpreadSelect(" & Chr(34) & "Any Card" & Chr(34) & _
                ", " & paramreverse & ")"
            SessionListBox.AddItem SessionCommand
            SessionStatusUpdate (0)
        End If
        ReturnCard (Val(FreeChoiceSpreadReturnPositionTextBox.Text))
        'SessionRecord
        If SessionRecordMode Then
            SessionCommand = "ReturnCard(" & FreeChoiceSpreadReturnPositionTextBox.Text & ")"
            SessionListBox.AddItem SessionCommand
            SessionStatusUpdate (0)
        End If
    ElseIf FreeChoiceSpreadSelectAnyCardOption = True And _
        FreeChoiceSpreadReturnSectionOption = True Then
        FreeChoiceSelectSpecificCombo.Text = ""
        FreeChoiceSpreadReturnPositionTextBox.Text = ""
        Call FreeChoiceSpreadSelect("Any Card", paramreverse)
        'SessionRecord
        If SessionRecordMode Then
            SessionCommand = "FreeChoiceSpreadSelect(" & Chr(34) & "Any Card" & Chr(34) & _
                ", " & paramreverse & ")"
            SessionListBox.AddItem SessionCommand
            SessionStatusUpdate (0)
        End If
        If FreeChoiceSpreadReturnSectionCombo.ListIndex = 0 Then
            ReturnCard ("Top Third")
            'SessionRecord
            If SessionRecordMode Then
                SessionCommand = "ReturnCard(" & FreeChoiceSpreadReturnSectionCombo.List(0) & ")"
                SessionListBox.AddItem SessionCommand
                SessionStatusUpdate (0)
            End If
        ElseIf FreeChoiceSpreadReturnSectionCombo.ListIndex = 1 Then
            ReturnCard ("Middle Third")
            'SessionRecord
            If SessionRecordMode Then
                SessionCommand = "ReturnCard(" & FreeChoiceSpreadReturnSectionCombo.List(1) & ")"
                SessionListBox.AddItem SessionCommand
                SessionStatusUpdate (0)
            End If
        ElseIf FreeChoiceSpreadReturnSectionCombo.ListIndex = 2 Then
            ReturnCard ("Bottom Third")
            'SessionRecord
            If SessionRecordMode Then
                SessionCommand = "ReturnCard(" & FreeChoiceSpreadReturnSectionCombo.List(2) & ")"
                SessionListBox.AddItem SessionCommand
                SessionStatusUpdate (0)
            End If
        End If
    ElseIf FreeChoiceSpreadSelectSpecificOption = True And _
        FreeChoiceSpreadReturnOriginalOption = True Then
        FreeChoiceSpreadReturnPositionTextBox.Text = ""
        FreeChoiceSpreadReturnSectionCombo.Text = ""
        If FreeChoiceSelectSpecificCombo.ListIndex = 0 Then
            Call FreeChoiceSpreadSelect("Top Third", paramreverse)
            'SessionRecord
            If SessionRecordMode Then
                SessionCommand = "FreeChoiceSpreadSelect(" & FreeChoiceSelectSpecificCombo.List(0) & _
                    ", " & paramreverse & ")"
                SessionListBox.AddItem SessionCommand
                SessionStatusUpdate (0)
            End If
        ElseIf FreeChoiceSelectSpecificCombo.ListIndex = 1 Then
            Call FreeChoiceSpreadSelect("Middle Third", paramreverse)
            'SessionRecord
            If SessionRecordMode Then
                SessionCommand = "FreeChoiceSpreadSelect(" & FreeChoiceSelectSpecificCombo.List(1) & _
                    ", " & paramreverse & ")"
                SessionListBox.AddItem SessionCommand
                SessionStatusUpdate (0)
            End If
        ElseIf FreeChoiceSelectSpecificCombo.ListIndex = 2 Then
            Call FreeChoiceSpreadSelect("Bottom Third", paramreverse)
            'SessionRecord
            If SessionRecordMode Then
                SessionCommand = "FreeChoiceSpreadSelect(" & FreeChoiceSelectSpecificCombo.List(2) & _
                    ", " & paramreverse & ")"
                SessionListBox.AddItem SessionCommand
                SessionStatusUpdate (0)
            End If
        End If
        ReturnCard ("Original Position")
        'SessionRecord
        If SessionRecordMode Then
            SessionCommand = "ReturnCard(" & Chr(34) & "Original Position" & Chr(34) & ")"
            SessionListBox.AddItem SessionCommand
            SessionStatusUpdate (0)
        End If
    ElseIf FreeChoiceSpreadSelectSpecificOption = True And _
        FreeChoiceSpreadReturnAnywhereOption = True Then
        FreeChoiceSpreadReturnPositionTextBox.Text = ""
        FreeChoiceSpreadReturnSectionCombo.Text = ""
        If FreeChoiceSelectSpecificCombo.ListIndex = 0 Then
            Call FreeChoiceSpreadSelect("Top Third", paramreverse)
            'SessionRecord
            If SessionRecordMode Then
                SessionCommand = "FreeChoiceSpreadSelect(" & FreeChoiceSelectSpecificCombo.List(0) & _
                    ", " & paramreverse & ")"
                SessionListBox.AddItem SessionCommand
                SessionStatusUpdate (0)
            End If
        ElseIf FreeChoiceSelectSpecificCombo.ListIndex = 1 Then
            Call FreeChoiceSpreadSelect("Middle Third", paramreverse)
            'SessionRecord
            If SessionRecordMode Then
                SessionCommand = "FreeChoiceSpreadSelect(" & FreeChoiceSelectSpecificCombo.List(1) & _
                    ", " & paramreverse & ")"
                SessionListBox.AddItem SessionCommand
                SessionStatusUpdate (0)
            End If
        ElseIf FreeChoiceSelectSpecificCombo.ListIndex = 2 Then
            Call FreeChoiceSpreadSelect("Bottom Third", paramreverse)
            'SessionRecord
            If SessionRecordMode Then
                SessionCommand = "FreeChoiceSpreadSelect(" & FreeChoiceSelectSpecificCombo.List(2) & _
                    ", " & paramreverse & ")"
                SessionListBox.AddItem SessionCommand
                SessionStatusUpdate (0)
            End If
        End If
        ReturnCard ("Anywhere")
        'SessionRecord
        If SessionRecordMode Then
            SessionCommand = "ReturnCard(" & Chr(34) & "Anywhere" & Chr(34) & ")"
            SessionListBox.AddItem SessionCommand
            SessionStatusUpdate (0)
        End If
    ElseIf FreeChoiceSpreadSelectSpecificOption = True And _
        FreeChoiceSpreadReturnPositionOption = True Then
        FreeChoiceSpreadReturnSectionCombo.Text = ""
        If FreeChoiceSelectSpecificCombo.ListIndex = 0 Then
            Call FreeChoiceSpreadSelect("Top Third", paramreverse)
            'SessionRecord
            If SessionRecordMode Then
                SessionCommand = "FreeChoiceSpreadSelect(" & FreeChoiceSelectSpecificCombo.List(0) & _
                    ", " & paramreverse & ")"
                SessionListBox.AddItem SessionCommand
                SessionStatusUpdate (0)
            End If
        ElseIf FreeChoiceSelectSpecificCombo.ListIndex = 1 Then
            Call FreeChoiceSpreadSelect("Middle Third", paramreverse)
            'SessionRecord
            If SessionRecordMode Then
                SessionCommand = "FreeChoiceSpreadSelect(" & FreeChoiceSelectSpecificCombo.List(1) & _
                    ", " & paramreverse & ")"
                SessionListBox.AddItem SessionCommand
                SessionStatusUpdate (0)
            End If
        ElseIf FreeChoiceSelectSpecificCombo.ListIndex = 2 Then
            Call FreeChoiceSpreadSelect("Bottom Third", paramreverse)
            'SessionRecord
            If SessionRecordMode Then
                SessionCommand = "FreeChoiceSpreadSelect(" & FreeChoiceSelectSpecificCombo.List(2) & _
                    ", " & paramreverse & ")"
                SessionListBox.AddItem SessionCommand
                SessionStatusUpdate (0)
            End If
        End If
        ReturnCard (Val(FreeChoiceSpreadReturnPositionTextBox.Text))
        'SessionRecord
        If SessionRecordMode Then
            SessionCommand = "ReturnCard(" & FreeChoiceSpreadReturnPositionTextBox.Text & ")"
            SessionListBox.AddItem SessionCommand
            SessionStatusUpdate (0)
        End If
    ElseIf FreeChoiceSpreadSelectSpecificOption = True And _
        FreeChoiceSpreadReturnSectionOption = True Then
        FreeChoiceSpreadReturnPositionTextBox.Text = ""
        If FreeChoiceSelectSpecificCombo.ListIndex = 0 Then
            Call FreeChoiceSpreadSelect("Top Third", paramreverse)
            'SessionRecord
            If SessionRecordMode Then
                SessionCommand = "FreeChoiceSpreadSelect(" & FreeChoiceSelectSpecificCombo.List(0) & _
                    ", " & paramreverse & ")"
                SessionListBox.AddItem SessionCommand
                SessionStatusUpdate (0)
            End If
        ElseIf FreeChoiceSelectSpecificCombo.ListIndex = 1 Then
            Call FreeChoiceSpreadSelect("Middle Third", paramreverse)
            'SessionRecord
            If SessionRecordMode Then
                SessionCommand = "FreeChoiceSpreadSelect(" & FreeChoiceSelectSpecificCombo.List(1) & _
                    ", " & paramreverse & ")"
                SessionListBox.AddItem SessionCommand
                SessionStatusUpdate (0)
            End If
        ElseIf FreeChoiceSelectSpecificCombo.ListIndex = 2 Then
            Call FreeChoiceSpreadSelect("Bottom Third", paramreverse)
            'SessionRecord
            If SessionRecordMode Then
                SessionCommand = "FreeChoiceSpreadSelect(" & FreeChoiceSelectSpecificCombo.List(2) & _
                    ", " & paramreverse & ")"
                SessionListBox.AddItem SessionCommand
                SessionStatusUpdate (0)
            End If
        End If
        If FreeChoiceSpreadReturnSectionCombo.ListIndex = 0 Then
            ReturnCard ("Top Third")
            'SessionRecord
            If SessionRecordMode Then
                SessionCommand = "ReturnCard(" & FreeChoiceSpreadReturnSectionCombo.List(0) & ")"
                SessionListBox.AddItem SessionCommand
                SessionStatusUpdate (0)
            End If
        ElseIf FreeChoiceSpreadReturnSectionCombo.ListIndex = 1 Then
            ReturnCard ("Middle Third")
            'SessionRecord
            If SessionRecordMode Then
                SessionCommand = "ReturnCard(" & FreeChoiceSpreadReturnSectionCombo.List(1) & ")"
                SessionListBox.AddItem SessionCommand
                SessionStatusUpdate (0)
            End If
        ElseIf FreeChoiceSpreadReturnSectionCombo.ListIndex = 2 Then
            ReturnCard ("Bottom Third")
            'SessionRecord
            If SessionRecordMode Then
                SessionCommand = "ReturnCard(" & FreeChoiceSpreadReturnSectionCombo.List(2) & ")"
                SessionListBox.AddItem SessionCommand
                SessionStatusUpdate (0)
            End If
        End If
    End If
End Sub

Public Sub FreeChoiceSpreadReturnAnywhereOption_Click()
    FreeChoiceSpreadReturnPositionTextBox.Text = ""
    FreeChoiceSpreadReturnSectionCombo.Text = ""
End Sub

Public Sub FreeChoiceSpreadReturnOriginalOption_Click()
    FreeChoiceSpreadReturnPositionTextBox.Text = ""
    FreeChoiceSpreadReturnSectionCombo.Text = ""
End Sub

Public Sub FreeChoiceSpreadReturnPositionOption_Click()
    FreeChoiceSpreadReturnSectionCombo.Text = ""
End Sub



Public Sub FreeChoiceSpreadReturnPositionTextBox_GotFocus()
    FreeChoiceSpreadReturnPositionOption.Value = True
End Sub



Public Sub FreeChoiceSpreadReturnSectionCombo_GotFocus()
    FreeChoiceSpreadReturnSectionOption.Value = True
End Sub

Public Sub FreeChoiceSpreadReturnSectionOption_Click()
    FreeChoiceSpreadReturnPositionTextBox.Text = ""
    FreeChoiceSpreadReturnSectionCombo.Text = "Identify selection area"
End Sub

Public Sub FreeChoiceSpreadSelectAnyCardOption_Click()
FreeChoiceSelectSpecificCombo.Text = ""
End Sub

Public Sub FreeChoiceSpreadSelectSpecificOption_Click()
FreeChoiceSelectSpecificCombo.Text = "Identify selection area"
End Sub

Public Sub HighlightSelectionsCheck_Click()
If HighlightSelectionsCheck.Value = 0 Then
    SelectionsTextBox.PasswordChar = "*"
ElseIf HighlightSelectionsCheck.Value = 1 Then
    SelectionsTextBox.PasswordChar = Empty
End If
If PilesShown = 1 Then
    ShowPiles
ElseIf PokerCardsDealt = 1 Then
    ShowDeal
Else
    ShowCards
End If
End Sub

Private Sub InFaro()
For i% = 1 To 26
    For z% = 1 To DeckProperties
        ChangedDeck(z%, 2 * i% - 1) = Deck(z%, i% + 26)
        ChangedDeck(z%, 2 * i%) = Deck(z%, i%)
    Next z%
Next i%
For m% = 1 To DeckCount
    For z% = 1 To DeckProperties
        Deck(z%, m%) = ChangedDeck(z%, m%)
    Next z%
Next m%
ShowCards
End Sub

Private Sub InFaroReverse()
ReverseBlock (26)
For i% = 1 To 26
    For z% = 1 To DeckProperties
        ChangedDeck(z%, 2 * i% - 1) = Deck(z%, i% + 26)
        ChangedDeck(z%, 2 * i%) = Deck(z%, i%)
    Next z%
Next i%
For m% = 1 To DeckCount
    For z% = 1 To DeckProperties
        Deck(z%, m%) = ChangedDeck(z%, m%)
    Next z%
Next m%
ShowCards
End Sub


Public Sub InFaroButton_Click()
If InFaroStandardOption.Value = False And _
    (Not IsNumeric(InFaroFromTopTextBox.Text) Or _
    Val(InFaroFromTopTextBox.Text) < 1 Or _
    Val(InFaroFromTopTextBox.Text) > 52) Then
    InFaroFromTopTextBox.Text = Empty
    MsgBox "Please enter a valid card position (1 to 52)" & Chr(13) _
        & "in the 'In Faro: From Top' Input Box"
    Exit Sub
End If
If InFaroSpecialOption.Value = True And _
    InFaroStartWeaveTopOption.Value = False And _
    InFaroStartWeaveBottomOption.Value = False Then
    MsgBox "Please select an 'In Faro: Start Weave' option"
    Exit Sub
End If
If InFaroStandardOption.Value = False And _
    InFaroInteriorTextBox.Text <> Empty And _
    (Not IsNumeric(InFaroInteriorTextBox.Text) Or _
    (Val(InFaroInteriorTextBox.Text) + _
        Val(InFaroFromTopTextBox.Text)) > 52) Then
    InFaroInteriorTextBox.Text = Empty
    MsgBox "Please enter a valid card position in the" & Chr(13) _
        & "'In Faro: Interior Position' Input Box" & Chr(13) & Chr(13) _
        & "The sum of 'From Top' and 'Interior Position'" & Chr(13) _
        & "must not be greater than 52."
    Exit Sub
End If

If InFaroInverseCheck Then
    If InFaroStandardOption.Value = True Then
        If InFaroReverseCheck Then
            InverseInFaroReverse
            'SessionRecord
            If SessionRecordMode Then
                SessionCommand = "InverseInFaroReverse"
                SessionListBox.AddItem SessionCommand
                SessionStatusUpdate (0)
            End If
        Else
            InverseInFaro
            'SessionRecord
            If SessionRecordMode Then
                SessionCommand = "InverseInFaro"
                SessionListBox.AddItem SessionCommand
                SessionStatusUpdate (0)
            End If
        End If
    ElseIf InFaroStartWeaveTopOption.Value = True Then
        If InFaroReverseCheck Then
            Call InverseInFaroSpecialTopReverse(Val(InFaroFromTopTextBox.Text), Val(InFaroInteriorTextBox.Text))
            'SessionRecord
            If SessionRecordMode Then
                If InFaroInteriorTextBox.Text <> Empty Then
                    SessionCommand = "InverseInFaroSpecialTopReverse(" & Val(InFaroFromTopTextBox.Text) _
                        & ", " & Val(InFaroInteriorTextBox.Text) & ")"
                Else
                    SessionCommand = "InverseInFaroSpecialTopReverse(" & Val(InFaroFromTopTextBox.Text) _
                        & ", 1)"
                End If
                SessionListBox.AddItem SessionCommand
                SessionStatusUpdate (0)
            End If
        Else
            Call InverseInFaroSpecialTop(Val(InFaroFromTopTextBox.Text), Val(InFaroInteriorTextBox.Text))
            'SessionRecord
            If SessionRecordMode Then
                If InFaroInteriorTextBox.Text <> Empty Then
                    SessionCommand = "InverseInFaroSpecialTop(" & Val(InFaroFromTopTextBox.Text) _
                        & ", " & Val(InFaroInteriorTextBox.Text) & ")"
                Else
                    SessionCommand = "InverseInFaroSpecialTop(" & Val(InFaroFromTopTextBox.Text) _
                        & ", 1)"
                End If
                SessionListBox.AddItem SessionCommand
                SessionStatusUpdate (0)
            End If
        End If
    ElseIf InFaroStartWeaveBottomOption.Value = True Then
        If InFaroReverseCheck Then
            Call InverseInFaroSpecialBottomReverse(Val(InFaroFromTopTextBox.Text), Val(InFaroInteriorTextBox.Text))
            'SessionRecord
            If SessionRecordMode Then
                If InFaroInteriorTextBox.Text <> Empty Then
                    SessionCommand = "InverseInFaroSpecialBottomReverse(" & Val(InFaroFromTopTextBox.Text) _
                        & ", " & Val(InFaroInteriorTextBox.Text) & ")"
                Else
                    SessionCommand = "InverseInFaroSpecialBottomReverse(" & Val(InFaroFromTopTextBox.Text) _
                        & ", 1)"
                End If
                SessionListBox.AddItem SessionCommand
                SessionStatusUpdate (0)
            End If
        Else
            Call InverseInFaroSpecialBottom(Val(InFaroFromTopTextBox.Text), Val(InFaroInteriorTextBox.Text))
            'SessionRecord
            If SessionRecordMode Then
                If InFaroInteriorTextBox.Text <> Empty Then
                    SessionCommand = "InverseInFaroSpecialBottom(" & Val(InFaroFromTopTextBox.Text) _
                        & ", " & Val(InFaroInteriorTextBox.Text) & ")"
                Else
                    SessionCommand = "InverseInFaroSpecialBottom(" & Val(InFaroFromTopTextBox.Text) _
                        & ", 1)"
                End If
                SessionListBox.AddItem SessionCommand
                SessionStatusUpdate (0)
            End If
        End If
    End If
Else
    If InFaroStandardOption.Value = True Then
        If InFaroReverseCheck Then
            InFaroReverse
            'SessionRecord
            If SessionRecordMode Then
                SessionCommand = "InFaroReverse"
                SessionListBox.AddItem SessionCommand
                SessionStatusUpdate (0)
            End If
        Else
            InFaro
            'SessionRecord
            If SessionRecordMode Then
                SessionCommand = "InFaro"
                SessionListBox.AddItem SessionCommand
                SessionStatusUpdate (0)
            End If
        End If
    ElseIf InFaroStartWeaveTopOption.Value = True Then
        If InFaroReverseCheck Then
            Call InFaroSpecialTopReverse(Val(InFaroFromTopTextBox.Text), Val(InFaroInteriorTextBox.Text))
            'SessionRecord
            If SessionRecordMode Then
                If InFaroInteriorTextBox.Text <> Empty Then
                    SessionCommand = "InFaroSpecialTopReverse(" & Val(InFaroFromTopTextBox.Text) _
                        & ", " & Val(InFaroInteriorTextBox.Text) & ")"
                Else
                    SessionCommand = "InFaroSpecialTopReverse(" & Val(InFaroFromTopTextBox.Text) _
                        & ", 1)"
                End If
                SessionListBox.AddItem SessionCommand
                SessionStatusUpdate (0)
            End If
        Else
            Call InFaroSpecialTop(Val(InFaroFromTopTextBox.Text), Val(InFaroInteriorTextBox.Text))
            'SessionRecord
            If SessionRecordMode Then
                If InFaroInteriorTextBox.Text <> Empty Then
                    SessionCommand = "InFaroSpecialTop(" & Val(InFaroFromTopTextBox.Text) _
                        & ", " & Val(InFaroInteriorTextBox.Text) & ")"
                Else
                    SessionCommand = "InFaroSpecialTop(" & Val(InFaroFromTopTextBox.Text) _
                        & ", 1)"
                End If
                SessionListBox.AddItem SessionCommand
                SessionStatusUpdate (0)
            End If
        End If
    ElseIf InFaroStartWeaveBottomOption.Value = True Then
        If InFaroReverseCheck Then
            Call InFaroSpecialBottomReverse(Val(InFaroFromTopTextBox.Text), Val(InFaroInteriorTextBox.Text))
            'SessionRecord
            If SessionRecordMode Then
                If InFaroInteriorTextBox.Text <> Empty Then
                    SessionCommand = "InFaroSpecialBottomReverse(" & Val(InFaroFromTopTextBox.Text) _
                        & ", " & Val(InFaroInteriorTextBox.Text) & ")"
                Else
                    SessionCommand = "InFaroSpecialBottomReverse(" & Val(InFaroFromTopTextBox.Text) _
                        & ", 1)"
                End If
                SessionListBox.AddItem SessionCommand
                SessionStatusUpdate (0)
            End If
        Else
            Call InFaroSpecialBottom(Val(InFaroFromTopTextBox.Text), Val(InFaroInteriorTextBox.Text))
            'SessionRecord
            If SessionRecordMode Then
                If InFaroInteriorTextBox.Text <> Empty Then
                    SessionCommand = "InFaroSpecialBottom(" & Val(InFaroFromTopTextBox.Text) _
                        & ", " & Val(InFaroInteriorTextBox.Text) & ")"
                Else
                    SessionCommand = "InFaroSpecialBottom(" & Val(InFaroFromTopTextBox.Text) _
                        & ", 1)"
                End If
                SessionListBox.AddItem SessionCommand
                SessionStatusUpdate (0)
            End If
        End If
    End If
End If
End Sub

Public Sub InFaroFromTopTextBox_GotFocus()
    InFaroSpecialOption.Value = True
End Sub

Private Sub InFaroSpecialBottom(isbnumber, ifinumber)
ProtectedBlock = Val(isbnumber)
InteriorCard = Val(ifinumber)
Call CutDeckPrecise(ProtectedBlock, "X")
'The 'from bottom' first requires a cut of the Protected Block.
'Then do the outfaro version (with a Inversed deck) to accomplish a
'proper infaro offset block mesh
InverseDeck
If ProtectedBlock = 26 And InteriorCard = Empty Then
    OutFaro
    '(because the deck was Inversed, and the CutPrecise was employed,
    ' I need to do an OutFaro to have the resultant InFaro come out correctly)
ElseIf ProtectedBlock < 26 Then
    If InteriorCard = Empty Then
        For i% = 1 To ProtectedBlock
            For z% = 1 To DeckProperties
                ChangedDeck(z%, 2 * i% - 1) = Deck(z%, i%)
                ChangedDeck(z%, 2 * i%) = Deck(z%, i% + ProtectedBlock)
            Next z%
        Next i%
        MeshedBlock = 2 * ProtectedBlock
        For k% = 1 To DeckCount - MeshedBlock
            For z% = 1 To DeckProperties
                ChangedDeck(z%, k% + MeshedBlock) = Deck(z%, k% + MeshedBlock)
            Next z%
        Next k%
        For m% = 1 To DeckCount
            For z% = 1 To DeckProperties
                Deck(z%, m%) = ChangedDeck(z%, m%)
            Next z%
        Next m%
    Else
        InteriorPosition = InteriorCard + ProtectedBlock
        If ProtectedBlock + InteriorPosition <= 52 Then
            For i% = 1 To InteriorPosition - ProtectedBlock - 1
            'added the "- 1" from InFaro solution
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, i%) = Deck(z%, ProtectedBlock + i%)
                Next z%
            Next i%
            MeshedBlockPosition = InteriorPosition - ProtectedBlock - 1
            'added the "- 1" from InFaro solution
            For i% = 1 To ProtectedBlock
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, MeshedBlockPosition + (2 * i% - 1)) = _
                        Deck(z%, i%)
                    ChangedDeck(z%, MeshedBlockPosition + (2 * i%)) = _
                        Deck(z%, i% + InteriorPosition - 1)
                        'added the "- 1" from InFaro solution
                Next z%
            Next i%
            For i% = ProtectedBlock * 2 + InteriorPosition - ProtectedBlock To DeckCount
            'took out a "+ 1" right before the "To" from InFaro solution
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, i%) = Deck(z%, i%)
                Next z%
            Next i%
            For m% = 1 To DeckCount
                For z% = 1 To DeckProperties
                    Deck(z%, m%) = ChangedDeck(z%, m%)
                Next z%
            Next m%
        Else
            For i% = 1 To InteriorPosition - ProtectedBlock - 1
            'added the "- 1" from InFaro solution
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, i%) = Deck(z%, ProtectedBlock + i%)
                Next z%
            Next i%
            MeshedBlockPosition = InteriorPosition - ProtectedBlock - 1
            'added the "- 1" from InFaro solution
            MeshedBlock = 2 * (DeckCount - InteriorPosition + 1)
            'added the "+ 1" from InFaro solution
            For i% = 1 To DeckCount - InteriorPosition + 1
            'added the "+ 1" from InFaro solution
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, MeshedBlockPosition + (2 * i% - 1)) = _
                        Deck(z%, i%)
                    ChangedDeck(z%, MeshedBlockPosition + (2 * i%)) = _
                        Deck(z%, i% + InteriorPosition - 1)
                        'added the "- 1" from InFaro solution
                Next z%
            Next i%
            For i% = 1 To ProtectedBlock - (DeckCount - InteriorPosition + 1)
            'added the "+ 1" from InFaro solution
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, MeshedBlockPosition + MeshedBlock + i%) = _
                        Deck(z%, DeckCount - InteriorPosition + 1 + i%)
                        'added the "+ 1" from InFaro solution
                Next z%
            Next i%
            For m% = 1 To DeckCount
                For z% = 1 To DeckProperties
                    Deck(z%, m%) = ChangedDeck(z%, m%)
                Next z%
            Next m%
        End If
    End If
ElseIf ProtectedBlock >= 26 Then
    If InteriorCard = Empty Then
        For i% = 1 To DeckCount - ProtectedBlock
            For z% = 1 To DeckProperties
                ChangedDeck(z%, 2 * i% - 1) = Deck(z%, i%)
                ChangedDeck(z%, 2 * i%) = Deck(z%, i% + ProtectedBlock)
            Next z%
        Next i%
        MeshedBlock = 2 * (DeckCount - ProtectedBlock)
        For k% = 1 To DeckCount - MeshedBlock
            For z% = 1 To DeckProperties
                ChangedDeck(z%, k% + MeshedBlock) = Deck(z%, k% + DeckCount - ProtectedBlock)
            Next z%
        Next k%
        For m% = 1 To DeckCount
            For z% = 1 To DeckProperties
                Deck(z%, m%) = ChangedDeck(z%, m%)
            Next z%
        Next m%
    Else
        InteriorPosition = InteriorCard + ProtectedBlock
        If ProtectedBlock + InteriorPosition <= 52 Then
            For i% = 1 To InteriorPosition - ProtectedBlock - 1
            'added the "- 1" from InFaro solution
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, i%) = Deck(z%, ProtectedBlock + i%)
                Next z%
            Next i%
            MeshedBlockPosition = InteriorPosition - ProtectedBlock - 1
            'added the "- 1" from InFaro solution
            For i% = 1 To ProtectedBlock
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, MeshedBlockPosition + (2 * i% - 1)) = _
                        Deck(z%, i%)
                    ChangedDeck(z%, MeshedBlockPosition + (2 * i%)) = _
                        Deck(z%, i% + InteriorPosition - 1)
                        'added the "- 1" from InFaro solution
                Next z%
            Next i%
            For i% = ProtectedBlock * 2 + InteriorPosition - ProtectedBlock To DeckCount
            'took out a "+ 1" right before the "To" from InFaro solution
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, i%) = Deck(z%, i%)
                Next z%
            Next i%
            For m% = 1 To DeckCount
                For z% = 1 To DeckProperties
                    Deck(z%, m%) = ChangedDeck(z%, m%)
                Next z%
            Next m%
        Else
            For i% = 1 To InteriorPosition - ProtectedBlock - 1
            'added the "- 1" from InFaro solution
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, i%) = Deck(z%, ProtectedBlock + i%)
                Next z%
            Next i%
            MeshedBlockPosition = InteriorPosition - ProtectedBlock - 1
            'added the "- 1" from InFaro solution
            MeshedBlock = 2 * (DeckCount - InteriorPosition + 1)
            'added the "+ 1" from InFaro solution
            For i% = 1 To DeckCount - InteriorPosition + 1
            'added the "+ 1" from InFaro solution
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, MeshedBlockPosition + (2 * i% - 1)) = _
                        Deck(z%, i%)
                    ChangedDeck(z%, MeshedBlockPosition + (2 * i%)) = _
                        Deck(z%, i% + InteriorPosition - 1)
                        'added the "- 1" from InFaro solution
                Next z%
            Next i%
            For i% = 1 To ProtectedBlock - (DeckCount - InteriorPosition + 1)
            'added the "+ 1" from InFaro solution
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, MeshedBlockPosition + MeshedBlock + i%) = _
                        Deck(z%, DeckCount - InteriorPosition + 1 + i%)
                        'added the "+ 1" from InFaro solution
                Next z%
            Next i%
            For m% = 1 To DeckCount
                For z% = 1 To DeckProperties
                    Deck(z%, m%) = ChangedDeck(z%, m%)
                Next z%
            Next m%
        End If
    End If
End If
InverseDeck
ShowCards
'ORIGINAL CODE BELOW
'InverseDeck
'ProtectedBlock = DeckCount - Val(isbnumber)
'If ProtectedBlock = 26 Then
'    InFaro
'ElseIf ProtectedBlock < 26 Then
'    For i% = 1 To ProtectedBlock
'        For z% = 1 To DeckProperties
'            ChangedDeck(z%, 2 * i% - 1) = Deck(z%, i% + ProtectedBlock)
'            ChangedDeck(z%, 2 * i%) = Deck(z%, i%)
'        Next z%
'    Next i%
'    MeshedBlock = 2 * ProtectedBlock
'    For k% = 1 To DeckCount - MeshedBlock
'        For z% = 1 To DeckProperties
'            ChangedDeck(z%, k% + MeshedBlock) = Deck(z%, k% + MeshedBlock)
'        Next z%
'    Next k%
'    For m% = 1 To DeckCount
'        For z% = 1 To DeckProperties
'            Deck(z%, m%) = ChangedDeck(z%, m%)
'        Next z%
'    Next m%
'ElseIf ProtectedBlock > 26 Then
'     For i% = 1 To DeckCount - ProtectedBlock
'        For z% = 1 To DeckProperties
'            ChangedDeck(z%, 2 * i% - 1) = Deck(z%, i% + ProtectedBlock)
'            ChangedDeck(z%, 2 * i%) = Deck(z%, i%)
'        Next z%
'    Next i%
'    MeshedBlock = 2 * (DeckCount - ProtectedBlock)
'    For k% = 1 To DeckCount - MeshedBlock
'        For z% = 1 To DeckProperties
'            ChangedDeck(z%, k% + MeshedBlock) = Deck(z%, k% + DeckCount - ProtectedBlock)
'        Next z%
'    Next k%
'    For m% = 1 To DeckCount
'        For z% = 1 To DeckProperties
'            Deck(z%, m%) = ChangedDeck(z%, m%)
'        Next z%
'    Next m%
'End If
'InverseDeck
'ShowCards
End Sub

Private Sub InFaroSpecialBottomReverse(isbnumber, ifinumber)
ProtectedBlock = Val(isbnumber)
InteriorCard = Val(ifinumber)
Call CutDeckPrecise(ProtectedBlock, "X")
'The 'from bottom' first requires a cut of the Protected Block.
'Then do the outfaro version (with a Inversed deck) to accomplish a
'proper infaro offset block mesh
InverseDeck
If ProtectedBlock = 26 And InteriorCard = Empty Then
    OutFaroReverse
    '(because the deck was Inversed, and the CutPrecise was employed,
    ' I need to do an OutFaro to have the resultant InFaro come out correctly)
ElseIf ProtectedBlock < 26 Then
    If InteriorCard = Empty Then
        ReverseBlock (ProtectedBlock)
        For i% = 1 To ProtectedBlock
            For z% = 1 To DeckProperties
                ChangedDeck(z%, 2 * i% - 1) = Deck(z%, i%)
                ChangedDeck(z%, 2 * i%) = Deck(z%, i% + ProtectedBlock)
            Next z%
        Next i%
        MeshedBlock = 2 * ProtectedBlock
        For k% = 1 To DeckCount - MeshedBlock
            For z% = 1 To DeckProperties
                ChangedDeck(z%, k% + MeshedBlock) = Deck(z%, k% + MeshedBlock)
            Next z%
        Next k%
        For m% = 1 To DeckCount
            For z% = 1 To DeckProperties
                Deck(z%, m%) = ChangedDeck(z%, m%)
            Next z%
        Next m%
    Else
        InteriorPosition = InteriorCard + ProtectedBlock
        If ProtectedBlock + InteriorPosition <= 52 Then
            ReverseBlock (ProtectedBlock)
            For i% = 1 To InteriorPosition - ProtectedBlock - 1
            'added the "- 1" from InFaro solution
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, i%) = Deck(z%, ProtectedBlock + i%)
                Next z%
            Next i%
            MeshedBlockPosition = InteriorPosition - ProtectedBlock - 1
            'added the "- 1" from InFaro solution
            For i% = 1 To ProtectedBlock
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, MeshedBlockPosition + (2 * i% - 1)) = _
                        Deck(z%, i%)
                    ChangedDeck(z%, MeshedBlockPosition + (2 * i%)) = _
                        Deck(z%, i% + InteriorPosition - 1)
                        'added the "- 1" from InFaro solution
                Next z%
            Next i%
            For i% = ProtectedBlock * 2 + InteriorPosition - ProtectedBlock To DeckCount
            'took out a "+ 1" right before the "To" from InFaro solution
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, i%) = Deck(z%, i%)
                Next z%
            Next i%
            For m% = 1 To DeckCount
                For z% = 1 To DeckProperties
                    Deck(z%, m%) = ChangedDeck(z%, m%)
                Next z%
            Next m%
        Else
            ReverseBlock (ProtectedBlock)
            For i% = 1 To InteriorPosition - ProtectedBlock - 1
            'added the "- 1" from InFaro solution
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, i%) = Deck(z%, ProtectedBlock + i%)
                Next z%
            Next i%
            MeshedBlockPosition = InteriorPosition - ProtectedBlock - 1
            'added the "- 1" from InFaro solution
            MeshedBlock = 2 * (DeckCount - InteriorPosition + 1)
            'added the "+ 1" from InFaro solution
            For i% = 1 To DeckCount - InteriorPosition + 1
            'added the "+ 1" from InFaro solution
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, MeshedBlockPosition + (2 * i% - 1)) = _
                        Deck(z%, i%)
                    ChangedDeck(z%, MeshedBlockPosition + (2 * i%)) = _
                        Deck(z%, i% + InteriorPosition - 1)
                        'added the "- 1" from InFaro solution
                Next z%
            Next i%
            For i% = 1 To ProtectedBlock - (DeckCount - InteriorPosition + 1)
            'added the "+ 1" from InFaro solution
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, MeshedBlockPosition + MeshedBlock + i%) = _
                        Deck(z%, DeckCount - InteriorPosition + 1 + i%)
                        'added the "+ 1" from InFaro solution
                Next z%
            Next i%
            For m% = 1 To DeckCount
                For z% = 1 To DeckProperties
                    Deck(z%, m%) = ChangedDeck(z%, m%)
                Next z%
            Next m%
        End If
    End If
ElseIf ProtectedBlock >= 26 Then
    If InteriorCard = Empty Then
        ReverseBlock (ProtectedBlock)
        For i% = 1 To DeckCount - ProtectedBlock
            For z% = 1 To DeckProperties
                ChangedDeck(z%, 2 * i% - 1) = Deck(z%, i%)
                ChangedDeck(z%, 2 * i%) = Deck(z%, i% + ProtectedBlock)
            Next z%
        Next i%
        MeshedBlock = 2 * (DeckCount - ProtectedBlock)
        For k% = 1 To DeckCount - MeshedBlock
            For z% = 1 To DeckProperties
                ChangedDeck(z%, k% + MeshedBlock) = Deck(z%, k% + DeckCount - ProtectedBlock)
            Next z%
        Next k%
        For m% = 1 To DeckCount
            For z% = 1 To DeckProperties
                Deck(z%, m%) = ChangedDeck(z%, m%)
            Next z%
        Next m%
    Else
        InteriorPosition = InteriorCard + ProtectedBlock
        If ProtectedBlock + InteriorPosition <= 52 Then
            ReverseBlock (ProtectedBlock)
            For i% = 1 To InteriorPosition - ProtectedBlock - 1
            'added the "- 1" from InFaro solution
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, i%) = Deck(z%, ProtectedBlock + i%)
                Next z%
            Next i%
            MeshedBlockPosition = InteriorPosition - ProtectedBlock - 1
            'added the "- 1" from InFaro solution
            For i% = 1 To ProtectedBlock
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, MeshedBlockPosition + (2 * i% - 1)) = _
                        Deck(z%, i%)
                    ChangedDeck(z%, MeshedBlockPosition + (2 * i%)) = _
                        Deck(z%, i% + InteriorPosition - 1)
                        'added the "- 1" from InFaro solution
                Next z%
            Next i%
            For i% = ProtectedBlock * 2 + InteriorPosition - ProtectedBlock To DeckCount
            'took out a "+ 1" right before the "To" from InFaro solution
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, i%) = Deck(z%, i%)
                Next z%
            Next i%
            For m% = 1 To DeckCount
                For z% = 1 To DeckProperties
                    Deck(z%, m%) = ChangedDeck(z%, m%)
                Next z%
            Next m%
        Else
            ReverseBlock (ProtectedBlock)
            For i% = 1 To InteriorPosition - ProtectedBlock - 1
            'added the "- 1" from InFaro solution
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, i%) = Deck(z%, ProtectedBlock + i%)
                Next z%
            Next i%
            MeshedBlockPosition = InteriorPosition - ProtectedBlock - 1
            'added the "- 1" from InFaro solution
            MeshedBlock = 2 * (DeckCount - InteriorPosition + 1)
            'added the "+ 1" from InFaro solution
            For i% = 1 To DeckCount - InteriorPosition + 1
            'added the "+ 1" from InFaro solution
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, MeshedBlockPosition + (2 * i% - 1)) = _
                        Deck(z%, i%)
                    ChangedDeck(z%, MeshedBlockPosition + (2 * i%)) = _
                        Deck(z%, i% + InteriorPosition - 1)
                        'added the "- 1" from InFaro solution
                Next z%
            Next i%
            For i% = 1 To ProtectedBlock - (DeckCount - InteriorPosition + 1)
            'added the "+ 1" from InFaro solution
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, MeshedBlockPosition + MeshedBlock + i%) = _
                        Deck(z%, DeckCount - InteriorPosition + 1 + i%)
                        'added the "+ 1" from InFaro solution
                Next z%
            Next i%
            For m% = 1 To DeckCount
                For z% = 1 To DeckProperties
                    Deck(z%, m%) = ChangedDeck(z%, m%)
                Next z%
            Next m%
        End If
    End If
End If
InverseDeck
ShowCards
End Sub


Private Sub InFaroSpecialTop(istnumber, ifinumber)
ProtectedBlock = Val(istnumber)
InteriorCard = Val(ifinumber)
If ProtectedBlock = 26 And InteriorCard = Empty Then
    InFaro
ElseIf ProtectedBlock < 26 Then
    If InteriorCard = Empty Then
        For i% = 1 To ProtectedBlock
            For z% = 1 To DeckProperties
                ChangedDeck(z%, 2 * i% - 1) = Deck(z%, i% + ProtectedBlock)
                ChangedDeck(z%, 2 * i%) = Deck(z%, i%)
            Next z%
        Next i%
        MeshedBlock = 2 * ProtectedBlock
        For k% = 1 To DeckCount - MeshedBlock
            For z% = 1 To DeckProperties
                ChangedDeck(z%, k% + MeshedBlock) = Deck(z%, k% + MeshedBlock)
            Next z%
        Next k%
        For m% = 1 To DeckCount
            For z% = 1 To DeckProperties
                Deck(z%, m%) = ChangedDeck(z%, m%)
            Next z%
        Next m%
    Else
        InteriorPosition = InteriorCard + ProtectedBlock
        If ProtectedBlock + InteriorPosition <= 52 Then
            For i% = 1 To InteriorPosition - ProtectedBlock
            'added the "- 1" from InFaro solution
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, i%) = Deck(z%, ProtectedBlock + i%)
                Next z%
            Next i%
            MeshedBlockPosition = InteriorPosition - ProtectedBlock
            'added the "- 1" from InFaro solution
            For i% = 1 To ProtectedBlock
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, MeshedBlockPosition + (2 * i% - 1)) = _
                        Deck(z%, i%)
                    ChangedDeck(z%, MeshedBlockPosition + (2 * i%)) = _
                        Deck(z%, i% + InteriorPosition)
                        'added the "- 1" from InFaro solution
                Next z%
            Next i%
            For i% = ProtectedBlock * 2 + InteriorPosition - ProtectedBlock + 1 To DeckCount
            'took out a "+ 1" right before the "To" from InFaro solution
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, i%) = Deck(z%, i%)
                Next z%
            Next i%
            For m% = 1 To DeckCount
                For z% = 1 To DeckProperties
                    Deck(z%, m%) = ChangedDeck(z%, m%)
                Next z%
            Next m%
        Else
            For i% = 1 To InteriorPosition - ProtectedBlock
            'added the "- 1" from InFaro solution
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, i%) = Deck(z%, ProtectedBlock + i%)
                Next z%
            Next i%
            MeshedBlockPosition = InteriorPosition - ProtectedBlock
            'added the "- 1" from InFaro solution
            MeshedBlock = 2 * (DeckCount - InteriorPosition)
            'added the "+ 1" from InFaro solution (inside parens after Interior Position)
            For i% = 1 To DeckCount - InteriorPosition
            'added the "+ 1" from InFaro solution
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, MeshedBlockPosition + (2 * i% - 1)) = _
                        Deck(z%, i%)
                    ChangedDeck(z%, MeshedBlockPosition + (2 * i%)) = _
                        Deck(z%, i% + InteriorPosition)
                        'added the "- 1" from InFaro solution (after Interior Position)
                Next z%
            Next i%
            For i% = 1 To ProtectedBlock - (DeckCount - InteriorPosition)
            'added the "+ 1" from InFaro solution (inside parens after Interior Position)
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, MeshedBlockPosition + MeshedBlock + i%) = _
                        Deck(z%, DeckCount - InteriorPosition + i%)
                        'added the "+ 1" from InFaro solution after Interior Position
                Next z%
            Next i%
            For m% = 1 To DeckCount
                For z% = 1 To DeckProperties
                    Deck(z%, m%) = ChangedDeck(z%, m%)
                Next z%
            Next m%
        End If
    End If
ElseIf ProtectedBlock >= 26 Then
    If InteriorCard = Empty Then
        For i% = 1 To DeckCount - ProtectedBlock
            For z% = 1 To DeckProperties
                ChangedDeck(z%, 2 * i% - 1) = Deck(z%, i% + ProtectedBlock)
                ChangedDeck(z%, 2 * i%) = Deck(z%, i%)
            Next z%
        Next i%
        MeshedBlock = 2 * (DeckCount - ProtectedBlock)
        For k% = 1 To DeckCount - MeshedBlock
            For z% = 1 To DeckProperties
                ChangedDeck(z%, k% + MeshedBlock) = Deck(z%, k% + DeckCount - ProtectedBlock)
            Next z%
        Next k%
        For m% = 1 To DeckCount
            For z% = 1 To DeckProperties
                Deck(z%, m%) = ChangedDeck(z%, m%)
            Next z%
        Next m%
    Else
        InteriorPosition = InteriorCard + ProtectedBlock
        If ProtectedBlock + InteriorPosition <= 52 Then
            For i% = 1 To InteriorPosition - ProtectedBlock
            'added the "- 1" from InFaro solution
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, i%) = Deck(z%, ProtectedBlock + i%)
                Next z%
            Next i%
            MeshedBlockPosition = InteriorPosition - ProtectedBlock
            'added the "- 1" from InFaro solution
            For i% = 1 To ProtectedBlock
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, MeshedBlockPosition + (2 * i% - 1)) = _
                        Deck(z%, i%)
                    ChangedDeck(z%, MeshedBlockPosition + (2 * i%)) = _
                        Deck(z%, i% + InteriorPosition)
                        'added the "- 1" from InFaro solution (inside parens after Interior Position)
                Next z%
            Next i%
            For i% = ProtectedBlock * 2 + InteriorPosition - ProtectedBlock + 1 To DeckCount
            'took out a "+ 1" right before the "To" from InFaro solution
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, i%) = Deck(z%, i%)
                Next z%
            Next i%
            For m% = 1 To DeckCount
                For z% = 1 To DeckProperties
                    Deck(z%, m%) = ChangedDeck(z%, m%)
                Next z%
            Next m%
        Else
            For i% = 1 To InteriorPosition - ProtectedBlock
            'added the "- 1" from InFaro solution
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, i%) = Deck(z%, ProtectedBlock + i%)
                Next z%
            Next i%
            MeshedBlockPosition = InteriorPosition - ProtectedBlock
            'added the "- 1" from InFaro solution
            MeshedBlock = 2 * (DeckCount - InteriorPosition)
            'added the "+ 1" from InFaro solution inside parens after InteriorPosition)
            For i% = 1 To DeckCount - InteriorPosition
            'added the "+ 1" from InFaro solution
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, MeshedBlockPosition + (2 * i% - 1)) = _
                        Deck(z%, i%)
                    ChangedDeck(z%, MeshedBlockPosition + (2 * i%)) = _
                        Deck(z%, i% + InteriorPosition)
                        'added the "- 1" from InFaro solution inside parens after Interior Position
                Next z%
            Next i%
            For i% = 1 To ProtectedBlock - (DeckCount - InteriorPosition)
            'added the "+ 1" from InFaro solution inside parens after InteriorPosition)
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, MeshedBlockPosition + MeshedBlock + i%) = _
                        Deck(z%, DeckCount - InteriorPosition + i%)
                        'added the "+ 1" from InFaro solution after InteriorPosition
                Next z%
            Next i%
            For m% = 1 To DeckCount
                For z% = 1 To DeckProperties
                    Deck(z%, m%) = ChangedDeck(z%, m%)
                Next z%
            Next m%
        End If
    End If
End If
ShowCards
End Sub

Private Sub InFaroSpecialTopReverse(istnumber, ifinumber)
ProtectedBlock = Val(istnumber)
InteriorCard = Val(ifinumber)
If ProtectedBlock = 26 And InteriorCard = Empty Then
    'ReverseBlock (26)
    InFaroReverse
ElseIf ProtectedBlock < 26 Then
    If InteriorCard = Empty Then
        ReverseBlock (ProtectedBlock)
        For i% = 1 To ProtectedBlock
            For z% = 1 To DeckProperties
                ChangedDeck(z%, 2 * i% - 1) = Deck(z%, i% + ProtectedBlock)
                ChangedDeck(z%, 2 * i%) = Deck(z%, i%)
            Next z%
        Next i%
        MeshedBlock = 2 * ProtectedBlock
        For k% = 1 To DeckCount - MeshedBlock
            For z% = 1 To DeckProperties
                ChangedDeck(z%, k% + MeshedBlock) = Deck(z%, k% + MeshedBlock)
            Next z%
        Next k%
        For m% = 1 To DeckCount
            For z% = 1 To DeckProperties
                Deck(z%, m%) = ChangedDeck(z%, m%)
            Next z%
        Next m%
    Else
        InteriorPosition = InteriorCard + ProtectedBlock
        If ProtectedBlock + InteriorPosition <= 52 Then
            ReverseBlock (ProtectedBlock)
            For i% = 1 To InteriorPosition - ProtectedBlock
            'added the "- 1" from InFaro solution
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, i%) = Deck(z%, ProtectedBlock + i%)
                Next z%
            Next i%
            MeshedBlockPosition = InteriorPosition - ProtectedBlock
            'added the "- 1" from InFaro solution
            For i% = 1 To ProtectedBlock
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, MeshedBlockPosition + (2 * i% - 1)) = _
                        Deck(z%, i%)
                    ChangedDeck(z%, MeshedBlockPosition + (2 * i%)) = _
                        Deck(z%, i% + InteriorPosition)
                        'added the "- 1" from InFaro solution
                Next z%
            Next i%
            For i% = ProtectedBlock * 2 + InteriorPosition - ProtectedBlock + 1 To DeckCount
            'took out a "+ 1" right before the "To" from InFaro solution
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, i%) = Deck(z%, i%)
                Next z%
            Next i%
            For m% = 1 To DeckCount
                For z% = 1 To DeckProperties
                    Deck(z%, m%) = ChangedDeck(z%, m%)
                Next z%
            Next m%
        Else
            ReverseBlock (ProtectedBlock)
            For i% = 1 To InteriorPosition - ProtectedBlock
            'added the "- 1" from InFaro solution
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, i%) = Deck(z%, ProtectedBlock + i%)
                Next z%
            Next i%
            MeshedBlockPosition = InteriorPosition - ProtectedBlock
            'added the "- 1" from InFaro solution
            MeshedBlock = 2 * (DeckCount - InteriorPosition)
            'added the "+ 1" from InFaro solution (inside parens after Interior Position)
            For i% = 1 To DeckCount - InteriorPosition
            'added the "+ 1" from InFaro solution
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, MeshedBlockPosition + (2 * i% - 1)) = _
                        Deck(z%, i%)
                    ChangedDeck(z%, MeshedBlockPosition + (2 * i%)) = _
                        Deck(z%, i% + InteriorPosition)
                        'added the "- 1" from InFaro solution (after Interior Position)
                Next z%
            Next i%
            For i% = 1 To ProtectedBlock - (DeckCount - InteriorPosition)
            'added the "+ 1" from InFaro solution (inside parens after Interior Position)
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, MeshedBlockPosition + MeshedBlock + i%) = _
                        Deck(z%, DeckCount - InteriorPosition + i%)
                        'added the "+ 1" from InFaro solution after Interior Position
                Next z%
            Next i%
            For m% = 1 To DeckCount
                For z% = 1 To DeckProperties
                    Deck(z%, m%) = ChangedDeck(z%, m%)
                Next z%
            Next m%
        End If
    End If
ElseIf ProtectedBlock >= 26 Then
    If InteriorCard = Empty Then
        ReverseBlock (ProtectedBlock)
        For i% = 1 To DeckCount - ProtectedBlock
            For z% = 1 To DeckProperties
                ChangedDeck(z%, 2 * i% - 1) = Deck(z%, i% + ProtectedBlock)
                ChangedDeck(z%, 2 * i%) = Deck(z%, i%)
            Next z%
        Next i%
        MeshedBlock = 2 * (DeckCount - ProtectedBlock)
        For k% = 1 To DeckCount - MeshedBlock
            For z% = 1 To DeckProperties
                ChangedDeck(z%, k% + MeshedBlock) = Deck(z%, k% + DeckCount - ProtectedBlock)
            Next z%
        Next k%
        For m% = 1 To DeckCount
            For z% = 1 To DeckProperties
                Deck(z%, m%) = ChangedDeck(z%, m%)
            Next z%
        Next m%
    Else
        InteriorPosition = InteriorCard + ProtectedBlock
        If ProtectedBlock + InteriorPosition <= 52 Then
            ReverseBlock (ProtectedBlock)
            For i% = 1 To InteriorPosition - ProtectedBlock
            'added the "- 1" from InFaro solution
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, i%) = Deck(z%, ProtectedBlock + i%)
                Next z%
            Next i%
            MeshedBlockPosition = InteriorPosition - ProtectedBlock
            'added the "- 1" from InFaro solution
            For i% = 1 To ProtectedBlock
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, MeshedBlockPosition + (2 * i% - 1)) = _
                        Deck(z%, i%)
                    ChangedDeck(z%, MeshedBlockPosition + (2 * i%)) = _
                        Deck(z%, i% + InteriorPosition)
                        'added the "- 1" from InFaro solution (inside parens after Interior Position)
                Next z%
            Next i%
            For i% = ProtectedBlock * 2 + InteriorPosition - ProtectedBlock + 1 To DeckCount
            'took out a "+ 1" right before the "To" from InFaro solution
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, i%) = Deck(z%, i%)
                Next z%
            Next i%
            For m% = 1 To DeckCount
                For z% = 1 To DeckProperties
                    Deck(z%, m%) = ChangedDeck(z%, m%)
                Next z%
            Next m%
        Else
            ReverseBlock (ProtectedBlock)
            For i% = 1 To InteriorPosition - ProtectedBlock
            'added the "- 1" from InFaro solution
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, i%) = Deck(z%, ProtectedBlock + i%)
                Next z%
            Next i%
            MeshedBlockPosition = InteriorPosition - ProtectedBlock
            'added the "- 1" from InFaro solution
            MeshedBlock = 2 * (DeckCount - InteriorPosition)
            'added the "+ 1" from InFaro solution inside parens after InteriorPosition)
            For i% = 1 To DeckCount - InteriorPosition
            'added the "+ 1" from InFaro solution
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, MeshedBlockPosition + (2 * i% - 1)) = _
                        Deck(z%, i%)
                    ChangedDeck(z%, MeshedBlockPosition + (2 * i%)) = _
                        Deck(z%, i% + InteriorPosition)
                        'added the "- 1" from InFaro solution inside parens after Interior Position
                Next z%
            Next i%
            For i% = 1 To ProtectedBlock - (DeckCount - InteriorPosition)
            'added the "+ 1" from InFaro solution inside parens after InteriorPosition)
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, MeshedBlockPosition + MeshedBlock + i%) = _
                        Deck(z%, DeckCount - InteriorPosition + i%)
                        'added the "+ 1" from InFaro solution after InteriorPosition
                Next z%
            Next i%
            For m% = 1 To DeckCount
                For z% = 1 To DeckProperties
                    Deck(z%, m%) = ChangedDeck(z%, m%)
                Next z%
            Next m%
        End If
    End If
End If
ShowCards
End Sub


Public Sub InFaroFromTopTextBox_LostFocus()
If InFaroFromTopTextBox.Text <> Empty And _
    InFaroStandardOption.Value = False And _
    (Not IsNumeric(InFaroFromTopTextBox.Text) Or _
    Val(InFaroFromTopTextBox.Text) < 1 Or _
    Val(InFaroFromTopTextBox.Text) > 52) Then
    InFaroFromTopTextBox.Text = Empty
    MsgBox "Please enter a valid card position (1 to 52)" & Chr(13) _
        & "in the 'In Faro: From Top' Input Box"
    InFaroFromTopTextBox.SetFocus
    Exit Sub
End If
End Sub

Public Sub InFaroInteriorTextBox_GotFocus()
    InFaroSpecialOption.Value = True
End Sub

Public Sub InFaroInteriorTextBox_LostFocus()
If InFaroInteriorTextBox.Text <> Empty And _
    InFaroStandardOption.Value = False And _
    (Not IsNumeric(InFaroInteriorTextBox.Text) Or _
    Val(InFaroInteriorTextBox.Text) < 1 Or _
    Val(InFaroInteriorTextBox.Text) > 52) Then
    InFaroInteriorTextBox.Text = Empty
    MsgBox "Please enter a valid card position (1 to 52)" & Chr(13) _
        & "in the 'In Faro: Interior Position' Input Box" & Chr(13) _
        & "(or leave it blank for a default value of 1)"
    InFaroInteriorTextBox.SetFocus
    Exit Sub
End If
End Sub


Public Sub InFaroStandardOption_Click()
    InFaroStartWeaveTopOption.Value = False
    InFaroStartWeaveBottomOption.Value = False
    InFaroFromTopTextBox.Text = Empty
    InFaroInteriorTextBox.Text = Empty
End Sub





Public Sub InFaroStartWeaveBottomOption_GotFocus()
    InFaroSpecialOption.Value = True
End Sub

Public Sub InFaroStartWeaveTopOption_GotFocus()
    InFaroSpecialOption.Value = True
End Sub



Public Sub OHShuffle()
StackIndex = 1
ChangedDeckIndex = 52
Do While StackIndex < 53
    Block = Int((Rnd ^ 2) * 10) + 1
    If Block > 53 - StackIndex Then
        Block = 52 - StackIndex
    End If
    tempIndex = ChangedDeckIndex + 1 - Block
    For i% = tempIndex To ChangedDeckIndex
        For z% = 1 To DeckProperties
            ChangedDeck(z%, i%) = Deck(z%, StackIndex + i% - tempIndex)
        Next z%
    Next i%
    ChangedDeckIndex = ChangedDeckIndex - Block
    StackIndex = StackIndex + Block
Loop
For m% = 1 To DeckCount
    For z% = 1 To DeckProperties
        Deck(z%, m%) = ChangedDeck(z%, m%)
    Next z%
Next m%
ShowCards
End Sub

Public Sub OHShuffleBottom(ohparameter)
' first section cuts the protected cards from the bottom
' to the top
CutDepth = 52 - Val(ohparameter)
s = 1
For j% = CutDepth + 1 To DeckCount
    For z% = 1 To DeckProperties
        ChangedDeck(z%, s) = Deck(z%, j%)
    Next z%
    s = s + 1
Next j%
For k% = 1 To CutDepth
    For z% = 1 To DeckProperties
        ChangedDeck(z%, s) = Deck(z%, k%)
    Next z%
    s = s + 1
Next k%
For m% = 1 To DeckCount
    For z% = 1 To DeckProperties
        Deck(z%, m%) = ChangedDeck(z%, m%)
    Next z%
Next m%
' second section pulls off the desired block amount
' to the bottom of the deck
StackIndex = 1
ChangedDeckIndex = 52
Block = Val(ohparameter)
tempIndex = ChangedDeckIndex + 1 - Block
For i% = tempIndex To ChangedDeckIndex
    For z% = 1 To DeckProperties
        ChangedDeck(z%, i%) = Deck(z%, StackIndex + i% - tempIndex)
    Next z%
Next i%
ChangedDeckIndex = ChangedDeckIndex - Block
StackIndex = StackIndex + Block
' third section shuffles the remainder of the deck
Do While StackIndex < 53
    Block = Int((Rnd ^ 2) * 10) + 1
    If Block > 53 - StackIndex Then
        Block = 52 - StackIndex
    End If
    tempIndex = ChangedDeckIndex + 1 - Block
    For i% = tempIndex To ChangedDeckIndex
        For z% = 1 To DeckProperties
            ChangedDeck(z%, i%) = Deck(z%, StackIndex + i% - tempIndex)
        Next z%
    Next i%
    ChangedDeckIndex = ChangedDeckIndex - Block
    StackIndex = StackIndex + Block
Loop
For m% = 1 To DeckCount
    For z% = 1 To DeckProperties
        Deck(z%, m%) = ChangedDeck(z%, m%)
    Next z%
Next m%
ShowCards
End Sub

Public Sub OHShuffleTop(ohparameter)
' first section pulls off the desired block amount
' to the bottom of the deck
StackIndex = 1
ChangedDeckIndex = 52
Block = Val(ohparameter)
tempIndex = ChangedDeckIndex + 1 - Block
For i% = tempIndex To ChangedDeckIndex
    For z% = 1 To DeckProperties
        ChangedDeck(z%, i%) = Deck(z%, StackIndex + i% - tempIndex)
    Next z%
Next i%
ChangedDeckIndex = ChangedDeckIndex - Block
StackIndex = StackIndex + Block
' second section shuffles the remainder of the deck
Do While StackIndex < 53
    Block = Int((Rnd ^ 2) * 10) + 1
    If Block > 53 - StackIndex Then
        Block = 52 - StackIndex
    End If
    tempIndex = ChangedDeckIndex + 1 - Block
    For i% = tempIndex To ChangedDeckIndex
        For z% = 1 To DeckProperties
            ChangedDeck(z%, i%) = Deck(z%, StackIndex + i% - tempIndex)
        Next z%
    Next i%
    ChangedDeckIndex = ChangedDeckIndex - Block
    StackIndex = StackIndex + Block
Loop
For m% = 1 To DeckCount
    For z% = 1 To DeckProperties
        Deck(z%, m%) = ChangedDeck(z%, m%)
    Next z%
Next m%
' third section cuts the protected cards from the bottom
' to the top
CutDepth = 52 - ohparameter
i = 1
For j% = CutDepth + 1 To DeckCount
    For z% = 1 To DeckProperties
        ChangedDeck(z%, i) = Deck(z%, j%)
    Next z%
    i = i + 1
Next j%
For k% = 1 To CutDepth
    For z% = 1 To DeckProperties
        ChangedDeck(z%, i) = Deck(z%, k%)
    Next z%
    i = i + 1
Next k%
For m% = 1 To DeckCount
    For z% = 1 To DeckProperties
        Deck(z%, m%) = ChangedDeck(z%, m%)
    Next z%
Next m%
ShowCards
End Sub

Private Sub OutFaro()
For i% = 1 To 26
    For z% = 1 To DeckProperties
        ChangedDeck(z%, 2 * i% - 1) = Deck(z%, i%)
        ChangedDeck(z%, 2 * i%) = Deck(z%, i% + 26)
    Next z%
Next i%
For m% = 1 To DeckCount
    For z% = 1 To DeckProperties
        Deck(z%, m%) = ChangedDeck(z%, m%)
    Next z%
Next m%
ShowCards
End Sub

Private Sub OutFaroReverse()
ReverseBlock (26)
For i% = 1 To 26
    For z% = 1 To DeckProperties
        ChangedDeck(z%, 2 * i% - 1) = Deck(z%, i%)
        ChangedDeck(z%, 2 * i%) = Deck(z%, i% + 26)
    Next z%
Next i%
For m% = 1 To DeckCount
    For z% = 1 To DeckProperties
        Deck(z%, m%) = ChangedDeck(z%, m%)
    Next z%
Next m%
ShowCards
End Sub




Public Sub MoveCardButton_Click()
If MoveCardFromTextBox.Text = Empty Or _
    MoveCardToTextBox.Text = Empty Or _
    Not IsNumeric(MoveCardFromTextBox.Text) Or _
    Not IsNumeric(MoveCardToTextBox.Text) Then
    MoveCardFromTextBox.Text = Empty
    MoveCardToTextBox.Text = Empty
    MsgBox "Please enter a valid card position in the" & Chr(13) _
        & "'From' and 'To' Input Boxes"
    Exit Sub
End If
If MoveCardInverseCheck Then
    If MoveCardReverseCheck Then
        Call InverseMoveCardReverse(Val(MoveCardFromTextBox.Text), Val(MoveCardToTextBox.Text))
        'SessionRecord
        If SessionRecordMode Then
            SessionCommand = "InverseMoveCardReverse(" & Val(MoveCardFromTextBox.Text) _
            & ", " & Val(MoveCardToTextBox.Text) & ")"
            SessionListBox.AddItem SessionCommand
            SessionStatusUpdate (0)
        End If
    Else
        Call InverseMoveCard(Val(MoveCardFromTextBox.Text), Val(MoveCardToTextBox.Text))
        'SessionRecord
        If SessionRecordMode Then
            SessionCommand = "InverseMoveCard(" & Val(MoveCardFromTextBox.Text) _
            & ", " & Val(MoveCardToTextBox.Text) & ")"
            SessionListBox.AddItem SessionCommand
            SessionStatusUpdate (0)
        End If
    End If
Else
    If MoveCardReverseCheck Then
        Call MoveCardReverse(Val(MoveCardFromTextBox.Text), Val(MoveCardToTextBox.Text))
        'SessionRecord
        If SessionRecordMode Then
            SessionCommand = "MoveCardReverse(" & Val(MoveCardFromTextBox.Text) _
            & ", " & Val(MoveCardToTextBox.Text) & ")"
            SessionListBox.AddItem SessionCommand
            SessionStatusUpdate (0)
        End If
    Else
        Call MoveCard(Val(MoveCardFromTextBox.Text), Val(MoveCardToTextBox.Text))
        'SessionRecord
        If SessionRecordMode Then
            SessionCommand = "MoveCard(" & Val(MoveCardFromTextBox.Text) _
            & ", " & Val(MoveCardToTextBox.Text) & ")"
            SessionListBox.AddItem SessionCommand
            SessionStatusUpdate (0)
        End If
    End If
End If
End Sub

Private Sub MoveCard(fromcard, tocard)
fromCardParam = Val(fromcard)
toCardParam = Val(tocard)
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
ShowCards
End Sub

Private Sub InverseMoveCard(fromcard, tocard)
fromCardParam = Val(tocard)
toCardParam = Val(fromcard)
'the "to" and "from" are interchanged for the Inverse request
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
ShowCards
End Sub


Private Sub MoveCardReverse(fromcard, tocard)
fromCardParam = Val(fromcard)
toCardParam = Val(tocard)
Deck(6, fromCardParam) = Not Deck(6, fromCardParam)
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
ShowCards
End Sub

Private Sub InverseMoveCardReverse(fromcard, tocard)
fromCardParam = Val(tocard)
toCardParam = Val(fromcard)
'the "to" and "from" are interchanged for the Inverse request
Deck(6, fromCardParam) = Not Deck(6, fromCardParam)
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
ShowCards
End Sub


Public Sub MoveCardFromTextBox_LostFocus()
If MoveCardFromTextBox.Text <> Empty And _
    (Not IsNumeric(MoveCardFromTextBox.Text) Or _
    Val(MoveCardFromTextBox.Text) < 1 Or _
    Val(MoveCardFromTextBox.Text) > 52) Then
    MoveCardFromTextBox.Text = Empty
    MsgBox "Please enter a valid card position (1 to 52)" & Chr(13) _
        & "in the 'From' Input Box"
    MoveCardFromTextBox.SetFocus
    Exit Sub
End If
End Sub

Public Sub MoveCardToTextBox_LostFocus()
If MoveCardToTextBox.Text <> Empty And _
    (Not IsNumeric(MoveCardToTextBox.Text) Or _
    Val(MoveCardToTextBox.Text) < 1 Or _
    Val(MoveCardToTextBox.Text) > 52) Then
    MoveCardToTextBox.Text = Empty
    MsgBox "Please enter a valid card position (1 to 52)" & Chr(13) _
        & "in the 'To' Input Box"
    MoveCardToTextBox.SetFocus
    Exit Sub
End If
End Sub

Public Sub OutFaroButton_Click()
If OutFaroStandardOption.Value = False And _
    (Not IsNumeric(OutFaroFromTopTextBox.Text) Or _
    Val(OutFaroFromTopTextBox.Text) < 1 Or _
    Val(OutFaroFromTopTextBox.Text) > 52) Then
    OutFaroFromTopTextBox.Text = Empty
    MsgBox "Please enter a valid card position (1 to 52)" & Chr(13) _
        & "in the 'Out Faro: From Top' Input Box"
    Exit Sub
End If
If OutFaroSpecialOption.Value = True And _
    OutFaroStartWeaveTopOption.Value = False And _
    OutFaroStartWeaveBottomOption.Value = False Then
    MsgBox "Please select an 'Out Faro: Start Weave' option"
    Exit Sub
End If
If OutFaroStandardOption.Value = False And _
    OutFaroInteriorTextBox.Text <> Empty And _
    (Not IsNumeric(OutFaroInteriorTextBox.Text) Or _
    (Val(OutFaroInteriorTextBox.Text) + _
        Val(OutFaroFromTopTextBox.Text)) > 52) Then
    OutFaroInteriorTextBox.Text = Empty
    MsgBox "Please enter a valid card position in the" & Chr(13) _
        & "'Out Faro: Interior Position' Input Box" & Chr(13) & Chr(13) _
        & "The sum of 'From Top' and 'Interior Position'" & Chr(13) _
        & "must not be greater than 52."
    Exit Sub
End If

If OutFaroInverseCheck Then
    If OutFaroStandardOption.Value = True Then
        If OutFaroReverseCheck Then
            InverseOutFaroReverse
            'SessionRecord
            If SessionRecordMode Then
                SessionCommand = "InverseOutFaroReverse"
                SessionListBox.AddItem SessionCommand
                SessionStatusUpdate (0)
            End If
        Else
            InverseOutFaro
            'SessionRecord
            If SessionRecordMode Then
                SessionCommand = "InverseOutFaro"
                SessionListBox.AddItem SessionCommand
                SessionStatusUpdate (0)
            End If
        End If
    ElseIf OutFaroStartWeaveTopOption.Value = True Then
        If OutFaroReverseCheck Then
            Call InverseOutFaroSpecialTopReverse(Val(OutFaroFromTopTextBox.Text), Val(OutFaroInteriorTextBox.Text))
            'SessionRecord
            If SessionRecordMode Then
                If OutFaroInteriorTextBox.Text <> Empty Then
                    SessionCommand = "InverseOutFaroSpecialTopReverse(" & Val(OutFaroFromTopTextBox.Text) _
                        & ", " & Val(OutFaroInteriorTextBox.Text) & ")"
                Else
                    SessionCommand = "InverseOutFaroSpecialTopReverse(" & Val(OutFaroFromTopTextBox.Text) _
                        & ", 1)"
                End If
                SessionListBox.AddItem SessionCommand
                SessionStatusUpdate (0)
            End If
        Else
            Call InverseOutFaroSpecialTop(Val(OutFaroFromTopTextBox.Text), Val(OutFaroInteriorTextBox.Text))
            'SessionRecord
            If SessionRecordMode Then
                If OutFaroInteriorTextBox.Text <> Empty Then
                    SessionCommand = "InverseOutFaroSpecialTop(" & Val(OutFaroFromTopTextBox.Text) _
                        & ", " & Val(OutFaroInteriorTextBox.Text) & ")"
                Else
                    SessionCommand = "InverseOutFaroSpecialTop(" & Val(OutFaroFromTopTextBox.Text) _
                        & ", 1)"
                End If
                SessionListBox.AddItem SessionCommand
                SessionStatusUpdate (0)
            End If
        End If
    ElseIf OutFaroStartWeaveBottomOption.Value = True Then
        If OutFaroReverseCheck Then
            Call InverseOutFaroSpecialBottomReverse(Val(OutFaroFromTopTextBox.Text), Val(OutFaroInteriorTextBox.Text))
            'SessionRecord
            If SessionRecordMode Then
                If OutFaroInteriorTextBox.Text <> Empty Then
                    SessionCommand = "InverseOutFaroSpecialBottomReverse(" & Val(OutFaroFromTopTextBox.Text) _
                        & ", " & Val(OutFaroInteriorTextBox.Text) & ")"
                Else
                    SessionCommand = "InverseOutFaroSpecialBottomReverse(" & Val(OutFaroFromTopTextBox.Text) _
                        & ", 1)"
                End If
                SessionListBox.AddItem SessionCommand
                SessionStatusUpdate (0)
            End If
        Else
            Call InverseOutFaroSpecialBottom(Val(OutFaroFromTopTextBox.Text), Val(OutFaroInteriorTextBox.Text))
            'SessionRecord
            If SessionRecordMode Then
                If OutFaroInteriorTextBox.Text <> Empty Then
                    SessionCommand = "InverseOutFaroSpecialBottom(" & Val(OutFaroFromTopTextBox.Text) _
                        & ", " & Val(OutFaroInteriorTextBox.Text) & ")"
                Else
                    SessionCommand = "InverseOutFaroSpecialBottom(" & Val(OutFaroFromTopTextBox.Text) _
                        & ", 1)"
                End If
                SessionListBox.AddItem SessionCommand
                SessionStatusUpdate (0)
            End If
        End If
    End If
Else
    If OutFaroStandardOption.Value = True Then
        If OutFaroReverseCheck Then
            OutFaroReverse
            'SessionRecord
            If SessionRecordMode Then
                SessionCommand = "OutFaroReverse"
                SessionListBox.AddItem SessionCommand
                SessionStatusUpdate (0)
            End If
        Else
            OutFaro
            'SessionRecord
            If SessionRecordMode Then
                SessionCommand = "OutFaro"
                SessionListBox.AddItem SessionCommand
                SessionStatusUpdate (0)
            End If
        End If
    ElseIf OutFaroStartWeaveTopOption.Value = True Then
        If OutFaroReverseCheck Then
            Call OutFaroSpecialTopReverse(Val(OutFaroFromTopTextBox.Text), Val(OutFaroInteriorTextBox.Text))
            'SessionRecord
            If SessionRecordMode Then
                If OutFaroInteriorTextBox.Text <> Empty Then
                    SessionCommand = "OutFaroSpecialTopReverse(" & Val(OutFaroFromTopTextBox.Text) _
                        & ", " & Val(OutFaroInteriorTextBox.Text) & ")"
                Else
                    SessionCommand = "OutFaroSpecialTopReverse(" & Val(OutFaroFromTopTextBox.Text) _
                        & ", 1)"
                End If
                SessionListBox.AddItem SessionCommand
                SessionStatusUpdate (0)
            End If
        Else
            Call OutFaroSpecialTop(Val(OutFaroFromTopTextBox.Text), Val(OutFaroInteriorTextBox.Text))
            'SessionRecord
            If SessionRecordMode Then
                If OutFaroInteriorTextBox.Text <> Empty Then
                    SessionCommand = "OutFaroSpecialTop(" & Val(OutFaroFromTopTextBox.Text) _
                        & ", " & Val(OutFaroInteriorTextBox.Text) & ")"
                Else
                    SessionCommand = "OutFaroSpecialTop(" & Val(OutFaroFromTopTextBox.Text) _
                        & ", 1)"
                End If
                SessionListBox.AddItem SessionCommand
                SessionStatusUpdate (0)
            End If
        End If
    ElseIf OutFaroStartWeaveBottomOption.Value = True Then
        If OutFaroReverseCheck Then
            Call OutFaroSpecialBottomReverse(Val(OutFaroFromTopTextBox.Text), Val(OutFaroInteriorTextBox.Text))
            'SessionRecord
            If SessionRecordMode Then
                If OutFaroInteriorTextBox.Text <> Empty Then
                    SessionCommand = "OutFaroSpecialBottomReverse(" & Val(OutFaroFromTopTextBox.Text) _
                        & ", " & Val(OutFaroInteriorTextBox.Text) & ")"
                Else
                    SessionCommand = "OutFaroSpecialBottomReverse(" & Val(OutFaroFromTopTextBox.Text) _
                        & ", 1)"
                End If
                SessionListBox.AddItem SessionCommand
                SessionStatusUpdate (0)
            End If
        Else
            Call OutFaroSpecialBottom(Val(OutFaroFromTopTextBox.Text), Val(OutFaroInteriorTextBox.Text))
            'SessionRecord
            If SessionRecordMode Then
                If OutFaroInteriorTextBox.Text <> Empty Then
                    SessionCommand = "OutFaroSpecialBottom(" & Val(OutFaroFromTopTextBox.Text) _
                        & ", " & Val(OutFaroInteriorTextBox.Text) & ")"
                Else
                    SessionCommand = "OutFaroSpecialBottom(" & Val(OutFaroFromTopTextBox.Text) _
                        & ", 1)"
                End If
                SessionListBox.AddItem SessionCommand
                SessionStatusUpdate (0)
            End If
        End If
    End If
End If
End Sub

Public Sub OutFaroFromTopTextBox_GotFocus()
    OutFaroSpecialOption.Value = True
End Sub

Private Sub OutFaroSpecialBottom(ofbnumber, ofinumber)
'COMPLETE
ProtectedBlock = Val(ofbnumber)
InteriorCard = Val(ofinumber)
Call CutDeckPrecise(ProtectedBlock, "X")
'The 'from bottom' first requires a cut of the Protected Block.
'Then do the infaro version (with a Inversed deck) to accomplish a
'proper outfaro offset block mesh
InverseDeck
If ProtectedBlock = 26 And InteriorCard = 0 Then
    InFaro
    '(because the deck was Inversed, and the CutPrecise was employed,
    ' I need to do an InFaro to have the resultant OutFaro come out correctly)
ElseIf ProtectedBlock < 26 Then
    If InteriorCard = 0 Then
        'complete
        For i% = 1 To ProtectedBlock
            For z% = 1 To DeckProperties
                ChangedDeck(z%, 2 * i%) = Deck(z%, i%)
                ChangedDeck(z%, 2 * i% - 1) = Deck(z%, i% + ProtectedBlock)
            Next z%
        Next i%
        MeshedBlock = 2 * ProtectedBlock
        For k% = 1 To DeckCount - MeshedBlock
            For z% = 1 To DeckProperties
                ChangedDeck(z%, k% + MeshedBlock) = Deck(z%, k% + MeshedBlock)
            Next z%
        Next k%
        For m% = 1 To DeckCount
            For z% = 1 To DeckProperties
                Deck(z%, m%) = ChangedDeck(z%, m%)
            Next z%
        Next m%
    Else
        InteriorPosition = InteriorCard + ProtectedBlock
        If ProtectedBlock + InteriorPosition <= 52 Then
            For i% = 1 To InteriorPosition - ProtectedBlock
            'added the "- 1" from InFaro solution
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, i%) = Deck(z%, ProtectedBlock + i%)
                Next z%
            Next i%
            MeshedBlockPosition = InteriorPosition - ProtectedBlock
            'added the "- 1" from InFaro solution
            For i% = 1 To ProtectedBlock
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, MeshedBlockPosition + (2 * i% - 1)) = _
                        Deck(z%, i%)
                    ChangedDeck(z%, MeshedBlockPosition + (2 * i%)) = _
                        Deck(z%, i% + InteriorPosition)
                        'added the "- 1" from InFaro solution
                Next z%
            Next i%
            For i% = ProtectedBlock * 2 + InteriorPosition - ProtectedBlock + 1 To DeckCount
            'took out a "+ 1" right before the "To" from InFaro solution
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, i%) = Deck(z%, i%)
                Next z%
            Next i%
            For m% = 1 To DeckCount
                For z% = 1 To DeckProperties
                    Deck(z%, m%) = ChangedDeck(z%, m%)
                Next z%
            Next m%
        Else
            For i% = 1 To InteriorPosition - ProtectedBlock
            'added the "- 1" from InFaro solution
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, i%) = Deck(z%, ProtectedBlock + i%)
                Next z%
            Next i%
            MeshedBlockPosition = InteriorPosition - ProtectedBlock
            'added the "- 1" from InFaro solution
            MeshedBlock = 2 * (DeckCount - InteriorPosition)
            'added the "+ 1" from InFaro solution (inside parens after Interior Position)
            For i% = 1 To DeckCount - InteriorPosition
            'added the "+ 1" from InFaro solution
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, MeshedBlockPosition + (2 * i% - 1)) = _
                        Deck(z%, i%)
                    ChangedDeck(z%, MeshedBlockPosition + (2 * i%)) = _
                        Deck(z%, i% + InteriorPosition)
                        'added the "- 1" from InFaro solution (after Interior Position)
                Next z%
            Next i%
            For i% = 1 To ProtectedBlock - (DeckCount - InteriorPosition)
            'added the "+ 1" from InFaro solution (inside parens after Interior Position)
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, MeshedBlockPosition + MeshedBlock + i%) = _
                        Deck(z%, DeckCount - InteriorPosition + i%)
                        'added the "+ 1" from InFaro solution after Interior Position
                Next z%
            Next i%
            For m% = 1 To DeckCount
                For z% = 1 To DeckProperties
                    Deck(z%, m%) = ChangedDeck(z%, m%)
                Next z%
            Next m%
        End If
    End If
ElseIf ProtectedBlock >= 26 Then
    If InteriorCard = 0 Then
        For i% = 1 To DeckCount - ProtectedBlock
            For z% = 1 To DeckProperties
                ChangedDeck(z%, 2 * i%) = Deck(z%, i%)
                ChangedDeck(z%, 2 * i% - 1) = Deck(z%, i% + ProtectedBlock)
            Next z%
        Next i%
        MeshedBlock = 2 * (DeckCount - ProtectedBlock)
        For k% = 1 To DeckCount - MeshedBlock
            For z% = 1 To DeckProperties
                ChangedDeck(z%, k% + MeshedBlock) = Deck(z%, k% + DeckCount - ProtectedBlock)
            Next z%
        Next k%
        For m% = 1 To DeckCount
            For z% = 1 To DeckProperties
                Deck(z%, m%) = ChangedDeck(z%, m%)
            Next z%
        Next m%
    Else
        InteriorPosition = InteriorCard + ProtectedBlock
        If ProtectedBlock + InteriorPosition <= 52 Then
            For i% = 1 To InteriorPosition - ProtectedBlock
            'added the "- 1" from InFaro solution
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, i%) = Deck(z%, ProtectedBlock + i%)
                Next z%
            Next i%
            MeshedBlockPosition = InteriorPosition - ProtectedBlock
            'added the "- 1" from InFaro solution
            For i% = 1 To ProtectedBlock
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, MeshedBlockPosition + (2 * i% - 1)) = _
                        Deck(z%, i%)
                    ChangedDeck(z%, MeshedBlockPosition + (2 * i%)) = _
                        Deck(z%, i% + InteriorPosition)
                        'added the "- 1" from InFaro solution (inside parens after Interior Position)
                Next z%
            Next i%
            For i% = ProtectedBlock * 2 + InteriorPosition - ProtectedBlock + 1 To DeckCount
            'took out a "+ 1" right before the "To" from InFaro solution
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, i%) = Deck(z%, i%)
                Next z%
            Next i%
            For m% = 1 To DeckCount
                For z% = 1 To DeckProperties
                    Deck(z%, m%) = ChangedDeck(z%, m%)
                Next z%
            Next m%
        Else
            For i% = 1 To InteriorPosition - ProtectedBlock
            'added the "- 1" from InFaro solution
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, i%) = Deck(z%, ProtectedBlock + i%)
                Next z%
            Next i%
            MeshedBlockPosition = InteriorPosition - ProtectedBlock
            'added the "- 1" from InFaro solution
            MeshedBlock = 2 * (DeckCount - InteriorPosition)
            'added the "+ 1" from InFaro solution inside parens after InteriorPosition)
            For i% = 1 To DeckCount - InteriorPosition
            'added the "+ 1" from InFaro solution
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, MeshedBlockPosition + (2 * i% - 1)) = _
                        Deck(z%, i%)
                    ChangedDeck(z%, MeshedBlockPosition + (2 * i%)) = _
                        Deck(z%, i% + InteriorPosition)
                        'added the "- 1" from InFaro solution inside parens after Interior Position
                Next z%
            Next i%
            For i% = 1 To ProtectedBlock - (DeckCount - InteriorPosition)
            'added the "+ 1" from InFaro solution inside parens after InteriorPosition)
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, MeshedBlockPosition + MeshedBlock + i%) = _
                        Deck(z%, DeckCount - InteriorPosition + i%)
                        'added the "+ 1" from InFaro solution after InteriorPosition
                Next z%
            Next i%
            For m% = 1 To DeckCount
                For z% = 1 To DeckProperties
                    Deck(z%, m%) = ChangedDeck(z%, m%)
                Next z%
            Next m%
        End If
    End If
End If
InverseDeck
ShowCards
End Sub

Private Sub OutFaroSpecialBottomReverse(ofbnumber, ofinumber)
'COMPLETE
ProtectedBlock = Val(ofbnumber)
InteriorCard = Val(ofinumber)
Call CutDeckPrecise(ProtectedBlock, "X")
'The 'from bottom' first requires a cut of the Protected Block.
'Then do the infaro version (with a Inversed deck) to accomplish a
'proper outfaro offset block mesh
InverseDeck
If ProtectedBlock = 26 And InteriorCard = 0 Then
    InFaroReverse
    '(because the deck was Inversed, and the CutPrecise was employed,
    ' I need to do an InFaro to have the resultant OutFaro come out correctly)
ElseIf ProtectedBlock < 26 Then
    If InteriorCard = 0 Then
        ReverseBlock (ProtectedBlock)
        'complete
        For i% = 1 To ProtectedBlock
            For z% = 1 To DeckProperties
                ChangedDeck(z%, 2 * i%) = Deck(z%, i%)
                ChangedDeck(z%, 2 * i% - 1) = Deck(z%, i% + ProtectedBlock)
            Next z%
        Next i%
        MeshedBlock = 2 * ProtectedBlock
        For k% = 1 To DeckCount - MeshedBlock
            For z% = 1 To DeckProperties
                ChangedDeck(z%, k% + MeshedBlock) = Deck(z%, k% + MeshedBlock)
            Next z%
        Next k%
        For m% = 1 To DeckCount
            For z% = 1 To DeckProperties
                Deck(z%, m%) = ChangedDeck(z%, m%)
            Next z%
        Next m%
    Else
        InteriorPosition = InteriorCard + ProtectedBlock
        If ProtectedBlock + InteriorPosition <= 52 Then
            ReverseBlock (ProtectedBlock)
            For i% = 1 To InteriorPosition - ProtectedBlock
            'added the "- 1" from InFaro solution
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, i%) = Deck(z%, ProtectedBlock + i%)
                Next z%
            Next i%
            MeshedBlockPosition = InteriorPosition - ProtectedBlock
            'added the "- 1" from InFaro solution
            For i% = 1 To ProtectedBlock
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, MeshedBlockPosition + (2 * i% - 1)) = _
                        Deck(z%, i%)
                    ChangedDeck(z%, MeshedBlockPosition + (2 * i%)) = _
                        Deck(z%, i% + InteriorPosition)
                        'added the "- 1" from InFaro solution
                Next z%
            Next i%
            For i% = ProtectedBlock * 2 + InteriorPosition - ProtectedBlock + 1 To DeckCount
            'took out a "+ 1" right before the "To" from InFaro solution
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, i%) = Deck(z%, i%)
                Next z%
            Next i%
            For m% = 1 To DeckCount
                For z% = 1 To DeckProperties
                    Deck(z%, m%) = ChangedDeck(z%, m%)
                Next z%
            Next m%
        Else
            ReverseBlock (ProtectedBlock)
            For i% = 1 To InteriorPosition - ProtectedBlock
            'added the "- 1" from InFaro solution
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, i%) = Deck(z%, ProtectedBlock + i%)
                Next z%
            Next i%
            MeshedBlockPosition = InteriorPosition - ProtectedBlock
            'added the "- 1" from InFaro solution
            MeshedBlock = 2 * (DeckCount - InteriorPosition)
            'added the "+ 1" from InFaro solution (inside parens after Interior Position)
            For i% = 1 To DeckCount - InteriorPosition
            'added the "+ 1" from InFaro solution
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, MeshedBlockPosition + (2 * i% - 1)) = _
                        Deck(z%, i%)
                    ChangedDeck(z%, MeshedBlockPosition + (2 * i%)) = _
                        Deck(z%, i% + InteriorPosition)
                        'added the "- 1" from InFaro solution (after Interior Position)
                Next z%
            Next i%
            For i% = 1 To ProtectedBlock - (DeckCount - InteriorPosition)
            'added the "+ 1" from InFaro solution (inside parens after Interior Position)
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, MeshedBlockPosition + MeshedBlock + i%) = _
                        Deck(z%, DeckCount - InteriorPosition + i%)
                        'added the "+ 1" from InFaro solution after Interior Position
                Next z%
            Next i%
            For m% = 1 To DeckCount
                For z% = 1 To DeckProperties
                    Deck(z%, m%) = ChangedDeck(z%, m%)
                Next z%
            Next m%
        End If
    End If
ElseIf ProtectedBlock >= 26 Then
    If InteriorCard = 0 Then
        ReverseBlock (ProtectedBlock)
        For i% = 1 To DeckCount - ProtectedBlock
            For z% = 1 To DeckProperties
                ChangedDeck(z%, 2 * i%) = Deck(z%, i%)
                ChangedDeck(z%, 2 * i% - 1) = Deck(z%, i% + ProtectedBlock)
            Next z%
        Next i%
        MeshedBlock = 2 * (DeckCount - ProtectedBlock)
        For k% = 1 To DeckCount - MeshedBlock
            For z% = 1 To DeckProperties
                ChangedDeck(z%, k% + MeshedBlock) = Deck(z%, k% + DeckCount - ProtectedBlock)
            Next z%
        Next k%
        For m% = 1 To DeckCount
            For z% = 1 To DeckProperties
                Deck(z%, m%) = ChangedDeck(z%, m%)
            Next z%
        Next m%
    Else
        InteriorPosition = InteriorCard + ProtectedBlock
        If ProtectedBlock + InteriorPosition <= 52 Then
            ReverseBlock (ProtectedBlock)
            For i% = 1 To InteriorPosition - ProtectedBlock
            'added the "- 1" from InFaro solution
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, i%) = Deck(z%, ProtectedBlock + i%)
                Next z%
            Next i%
            MeshedBlockPosition = InteriorPosition - ProtectedBlock
            'added the "- 1" from InFaro solution
            For i% = 1 To ProtectedBlock
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, MeshedBlockPosition + (2 * i% - 1)) = _
                        Deck(z%, i%)
                    ChangedDeck(z%, MeshedBlockPosition + (2 * i%)) = _
                        Deck(z%, i% + InteriorPosition)
                        'added the "- 1" from InFaro solution (inside parens after Interior Position)
                Next z%
            Next i%
            For i% = ProtectedBlock * 2 + InteriorPosition - ProtectedBlock + 1 To DeckCount
            'took out a "+ 1" right before the "To" from InFaro solution
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, i%) = Deck(z%, i%)
                Next z%
            Next i%
            For m% = 1 To DeckCount
                For z% = 1 To DeckProperties
                    Deck(z%, m%) = ChangedDeck(z%, m%)
                Next z%
            Next m%
        Else
            ReverseBlock (ProtectedBlock)
            For i% = 1 To InteriorPosition - ProtectedBlock
            'added the "- 1" from InFaro solution
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, i%) = Deck(z%, ProtectedBlock + i%)
                Next z%
            Next i%
            MeshedBlockPosition = InteriorPosition - ProtectedBlock
            'added the "- 1" from InFaro solution
            MeshedBlock = 2 * (DeckCount - InteriorPosition)
            'added the "+ 1" from InFaro solution inside parens after InteriorPosition)
            For i% = 1 To DeckCount - InteriorPosition
            'added the "+ 1" from InFaro solution
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, MeshedBlockPosition + (2 * i% - 1)) = _
                        Deck(z%, i%)
                    ChangedDeck(z%, MeshedBlockPosition + (2 * i%)) = _
                        Deck(z%, i% + InteriorPosition)
                        'added the "- 1" from InFaro solution inside parens after Interior Position
                Next z%
            Next i%
            For i% = 1 To ProtectedBlock - (DeckCount - InteriorPosition)
            'added the "+ 1" from InFaro solution inside parens after InteriorPosition)
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, MeshedBlockPosition + MeshedBlock + i%) = _
                        Deck(z%, DeckCount - InteriorPosition + i%)
                        'added the "+ 1" from InFaro solution after InteriorPosition
                Next z%
            Next i%
            For m% = 1 To DeckCount
                For z% = 1 To DeckProperties
                    Deck(z%, m%) = ChangedDeck(z%, m%)
                Next z%
            Next m%
        End If
    End If
End If
InverseDeck
ShowCards
End Sub


Private Sub OutFaroSpecialTop(oftnumber, ofinumber)
'COMPLETE
ProtectedBlock = Val(oftnumber)
InteriorCard = Val(ofinumber)
If ProtectedBlock = 26 And InteriorCard = Empty Then
    OutFaro
ElseIf ProtectedBlock < 26 Then
    If InteriorCard = Empty Then
        For i% = 1 To ProtectedBlock
            For z% = 1 To DeckProperties
                ChangedDeck(z%, 2 * i% - 1) = Deck(z%, i%)
                ChangedDeck(z%, 2 * i%) = Deck(z%, i% + ProtectedBlock)
            Next z%
        Next i%
        MeshedBlock = 2 * ProtectedBlock
        For k% = 1 To DeckCount - MeshedBlock
            For z% = 1 To DeckProperties
                ChangedDeck(z%, k% + MeshedBlock) = Deck(z%, k% + MeshedBlock)
            Next z%
        Next k%
        For m% = 1 To DeckCount
            For z% = 1 To DeckProperties
                Deck(z%, m%) = ChangedDeck(z%, m%)
            Next z%
        Next m%
    Else
        InteriorPosition = InteriorCard + ProtectedBlock
        If ProtectedBlock + InteriorPosition <= 52 Then
            For i% = 1 To InteriorPosition - ProtectedBlock - 1
            'added the "- 1" from InFaro solution
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, i%) = Deck(z%, ProtectedBlock + i%)
                Next z%
            Next i%
            MeshedBlockPosition = InteriorPosition - ProtectedBlock - 1
            'added the "- 1" from InFaro solution
            For i% = 1 To ProtectedBlock
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, MeshedBlockPosition + (2 * i% - 1)) = _
                        Deck(z%, i%)
                    ChangedDeck(z%, MeshedBlockPosition + (2 * i%)) = _
                        Deck(z%, i% + InteriorPosition - 1)
                        'added the "- 1" from InFaro solution
                Next z%
            Next i%
            For i% = ProtectedBlock * 2 + InteriorPosition - ProtectedBlock To DeckCount
            'took out a "+ 1" right before the "To" from InFaro solution
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, i%) = Deck(z%, i%)
                Next z%
            Next i%
            For m% = 1 To DeckCount
                For z% = 1 To DeckProperties
                    Deck(z%, m%) = ChangedDeck(z%, m%)
                Next z%
            Next m%
        Else
            For i% = 1 To InteriorPosition - ProtectedBlock - 1
            'added the "- 1" from InFaro solution
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, i%) = Deck(z%, ProtectedBlock + i%)
                Next z%
            Next i%
            MeshedBlockPosition = InteriorPosition - ProtectedBlock - 1
            'added the "- 1" from InFaro solution
            MeshedBlock = 2 * (DeckCount - InteriorPosition + 1)
            'added the "+ 1" from InFaro solution
            For i% = 1 To DeckCount - InteriorPosition + 1
            'added the "+ 1" from InFaro solution
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, MeshedBlockPosition + (2 * i% - 1)) = _
                        Deck(z%, i%)
                    ChangedDeck(z%, MeshedBlockPosition + (2 * i%)) = _
                        Deck(z%, i% + InteriorPosition - 1)
                        'added the "- 1" from InFaro solution
                Next z%
            Next i%
            For i% = 1 To ProtectedBlock - (DeckCount - InteriorPosition + 1)
            'added the "+ 1" from InFaro solution
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, MeshedBlockPosition + MeshedBlock + i%) = _
                        Deck(z%, DeckCount - InteriorPosition + 1 + i%)
                        'added the "+ 1" from InFaro solution
                Next z%
            Next i%
            For m% = 1 To DeckCount
                For z% = 1 To DeckProperties
                    Deck(z%, m%) = ChangedDeck(z%, m%)
                Next z%
            Next m%
        End If
    End If
ElseIf ProtectedBlock >= 26 Then
    If InteriorCard = Empty Then
        For i% = 1 To DeckCount - ProtectedBlock
            For z% = 1 To DeckProperties
                ChangedDeck(z%, 2 * i% - 1) = Deck(z%, i%)
                ChangedDeck(z%, 2 * i%) = Deck(z%, i% + ProtectedBlock)
            Next z%
        Next i%
        MeshedBlock = 2 * (DeckCount - ProtectedBlock)
        For k% = 1 To DeckCount - MeshedBlock
            For z% = 1 To DeckProperties
                ChangedDeck(z%, k% + MeshedBlock) = Deck(z%, k% + DeckCount - ProtectedBlock)
            Next z%
        Next k%
        For m% = 1 To DeckCount
            For z% = 1 To DeckProperties
                Deck(z%, m%) = ChangedDeck(z%, m%)
            Next z%
        Next m%
    Else
        InteriorPosition = InteriorCard + ProtectedBlock
        If ProtectedBlock + InteriorPosition <= 52 Then
            For i% = 1 To InteriorPosition - ProtectedBlock - 1
            'added the "- 1" from InFaro solution
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, i%) = Deck(z%, ProtectedBlock + i%)
                Next z%
            Next i%
            MeshedBlockPosition = InteriorPosition - ProtectedBlock - 1
            'added the "- 1" from InFaro solution
            For i% = 1 To ProtectedBlock
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, MeshedBlockPosition + (2 * i% - 1)) = _
                        Deck(z%, i%)
                    ChangedDeck(z%, MeshedBlockPosition + (2 * i%)) = _
                        Deck(z%, i% + InteriorPosition - 1)
                        'added the "- 1" from InFaro solution
                Next z%
            Next i%
            For i% = ProtectedBlock * 2 + InteriorPosition - ProtectedBlock To DeckCount
            'took out a "+ 1" right before the "To" from InFaro solution
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, i%) = Deck(z%, i%)
                Next z%
            Next i%
            For m% = 1 To DeckCount
                For z% = 1 To DeckProperties
                    Deck(z%, m%) = ChangedDeck(z%, m%)
                Next z%
            Next m%
        Else
            For i% = 1 To InteriorPosition - ProtectedBlock - 1
            'added the "- 1" from InFaro solution
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, i%) = Deck(z%, ProtectedBlock + i%)
                Next z%
            Next i%
            MeshedBlockPosition = InteriorPosition - ProtectedBlock - 1
            'added the "- 1" from InFaro solution
            MeshedBlock = 2 * (DeckCount - InteriorPosition + 1)
            'added the "+ 1" from InFaro solution
            For i% = 1 To DeckCount - InteriorPosition + 1
            'added the "+ 1" from InFaro solution
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, MeshedBlockPosition + (2 * i% - 1)) = _
                        Deck(z%, i%)
                    ChangedDeck(z%, MeshedBlockPosition + (2 * i%)) = _
                        Deck(z%, i% + InteriorPosition - 1)
                        'added the "- 1" from InFaro solution
                Next z%
            Next i%
            For i% = 1 To ProtectedBlock - (DeckCount - InteriorPosition + 1)
            'added the "+ 1" from InFaro solution
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, MeshedBlockPosition + MeshedBlock + i%) = _
                        Deck(z%, DeckCount - InteriorPosition + 1 + i%)
                        'added the "+ 1" from InFaro solution
                Next z%
            Next i%
            For m% = 1 To DeckCount
                For z% = 1 To DeckProperties
                    Deck(z%, m%) = ChangedDeck(z%, m%)
                Next z%
            Next m%
        End If
    End If
End If
ShowCards
End Sub

Private Sub OutFaroSpecialTopReverse(oftnumber, ofinumber)
'COMPLETE
ProtectedBlock = Val(oftnumber)
InteriorCard = Val(ofinumber)
If ProtectedBlock = 26 And InteriorCard = Empty Then
    OutFaroReverse
ElseIf ProtectedBlock < 26 Then
    If InteriorCard = Empty Then
        ReverseBlock (ProtectedBlock)
        For i% = 1 To ProtectedBlock
            For z% = 1 To DeckProperties
                ChangedDeck(z%, 2 * i% - 1) = Deck(z%, i%)
                ChangedDeck(z%, 2 * i%) = Deck(z%, i% + ProtectedBlock)
            Next z%
        Next i%
        MeshedBlock = 2 * ProtectedBlock
        For k% = 1 To DeckCount - MeshedBlock
            For z% = 1 To DeckProperties
                ChangedDeck(z%, k% + MeshedBlock) = Deck(z%, k% + MeshedBlock)
            Next z%
        Next k%
        For m% = 1 To DeckCount
            For z% = 1 To DeckProperties
                Deck(z%, m%) = ChangedDeck(z%, m%)
            Next z%
        Next m%
    Else
        InteriorPosition = InteriorCard + ProtectedBlock
        If ProtectedBlock + InteriorPosition <= 52 Then
            ReverseBlock (ProtectedBlock)
            For i% = 1 To InteriorPosition - ProtectedBlock - 1
            'added the "- 1" from InFaro solution
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, i%) = Deck(z%, ProtectedBlock + i%)
                Next z%
            Next i%
            MeshedBlockPosition = InteriorPosition - ProtectedBlock - 1
            'added the "- 1" from InFaro solution
            For i% = 1 To ProtectedBlock
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, MeshedBlockPosition + (2 * i% - 1)) = _
                        Deck(z%, i%)
                    ChangedDeck(z%, MeshedBlockPosition + (2 * i%)) = _
                        Deck(z%, i% + InteriorPosition - 1)
                        'added the "- 1" from InFaro solution
                Next z%
            Next i%
            For i% = ProtectedBlock * 2 + InteriorPosition - ProtectedBlock To DeckCount
            'took out a "+ 1" right before the "To" from InFaro solution
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, i%) = Deck(z%, i%)
                Next z%
            Next i%
            For m% = 1 To DeckCount
                For z% = 1 To DeckProperties
                    Deck(z%, m%) = ChangedDeck(z%, m%)
                Next z%
            Next m%
        Else
            ReverseBlock (ProtectedBlock)
            For i% = 1 To InteriorPosition - ProtectedBlock - 1
            'added the "- 1" from InFaro solution
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, i%) = Deck(z%, ProtectedBlock + i%)
                Next z%
            Next i%
            MeshedBlockPosition = InteriorPosition - ProtectedBlock - 1
            'added the "- 1" from InFaro solution
            MeshedBlock = 2 * (DeckCount - InteriorPosition + 1)
            'added the "+ 1" from InFaro solution
            For i% = 1 To DeckCount - InteriorPosition + 1
            'added the "+ 1" from InFaro solution
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, MeshedBlockPosition + (2 * i% - 1)) = _
                        Deck(z%, i%)
                    ChangedDeck(z%, MeshedBlockPosition + (2 * i%)) = _
                        Deck(z%, i% + InteriorPosition - 1)
                        'added the "- 1" from InFaro solution
                Next z%
            Next i%
            For i% = 1 To ProtectedBlock - (DeckCount - InteriorPosition + 1)
            'added the "+ 1" from InFaro solution
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, MeshedBlockPosition + MeshedBlock + i%) = _
                        Deck(z%, DeckCount - InteriorPosition + 1 + i%)
                        'added the "+ 1" from InFaro solution
                Next z%
            Next i%
            For m% = 1 To DeckCount
                For z% = 1 To DeckProperties
                    Deck(z%, m%) = ChangedDeck(z%, m%)
                Next z%
            Next m%
        End If
    End If
ElseIf ProtectedBlock >= 26 Then
    If InteriorCard = Empty Then
        ReverseBlock (ProtectedBlock)
        For i% = 1 To DeckCount - ProtectedBlock
            For z% = 1 To DeckProperties
                ChangedDeck(z%, 2 * i% - 1) = Deck(z%, i%)
                ChangedDeck(z%, 2 * i%) = Deck(z%, i% + ProtectedBlock)
            Next z%
        Next i%
        MeshedBlock = 2 * (DeckCount - ProtectedBlock)
        For k% = 1 To DeckCount - MeshedBlock
            For z% = 1 To DeckProperties
                ChangedDeck(z%, k% + MeshedBlock) = Deck(z%, k% + DeckCount - ProtectedBlock)
            Next z%
        Next k%
        For m% = 1 To DeckCount
            For z% = 1 To DeckProperties
                Deck(z%, m%) = ChangedDeck(z%, m%)
            Next z%
        Next m%
    Else
        InteriorPosition = InteriorCard + ProtectedBlock
        If ProtectedBlock + InteriorPosition <= 52 Then
            ReverseBlock (ProtectedBlock)
            For i% = 1 To InteriorPosition - ProtectedBlock - 1
            'added the "- 1" from InFaro solution
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, i%) = Deck(z%, ProtectedBlock + i%)
                Next z%
            Next i%
            MeshedBlockPosition = InteriorPosition - ProtectedBlock - 1
            'added the "- 1" from InFaro solution
            For i% = 1 To ProtectedBlock
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, MeshedBlockPosition + (2 * i% - 1)) = _
                        Deck(z%, i%)
                    ChangedDeck(z%, MeshedBlockPosition + (2 * i%)) = _
                        Deck(z%, i% + InteriorPosition - 1)
                        'added the "- 1" from InFaro solution
                Next z%
            Next i%
            For i% = ProtectedBlock * 2 + InteriorPosition - ProtectedBlock To DeckCount
            'took out a "+ 1" right before the "To" from InFaro solution
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, i%) = Deck(z%, i%)
                Next z%
            Next i%
            For m% = 1 To DeckCount
                For z% = 1 To DeckProperties
                    Deck(z%, m%) = ChangedDeck(z%, m%)
                Next z%
            Next m%
        Else
            ReverseBlock (ProtectedBlock)
            For i% = 1 To InteriorPosition - ProtectedBlock - 1
            'added the "- 1" from InFaro solution
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, i%) = Deck(z%, ProtectedBlock + i%)
                Next z%
            Next i%
            MeshedBlockPosition = InteriorPosition - ProtectedBlock - 1
            'added the "- 1" from InFaro solution
            MeshedBlock = 2 * (DeckCount - InteriorPosition + 1)
            'added the "+ 1" from InFaro solution
            For i% = 1 To DeckCount - InteriorPosition + 1
            'added the "+ 1" from InFaro solution
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, MeshedBlockPosition + (2 * i% - 1)) = _
                        Deck(z%, i%)
                    ChangedDeck(z%, MeshedBlockPosition + (2 * i%)) = _
                        Deck(z%, i% + InteriorPosition - 1)
                        'added the "- 1" from InFaro solution
                Next z%
            Next i%
            For i% = 1 To ProtectedBlock - (DeckCount - InteriorPosition + 1)
            'added the "+ 1" from InFaro solution
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, MeshedBlockPosition + MeshedBlock + i%) = _
                        Deck(z%, DeckCount - InteriorPosition + 1 + i%)
                        'added the "+ 1" from InFaro solution
                Next z%
            Next i%
            For m% = 1 To DeckCount
                For z% = 1 To DeckProperties
                    Deck(z%, m%) = ChangedDeck(z%, m%)
                Next z%
            Next m%
        End If
    End If
End If
ShowCards
End Sub


Public Sub OutFaroFromTopTextBox_LostFocus()
If OutFaroFromTopTextBox.Text <> Empty And _
    OutFaroStandardOption.Value = False And _
    (Not IsNumeric(OutFaroFromTopTextBox.Text) Or _
    Val(OutFaroFromTopTextBox.Text) < 1 Or _
    Val(OutFaroFromTopTextBox.Text) > 52) Then
    OutFaroFromTopTextBox.Text = Empty
    MsgBox "Please enter a valid card position (1 to 52)" & Chr(13) _
        & "in the 'Out Faro: From Top' Input Box"
    OutFaroFromTopTextBox.SetFocus
    Exit Sub
End If
End Sub

Public Sub OutFaroInteriorTextBox_GotFocus()
    OutFaroSpecialOption.Value = True
End Sub

Public Sub OutFaroInteriorTextBox_LostFocus()
If OutFaroInteriorTextBox.Text <> Empty And _
    OutFaroStandardOption.Value = False And _
    (Not IsNumeric(OutFaroInteriorTextBox.Text) Or _
    Val(OutFaroInteriorTextBox.Text) < 1 Or _
    Val(OutFaroInteriorTextBox.Text) > 52) Then
    OutFaroInteriorTextBox.Text = Empty
    MsgBox "Please enter a valid card position (1 to 52)" & Chr(13) _
        & "in the 'Out Faro: Interior Position' Input Box" & Chr(13) _
        & "(or leave it blank for a default value of 1)"
    OutFaroInteriorTextBox.SetFocus
    Exit Sub
End If
End Sub



Public Sub OutFaroStandardOption_Click()
    OutFaroStartWeaveTopOption.Value = False
    OutFaroStartWeaveBottomOption.Value = False
    OutFaroFromTopTextBox.Text = Empty
    OutFaroInteriorTextBox.Text = Empty
End Sub




Public Sub OutFaroStartWeaveBottomOption_GotFocus()
    OutFaroSpecialOption.Value = True
End Sub

Public Sub OutFaroStartWeaveTopOption_GotFocus()
    OutFaroSpecialOption.Value = True
End Sub





Public Sub OverhandLocationCombo_GotFocus()
    OverhandProtectOption.Value = True
End Sub

Public Sub OverhandNumberTextBox_GotFocus()
    OverhandProtectOption.Value = True
End Sub

Public Sub OverhandRandomOption_Click()
    OverhandNumberTextBox.Text = Empty
    OverhandLocationCombo.Text = Empty
End Sub

Public Sub OverhandShuffleButton_Click()
If OverhandProtectOption.Value = True And _
    (Not IsNumeric(OverhandNumberTextBox.Text) Or _
    Val(OverhandNumberTextBox.Text) < 1 Or _
    Val(OverhandNumberTextBox.Text) > 52) Then
    OverhandNumberTextBox.Text = Empty
    MsgBox "Please enter a valid number of cards (1 to 52)" & Chr(13) _
        & "in the 'Overhand: Number' Input Box"
    Exit Sub
End If
If OverhandProtectOption = True And _
    OverhandLocationCombo.ListIndex = -1 Then
    OverhandLocationCombo.Text = "Select area"
    MsgBox "Please select a valid response from" & Chr(13) _
        & "the 'Overhand: Location' Dropdown Box"
    Exit Sub
End If
If OverhandRandomOption.Value = True Then
    OHShuffle
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "OHShuffle"
        SessionListBox.AddItem SessionCommand
        SessionStatusUpdate (0)
    End If
ElseIf OverhandProtectOption.Value = True And _
    OverhandLocationCombo.ListIndex = 0 Then
    OHShuffleTop (Val(OverhandNumberTextBox.Text))
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "OHShuffleTop(" & OverhandNumberTextBox.Text & ")"
        SessionListBox.AddItem SessionCommand
        SessionStatusUpdate (0)
    End If
ElseIf OverhandProtectOption.Value = True And _
    OverhandLocationCombo.ListIndex = 1 Then
    OHShuffleBottom (Val(OverhandNumberTextBox.Text))
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "OHShuffleBottom(" & OverhandNumberTextBox.Text & ")"
        SessionListBox.AddItem SessionCommand
        SessionStatusUpdate (0)
    End If
End If
End Sub
Private Sub ResetCurrentDeck()
For k% = 1 To 52
    For m% = 1 To 52
        If Deck(1, m%) = k% Then
            For p% = 1 To DeckProperties
                ChangedDeck(p%, k%) = Deck(p%, m%)
            Next p%
        End If
    Next m%
Next k%
For m% = 1 To DeckCount
    For z% = 1 To DeckProperties
        Deck(z%, m%) = ChangedDeck(z%, m%)
    Next z%
Next m%
'For p% = 1 To DeckCount
'    Deck(6, p%) = False
'    ' a deck reset should turn all the cards face up again
'Next p%
ShowCards
End Sub


Private Sub RefreshView_Click()
ShowCards
End Sub

Public Sub ResetCurrentDeckButton_Click()
ResetCurrentDeck
'SessionRecord
If SessionRecordMode Then
    SessionCommand = "ResetCurrentDeck"
    SessionListBox.AddItem SessionCommand
    SessionStatusUpdate (0)
End If
End Sub


Private Sub ReverseSelections_Click()
For i% = 1 To DeckCount
    If Deck(4, i%) = "Selected" Then
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
        ShowPiles
    End If
ElseIf PokerCardsDealt = 1 Then
    ShowDeal
Else
    ShowCards
End If
End Sub

Private Sub RiffleReverseTopCheck_Click()
If RiffleReverseTopCheck.Value = 1 Then
    If RiffleReverseBottomCheck.Value = 1 Then
        RiffleReverseBottomCheck.Value = 0
    End If
End If
End Sub

Private Sub RiffleReverseBottomCheck_Click()
If RiffleReverseBottomCheck.Value = 1 Then
    If RiffleReverseTopCheck.Value = 1 Then
        RiffleReverseTopCheck.Value = 0
    End If
End If
End Sub

Private Sub CutReverseTopCheck_Click()
If CutReverseTopCheck.Value = 1 Then
    If CutReverseBottomCheck.Value = 1 Then
        CutReverseBottomCheck.Value = 0
    End If
End If
End Sub

Private Sub CutReverseBottomCheck_Click()
If CutReverseBottomCheck.Value = 1 Then
    If CutReverseTopCheck.Value = 1 Then
        CutReverseTopCheck.Value = 0
    End If
End If
End Sub

Private Sub SessionInsertMacro_Click()
InsertMacroError = False
'the call is to frmMain because of the dlgCommonDialog object being there
Call frmMain.SessionInsertMacro
End Sub

Public Sub SessionRecordToggle_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If SessionRecordMode Then
    SessionRecordToggle(0).Visible = False
    SessionRecordToggle(1).Visible = False
    SessionRecordToggle(2).Visible = False
    SessionRecordToggle(3).Visible = True
Else
    SessionRecordToggle(0).Visible = False
    SessionRecordToggle(1).Visible = True
    SessionRecordToggle(2).Visible = False
    SessionRecordToggle(3).Visible = False
End If
End Sub

Public Sub SessionRecordToggle_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If SessionRecordMode Then
    SessionRecordToggle(0).Visible = True
    SessionRecordToggle(1).Visible = False
    SessionRecordToggle(2).Visible = False
    SessionRecordToggle(3).Visible = False
Else
    SessionRecordToggle(0).Visible = False
    SessionRecordToggle(1).Visible = False
    SessionRecordToggle(2).Visible = True
    SessionRecordToggle(3).Visible = False
End If
SessionRecordMode = Not SessionRecordMode
SessionRecordButtons
'frmMain.mnuRecord_Click
'Toggle logical value when button pressed
If SessionRecordMode Then
    SessionRecordLabel.Caption = "Stop Recording"
    SessionRecordLabel.ForeColor = &HFF&
    'set color to RED
    SessionRecordingStatus.Caption = "Currently Recording"
    'frmMain.mnuRecord.Checked = True 'check the menu item (but it causes an error)
Else
    SessionRecordLabel.Caption = "Start Recording"
    SessionRecordLabel.ForeColor = &H8000&
    'set color to GREEN
    SessionRecordingStatus.Caption = "Not Recording"
    'frmMain.mnuRecord.Checked = False 'uncheck the menu item (but it causes an error)
End If
End Sub
Public Sub SessionRecordButtons()
If SessionRecordMode Then
    'turn all active events RED while recording
    ResetDeckButton.BackColor = &HFF&
    ResetCurrentDeckButton.BackColor = &HFF&
    PokerDealButton.BackColor = &HFF&
    AssemblePokerDealButton.BackColor = &HFF&
    RiffleShuffleButton.BackColor = &HFF&
    OverhandShuffleButton.BackColor = &HFF&
    CutShuffleButton.BackColor = &HFF&
    ShiftTopBlockButton.BackColor = &HFF&
    MoveCardButton.BackColor = &HFF&
    RunSingleButton.BackColor = &HFF&
    InFaroButton.BackColor = &HFF&
    OutFaroButton.BackColor = &HFF&
    ForceButton.BackColor = &HFF&
    FreeChoiceSpreadButton.BackColor = &HFF&
    FreeChoiceHandlingButton.BackColor = &HFF&
    SwapCardsButton.BackColor = &HFF&
    If frmMain.mnuPiles.Checked = True Then
    'If frmPiles.Visible Then
        frmPiles.CreatePilesButton.BackColor = &HFF&
        frmPiles.SwapPilesButton.BackColor = &HFF&
        frmPiles.AustralianDealButton.BackColor = &HFF&
        frmPiles.SpecialButton.BackColor = &HFF&
        frmPiles.CombinePilesButton.BackColor = &HFF&
        frmPiles.RiffleShufflePileButton.BackColor = &HFF&
        frmPiles.SelectReturnButton.BackColor = &HFF&
        frmPiles.CutCardsButton.BackColor = &HFF&
    End If
Else
    'turn all active events back to basic button color when not recording
    ResetDeckButton.BackColor = &H8000000F
    ResetCurrentDeckButton.BackColor = &H8000000F
    PokerDealButton.BackColor = &H8000000F
    AssemblePokerDealButton.BackColor = &H8000000F
    RiffleShuffleButton.BackColor = &H8000000F
    OverhandShuffleButton.BackColor = &H8000000F
    CutShuffleButton.BackColor = &H8000000F
    ShiftTopBlockButton.BackColor = &H8000000F
    MoveCardButton.BackColor = &H8000000F
    RunSingleButton.BackColor = &H8000000F
    InFaroButton.BackColor = &H8000000F
    OutFaroButton.BackColor = &H8000000F
    ForceButton.BackColor = &H8000000F
    FreeChoiceSpreadButton.BackColor = &H8000000F
    FreeChoiceHandlingButton.BackColor = &H8000000F
    SwapCardsButton.BackColor = &H8000000F
    If frmMain.mnuPiles.Checked = True Then
        frmPiles.CreatePilesButton.BackColor = &HFFC0C0
        frmPiles.SwapPilesButton.BackColor = &H8000000F
        frmPiles.AustralianDealButton.BackColor = &H8000000F
        frmPiles.SpecialButton.BackColor = &H8000000F
        frmPiles.CombinePilesButton.BackColor = &H8000000F
        frmPiles.RiffleShufflePileButton.BackColor = &H8000000F
        frmPiles.SelectReturnButton.BackColor = &H8000000F
        frmPiles.CutCardsButton.BackColor = &H8000000F
    End If
End If
End Sub

Public Sub PokerDealButton_Click()
If PokerDealCombo.ListIndex = -1 Then
    PokerDealCombo.Text = "Select a Poker deal"
    MsgBox "Please select a valid response from" & Chr(13) _
        & "the 'Poker Deal' Dropdown Box"
    Exit Sub
End If
If PokerDealCombo.ListIndex = 0 Then
    tmp = 2
ElseIf PokerDealCombo.ListIndex = 1 Then
    tmp = 3
ElseIf PokerDealCombo.ListIndex = 2 Then
    tmp = 4
ElseIf PokerDealCombo.ListIndex = 3 Then
    tmp = 5
ElseIf PokerDealCombo.ListIndex = 4 Then
    tmp = 6
ElseIf PokerDealCombo.ListIndex = 5 Then
    tmp = 7
ElseIf PokerDealCombo.ListIndex = 6 Then
    tmp = 8
ElseIf PokerDealCombo.ListIndex = 7 Then
    tmp = 9
ElseIf PokerDealCombo.ListIndex = 8 Then
    tmp = 10
End If
PokerDeal (tmp)
'SessionRecord
If SessionRecordMode Then
    SessionCommand = "PokerDeal(" & tmp & ")"
    SessionListBox.AddItem SessionCommand
    SessionStatusUpdate (0)
End If
End Sub

Public Sub PokerDeal(listhands)
    Hands = listhands
    For m% = 1 To DeckCount
        For z% = 1 To DeckProperties
            UnwindDeck(z%, m%) = Deck(z%, m%)
        Next z%
    Next m%
    CalculateDeal
    ShowDeal
End Sub

Public Sub CalculateDeal()
    currentPosition = 1
    For i% = 1 To Hands
        For j% = 1 To 5
            For z% = 1 To DeckProperties
                ChangedDeck(z%, currentPosition) = Deck(z%, i% + (j% - 1) * Hands)
            Next z%
            currentPosition = currentPosition + 1
        Next j%
    Next i%
    For m% = 1 To 5 * Hands
        For z% = 1 To DeckProperties
            Deck(z%, m%) = ChangedDeck(z%, m%)
        Next z%
    Next m%
End Sub
Public Sub ShowDeal()
PokerCardsDealt = 1
PilesShown = 0
Call frmDeck.DisplayDeal
End Sub


Public Sub ResetDeckButton_Click()
If SetStackCombo.ListIndex = -1 Then
    SetStackCombo.Text = "Select a stack to use"
    MsgBox "Please select a valid response from" & Chr(13) _
        & "the 'Set Stack' Dropdown Box"
    Exit Sub
End If
If SetStackCombo.ListIndex = 0 Then
    SetStack ("Default")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "SetStack(" & Chr(34) & "Default" & Chr(34) & ")"
        SessionListBox.AddItem SessionCommand
        SessionStatusUpdate (0)
    End If
ElseIf SetStackCombo.ListIndex = 1 Then
    SetStack ("Aronson")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "SetStack(" & Chr(34) & "Aronson" & Chr(34) & ")"
        SessionListBox.AddItem SessionCommand
        SessionStatusUpdate (0)
    End If
ElseIf SetStackCombo.ListIndex = 2 Then
    SetStack ("Eight Kings")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "SetStack(" & Chr(34) & "Eight Kings" & Chr(34) & ")"
        SessionListBox.AddItem SessionCommand
        SessionStatusUpdate (0)
    End If
ElseIf SetStackCombo.ListIndex = 3 Then
    SetStack ("Ireland")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "SetStack(" & Chr(34) & "Ireland" & Chr(34) & ")"
        SessionListBox.AddItem SessionCommand
        SessionStatusUpdate (0)
    End If
ElseIf SetStackCombo.ListIndex = 4 Then
    SetStack ("Joyal (CHaSeD)")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "SetStack(" & Chr(34) & "Joyal (CHaSeD)" & Chr(34) & ")"
        SessionListBox.AddItem SessionCommand
        SessionStatusUpdate (0)
    End If
ElseIf SetStackCombo.ListIndex = 5 Then
    SetStack ("Joyal (SHoCkeD)")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "SetStack(" & Chr(34) & "Joyal (SHoCkeD)" & Chr(34) & ")"
        SessionListBox.AddItem SessionCommand
        SessionStatusUpdate (0)
    End If
ElseIf SetStackCombo.ListIndex = 6 Then
    SetStack ("New Deck (Bicycle)")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "SetStack(" & Chr(34) & "New Deck (Bicycle)" & Chr(34) & ")"
        SessionListBox.AddItem SessionCommand
        SessionStatusUpdate (0)
    End If
ElseIf SetStackCombo.ListIndex = 7 Then
    SetStack ("New Deck (Fournier)")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "SetStack(" & Chr(34) & "New Deck (Fournier)" & Chr(34) & ")"
        SessionListBox.AddItem SessionCommand
        SessionStatusUpdate (0)
    End If
ElseIf SetStackCombo.ListIndex = 8 Then
    SetStack ("Nikola")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "SetStack(" & Chr(34) & "Nikola" & Chr(34) & ")"
        SessionListBox.AddItem SessionCommand
        SessionStatusUpdate (0)
    End If
ElseIf SetStackCombo.ListIndex = 9 Then
    SetStack ("Osterlind")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "SetStack(" & Chr(34) & "Osterlind" & Chr(34) & ")"
        SessionListBox.AddItem SessionCommand
        SessionStatusUpdate (0)
    End If
ElseIf SetStackCombo.ListIndex = 10 Then
    SetStack ("Si Stebbins (3)")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "SetStack(" & Chr(34) & "Si Stebbins (3)" & Chr(34) & ")"
        SessionListBox.AddItem SessionCommand
        SessionStatusUpdate (0)
    End If
ElseIf SetStackCombo.ListIndex = 11 Then
    SetStack ("Si Stebbins (4)")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "SetStack(" & Chr(34) & "Si Stebbins (4)" & Chr(34) & ")"
        SessionListBox.AddItem SessionCommand
        SessionStatusUpdate (0)
    End If
ElseIf SetStackCombo.ListIndex = 12 Then
    SetStack ("Stanyon")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "SetStack(" & Chr(34) & "Stanyon" & Chr(34) & ")"
        SessionListBox.AddItem SessionCommand
        SessionStatusUpdate (0)
    End If
ElseIf SetStackCombo.ListIndex = 13 Then
    SetStack ("Tamariz")
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "SetStack(" & Chr(34) & "Tamariz" & Chr(34) & ")"
        SessionListBox.AddItem SessionCommand
        SessionStatusUpdate (0)
    End If
End If
End Sub

Public Sub SetStack(deckname)
If deckname = Chr(34) & "Default" & Chr(34) Or _
    deckname = "Default" Then
    On Error GoTo DefaultLoad
    Dim fso As New FileSystemObject, deckfile As File, ts As TextStream
    Set deckfile = fso.GetFile(App.Path & "\stackview.svf")
    Set ts = deckfile.OpenAsTextStream(ForReading)
    For i% = 1 To 52
        Deck(1, i%) = Val(ts.ReadLine)
        Deck(2, i%) = ts.ReadLine
    Next i%
    'XXX
    Dim sProperty As String
    If Not ts.AtEndOfStream Then
        sProperty = ts.ReadLine
        If Left(sProperty, 10) = "BackDesign" Then
            BackDesignCurrent = Right(sProperty, Len(sProperty) - 11)
        End If
    End If
    Call frmBackDesignDialog.LoadBackDesign(BackDesignCurrent)
    'XXX
    ts.Close
    For k% = 1 To 52
        For m% = 1 To 52
            If Deck(1, m%) = k% Then
                TestOriginalDeck(1, k%) = Deck(1, m%)
                TestOriginalDeck(2, k%) = Deck(2, m%)
            End If
        Next m%
    Next k%
    'the above For/Next sets the current deck to its
    'original position for testing
    ClearSelections_Click
    NumberOfSelectedCards = 0
    SelectionsTextBox.Text = Empty
    DeckProperties = 6
    DeckCount = 52
    ShowCards
    GoTo FinishLoad
DefaultLoad:
    SetDeckNewBicycle
    For k% = 1 To 52
        For m% = 1 To 52
            If Deck(1, m%) = k% Then
                TestOriginalDeck(1, k%) = Deck(1, m%)
                TestOriginalDeck(2, k%) = Deck(2, m%)
            End If
        Next m%
    Next k%
    NumberOfSelectedCards = 0
    SelectionsTextBox.Text = Empty
    DeckProperties = 6
ElseIf deckname = Chr(34) & "Aronson" & Chr(34) Or _
    deckname = "Aronson" Then
    SetDeckAronson
ElseIf deckname = Chr(34) & "Ireland" & Chr(34) Or _
    deckname = "Ireland" Then
    SetDeckIreland
ElseIf deckname = Chr(34) & "Eight Kings" & Chr(34) Or _
    deckname = "Eight Kings" Then
    SetDeckEightKings
ElseIf deckname = Chr(34) & "Joyal (CHaSeD)" & Chr(34) Or _
    deckname = "Joyal (CHaSeD)" Then
    SetDeckJoyalCHSD
ElseIf deckname = Chr(34) & "Joyal (SHoCkeD)" & Chr(34) Or _
    deckname = "Joyal (SHoCkeD)" Then
    SetDeckJoyalSHCD
ElseIf deckname = Chr(34) & "New Deck (Bicycle)" & Chr(34) Or _
    deckname = "New Deck (Bicycle)" Then
    SetDeckNewBicycle
ElseIf deckname = Chr(34) & "New Deck (Fournier)" & Chr(34) Or _
    deckname = "New Deck (Fournier)" Then
    SetDeckNewFournier
ElseIf deckname = Chr(34) & "Nikola" & Chr(34) Or _
    deckname = "Nikola" Then
    SetDeckNikola
ElseIf deckname = Chr(34) & "Osterlind" & Chr(34) Or _
    deckname = "Osterlind" Then
    SetDeckOsterlind
ElseIf deckname = Chr(34) & "Si Stebbins (3)" & Chr(34) Or _
    deckname = "Si Stebbins (3)" Then
    SetDeckSiStebbins3
ElseIf deckname = Chr(34) & "Si Stebbins (4)" & Chr(34) Or _
    deckname = "Si Stebbins (4)" Then
    SetDeckSiStebbins4
ElseIf deckname = Chr(34) & "Stanyon" & Chr(34) Or _
    deckname = "Stanyon" Then
    SetDeckStanyon
ElseIf deckname = Chr(34) & "Tamariz" & Chr(34) Or _
    deckname = "Tamariz" Then
    SetDeckTamariz
End If
For k% = 1 To 52
    For m% = 1 To 52
        If Deck(1, m%) = k% Then
            TestOriginalDeck(1, k%) = Deck(1, m%)
            TestOriginalDeck(2, k%) = Deck(2, m%)
        End If
    Next m%
Next k%
'the above For/Next sets the current deck to its
'original position for testing
FinishLoad:
For i% = 1 To DeckCount
    Deck(3, i%) = "Card" & Deck(2, i%)
Next i%
ClearSelections_Click
ClearReversedCards_Click
Call ResetCurrentDeck
End Sub

Private Sub InverseDeck()
For i% = 1 To DeckCount
    For z% = 1 To DeckProperties
        ChangedDeck(z%, i%) = Deck(z%, DeckCount + 1 - i%)
    Next z%
Next i%
For m% = 1 To DeckCount
    For z% = 1 To DeckProperties
        Deck(z%, m%) = ChangedDeck(z%, m%)
    Next z%
Next m%
End Sub

Private Sub InverseInFaroReverse()
For i% = 1 To 26
    For z% = 1 To DeckProperties
        ChangedDeck(z%, i% + 26) = Deck(z%, 2 * i% - 1)
        ChangedDeck(z%, i%) = Deck(z%, 2 * i%)
    Next z%
Next i%
For m% = 1 To DeckCount
    For z% = 1 To DeckProperties
        Deck(z%, m%) = ChangedDeck(z%, m%)
    Next z%
Next m%
ReverseBlock (26)
ShowCards
End Sub


Private Sub InverseInFaro()
For i% = 1 To 26
    For z% = 1 To DeckProperties
        ChangedDeck(z%, i% + 26) = Deck(z%, 2 * i% - 1)
        ChangedDeck(z%, i%) = Deck(z%, 2 * i%)
    Next z%
Next i%
For m% = 1 To DeckCount
    For z% = 1 To DeckProperties
        Deck(z%, m%) = ChangedDeck(z%, m%)
    Next z%
Next m%
ShowCards
End Sub

Private Sub InverseInFaroSpecialBottom(rifbnumber, rifinumber)
ProtectedBlock = Val(rifbnumber)
InteriorCard = Val(rifinumber)
InverseDeck
'COMPLETE
If ProtectedBlock = 26 And InteriorCard = 0 Then
    'complete
    InverseOutFaro
    'need the opposite faro for Inverse bottom
ElseIf ProtectedBlock < 26 Then
    If InteriorCard = 0 Then
        'complete
        MeshedBlock = 2 * ProtectedBlock
        For i% = 1 To DeckCount - MeshedBlock
            For z% = 1 To DeckProperties
                ChangedDeck(z%, i% + MeshedBlock) = Deck(z%, i% + MeshedBlock)
            Next z%
        Next i%
        For k% = 1 To ProtectedBlock
            For z% = 1 To DeckProperties
                ChangedDeck(z%, k%) = Deck(z%, 2 * k% - 1)
                'for infaro, tool out the -1 after k% in Deck only
                ChangedDeck(z%, k% + ProtectedBlock) = Deck(z%, 2 * k%)
                'for infaro added the -1 after k%
            Next z%
        Next k%
        For m% = 1 To DeckCount
            For z% = 1 To DeckProperties
                Deck(z%, m%) = ChangedDeck(z%, m%)
            Next z%
        Next m%
    Else
        'complete
        InteriorPosition = InteriorCard + ProtectedBlock
        If ProtectedBlock + InteriorPosition <= 52 Then
            For j% = 1 To ProtectedBlock
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, j%) = Deck(z%, InteriorCard + 2 * (j% - 1))
                    'for infaro, removed the -1 after (j%) in Deck only
                Next z%
            Next j%
            For k% = 1 To InteriorCard - 1
                'for infaro removed the -1 after InteriorCard
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, ProtectedBlock + k%) = Deck(z%, k%)
                Next z%
            Next k%
            For b% = 1 To ProtectedBlock
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, InteriorPosition + b% - 1) = Deck(z%, InteriorCard + (2 * b%) - 1)
                    'for infaro took out the -1 after b%) in both terms
                Next z%
            Next b%
            For p% = InteriorPosition + ProtectedBlock To DeckCount
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, p%) = Deck(z%, p%)
                Next z%
            Next p%
            For m% = 1 To DeckCount
                For z% = 1 To DeckProperties
                    Deck(z%, m%) = ChangedDeck(z%, m%)
                Next z%
            Next m%
        Else
            'complete
            For j% = 1 To DeckCount - InteriorPosition + 1
            'for infaro removed the + 1 after InteriorPosition
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, j%) = Deck(z%, InteriorCard + 2 * (j% - 1))
                    'for infaro changed from 2* (j% - 1)
                Next z%
            Next j%
            For k% = 1 To ProtectedBlock - (DeckCount - InteriorPosition + 1)
            'for infaro removed the + 1 after InteriorPosition
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, DeckCount - InteriorPosition + 1 + k%) = _
                        Deck(z%, 2 * DeckCount - 2 * ProtectedBlock - InteriorCard + 1 + k%)
                        'for infaro took out + 1 after InteriorCard in both terms
                Next z%
            Next k%
            For b% = 1 To InteriorCard - 1
            'for infaro removed the -1 after InteriorCard
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, ProtectedBlock + b%) = _
                        Deck(z%, b%)
                Next z%
            Next b%
            For p% = 1 To DeckCount - InteriorPosition + 1
                'for infaro removed the + 1 after InteriorPosition
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, InteriorPosition - 1 + p%) = _
                        Deck(z%, InteriorCard + 2 * p% - 1)
                        'for infaro removed the -1 after InteriorPosition
                        'and changed from 2 * p% - 1
                Next z%
            Next p%
            For m% = 1 To DeckCount
                For z% = 1 To DeckProperties
                    Deck(z%, m%) = ChangedDeck(z%, m%)
                Next z%
            Next m%
        End If
    End If
ElseIf ProtectedBlock >= 26 Then
    If InteriorCard = 0 Then
        'complete
        MeshedBlock = 2 * (DeckCount - ProtectedBlock)
        For i% = 1 To DeckCount - ProtectedBlock
            For z% = 1 To DeckProperties
                ChangedDeck(z%, i%) = Deck(z%, 2 * i% - 1)
                'for infaro removed the -1 after i% in Deck only
            Next z%
        Next i%
        For j% = 1 To 2 * ProtectedBlock - DeckCount
            For z% = 1 To DeckProperties
                ChangedDeck(z%, j% + DeckCount - ProtectedBlock) = Deck(z%, j% + MeshedBlock)
            Next z%
        Next j%
        For k% = 1 To DeckCount - ProtectedBlock
            For z% = 1 To DeckProperties
                ChangedDeck(z%, k% + ProtectedBlock) = Deck(z%, 2 * k%)
                'for infaro added the - 1 after k% in Deck only
            Next z%
        Next k%
        For m% = 1 To DeckCount
            For z% = 1 To DeckProperties
                Deck(z%, m%) = ChangedDeck(z%, m%)
            Next z%
        Next m%
    Else
        'complete
        InteriorPosition = InteriorCard + ProtectedBlock
            For j% = 1 To DeckCount - InteriorPosition + 1
            'for infaro removed the + 1 after InteriorPostion
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, j%) = Deck(z%, InteriorCard + 2 * (j% - 1))
                    'for infaro changed from 2 * (j% - 1)
                Next z%
            Next j%
            For k% = 1 To ProtectedBlock - (DeckCount - InteriorPosition + 1)
                    'for infaro removed the + 1 after InteriorPostion
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, DeckCount - InteriorPosition + 1 + k%) = _
                        Deck(z%, 2 * DeckCount - 2 * ProtectedBlock - InteriorCard + 1 + k%)
                        'for infaro removed the + 1 after InteriorPostion and Interior card
                Next z%
            Next k%
            For b% = 1 To InteriorCard - 1
            'for infaro removed the - 1 after InteriorCard
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, ProtectedBlock + b%) = _
                        Deck(z%, b%)
                Next z%
            Next b%
            For p% = 1 To DeckCount - InteriorPosition + 1
            'for infaro removed the + 1 after InteriorPostion
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, InteriorPosition - 1 + p%) = _
                        Deck(z%, InteriorCard + 2 * p% - 1)
                        'for infaro removed the -1 after InteriorPosition
                        'and changed from 2 * p% - 1
                Next z%
            Next p%
            For m% = 1 To DeckCount
                For z% = 1 To DeckProperties
                    Deck(z%, m%) = ChangedDeck(z%, m%)
                Next z%
            Next m%
    End If
End If
Call CutDeckPrecise(ProtectedBlock, "X")
InverseDeck
ShowCards
End Sub

Private Sub InverseInFaroSpecialBottomReverse(rifbnumber, rifinumber)
ProtectedBlock = Val(rifbnumber)
InteriorCard = Val(rifinumber)
InverseDeck
'COMPLETE
If ProtectedBlock = 26 And InteriorCard = 0 Then
    'complete
    InverseOutFaroReverse
    'need the opposite faro for Inverse bottom
ElseIf ProtectedBlock < 26 Then
    If InteriorCard = 0 Then
        'complete
        MeshedBlock = 2 * ProtectedBlock
        For i% = 1 To DeckCount - MeshedBlock
            For z% = 1 To DeckProperties
                ChangedDeck(z%, i% + MeshedBlock) = Deck(z%, i% + MeshedBlock)
            Next z%
        Next i%
        For k% = 1 To ProtectedBlock
            For z% = 1 To DeckProperties
                ChangedDeck(z%, k%) = Deck(z%, 2 * k% - 1)
                'for infaro, tool out the -1 after k% in Deck only
                ChangedDeck(z%, k% + ProtectedBlock) = Deck(z%, 2 * k%)
                'for infaro added the -1 after k%
            Next z%
        Next k%
        For m% = 1 To DeckCount
            For z% = 1 To DeckProperties
                Deck(z%, m%) = ChangedDeck(z%, m%)
            Next z%
        Next m%
        ReverseBlock (ProtectedBlock)
    Else
        'complete
        InteriorPosition = InteriorCard + ProtectedBlock
        If ProtectedBlock + InteriorPosition <= 52 Then
            For j% = 1 To ProtectedBlock
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, j%) = Deck(z%, InteriorCard + 2 * (j% - 1))
                    'for infaro, removed the -1 after (j%) in Deck only
                Next z%
            Next j%
            For k% = 1 To InteriorCard - 1
                'for infaro removed the -1 after InteriorCard
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, ProtectedBlock + k%) = Deck(z%, k%)
                Next z%
            Next k%
            For b% = 1 To ProtectedBlock
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, InteriorPosition + b% - 1) = Deck(z%, InteriorCard + (2 * b%) - 1)
                    'for infaro took out the -1 after b%) in both terms
                Next z%
            Next b%
            For p% = InteriorPosition + ProtectedBlock To DeckCount
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, p%) = Deck(z%, p%)
                Next z%
            Next p%
            For m% = 1 To DeckCount
                For z% = 1 To DeckProperties
                    Deck(z%, m%) = ChangedDeck(z%, m%)
                Next z%
            Next m%
            ReverseBlock (ProtectedBlock)
        Else
            'complete
            For j% = 1 To DeckCount - InteriorPosition + 1
            'for infaro removed the + 1 after InteriorPosition
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, j%) = Deck(z%, InteriorCard + 2 * (j% - 1))
                    'for infaro changed from 2* (j% - 1)
                Next z%
            Next j%
            For k% = 1 To ProtectedBlock - (DeckCount - InteriorPosition + 1)
            'for infaro removed the + 1 after InteriorPosition
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, DeckCount - InteriorPosition + 1 + k%) = _
                        Deck(z%, 2 * DeckCount - 2 * ProtectedBlock - InteriorCard + 1 + k%)
                        'for infaro took out + 1 after InteriorCard in both terms
                Next z%
            Next k%
            For b% = 1 To InteriorCard - 1
            'for infaro removed the -1 after InteriorCard
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, ProtectedBlock + b%) = _
                        Deck(z%, b%)
                Next z%
            Next b%
            For p% = 1 To DeckCount - InteriorPosition + 1
                'for infaro removed the + 1 after InteriorPosition
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, InteriorPosition - 1 + p%) = _
                        Deck(z%, InteriorCard + 2 * p% - 1)
                        'for infaro removed the -1 after InteriorPosition
                        'and changed from 2 * p% - 1
                Next z%
            Next p%
            For m% = 1 To DeckCount
                For z% = 1 To DeckProperties
                    Deck(z%, m%) = ChangedDeck(z%, m%)
                Next z%
            Next m%
            ReverseBlock (ProtectedBlock)
        End If
    End If
ElseIf ProtectedBlock >= 26 Then
    If InteriorCard = 0 Then
        'complete
        MeshedBlock = 2 * (DeckCount - ProtectedBlock)
        For i% = 1 To DeckCount - ProtectedBlock
            For z% = 1 To DeckProperties
                ChangedDeck(z%, i%) = Deck(z%, 2 * i% - 1)
                'for infaro removed the -1 after i% in Deck only
            Next z%
        Next i%
        For j% = 1 To 2 * ProtectedBlock - DeckCount
            For z% = 1 To DeckProperties
                ChangedDeck(z%, j% + DeckCount - ProtectedBlock) = Deck(z%, j% + MeshedBlock)
            Next z%
        Next j%
        For k% = 1 To DeckCount - ProtectedBlock
            For z% = 1 To DeckProperties
                ChangedDeck(z%, k% + ProtectedBlock) = Deck(z%, 2 * k%)
                'for infaro added the - 1 after k% in Deck only
            Next z%
        Next k%
        For m% = 1 To DeckCount
            For z% = 1 To DeckProperties
                Deck(z%, m%) = ChangedDeck(z%, m%)
            Next z%
        Next m%
        ReverseBlock (ProtectedBlock)
    Else
        'complete
        InteriorPosition = InteriorCard + ProtectedBlock
            For j% = 1 To DeckCount - InteriorPosition + 1
            'for infaro removed the + 1 after InteriorPostion
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, j%) = Deck(z%, InteriorCard + 2 * (j% - 1))
                    'for infaro changed from 2 * (j% - 1)
                Next z%
            Next j%
            For k% = 1 To ProtectedBlock - (DeckCount - InteriorPosition + 1)
                    'for infaro removed the + 1 after InteriorPostion
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, DeckCount - InteriorPosition + 1 + k%) = _
                        Deck(z%, 2 * DeckCount - 2 * ProtectedBlock - InteriorCard + 1 + k%)
                        'for infaro removed the + 1 after InteriorPostion and Interior card
                Next z%
            Next k%
            For b% = 1 To InteriorCard - 1
            'for infaro removed the - 1 after InteriorCard
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, ProtectedBlock + b%) = _
                        Deck(z%, b%)
                Next z%
            Next b%
            For p% = 1 To DeckCount - InteriorPosition + 1
            'for infaro removed the + 1 after InteriorPostion
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, InteriorPosition - 1 + p%) = _
                        Deck(z%, InteriorCard + 2 * p% - 1)
                        'for infaro removed the -1 after InteriorPosition
                        'and changed from 2 * p% - 1
                Next z%
            Next p%
            For m% = 1 To DeckCount
                For z% = 1 To DeckProperties
                    Deck(z%, m%) = ChangedDeck(z%, m%)
                Next z%
            Next m%
            ReverseBlock (ProtectedBlock)
    End If
End If
Call CutDeckPrecise(ProtectedBlock, "X")
InverseDeck
ShowCards
End Sub


Private Sub InverseInFaroSpecialTop(riftnumber, rifinumber)
'COMPLETE
ProtectedBlock = Val(riftnumber)
InteriorCard = Val(rifinumber)
If ProtectedBlock = 26 And InteriorCard = 0 Then
    'complete
    InverseInFaro
ElseIf ProtectedBlock < 26 Then
    If InteriorCard = 0 Then
        'complete
        MeshedBlock = 2 * ProtectedBlock
        For i% = 1 To DeckCount - MeshedBlock
            For z% = 1 To DeckProperties
                ChangedDeck(z%, i% + MeshedBlock) = Deck(z%, i% + MeshedBlock)
            Next z%
        Next i%
        For k% = 1 To ProtectedBlock
            For z% = 1 To DeckProperties
                ChangedDeck(z%, k%) = Deck(z%, 2 * k%)
                ChangedDeck(z%, k% + ProtectedBlock) = Deck(z%, 2 * k% - 1)
            Next z%
        Next k%
        For m% = 1 To DeckCount
            For z% = 1 To DeckProperties
                Deck(z%, m%) = ChangedDeck(z%, m%)
            Next z%
        Next m%
    Else
        'complete
        InteriorPosition = InteriorCard + ProtectedBlock
        If ProtectedBlock + InteriorPosition <= 52 Then
            For j% = 1 To ProtectedBlock
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, j%) = Deck(z%, InteriorCard + 2 * j% - 1)
                Next z%
            Next j%
            For k% = 1 To InteriorCard
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, ProtectedBlock + k%) = Deck(z%, k%)
                Next z%
            Next k%
            For b% = 1 To ProtectedBlock
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, InteriorPosition + b%) = Deck(z%, InteriorCard + (2 * b%))
                Next z%
            Next b%
            For p% = InteriorPosition + ProtectedBlock + 1 To DeckCount
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, p%) = Deck(z%, p%)
                Next z%
            Next p%
            For m% = 1 To DeckCount
                For z% = 1 To DeckProperties
                    Deck(z%, m%) = ChangedDeck(z%, m%)
                Next z%
            Next m%
        Else
            'complete
            For j% = 1 To DeckCount - InteriorPosition
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, j%) = Deck(z%, InteriorCard + 2 * j% - 1)
                Next z%
            Next j%
            For k% = 1 To ProtectedBlock - (DeckCount - InteriorPosition)
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, DeckCount - InteriorPosition + k%) = _
                        Deck(z%, 2 * DeckCount - 2 * ProtectedBlock - InteriorCard + k%)
                Next z%
            Next k%
            For b% = 1 To InteriorCard
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, ProtectedBlock + b%) = _
                        Deck(z%, b%)
                Next z%
            Next b%
            For p% = 1 To DeckCount - InteriorPosition
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, InteriorPosition + p%) = _
                        Deck(z%, InteriorCard + 2 * p%)
                Next z%
            Next p%
            For m% = 1 To DeckCount
                For z% = 1 To DeckProperties
                    Deck(z%, m%) = ChangedDeck(z%, m%)
                Next z%
            Next m%
        End If
    End If
ElseIf ProtectedBlock >= 26 Then
    If InteriorCard = 0 Then
        'complete
        MeshedBlock = 2 * (DeckCount - ProtectedBlock)
        For i% = 1 To DeckCount - ProtectedBlock
            For z% = 1 To DeckProperties
                ChangedDeck(z%, i%) = Deck(z%, 2 * i%)
            Next z%
        Next i%
        For j% = 1 To 2 * ProtectedBlock - DeckCount
            For z% = 1 To DeckProperties
                ChangedDeck(z%, j% + DeckCount - ProtectedBlock) = Deck(z%, j% + MeshedBlock)
            Next z%
        Next j%
        For k% = 1 To DeckCount - ProtectedBlock
            For z% = 1 To DeckProperties
                ChangedDeck(z%, k% + ProtectedBlock) = Deck(z%, 2 * k% - 1)
            Next z%
        Next k%
        For m% = 1 To DeckCount
            For z% = 1 To DeckProperties
                Deck(z%, m%) = ChangedDeck(z%, m%)
            Next z%
        Next m%
    Else
        'complete
        InteriorPosition = InteriorCard + ProtectedBlock
            For j% = 1 To DeckCount - InteriorPosition
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, j%) = Deck(z%, InteriorCard + 2 * j% - 1)
                Next z%
            Next j%
            For k% = 1 To ProtectedBlock - (DeckCount - InteriorPosition)
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, DeckCount - InteriorPosition + k%) = _
                        Deck(z%, 2 * DeckCount - 2 * ProtectedBlock - InteriorCard + k%)
                Next z%
            Next k%
            For b% = 1 To InteriorCard
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, ProtectedBlock + b%) = _
                        Deck(z%, b%)
                Next z%
            Next b%
            For p% = 1 To DeckCount - InteriorPosition
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, InteriorPosition + p%) = _
                        Deck(z%, InteriorCard + 2 * p%)
                Next z%
            Next p%
            For m% = 1 To DeckCount
                For z% = 1 To DeckProperties
                    Deck(z%, m%) = ChangedDeck(z%, m%)
                Next z%
            Next m%
    End If
End If
ShowCards
End Sub

Private Sub InverseInFaroSpecialTopReverse(riftnumber, rifinumber)
'COMPLETE
ProtectedBlock = Val(riftnumber)
InteriorCard = Val(rifinumber)
If ProtectedBlock = 26 And InteriorCard = 0 Then
    'complete
    InverseInFaroReverse
ElseIf ProtectedBlock < 26 Then
    If InteriorCard = 0 Then
        'complete
        MeshedBlock = 2 * ProtectedBlock
        For i% = 1 To DeckCount - MeshedBlock
            For z% = 1 To DeckProperties
                ChangedDeck(z%, i% + MeshedBlock) = Deck(z%, i% + MeshedBlock)
            Next z%
        Next i%
        For k% = 1 To ProtectedBlock
            For z% = 1 To DeckProperties
                ChangedDeck(z%, k%) = Deck(z%, 2 * k%)
                ChangedDeck(z%, k% + ProtectedBlock) = Deck(z%, 2 * k% - 1)
            Next z%
        Next k%
        For m% = 1 To DeckCount
            For z% = 1 To DeckProperties
                Deck(z%, m%) = ChangedDeck(z%, m%)
            Next z%
        Next m%
        ReverseBlock (ProtectedBlock)
    Else
        'complete
        InteriorPosition = InteriorCard + ProtectedBlock
        If ProtectedBlock + InteriorPosition <= 52 Then
            For j% = 1 To ProtectedBlock
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, j%) = Deck(z%, InteriorCard + 2 * j% - 1)
                Next z%
            Next j%
            For k% = 1 To InteriorCard
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, ProtectedBlock + k%) = Deck(z%, k%)
                Next z%
            Next k%
            For b% = 1 To ProtectedBlock
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, InteriorPosition + b%) = Deck(z%, InteriorCard + (2 * b%))
                Next z%
            Next b%
            For p% = InteriorPosition + ProtectedBlock + 1 To DeckCount
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, p%) = Deck(z%, p%)
                Next z%
            Next p%
            For m% = 1 To DeckCount
                For z% = 1 To DeckProperties
                    Deck(z%, m%) = ChangedDeck(z%, m%)
                Next z%
            Next m%
            ReverseBlock (ProtectedBlock)
        Else
            'complete
            For j% = 1 To DeckCount - InteriorPosition
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, j%) = Deck(z%, InteriorCard + 2 * j% - 1)
                Next z%
            Next j%
            For k% = 1 To ProtectedBlock - (DeckCount - InteriorPosition)
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, DeckCount - InteriorPosition + k%) = _
                        Deck(z%, 2 * DeckCount - 2 * ProtectedBlock - InteriorCard + k%)
                Next z%
            Next k%
            For b% = 1 To InteriorCard
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, ProtectedBlock + b%) = _
                        Deck(z%, b%)
                Next z%
            Next b%
            For p% = 1 To DeckCount - InteriorPosition
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, InteriorPosition + p%) = _
                        Deck(z%, InteriorCard + 2 * p%)
                Next z%
            Next p%
            For m% = 1 To DeckCount
                For z% = 1 To DeckProperties
                    Deck(z%, m%) = ChangedDeck(z%, m%)
                Next z%
            Next m%
            ReverseBlock (ProtectedBlock)
        End If
    End If
ElseIf ProtectedBlock >= 26 Then
    If InteriorCard = 0 Then
        'complete
        MeshedBlock = 2 * (DeckCount - ProtectedBlock)
        For i% = 1 To DeckCount - ProtectedBlock
            For z% = 1 To DeckProperties
                ChangedDeck(z%, i%) = Deck(z%, 2 * i%)
            Next z%
        Next i%
        For j% = 1 To 2 * ProtectedBlock - DeckCount
            For z% = 1 To DeckProperties
                ChangedDeck(z%, j% + DeckCount - ProtectedBlock) = Deck(z%, j% + MeshedBlock)
            Next z%
        Next j%
        For k% = 1 To DeckCount - ProtectedBlock
            For z% = 1 To DeckProperties
                ChangedDeck(z%, k% + ProtectedBlock) = Deck(z%, 2 * k% - 1)
            Next z%
        Next k%
        For m% = 1 To DeckCount
            For z% = 1 To DeckProperties
                Deck(z%, m%) = ChangedDeck(z%, m%)
            Next z%
        Next m%
        ReverseBlock (ProtectedBlock)
    Else
        'complete
        InteriorPosition = InteriorCard + ProtectedBlock
            For j% = 1 To DeckCount - InteriorPosition
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, j%) = Deck(z%, InteriorCard + 2 * j% - 1)
                Next z%
            Next j%
            For k% = 1 To ProtectedBlock - (DeckCount - InteriorPosition)
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, DeckCount - InteriorPosition + k%) = _
                        Deck(z%, 2 * DeckCount - 2 * ProtectedBlock - InteriorCard + k%)
                Next z%
            Next k%
            For b% = 1 To InteriorCard
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, ProtectedBlock + b%) = _
                        Deck(z%, b%)
                Next z%
            Next b%
            For p% = 1 To DeckCount - InteriorPosition
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, InteriorPosition + p%) = _
                        Deck(z%, InteriorCard + 2 * p%)
                Next z%
            Next p%
            For m% = 1 To DeckCount
                For z% = 1 To DeckProperties
                    Deck(z%, m%) = ChangedDeck(z%, m%)
                Next z%
            Next m%
            ReverseBlock (ProtectedBlock)
    End If
End If
ShowCards
End Sub


Private Sub InverseOutFaro()
For i% = 1 To 26
    For z% = 1 To DeckProperties
        ChangedDeck(z%, i%) = Deck(z%, 2 * i% - 1)
        ChangedDeck(z%, i% + 26) = Deck(z%, 2 * i%)
    Next z%
Next i%
For m% = 1 To DeckCount
    For z% = 1 To DeckProperties
        Deck(z%, m%) = ChangedDeck(z%, m%)
    Next z%
Next m%
ShowCards
End Sub

Private Sub InverseOutFaroReverse()
For i% = 1 To 26
    For z% = 1 To DeckProperties
        ChangedDeck(z%, i%) = Deck(z%, 2 * i% - 1)
        ChangedDeck(z%, i% + 26) = Deck(z%, 2 * i%)
    Next z%
Next i%
For m% = 1 To DeckCount
    For z% = 1 To DeckProperties
        Deck(z%, m%) = ChangedDeck(z%, m%)
    Next z%
Next m%
ReverseBlock (26)
ShowCards
End Sub




Private Sub InverseOutFaroSpecialBottom(rofbnumber, rofinumber)
ProtectedBlock = Val(rofbnumber)
InteriorCard = Val(rofinumber)
InverseDeck
'COMPLETE
If ProtectedBlock = 26 And InteriorCard = 0 Then
    'complete
    InverseInFaro
    'need the opposite faro for Inverse bottom
ElseIf ProtectedBlock < 26 Then
    If InteriorCard = 0 Then
        'complete
        MeshedBlock = 2 * ProtectedBlock
        For i% = 1 To DeckCount - MeshedBlock
            For z% = 1 To DeckProperties
                ChangedDeck(z%, i% + MeshedBlock) = Deck(z%, i% + MeshedBlock)
            Next z%
        Next i%
        For k% = 1 To ProtectedBlock
            For z% = 1 To DeckProperties
                ChangedDeck(z%, k%) = Deck(z%, 2 * k%)
                'for infaro, tool out the -1 after k% in Deck only
                ChangedDeck(z%, k% + ProtectedBlock) = Deck(z%, 2 * k% - 1)
                'for infaro added the -1 after k%
            Next z%
        Next k%
        For m% = 1 To DeckCount
            For z% = 1 To DeckProperties
                Deck(z%, m%) = ChangedDeck(z%, m%)
            Next z%
        Next m%
    Else
        'complete
        InteriorPosition = InteriorCard + ProtectedBlock
        If ProtectedBlock + InteriorPosition <= 52 Then
            For j% = 1 To ProtectedBlock
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, j%) = Deck(z%, InteriorCard + 2 * j% - 1)
                    'for infaro, removed the -1 after (j%) in Deck only
                Next z%
            Next j%
            For k% = 1 To InteriorCard
                'for infaro removed the -1 after InteriorCard
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, ProtectedBlock + k%) = Deck(z%, k%)
                Next z%
            Next k%
            For b% = 1 To ProtectedBlock
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, InteriorPosition + b%) = Deck(z%, InteriorCard + (2 * b%))
                    'for infaro took out the -1 after b%) in both terms
                Next z%
            Next b%
            For p% = InteriorPosition + ProtectedBlock To DeckCount
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, p%) = Deck(z%, p%)
                Next z%
            Next p%
            For m% = 1 To DeckCount
                For z% = 1 To DeckProperties
                    Deck(z%, m%) = ChangedDeck(z%, m%)
                Next z%
            Next m%
        Else
            'complete
            For j% = 1 To DeckCount - InteriorPosition
            'for infaro removed the + 1 after InteriorPosition
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, j%) = Deck(z%, InteriorCard + 2 * j% - 1)
                    'for infaro changed from 2* (j% - 1)
                Next z%
            Next j%
            For k% = 1 To ProtectedBlock - (DeckCount - InteriorPosition)
            'for infaro removed the + 1 after InteriorPosition
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, DeckCount - InteriorPosition + k%) = _
                        Deck(z%, 2 * DeckCount - 2 * ProtectedBlock - InteriorCard + k%)
                        'for infaro took out + 1 after InteriorCard in both terms
                Next z%
            Next k%
            For b% = 1 To InteriorCard
            'for infaro removed the -1 after InteriorCard
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, ProtectedBlock + b%) = _
                        Deck(z%, b%)
                Next z%
            Next b%
            For p% = 1 To DeckCount - InteriorPosition
                'for infaro removed the + 1 after InteriorPosition
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, InteriorPosition + p%) = _
                        Deck(z%, InteriorCard + 2 * p%)
                        'for infaro removed the -1 after InteriorPosition
                        'and changed from 2 * p% - 1
                Next z%
            Next p%
            For m% = 1 To DeckCount
                For z% = 1 To DeckProperties
                    Deck(z%, m%) = ChangedDeck(z%, m%)
                Next z%
            Next m%
        End If
    End If
ElseIf ProtectedBlock >= 26 Then
    If InteriorCard = 0 Then
        'complete
        MeshedBlock = 2 * (DeckCount - ProtectedBlock)
        For i% = 1 To DeckCount - ProtectedBlock
            For z% = 1 To DeckProperties
                ChangedDeck(z%, i%) = Deck(z%, 2 * i%)
                'for infaro removed the -1 after i% in Deck only
            Next z%
        Next i%
        For j% = 1 To 2 * ProtectedBlock - DeckCount
            For z% = 1 To DeckProperties
                ChangedDeck(z%, j% + DeckCount - ProtectedBlock) = Deck(z%, j% + MeshedBlock)
            Next z%
        Next j%
        For k% = 1 To DeckCount - ProtectedBlock
            For z% = 1 To DeckProperties
                ChangedDeck(z%, k% + ProtectedBlock) = Deck(z%, 2 * k% - 1)
                'for infaro added the - 1 after k% in Deck only
            Next z%
        Next k%
        For m% = 1 To DeckCount
            For z% = 1 To DeckProperties
                Deck(z%, m%) = ChangedDeck(z%, m%)
            Next z%
        Next m%
    Else
        'complete
        InteriorPosition = InteriorCard + ProtectedBlock
            For j% = 1 To DeckCount - InteriorPosition
            'for infaro removed the + 1 after InteriorPostion
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, j%) = Deck(z%, InteriorCard + 2 * j% - 1)
                    'for infaro changed from 2 * (j% - 1)
                Next z%
            Next j%
            For k% = 1 To ProtectedBlock - (DeckCount - InteriorPosition)
                    'for infaro removed the + 1 after InteriorPostion
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, DeckCount - InteriorPosition + k%) = _
                        Deck(z%, 2 * DeckCount - 2 * ProtectedBlock - InteriorCard + k%)
                        'for infaro removed the + 1 after InteriorPostion and Interior card
                Next z%
            Next k%
            For b% = 1 To InteriorCard
            'for infaro removed the - 1 after InteriorCard
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, ProtectedBlock + b%) = _
                        Deck(z%, b%)
                Next z%
            Next b%
            For p% = 1 To DeckCount - InteriorPosition
            'for infaro removed the + 1 after InteriorPostion
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, InteriorPosition + p%) = _
                        Deck(z%, InteriorCard + 2 * p%)
                        'for infaro removed the -1 after InteriorPosition
                        'and changed from 2 * p% - 1
                Next z%
            Next p%
            For m% = 1 To DeckCount
                For z% = 1 To DeckProperties
                    Deck(z%, m%) = ChangedDeck(z%, m%)
                Next z%
            Next m%
    End If
End If
Call CutDeckPrecise(ProtectedBlock, "X")
InverseDeck
ShowCards
End Sub

Private Sub InverseOutFaroSpecialBottomReverse(rofbnumber, rofinumber)
ProtectedBlock = Val(rofbnumber)
InteriorCard = Val(rofinumber)
InverseDeck
'COMPLETE
If ProtectedBlock = 26 And InteriorCard = 0 Then
    'complete
    InverseInFaroReverse
    'need the opposite faro for Inverse bottom
ElseIf ProtectedBlock < 26 Then
    If InteriorCard = 0 Then
        'complete
        MeshedBlock = 2 * ProtectedBlock
        For i% = 1 To DeckCount - MeshedBlock
            For z% = 1 To DeckProperties
                ChangedDeck(z%, i% + MeshedBlock) = Deck(z%, i% + MeshedBlock)
            Next z%
        Next i%
        For k% = 1 To ProtectedBlock
            For z% = 1 To DeckProperties
                ChangedDeck(z%, k%) = Deck(z%, 2 * k%)
                'for infaro, tool out the -1 after k% in Deck only
                ChangedDeck(z%, k% + ProtectedBlock) = Deck(z%, 2 * k% - 1)
                'for infaro added the -1 after k%
            Next z%
        Next k%
        For m% = 1 To DeckCount
            For z% = 1 To DeckProperties
                Deck(z%, m%) = ChangedDeck(z%, m%)
            Next z%
        Next m%
        ReverseBlock (ProtectedBlock)
    Else
        'complete
        InteriorPosition = InteriorCard + ProtectedBlock
        If ProtectedBlock + InteriorPosition <= 52 Then
            For j% = 1 To ProtectedBlock
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, j%) = Deck(z%, InteriorCard + 2 * j% - 1)
                    'for infaro, removed the -1 after (j%) in Deck only
                Next z%
            Next j%
            For k% = 1 To InteriorCard
                'for infaro removed the -1 after InteriorCard
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, ProtectedBlock + k%) = Deck(z%, k%)
                Next z%
            Next k%
            For b% = 1 To ProtectedBlock
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, InteriorPosition + b%) = Deck(z%, InteriorCard + (2 * b%))
                    'for infaro took out the -1 after b%) in both terms
                Next z%
            Next b%
            For p% = InteriorPosition + ProtectedBlock To DeckCount
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, p%) = Deck(z%, p%)
                Next z%
            Next p%
            For m% = 1 To DeckCount
                For z% = 1 To DeckProperties
                    Deck(z%, m%) = ChangedDeck(z%, m%)
                Next z%
            Next m%
            ReverseBlock (ProtectedBlock)
        Else
            'complete
            For j% = 1 To DeckCount - InteriorPosition
            'for infaro removed the + 1 after InteriorPosition
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, j%) = Deck(z%, InteriorCard + 2 * j% - 1)
                    'for infaro changed from 2* (j% - 1)
                Next z%
            Next j%
            For k% = 1 To ProtectedBlock - (DeckCount - InteriorPosition)
            'for infaro removed the + 1 after InteriorPosition
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, DeckCount - InteriorPosition + k%) = _
                        Deck(z%, 2 * DeckCount - 2 * ProtectedBlock - InteriorCard + k%)
                        'for infaro took out + 1 after InteriorCard in both terms
                Next z%
            Next k%
            For b% = 1 To InteriorCard
            'for infaro removed the -1 after InteriorCard
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, ProtectedBlock + b%) = _
                        Deck(z%, b%)
                Next z%
            Next b%
            For p% = 1 To DeckCount - InteriorPosition
                'for infaro removed the + 1 after InteriorPosition
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, InteriorPosition + p%) = _
                        Deck(z%, InteriorCard + 2 * p%)
                        'for infaro removed the -1 after InteriorPosition
                        'and changed from 2 * p% - 1
                Next z%
            Next p%
            For m% = 1 To DeckCount
                For z% = 1 To DeckProperties
                    Deck(z%, m%) = ChangedDeck(z%, m%)
                Next z%
            Next m%
            ReverseBlock (ProtectedBlock)
        End If
    End If
ElseIf ProtectedBlock >= 26 Then
    If InteriorCard = 0 Then
        'complete
        MeshedBlock = 2 * (DeckCount - ProtectedBlock)
        For i% = 1 To DeckCount - ProtectedBlock
            For z% = 1 To DeckProperties
                ChangedDeck(z%, i%) = Deck(z%, 2 * i%)
                'for infaro removed the -1 after i% in Deck only
            Next z%
        Next i%
        For j% = 1 To 2 * ProtectedBlock - DeckCount
            For z% = 1 To DeckProperties
                ChangedDeck(z%, j% + DeckCount - ProtectedBlock) = Deck(z%, j% + MeshedBlock)
            Next z%
        Next j%
        For k% = 1 To DeckCount - ProtectedBlock
            For z% = 1 To DeckProperties
                ChangedDeck(z%, k% + ProtectedBlock) = Deck(z%, 2 * k% - 1)
                'for infaro added the - 1 after k% in Deck only
            Next z%
        Next k%
        For m% = 1 To DeckCount
            For z% = 1 To DeckProperties
                Deck(z%, m%) = ChangedDeck(z%, m%)
            Next z%
        Next m%
        ReverseBlock (ProtectedBlock)
    Else
        'complete
        InteriorPosition = InteriorCard + ProtectedBlock
            For j% = 1 To DeckCount - InteriorPosition
            'for infaro removed the + 1 after InteriorPostion
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, j%) = Deck(z%, InteriorCard + 2 * j% - 1)
                    'for infaro changed from 2 * (j% - 1)
                Next z%
            Next j%
            For k% = 1 To ProtectedBlock - (DeckCount - InteriorPosition)
                    'for infaro removed the + 1 after InteriorPostion
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, DeckCount - InteriorPosition + k%) = _
                        Deck(z%, 2 * DeckCount - 2 * ProtectedBlock - InteriorCard + k%)
                        'for infaro removed the + 1 after InteriorPostion and Interior card
                Next z%
            Next k%
            For b% = 1 To InteriorCard
            'for infaro removed the - 1 after InteriorCard
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, ProtectedBlock + b%) = _
                        Deck(z%, b%)
                Next z%
            Next b%
            For p% = 1 To DeckCount - InteriorPosition
            'for infaro removed the + 1 after InteriorPostion
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, InteriorPosition + p%) = _
                        Deck(z%, InteriorCard + 2 * p%)
                        'for infaro removed the -1 after InteriorPosition
                        'and changed from 2 * p% - 1
                Next z%
            Next p%
            For m% = 1 To DeckCount
                For z% = 1 To DeckProperties
                    Deck(z%, m%) = ChangedDeck(z%, m%)
                Next z%
            Next m%
            ReverseBlock (ProtectedBlock)
    End If
End If
Call CutDeckPrecise(ProtectedBlock, "X")
InverseDeck
ShowCards
End Sub



Private Sub InverseOutFaroSpecialTop(roftnumber, rofinumber)
'COMPLETE
ProtectedBlock = Val(roftnumber)
InteriorCard = Val(rofinumber)
If ProtectedBlock = 26 And InteriorCard = 0 Then
    'complete
    InverseOutFaro
ElseIf ProtectedBlock < 26 Then
    If InteriorCard = 0 Then
        'complete
        MeshedBlock = 2 * ProtectedBlock
        For i% = 1 To DeckCount - MeshedBlock
            For z% = 1 To DeckProperties
                ChangedDeck(z%, i% + MeshedBlock) = Deck(z%, i% + MeshedBlock)
            Next z%
        Next i%
        For k% = 1 To ProtectedBlock
            For z% = 1 To DeckProperties
                ChangedDeck(z%, k%) = Deck(z%, 2 * k% - 1)
                ChangedDeck(z%, k% + ProtectedBlock) = Deck(z%, 2 * k%)
            Next z%
        Next k%
        For m% = 1 To DeckCount
            For z% = 1 To DeckProperties
                Deck(z%, m%) = ChangedDeck(z%, m%)
            Next z%
        Next m%
    Else
        'complete
        InteriorPosition = InteriorCard + ProtectedBlock
        If ProtectedBlock + InteriorPosition <= 52 Then
            For j% = 1 To ProtectedBlock
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, j%) = Deck(z%, InteriorCard + 2 * (j% - 1))
                Next z%
            Next j%
            For k% = 1 To InteriorCard - 1
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, ProtectedBlock + k%) = Deck(z%, k%)
                Next z%
            Next k%
            For b% = 1 To ProtectedBlock
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, InteriorPosition + b% - 1) = Deck(z%, InteriorCard + (2 * b%) - 1)
                Next z%
            Next b%
            For p% = InteriorPosition + ProtectedBlock To DeckCount
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, p%) = Deck(z%, p%)
                Next z%
            Next p%
            For m% = 1 To DeckCount
                For z% = 1 To DeckProperties
                    Deck(z%, m%) = ChangedDeck(z%, m%)
                Next z%
            Next m%
        Else
            'complete
            For j% = 1 To DeckCount - InteriorPosition + 1
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, j%) = Deck(z%, InteriorCard + 2 * (j% - 1))
                Next z%
            Next j%
            For k% = 1 To ProtectedBlock - (DeckCount - InteriorPosition + 1)
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, DeckCount - InteriorPosition + 1 + k%) = _
                        Deck(z%, 2 * DeckCount - 2 * ProtectedBlock - InteriorCard + 1 + k%)
                Next z%
            Next k%
            For b% = 1 To InteriorCard - 1
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, ProtectedBlock + b%) = _
                        Deck(z%, b%)
                Next z%
            Next b%
            For p% = 1 To DeckCount - InteriorPosition + 1
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, InteriorPosition - 1 + p%) = _
                        Deck(z%, InteriorCard + 2 * p% - 1)
                Next z%
            Next p%
            For m% = 1 To DeckCount
                For z% = 1 To DeckProperties
                    Deck(z%, m%) = ChangedDeck(z%, m%)
                Next z%
            Next m%
        End If
    End If
ElseIf ProtectedBlock >= 26 Then
    If InteriorCard = 0 Then
        'complete
        MeshedBlock = 2 * (DeckCount - ProtectedBlock)
        For i% = 1 To DeckCount - ProtectedBlock
            For z% = 1 To DeckProperties
                ChangedDeck(z%, i%) = Deck(z%, 2 * i% - 1)
            Next z%
        Next i%
        For j% = 1 To 2 * ProtectedBlock - DeckCount
            For z% = 1 To DeckProperties
                ChangedDeck(z%, j% + DeckCount - ProtectedBlock) = Deck(z%, j% + MeshedBlock)
            Next z%
        Next j%
        For k% = 1 To DeckCount - ProtectedBlock
            For z% = 1 To DeckProperties
                ChangedDeck(z%, k% + ProtectedBlock) = Deck(z%, 2 * k%)
            Next z%
        Next k%
        For m% = 1 To DeckCount
            For z% = 1 To DeckProperties
                Deck(z%, m%) = ChangedDeck(z%, m%)
            Next z%
        Next m%
    Else
        'complete
        InteriorPosition = InteriorCard + ProtectedBlock
            For j% = 1 To DeckCount - InteriorPosition + 1
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, j%) = Deck(z%, InteriorCard + 2 * (j% - 1))
                Next z%
            Next j%
            For k% = 1 To ProtectedBlock - (DeckCount - InteriorPosition + 1)
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, DeckCount - InteriorPosition + 1 + k%) = _
                        Deck(z%, 2 * DeckCount - 2 * ProtectedBlock - InteriorCard + 1 + k%)
                Next z%
            Next k%
            For b% = 1 To InteriorCard - 1
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, ProtectedBlock + b%) = _
                        Deck(z%, b%)
                Next z%
            Next b%
            For p% = 1 To DeckCount - InteriorPosition + 1
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, InteriorPosition - 1 + p%) = _
                        Deck(z%, InteriorCard + 2 * p% - 1)
                Next z%
            Next p%
            For m% = 1 To DeckCount
                For z% = 1 To DeckProperties
                    Deck(z%, m%) = ChangedDeck(z%, m%)
                Next z%
            Next m%
    End If
End If
ShowCards

'ORIGINAL CODE
'ProtectedBlock = Val(roftnumber)
'If ProtectedBlock = 26 Then
'    InverseOutFaro
'ElseIf ProtectedBlock < 26 Then
'    MeshedBlock = 2 * ProtectedBlock
'    For i% = 1 To DeckCount - MeshedBlock
'        For z% = 1 To DeckProperties
'            ChangedDeck(z%, i% + MeshedBlock) = Deck(z%, i% + MeshedBlock)
'        Next z%
'    Next i%
'    For k% = 1 To ProtectedBlock
'        For z% = 1 To DeckProperties
'            ChangedDeck(z%, k%) = Deck(z%, 2 * k% - 1)
'            ChangedDeck(z%, k% + ProtectedBlock) = Deck(z%, 2 * k%)
'        Next z%
'    Next k%
'    For m% = 1 To DeckCount
'        For z% = 1 To DeckProperties
'            Deck(z%, m%) = ChangedDeck(z%, m%)
'        Next z%
'    Next m%
'ElseIf ProtectedBlock > 26 Then
'    MeshedBlock = 2 * (DeckCount - ProtectedBlock)
'    For i% = 1 To DeckCount - ProtectedBlock
'        For z% = 1 To DeckProperties
'            ChangedDeck(z%, i%) = Deck(z%, 2 * i% - 1)
'        Next z%
'    Next i%
'    For j% = 1 To 2 * ProtectedBlock - DeckCount
'        For z% = 1 To DeckProperties
'            ChangedDeck(z%, j% + DeckCount - ProtectedBlock) = Deck(z%, j% + MeshedBlock)
'        Next z%
'    Next j%
'    For k% = 1 To DeckCount - ProtectedBlock
'        For z% = 1 To DeckProperties
'            ChangedDeck(z%, k% + ProtectedBlock) = Deck(z%, 2 * k%)
'        Next z%
'    Next k%
'    For m% = 1 To DeckCount
'        For z% = 1 To DeckProperties
'            Deck(z%, m%) = ChangedDeck(z%, m%)
'        Next z%
'    Next m%
'End If
'ShowCards
End Sub

Private Sub InverseOutFaroSpecialTopReverse(roftnumber, rofinumber)
'COMPLETE
ProtectedBlock = Val(roftnumber)
InteriorCard = Val(rofinumber)
If ProtectedBlock = 26 And InteriorCard = 0 Then
    'complete
    InverseOutFaroReverse
ElseIf ProtectedBlock < 26 Then
    If InteriorCard = 0 Then
        'complete
        MeshedBlock = 2 * ProtectedBlock
        For i% = 1 To DeckCount - MeshedBlock
            For z% = 1 To DeckProperties
                ChangedDeck(z%, i% + MeshedBlock) = Deck(z%, i% + MeshedBlock)
            Next z%
        Next i%
        For k% = 1 To ProtectedBlock
            For z% = 1 To DeckProperties
                ChangedDeck(z%, k%) = Deck(z%, 2 * k% - 1)
                ChangedDeck(z%, k% + ProtectedBlock) = Deck(z%, 2 * k%)
            Next z%
        Next k%
        For m% = 1 To DeckCount
            For z% = 1 To DeckProperties
                Deck(z%, m%) = ChangedDeck(z%, m%)
            Next z%
        Next m%
        ReverseBlock (ProtectedBlock)
    Else
        'complete
        InteriorPosition = InteriorCard + ProtectedBlock
        If ProtectedBlock + InteriorPosition <= 52 Then
            For j% = 1 To ProtectedBlock
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, j%) = Deck(z%, InteriorCard + 2 * (j% - 1))
                Next z%
            Next j%
            For k% = 1 To InteriorCard - 1
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, ProtectedBlock + k%) = Deck(z%, k%)
                Next z%
            Next k%
            For b% = 1 To ProtectedBlock
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, InteriorPosition + b% - 1) = Deck(z%, InteriorCard + (2 * b%) - 1)
                Next z%
            Next b%
            For p% = InteriorPosition + ProtectedBlock To DeckCount
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, p%) = Deck(z%, p%)
                Next z%
            Next p%
            For m% = 1 To DeckCount
                For z% = 1 To DeckProperties
                    Deck(z%, m%) = ChangedDeck(z%, m%)
                Next z%
            Next m%
            ReverseBlock (ProtectedBlock)
        Else
            'complete
            For j% = 1 To DeckCount - InteriorPosition + 1
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, j%) = Deck(z%, InteriorCard + 2 * (j% - 1))
                Next z%
            Next j%
            For k% = 1 To ProtectedBlock - (DeckCount - InteriorPosition + 1)
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, DeckCount - InteriorPosition + 1 + k%) = _
                        Deck(z%, 2 * DeckCount - 2 * ProtectedBlock - InteriorCard + 1 + k%)
                Next z%
            Next k%
            For b% = 1 To InteriorCard - 1
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, ProtectedBlock + b%) = _
                        Deck(z%, b%)
                Next z%
            Next b%
            For p% = 1 To DeckCount - InteriorPosition + 1
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, InteriorPosition - 1 + p%) = _
                        Deck(z%, InteriorCard + 2 * p% - 1)
                Next z%
            Next p%
            For m% = 1 To DeckCount
                For z% = 1 To DeckProperties
                    Deck(z%, m%) = ChangedDeck(z%, m%)
                Next z%
            Next m%
            ReverseBlock (ProtectedBlock)
        End If
    End If
ElseIf ProtectedBlock >= 26 Then
    If InteriorCard = 0 Then
        'complete
        MeshedBlock = 2 * (DeckCount - ProtectedBlock)
        For i% = 1 To DeckCount - ProtectedBlock
            For z% = 1 To DeckProperties
                ChangedDeck(z%, i%) = Deck(z%, 2 * i% - 1)
            Next z%
        Next i%
        For j% = 1 To 2 * ProtectedBlock - DeckCount
            For z% = 1 To DeckProperties
                ChangedDeck(z%, j% + DeckCount - ProtectedBlock) = Deck(z%, j% + MeshedBlock)
            Next z%
        Next j%
        For k% = 1 To DeckCount - ProtectedBlock
            For z% = 1 To DeckProperties
                ChangedDeck(z%, k% + ProtectedBlock) = Deck(z%, 2 * k%)
            Next z%
        Next k%
        For m% = 1 To DeckCount
            For z% = 1 To DeckProperties
                Deck(z%, m%) = ChangedDeck(z%, m%)
            Next z%
        Next m%
        ReverseBlock (ProtectedBlock)
    Else
        'complete
        InteriorPosition = InteriorCard + ProtectedBlock
            For j% = 1 To DeckCount - InteriorPosition + 1
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, j%) = Deck(z%, InteriorCard + 2 * (j% - 1))
                Next z%
            Next j%
            For k% = 1 To ProtectedBlock - (DeckCount - InteriorPosition + 1)
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, DeckCount - InteriorPosition + 1 + k%) = _
                        Deck(z%, 2 * DeckCount - 2 * ProtectedBlock - InteriorCard + 1 + k%)
                Next z%
            Next k%
            For b% = 1 To InteriorCard - 1
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, ProtectedBlock + b%) = _
                        Deck(z%, b%)
                Next z%
            Next b%
            For p% = 1 To DeckCount - InteriorPosition + 1
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, InteriorPosition - 1 + p%) = _
                        Deck(z%, InteriorCard + 2 * p% - 1)
                Next z%
            Next p%
            For m% = 1 To DeckCount
                For z% = 1 To DeckProperties
                    Deck(z%, m%) = ChangedDeck(z%, m%)
                Next z%
            Next m%
        ReverseBlock (ProtectedBlock)
    End If
End If
ShowCards
End Sub


'Public Sub RevInFaroOldButton()
'If RevInFaroStandardOption.Value = True Then
'    InverseInFaro
'    'SessionRecord
'    If SessionRecordMode Then
'        SessionCommand = "InverseInFaro"
'        SessionListBox.AddItem SessionCommand
'        SessionStatusUpdate (0)
'    End If
'ElseIf RevInFaroStartWeaveTopOption.Value = True Then
'    InverseInFaroSpecialTop (Val(RevInFaroFromTopTextBox.Text))
'    'SessionRecord
'    If SessionRecordMode Then
'        SessionCommand = "InverseInFaroSpecialTop(" & RevInFaroFromTopTextBox.Text & ")"
'        SessionListBox.AddItem SessionCommand
'        SessionStatusUpdate (0)
'    End If
'ElseIf RevInFaroStartWeaveBottomOption.Value = True Then
'    InverseInFaroSpecialBottom (Val(RevInFaroFromTopTextBox.Text))
'    'SessionRecord
'    If SessionRecordMode Then
'        SessionCommand = "InverseInFaroSpecialBottom(" & RevInFaroFromTopTextBox.Text & ")"
'        SessionListBox.AddItem SessionCommand
'        SessionStatusUpdate (0)
'    End If
'End If
'End Sub


'Public Sub RevOutFaroOldButton()
'If RevOutFaroStandardOption.Value = True Then
'    InverseOutFaro
'    'SessionRecord
'    If SessionRecordMode Then
'        SessionCommand = "InverseOutFaro"
'        SessionListBox.AddItem SessionCommand
'        SessionStatusUpdate (0)
'    End If
'ElseIf RevOutFaroStartWeaveTopOption.Value = True Then
'    InverseOutFaroSpecialTop (Val(RevOutFaroFromTopTextBox.Text))
'    'SessionRecord
'    If SessionRecordMode Then
'        SessionCommand = "InverseOutFaroSpecialTop(" & RevOutFaroFromTopTextBox.Text & ")"
'        SessionListBox.AddItem SessionCommand
'        SessionStatusUpdate (0)
'    End If
'ElseIf RevOutFaroStartWeaveBottomOption.Value = True Then
'    InverseOutFaroSpecialBottom (Val(RevOutFaroFromTopTextBox.Text))
'    'SessionRecord
'    If SessionRecordMode Then
'        SessionCommand = "InverseOutFaroSpecialBottom(" & RevOutFaroFromTopTextBox.Text & ")"
'        SessionListBox.AddItem SessionCommand
'        SessionStatusUpdate (0)
'    End If
'End If
'End Sub





Public Sub RiffleLocationCombo_GotFocus()
    RiffleProtectOption.Value = True
End Sub

Public Sub RiffleNumberTextBox_GotFocus()
    RiffleProtectOption.Value = True
End Sub

Public Sub RiffleRandomOption_Click()
    RiffleNumberTextBox.Text = Empty
    RiffleLocationCombo.Text = Empty
End Sub

Public Sub RiffleShuffleBottom(riffleshufflenumber, paramreverse)
Dim pReverse As String
Dim pCutDepth As Integer
pReverse = paramreverse
'establish the cut parameters
ProtectedBlock = Val(riffleshufflenumber)
RifflePortion = DeckCount - ProtectedBlock
CutError = Int(0.2 * RifflePortion) + 1
CutDepth = Int(Rnd * CutError) + Int((RifflePortion - CutError) / 2)
  'approximately half the unprotected deck for a riffle shuffle
If CutDepth < 0 Then
    CutDepth = 0
        ' need to avoid an error
End If
RemainingCut = RifflePortion - CutDepth
    'reverse a block if required
If pReverse = "T" Then
    pCutStartCard = 1
    pCutEndCard = CutDepth
ElseIf pReverse = "B" Then
    pCutStartCard = CutDepth + 1
    pCutEndCard = 52
End If
If Not pReverse = "X" Then
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

' first section cuts the protected cards from the bottom
' to the top
pCutDepth = 52 - Val(riffleshufflenumber)
s = 1
For j% = pCutDepth + 1 To DeckCount
    For z% = 1 To DeckProperties
        ChangedDeck(z%, s) = Deck(z%, j%)
    Next z%
    s = s + 1
Next j%
For k% = 1 To pCutDepth
    For z% = 1 To DeckProperties
        ChangedDeck(z%, s) = Deck(z%, k%)
    Next z%
    s = s + 1
Next k%
For m% = 1 To DeckCount
    For z% = 1 To DeckProperties
        Deck(z%, m%) = ChangedDeck(z%, m%)
    Next z%
Next m%
' second section does the riffle shuffle with the
' protected section on top
    ReDim TopCut(DeckProperties, CutDepth)
    ReDim BottomCut(DeckProperties, RemainingCut)
    For i% = 1 To CutDepth
        For z% = 1 To DeckProperties
            TopCut(z%, i%) = Deck(z%, ProtectedBlock + i%)
        Next z%
    Next i%
    For j% = 1 To RemainingCut
        For z% = 1 To DeckProperties
            BottomCut(z%, j%) = Deck(z%, ProtectedBlock + CutDepth + j%)
        Next z%
    Next j%
    For p% = 1 To ProtectedBlock
        For z% = 1 To DeckProperties
            ChangedDeck(z%, p%) = Deck(z%, p%)
        Next z%
    Next p%
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
                TopIndex = TopIndex + 1
            Else
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, ProtectedBlock + k%) = BottomCut(z%, BottomIndex)
                Next z%
                BottomIndex = BottomIndex + 1
            End If
        Else
            If BottomIndex <= RemainingCut Then
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, ProtectedBlock + k%) = BottomCut(z%, BottomIndex)
                Next z%
                BottomIndex = BottomIndex + 1
            Else
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, ProtectedBlock + k%) = TopCut(z%, TopIndex)
                Next z%
                TopIndex = TopIndex + 1
            End If
        End If
    Next k%
    For m% = 1 To DeckCount
        For z% = 1 To DeckProperties
            Deck(z%, m%) = ChangedDeck(z%, m%)
        Next z%
    Next m%
' third section cuts the protected cards from the top
' to the bottom
CutDepth = riffleshufflenumber
s = 1
For j% = CutDepth + 1 To DeckCount
    For z% = 1 To DeckProperties
        ChangedDeck(z%, s) = Deck(z%, j%)
    Next z%
    s = s + 1
Next j%
For k% = 1 To CutDepth
    For z% = 1 To DeckProperties
        ChangedDeck(z%, s) = Deck(z%, k%)
    Next z%
    s = s + 1
Next k%
For m% = 1 To DeckCount
    For z% = 1 To DeckProperties
        Deck(z%, m%) = ChangedDeck(z%, m%)
    Next z%
Next m%
ShowCards
End Sub

Public Sub RiffleShuffleButton_Click()
Dim pReverse As String
If RiffleReverseTopCheck.Value = 1 Then
    pReverse = "T"
ElseIf RiffleReverseBottomCheck.Value = 1 Then
    pReverse = "B"
Else
    pReverse = "X"
End If
If RiffleProtectOption.Value = True And _
    (Not IsNumeric(RiffleNumberTextBox.Text) Or _
    Val(RiffleNumberTextBox.Text) < 1 Or _
    Val(RiffleNumberTextBox.Text) > 52) Then
    RiffleNumberTextBox.Text = Empty
    MsgBox "Please enter a valid number of cards (1 to 52)" & Chr(13) _
        & "in the 'Riffle: Number' Input Box"
    Exit Sub
End If
If RiffleProtectOption = True And _
    RiffleLocationCombo.ListIndex = -1 Then
    RiffleLocationCombo.Text = "Select area"
    MsgBox "Please select a valid response from" & Chr(13) _
        & "the 'Riffle: Location' Dropdown Box"
    Exit Sub
End If
If RiffleRandomOption.Value = True Then
    Call RiffleShuffle(pReverse)
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "RiffleShuffle(" & pReverse & ")"
        SessionListBox.AddItem SessionCommand
        SessionStatusUpdate (0)
    End If
ElseIf RiffleProtectOption.Value = True And _
    RiffleLocationCombo.ListIndex = 0 Then
    Call RiffleShuffleTop(Val(RiffleNumberTextBox.Text), pReverse)
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "RiffleShuffleTop(" & _
            RiffleNumberTextBox.Text & ", " & pReverse & ")"
        SessionListBox.AddItem SessionCommand
        SessionStatusUpdate (0)
    End If
ElseIf RiffleProtectOption.Value = True And _
    RiffleLocationCombo.ListIndex = 1 Then
    Call RiffleShuffleBottom(Val(RiffleNumberTextBox.Text), pReverse)
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "RiffleShuffleBottom(" & _
            RiffleNumberTextBox.Text & ", " & pReverse & ")"
        SessionListBox.AddItem SessionCommand
        SessionStatusUpdate (0)
    End If
End If
End Sub

Public Sub RiffleShuffleTop(riffleshufflenumber, paramreverse)
    Dim pReverse As String
    pReverse = paramreverse
    Dim pCutStartCard As Integer
    Dim pCutEndCard As Integer
    ProtectedBlock = Val(riffleshufflenumber)
    RifflePortion = DeckCount - ProtectedBlock
    CutError = Int(0.2 * RifflePortion) + 1
    CutDepth = Int(Rnd * CutError) + Int((RifflePortion - CutError) / 2)
      'approximately half the unprotected deck for a riffle shuffle
    If CutDepth < 0 Then
        CutDepth = 0
            ' need to avoid an error
    End If
    RemainingCut = RifflePortion - CutDepth
        'reverse a block if required
    If pReverse = "T" Then
        pCutStartCard = 1
        pCutEndCard = CutDepth + ProtectedBlock
    ElseIf pReverse = "B" Then
        pCutStartCard = CutDepth + ProtectedBlock + 1
        pCutEndCard = 52
    End If
    If Not pReverse = "X" Then
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
    
    
    ReDim TopCut(DeckProperties, CutDepth)
    ReDim BottomCut(DeckProperties, RemainingCut)
    For i% = 1 To CutDepth
        For z% = 1 To DeckProperties
            TopCut(z%, i%) = Deck(z%, ProtectedBlock + i%)
        Next z%
    Next i%
    For j% = 1 To RemainingCut
        For z% = 1 To DeckProperties
            BottomCut(z%, j%) = Deck(z%, ProtectedBlock + CutDepth + j%)
        Next z%
    Next j%
    For p% = 1 To ProtectedBlock
        For z% = 1 To DeckProperties
            ChangedDeck(z%, p%) = Deck(z%, p%)
        Next z%
    Next p%
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
                TopIndex = TopIndex + 1
            Else
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, ProtectedBlock + k%) = BottomCut(z%, BottomIndex)
                Next z%
                BottomIndex = BottomIndex + 1
            End If
        Else
            If BottomIndex <= RemainingCut Then
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, ProtectedBlock + k%) = BottomCut(z%, BottomIndex)
                Next z%
                BottomIndex = BottomIndex + 1
            Else
                For z% = 1 To DeckProperties
                    ChangedDeck(z%, ProtectedBlock + k%) = TopCut(z%, TopIndex)
                Next z%
                TopIndex = TopIndex + 1
            End If
        End If
    Next k%
    For m% = 1 To DeckCount
        For z% = 1 To DeckProperties
            Deck(z%, m%) = ChangedDeck(z%, m%)
        Next z%
    Next m%
    ShowCards
End Sub

Public Sub SetDeckJoyalCHSD()
DeckCount = 52
StackedDeck(1, 1) = 1
StackedDeck(2, 1) = "JH"
StackedDeck(1, 2) = 2
StackedDeck(2, 2) = "6C"
StackedDeck(1, 3) = 3
StackedDeck(2, 3) = "6H"
StackedDeck(1, 4) = 4
StackedDeck(2, 4) = "4C"
StackedDeck(1, 5) = 5
StackedDeck(2, 5) = "10D"
StackedDeck(1, 6) = 6
StackedDeck(2, 6) = "AD"
StackedDeck(1, 7) = 7
StackedDeck(2, 7) = "7C"
StackedDeck(1, 8) = 8
StackedDeck(2, 8) = "4H"
StackedDeck(1, 9) = 9
StackedDeck(2, 9) = "9C"
StackedDeck(1, 10) = 10
StackedDeck(2, 10) = "5D"
StackedDeck(1, 11) = 11
StackedDeck(2, 11) = "QH"
StackedDeck(1, 12) = 12
StackedDeck(2, 12) = "AS"
StackedDeck(1, 13) = 13
StackedDeck(2, 13) = "KC"
StackedDeck(1, 14) = 14
StackedDeck(2, 14) = "7H"
StackedDeck(1, 15) = 15
StackedDeck(2, 15) = "10S"
StackedDeck(1, 16) = 16
StackedDeck(2, 16) = "4S"
StackedDeck(1, 17) = 17
StackedDeck(2, 17) = "JS"
StackedDeck(1, 18) = 18
StackedDeck(2, 18) = "9H"
StackedDeck(1, 19) = 19
StackedDeck(2, 19) = "KD"
StackedDeck(1, 20) = 20
StackedDeck(2, 20) = "5S"
StackedDeck(1, 21) = 21
StackedDeck(2, 21) = "7S"
StackedDeck(1, 22) = 22
StackedDeck(2, 22) = "2C"
StackedDeck(1, 23) = 23
StackedDeck(2, 23) = "QC"
StackedDeck(1, 24) = 24
StackedDeck(2, 24) = "AH"
StackedDeck(1, 25) = 25
StackedDeck(2, 25) = "10H"
StackedDeck(1, 26) = 26
StackedDeck(2, 26) = "6S"
StackedDeck(1, 27) = 27
StackedDeck(2, 27) = "9S"
StackedDeck(1, 28) = 28
StackedDeck(2, 28) = "7D"
StackedDeck(1, 29) = 29
StackedDeck(2, 29) = "QD"
StackedDeck(1, 30) = 30
StackedDeck(2, 30) = "5H"
StackedDeck(1, 31) = 31
StackedDeck(2, 31) = "KH"
StackedDeck(1, 32) = 32
StackedDeck(2, 32) = "4D"
StackedDeck(1, 33) = 33
StackedDeck(2, 33) = "3C"
StackedDeck(1, 34) = 34
StackedDeck(2, 34) = "3H"
StackedDeck(1, 35) = 35
StackedDeck(2, 35) = "10C"
StackedDeck(1, 36) = 36
StackedDeck(2, 36) = "9D"
StackedDeck(1, 37) = 37
StackedDeck(2, 37) = "QS"
StackedDeck(1, 38) = 38
StackedDeck(2, 38) = "3S"
StackedDeck(1, 39) = 39
StackedDeck(2, 39) = "3D"
StackedDeck(1, 40) = 40
StackedDeck(2, 40) = "2H"
StackedDeck(1, 41) = 41
StackedDeck(2, 41) = "8C"
StackedDeck(1, 42) = 42
StackedDeck(2, 42) = "2S"
StackedDeck(1, 43) = 43
StackedDeck(2, 43) = "JC"
StackedDeck(1, 44) = 44
StackedDeck(2, 44) = "2D"
StackedDeck(1, 45) = 45
StackedDeck(2, 45) = "8H"
StackedDeck(1, 46) = 46
StackedDeck(2, 46) = "8S"
StackedDeck(1, 47) = 47
StackedDeck(2, 47) = "KS"
StackedDeck(1, 48) = 48
StackedDeck(2, 48) = "AC"
StackedDeck(1, 49) = 49
StackedDeck(2, 49) = "JD"
StackedDeck(1, 50) = 50
StackedDeck(2, 50) = "5C"
StackedDeck(1, 51) = 51
StackedDeck(2, 51) = "8D"
StackedDeck(1, 52) = 52
StackedDeck(2, 52) = "6D"
CutDepth = 0
Cls
For i% = 1 To 52
    Deck(1, i%) = StackedDeck(1, i%)
    Deck(2, i%) = StackedDeck(2, i%)
    Deck(6, i%) = False
Next i%
ShowCards
End Sub
Public Sub SetDeckNewBicycle()
DeckCount = 52
StackedDeck(1, 1) = 1
StackedDeck(2, 1) = "AH"
StackedDeck(1, 2) = 2
StackedDeck(2, 2) = "2H"
StackedDeck(1, 3) = 3
StackedDeck(2, 3) = "3H"
StackedDeck(1, 4) = 4
StackedDeck(2, 4) = "4H"
StackedDeck(1, 5) = 5
StackedDeck(2, 5) = "5H"
StackedDeck(1, 6) = 6
StackedDeck(2, 6) = "6H"
StackedDeck(1, 7) = 7
StackedDeck(2, 7) = "7H"
StackedDeck(1, 8) = 8
StackedDeck(2, 8) = "8H"
StackedDeck(1, 9) = 9
StackedDeck(2, 9) = "9H"
StackedDeck(1, 10) = 10
StackedDeck(2, 10) = "10H"
StackedDeck(1, 11) = 11
StackedDeck(2, 11) = "JH"
StackedDeck(1, 12) = 12
StackedDeck(2, 12) = "QH"
StackedDeck(1, 13) = 13
StackedDeck(2, 13) = "KH"
StackedDeck(1, 14) = 14
StackedDeck(2, 14) = "AC"
StackedDeck(1, 15) = 15
StackedDeck(2, 15) = "2C"
StackedDeck(1, 16) = 16
StackedDeck(2, 16) = "3C"
StackedDeck(1, 17) = 17
StackedDeck(2, 17) = "4C"
StackedDeck(1, 18) = 18
StackedDeck(2, 18) = "5C"
StackedDeck(1, 19) = 19
StackedDeck(2, 19) = "6C"
StackedDeck(1, 20) = 20
StackedDeck(2, 20) = "7C"
StackedDeck(1, 21) = 21
StackedDeck(2, 21) = "8C"
StackedDeck(1, 22) = 22
StackedDeck(2, 22) = "9C"
StackedDeck(1, 23) = 23
StackedDeck(2, 23) = "10C"
StackedDeck(1, 24) = 24
StackedDeck(2, 24) = "JC"
StackedDeck(1, 25) = 25
StackedDeck(2, 25) = "QC"
StackedDeck(1, 26) = 26
StackedDeck(2, 26) = "KC"
StackedDeck(1, 27) = 27
StackedDeck(2, 27) = "KD"
StackedDeck(1, 28) = 28
StackedDeck(2, 28) = "QD"
StackedDeck(1, 29) = 29
StackedDeck(2, 29) = "JD"
StackedDeck(1, 30) = 30
StackedDeck(2, 30) = "10D"
StackedDeck(1, 31) = 31
StackedDeck(2, 31) = "9D"
StackedDeck(1, 32) = 32
StackedDeck(2, 32) = "8D"
StackedDeck(1, 33) = 33
StackedDeck(2, 33) = "7D"
StackedDeck(1, 34) = 34
StackedDeck(2, 34) = "6D"
StackedDeck(1, 35) = 35
StackedDeck(2, 35) = "5D"
StackedDeck(1, 36) = 36
StackedDeck(2, 36) = "4D"
StackedDeck(1, 37) = 37
StackedDeck(2, 37) = "3D"
StackedDeck(1, 38) = 38
StackedDeck(2, 38) = "2D"
StackedDeck(1, 39) = 39
StackedDeck(2, 39) = "AD"
StackedDeck(1, 40) = 40
StackedDeck(2, 40) = "KS"
StackedDeck(1, 41) = 41
StackedDeck(2, 41) = "QS"
StackedDeck(1, 42) = 42
StackedDeck(2, 42) = "JS"
StackedDeck(1, 43) = 43
StackedDeck(2, 43) = "10S"
StackedDeck(1, 44) = 44
StackedDeck(2, 44) = "9S"
StackedDeck(1, 45) = 45
StackedDeck(2, 45) = "8S"
StackedDeck(1, 46) = 46
StackedDeck(2, 46) = "7S"
StackedDeck(1, 47) = 47
StackedDeck(2, 47) = "6S"
StackedDeck(1, 48) = 48
StackedDeck(2, 48) = "5S"
StackedDeck(1, 49) = 49
StackedDeck(2, 49) = "4S"
StackedDeck(1, 50) = 50
StackedDeck(2, 50) = "3S"
StackedDeck(1, 51) = 51
StackedDeck(2, 51) = "2S"
StackedDeck(1, 52) = 52
StackedDeck(2, 52) = "AS"
CutDepth = 0
Cls
For i% = 1 To 52
    Deck(1, i%) = StackedDeck(1, i%)
    Deck(2, i%) = StackedDeck(2, i%)
    Deck(6, i%) = False
Next i%
ShowCards
End Sub

Public Sub SetDeckNewFournier()
DeckCount = 52
StackedDeck(1, 1) = 1
StackedDeck(2, 1) = "AS"
StackedDeck(1, 2) = 2
StackedDeck(2, 2) = "2S"
StackedDeck(1, 3) = 3
StackedDeck(2, 3) = "3S"
StackedDeck(1, 4) = 4
StackedDeck(2, 4) = "4S"
StackedDeck(1, 5) = 5
StackedDeck(2, 5) = "5S"
StackedDeck(1, 6) = 6
StackedDeck(2, 6) = "6S"
StackedDeck(1, 7) = 7
StackedDeck(2, 7) = "7S"
StackedDeck(1, 8) = 8
StackedDeck(2, 8) = "8S"
StackedDeck(1, 9) = 9
StackedDeck(2, 9) = "9S"
StackedDeck(1, 10) = 10
StackedDeck(2, 10) = "10S"
StackedDeck(1, 11) = 11
StackedDeck(2, 11) = "JS"
StackedDeck(1, 12) = 12
StackedDeck(2, 12) = "QS"
StackedDeck(1, 13) = 13
StackedDeck(2, 13) = "KS"
StackedDeck(1, 14) = 14
StackedDeck(2, 14) = "AH"
StackedDeck(1, 15) = 15
StackedDeck(2, 15) = "2H"
StackedDeck(1, 16) = 16
StackedDeck(2, 16) = "3H"
StackedDeck(1, 17) = 17
StackedDeck(2, 17) = "4H"
StackedDeck(1, 18) = 18
StackedDeck(2, 18) = "5H"
StackedDeck(1, 19) = 19
StackedDeck(2, 19) = "6H"
StackedDeck(1, 20) = 20
StackedDeck(2, 20) = "7H"
StackedDeck(1, 21) = 21
StackedDeck(2, 21) = "8H"
StackedDeck(1, 22) = 22
StackedDeck(2, 22) = "9H"
StackedDeck(1, 23) = 23
StackedDeck(2, 23) = "10H"
StackedDeck(1, 24) = 24
StackedDeck(2, 24) = "JH"
StackedDeck(1, 25) = 25
StackedDeck(2, 25) = "QH"
StackedDeck(1, 26) = 26
StackedDeck(2, 26) = "KH"
StackedDeck(1, 27) = 27
StackedDeck(2, 27) = "KD"
StackedDeck(1, 28) = 28
StackedDeck(2, 28) = "QD"
StackedDeck(1, 29) = 29
StackedDeck(2, 29) = "JD"
StackedDeck(1, 30) = 30
StackedDeck(2, 30) = "10D"
StackedDeck(1, 31) = 31
StackedDeck(2, 31) = "9D"
StackedDeck(1, 32) = 32
StackedDeck(2, 32) = "8D"
StackedDeck(1, 33) = 33
StackedDeck(2, 33) = "7D"
StackedDeck(1, 34) = 34
StackedDeck(2, 34) = "6D"
StackedDeck(1, 35) = 35
StackedDeck(2, 35) = "5D"
StackedDeck(1, 36) = 36
StackedDeck(2, 36) = "4D"
StackedDeck(1, 37) = 37
StackedDeck(2, 37) = "3D"
StackedDeck(1, 38) = 38
StackedDeck(2, 38) = "2D"
StackedDeck(1, 39) = 39
StackedDeck(2, 39) = "AD"
StackedDeck(1, 40) = 40
StackedDeck(2, 40) = "KC"
StackedDeck(1, 41) = 41
StackedDeck(2, 41) = "QC"
StackedDeck(1, 42) = 42
StackedDeck(2, 42) = "JC"
StackedDeck(1, 43) = 43
StackedDeck(2, 43) = "10C"
StackedDeck(1, 44) = 44
StackedDeck(2, 44) = "9C"
StackedDeck(1, 45) = 45
StackedDeck(2, 45) = "8C"
StackedDeck(1, 46) = 46
StackedDeck(2, 46) = "7C"
StackedDeck(1, 47) = 47
StackedDeck(2, 47) = "6C"
StackedDeck(1, 48) = 48
StackedDeck(2, 48) = "5C"
StackedDeck(1, 49) = 49
StackedDeck(2, 49) = "4C"
StackedDeck(1, 50) = 50
StackedDeck(2, 50) = "3C"
StackedDeck(1, 51) = 51
StackedDeck(2, 51) = "2C"
StackedDeck(1, 52) = 52
StackedDeck(2, 52) = "AC"
CutDepth = 0
Cls
For i% = 1 To 52
    Deck(1, i%) = StackedDeck(1, i%)
    Deck(2, i%) = StackedDeck(2, i%)
    Deck(6, i%) = False
Next i%
ShowCards
End Sub

Public Sub SetDeckJoyalSHCD()
DeckCount = 52
StackedDeck(1, 1) = 1
StackedDeck(2, 1) = "JH"
StackedDeck(1, 2) = 2
StackedDeck(2, 2) = "6S"
StackedDeck(1, 3) = 3
StackedDeck(2, 3) = "6H"
StackedDeck(1, 4) = 4
StackedDeck(2, 4) = "4S"
StackedDeck(1, 5) = 5
StackedDeck(2, 5) = "10D"
StackedDeck(1, 6) = 6
StackedDeck(2, 6) = "AD"
StackedDeck(1, 7) = 7
StackedDeck(2, 7) = "7S"
StackedDeck(1, 8) = 8
StackedDeck(2, 8) = "4H"
StackedDeck(1, 9) = 9
StackedDeck(2, 9) = "9S"
StackedDeck(1, 10) = 10
StackedDeck(2, 10) = "5D"
StackedDeck(1, 11) = 11
StackedDeck(2, 11) = "QH"
StackedDeck(1, 12) = 12
StackedDeck(2, 12) = "AC"
StackedDeck(1, 13) = 13
StackedDeck(2, 13) = "KC"
StackedDeck(1, 14) = 14
StackedDeck(2, 14) = "7H"
StackedDeck(1, 15) = 15
StackedDeck(2, 15) = "10C"
StackedDeck(1, 16) = 16
StackedDeck(2, 16) = "4C"
StackedDeck(1, 17) = 17
StackedDeck(2, 17) = "JS"
StackedDeck(1, 18) = 18
StackedDeck(2, 18) = "9H"
StackedDeck(1, 19) = 19
StackedDeck(2, 19) = "KD"
StackedDeck(1, 20) = 20
StackedDeck(2, 20) = "5C"
StackedDeck(1, 21) = 21
StackedDeck(2, 21) = "7C"
StackedDeck(1, 22) = 22
StackedDeck(2, 22) = "2S"
StackedDeck(1, 23) = 23
StackedDeck(2, 23) = "QC"
StackedDeck(1, 24) = 24
StackedDeck(2, 24) = "AH"
StackedDeck(1, 25) = 25
StackedDeck(2, 25) = "10H"
StackedDeck(1, 26) = 26
StackedDeck(2, 26) = "6C"
StackedDeck(1, 27) = 27
StackedDeck(2, 27) = "9C"
StackedDeck(1, 28) = 28
StackedDeck(2, 28) = "7D"
StackedDeck(1, 29) = 29
StackedDeck(2, 29) = "QD"
StackedDeck(1, 30) = 30
StackedDeck(2, 30) = "5H"
StackedDeck(1, 31) = 31
StackedDeck(2, 31) = "KH"
StackedDeck(1, 32) = 32
StackedDeck(2, 32) = "4D"
StackedDeck(1, 33) = 33
StackedDeck(2, 33) = "3S"
StackedDeck(1, 34) = 34
StackedDeck(2, 34) = "3H"
StackedDeck(1, 35) = 35
StackedDeck(2, 35) = "10S"
StackedDeck(1, 36) = 36
StackedDeck(2, 36) = "9D"
StackedDeck(1, 37) = 37
StackedDeck(2, 37) = "QS"
StackedDeck(1, 38) = 38
StackedDeck(2, 38) = "3C"
StackedDeck(1, 39) = 39
StackedDeck(2, 39) = "3D"
StackedDeck(1, 40) = 40
StackedDeck(2, 40) = "2H"
StackedDeck(1, 41) = 41
StackedDeck(2, 41) = "8S"
StackedDeck(1, 42) = 42
StackedDeck(2, 42) = "2C"
StackedDeck(1, 43) = 43
StackedDeck(2, 43) = "JC"
StackedDeck(1, 44) = 44
StackedDeck(2, 44) = "2D"
StackedDeck(1, 45) = 45
StackedDeck(2, 45) = "8H"
StackedDeck(1, 46) = 46
StackedDeck(2, 46) = "8C"
StackedDeck(1, 47) = 47
StackedDeck(2, 47) = "KS"
StackedDeck(1, 48) = 48
StackedDeck(2, 48) = "AS"
StackedDeck(1, 49) = 49
StackedDeck(2, 49) = "JD"
StackedDeck(1, 50) = 50
StackedDeck(2, 50) = "5S"
StackedDeck(1, 51) = 51
StackedDeck(2, 51) = "8D"
StackedDeck(1, 52) = 52
StackedDeck(2, 52) = "6D"
CutDepth = 0
Cls
For i% = 1 To 52
    Deck(1, i%) = StackedDeck(1, i%)
    Deck(2, i%) = StackedDeck(2, i%)
    Deck(6, i%) = False
Next i%
ShowCards
End Sub
Public Sub SetDeckNikola()
DeckCount = 52
StackedDeck(1, 1) = 1
StackedDeck(2, 1) = "6D"
StackedDeck(1, 2) = 2
StackedDeck(2, 2) = "5C"
StackedDeck(1, 3) = 3
StackedDeck(2, 3) = "KC"
StackedDeck(1, 4) = 4
StackedDeck(2, 4) = "JH"
StackedDeck(1, 5) = 5
StackedDeck(2, 5) = "5S"
StackedDeck(1, 6) = 6
StackedDeck(2, 6) = "9D"
StackedDeck(1, 7) = 7
StackedDeck(2, 7) = "9S"
StackedDeck(1, 8) = 8
StackedDeck(2, 8) = "QH"
StackedDeck(1, 9) = 9
StackedDeck(2, 9) = "3C"
StackedDeck(1, 10) = 10
StackedDeck(2, 10) = "10C"
StackedDeck(1, 11) = 11
StackedDeck(2, 11) = "KS"
StackedDeck(1, 12) = 12
StackedDeck(2, 12) = "AH"
StackedDeck(1, 13) = 13
StackedDeck(2, 13) = "4D"
StackedDeck(1, 14) = 14
StackedDeck(2, 14) = "JD"
StackedDeck(1, 15) = 15
StackedDeck(2, 15) = "KD"
StackedDeck(1, 16) = 16
StackedDeck(2, 16) = "KH"
StackedDeck(1, 17) = 17
StackedDeck(2, 17) = "2D"
StackedDeck(1, 18) = 18
StackedDeck(2, 18) = "QC"
StackedDeck(1, 19) = 19
StackedDeck(2, 19) = "9C"
StackedDeck(1, 20) = 20
StackedDeck(2, 20) = "10H"
StackedDeck(1, 21) = 21
StackedDeck(2, 21) = "8D"
StackedDeck(1, 22) = 22
StackedDeck(2, 22) = "2C"
StackedDeck(1, 23) = 23
StackedDeck(2, 23) = "AC"
StackedDeck(1, 24) = 24
StackedDeck(2, 24) = "7H"
StackedDeck(1, 25) = 25
StackedDeck(2, 25) = "7C"
StackedDeck(1, 26) = 26
StackedDeck(2, 26) = "4S"
StackedDeck(1, 27) = 27
StackedDeck(2, 27) = "7S"
StackedDeck(1, 28) = 28
StackedDeck(2, 28) = "9H"
StackedDeck(1, 29) = 29
StackedDeck(2, 29) = "8S"
StackedDeck(1, 30) = 30
StackedDeck(2, 30) = "6S"
StackedDeck(1, 31) = 31
StackedDeck(2, 31) = "6C"
StackedDeck(1, 32) = 32
StackedDeck(2, 32) = "2H"
StackedDeck(1, 33) = 33
StackedDeck(2, 33) = "AS"
StackedDeck(1, 34) = 34
StackedDeck(2, 34) = "JS"
StackedDeck(1, 35) = 35
StackedDeck(2, 35) = "4C"
StackedDeck(1, 36) = 36
StackedDeck(2, 36) = "5H"
StackedDeck(1, 37) = 37
StackedDeck(2, 37) = "10S"
StackedDeck(1, 38) = 38
StackedDeck(2, 38) = "AD"
StackedDeck(1, 39) = 39
StackedDeck(2, 39) = "JC"
StackedDeck(1, 40) = 40
StackedDeck(2, 40) = "4H"
StackedDeck(1, 41) = 41
StackedDeck(2, 41) = "2S"
StackedDeck(1, 42) = 42
StackedDeck(2, 42) = "7D"
StackedDeck(1, 43) = 43
StackedDeck(2, 43) = "QS"
StackedDeck(1, 44) = 44
StackedDeck(2, 44) = "3H"
StackedDeck(1, 45) = 45
StackedDeck(2, 45) = "3S"
StackedDeck(1, 46) = 46
StackedDeck(2, 46) = "8C"
StackedDeck(1, 47) = 47
StackedDeck(2, 47) = "10D"
StackedDeck(1, 48) = 48
StackedDeck(2, 48) = "6H"
StackedDeck(1, 49) = 49
StackedDeck(2, 49) = "5D"
StackedDeck(1, 50) = 50
StackedDeck(2, 50) = "3D"
StackedDeck(1, 51) = 51
StackedDeck(2, 51) = "QD"
StackedDeck(1, 52) = 52
StackedDeck(2, 52) = "8H"
CutDepth = 0
Cls
For i% = 1 To 52
    Deck(1, i%) = StackedDeck(1, i%)
    Deck(2, i%) = StackedDeck(2, i%)
    Deck(6, i%) = False
Next i%
ShowCards
End Sub

Public Sub SetDeckIreland()
DeckCount = 52
StackedDeck(1, 1) = 1
StackedDeck(2, 1) = "7H"
StackedDeck(1, 2) = 2
StackedDeck(2, 2) = "4H"
StackedDeck(1, 3) = 3
StackedDeck(2, 3) = "KH"
StackedDeck(1, 4) = 4
StackedDeck(2, 4) = "2C"
StackedDeck(1, 5) = 5
StackedDeck(2, 5) = "10S"
StackedDeck(1, 6) = 6
StackedDeck(2, 6) = "6C"
StackedDeck(1, 7) = 7
StackedDeck(2, 7) = "8H"
StackedDeck(1, 8) = 8
StackedDeck(2, 8) = "QD"
StackedDeck(1, 9) = 9
StackedDeck(2, 9) = "2D"
StackedDeck(1, 10) = 10
StackedDeck(2, 10) = "QS"
StackedDeck(1, 11) = 11
StackedDeck(2, 11) = "5D"
StackedDeck(1, 12) = 12
StackedDeck(2, 12) = "6H"
StackedDeck(1, 13) = 13
StackedDeck(2, 13) = "KC"
StackedDeck(1, 14) = 14
StackedDeck(2, 14) = "7S"
StackedDeck(1, 15) = 15
StackedDeck(2, 15) = "JS"
StackedDeck(1, 16) = 16
StackedDeck(2, 16) = "4S"
StackedDeck(1, 17) = 17
StackedDeck(2, 17) = "QH"
StackedDeck(1, 18) = 18
StackedDeck(2, 18) = "QC"
StackedDeck(1, 19) = 19
StackedDeck(2, 19) = "2S"
StackedDeck(1, 20) = 20
StackedDeck(2, 20) = "KS"
StackedDeck(1, 21) = 21
StackedDeck(2, 21) = "3H"
StackedDeck(1, 22) = 22
StackedDeck(2, 22) = "JH"
StackedDeck(1, 23) = 23
StackedDeck(2, 23) = "KD"
StackedDeck(1, 24) = 24
StackedDeck(2, 24) = "2H"
StackedDeck(1, 25) = 25
StackedDeck(2, 25) = "AS"
StackedDeck(1, 26) = 26
StackedDeck(2, 26) = "6S"
StackedDeck(1, 27) = 27
StackedDeck(2, 27) = "AC"
StackedDeck(1, 28) = 28
StackedDeck(2, 28) = "9C"
StackedDeck(1, 29) = 29
StackedDeck(2, 29) = "3C"
StackedDeck(1, 30) = 30
StackedDeck(2, 30) = "AD"
StackedDeck(1, 31) = 31
StackedDeck(2, 31) = "JC"
StackedDeck(1, 32) = 32
StackedDeck(2, 32) = "8D"
StackedDeck(1, 33) = 33
StackedDeck(2, 33) = "9H"
StackedDeck(1, 34) = 34
StackedDeck(2, 34) = "8C"
StackedDeck(1, 35) = 35
StackedDeck(2, 35) = "9S"
StackedDeck(1, 36) = 36
StackedDeck(2, 36) = "AH"
StackedDeck(1, 37) = 37
StackedDeck(2, 37) = "9D"
StackedDeck(1, 38) = 38
StackedDeck(2, 38) = "10H"
StackedDeck(1, 39) = 39
StackedDeck(2, 39) = "8S"
StackedDeck(1, 40) = 40
StackedDeck(2, 40) = "6D"
StackedDeck(1, 41) = 41
StackedDeck(2, 41) = "3S"
StackedDeck(1, 42) = 42
StackedDeck(2, 42) = "5H"
StackedDeck(1, 43) = 43
StackedDeck(2, 43) = "5C"
StackedDeck(1, 44) = 44
StackedDeck(2, 44) = "4D"
StackedDeck(1, 45) = 45
StackedDeck(2, 45) = "10D"
StackedDeck(1, 46) = 46
StackedDeck(2, 46) = "7C"
StackedDeck(1, 47) = 47
StackedDeck(2, 47) = "3D"
StackedDeck(1, 48) = 48
StackedDeck(2, 48) = "4C"
StackedDeck(1, 49) = 49
StackedDeck(2, 49) = "7D"
StackedDeck(1, 50) = 50
StackedDeck(2, 50) = "10C"
StackedDeck(1, 51) = 51
StackedDeck(2, 51) = "JD"
StackedDeck(1, 52) = 52
StackedDeck(2, 52) = "5S"
CutDepth = 0
Cls
For i% = 1 To 52
    Deck(1, i%) = StackedDeck(1, i%)
    Deck(2, i%) = StackedDeck(2, i%)
    Deck(6, i%) = False
Next i%
ShowCards
End Sub

Public Sub SetDeckAronson()
DeckCount = 52
StackedDeck(1, 1) = 1
StackedDeck(2, 1) = "JS"
StackedDeck(1, 2) = 2
StackedDeck(2, 2) = "KC"
StackedDeck(1, 3) = 3
StackedDeck(2, 3) = "5C"
StackedDeck(1, 4) = 4
StackedDeck(2, 4) = "2H"
StackedDeck(1, 5) = 5
StackedDeck(2, 5) = "9S"
StackedDeck(1, 6) = 6
StackedDeck(2, 6) = "AS"
StackedDeck(1, 7) = 7
StackedDeck(2, 7) = "3H"
StackedDeck(1, 8) = 8
StackedDeck(2, 8) = "6C"
StackedDeck(1, 9) = 9
StackedDeck(2, 9) = "8D"
StackedDeck(1, 10) = 10
StackedDeck(2, 10) = "AC"
StackedDeck(1, 11) = 11
StackedDeck(2, 11) = "10S"
StackedDeck(1, 12) = 12
StackedDeck(2, 12) = "5H"
StackedDeck(1, 13) = 13
StackedDeck(2, 13) = "2D"
StackedDeck(1, 14) = 14
StackedDeck(2, 14) = "KD"
StackedDeck(1, 15) = 15
StackedDeck(2, 15) = "7D"
StackedDeck(1, 16) = 16
StackedDeck(2, 16) = "8C"
StackedDeck(1, 17) = 17
StackedDeck(2, 17) = "3S"
StackedDeck(1, 18) = 18
StackedDeck(2, 18) = "AD"
StackedDeck(1, 19) = 19
StackedDeck(2, 19) = "7S"
StackedDeck(1, 20) = 20
StackedDeck(2, 20) = "5S"
StackedDeck(1, 21) = 21
StackedDeck(2, 21) = "QD"
StackedDeck(1, 22) = 22
StackedDeck(2, 22) = "AH"
StackedDeck(1, 23) = 23
StackedDeck(2, 23) = "8S"
StackedDeck(1, 24) = 24
StackedDeck(2, 24) = "3D"
StackedDeck(1, 25) = 25
StackedDeck(2, 25) = "7H"
StackedDeck(1, 26) = 26
StackedDeck(2, 26) = "QH"
StackedDeck(1, 27) = 27
StackedDeck(2, 27) = "5D"
StackedDeck(1, 28) = 28
StackedDeck(2, 28) = "7C"
StackedDeck(1, 29) = 29
StackedDeck(2, 29) = "4H"
StackedDeck(1, 30) = 30
StackedDeck(2, 30) = "KH"
StackedDeck(1, 31) = 31
StackedDeck(2, 31) = "4D"
StackedDeck(1, 32) = 32
StackedDeck(2, 32) = "10D"
StackedDeck(1, 33) = 33
StackedDeck(2, 33) = "JC"
StackedDeck(1, 34) = 34
StackedDeck(2, 34) = "JH"
StackedDeck(1, 35) = 35
StackedDeck(2, 35) = "10C"
StackedDeck(1, 36) = 36
StackedDeck(2, 36) = "JD"
StackedDeck(1, 37) = 37
StackedDeck(2, 37) = "4S"
StackedDeck(1, 38) = 38
StackedDeck(2, 38) = "10H"
StackedDeck(1, 39) = 39
StackedDeck(2, 39) = "6H"
StackedDeck(1, 40) = 40
StackedDeck(2, 40) = "3C"
StackedDeck(1, 41) = 41
StackedDeck(2, 41) = "2S"
StackedDeck(1, 42) = 42
StackedDeck(2, 42) = "9H"
StackedDeck(1, 43) = 43
StackedDeck(2, 43) = "KS"
StackedDeck(1, 44) = 44
StackedDeck(2, 44) = "6S"
StackedDeck(1, 45) = 45
StackedDeck(2, 45) = "4C"
StackedDeck(1, 46) = 46
StackedDeck(2, 46) = "8H"
StackedDeck(1, 47) = 47
StackedDeck(2, 47) = "9C"
StackedDeck(1, 48) = 48
StackedDeck(2, 48) = "QS"
StackedDeck(1, 49) = 49
StackedDeck(2, 49) = "6D"
StackedDeck(1, 50) = 50
StackedDeck(2, 50) = "QC"
StackedDeck(1, 51) = 51
StackedDeck(2, 51) = "2C"
StackedDeck(1, 52) = 52
StackedDeck(2, 52) = "9D"
CutDepth = 0
Cls
For i% = 1 To 52
    Deck(1, i%) = StackedDeck(1, i%)
    Deck(2, i%) = StackedDeck(2, i%)
    Deck(6, i%) = False
Next i%
ShowCards
End Sub
Public Sub SetDeckEightKings()
DeckCount = 52
StackedDeck(1, 1) = 1
StackedDeck(2, 1) = "8C"
StackedDeck(1, 2) = 2
StackedDeck(2, 2) = "KH"
StackedDeck(1, 3) = 3
StackedDeck(2, 3) = "3S"
StackedDeck(1, 4) = 4
StackedDeck(2, 4) = "10D"
StackedDeck(1, 5) = 5
StackedDeck(2, 5) = "2C"
StackedDeck(1, 6) = 6
StackedDeck(2, 6) = "7H"
StackedDeck(1, 7) = 7
StackedDeck(2, 7) = "9S"
StackedDeck(1, 8) = 8
StackedDeck(2, 8) = "5D"
StackedDeck(1, 9) = 9
StackedDeck(2, 9) = "QC"
StackedDeck(1, 10) = 10
StackedDeck(2, 10) = "4H"
StackedDeck(1, 11) = 11
StackedDeck(2, 11) = "AS"
StackedDeck(1, 12) = 12
StackedDeck(2, 12) = "6D"
StackedDeck(1, 13) = 13
StackedDeck(2, 13) = "JC"
StackedDeck(1, 14) = 14
StackedDeck(2, 14) = "8H"
StackedDeck(1, 15) = 15
StackedDeck(2, 15) = "KS"
StackedDeck(1, 16) = 16
StackedDeck(2, 16) = "3D"
StackedDeck(1, 17) = 17
StackedDeck(2, 17) = "10C"
StackedDeck(1, 18) = 18
StackedDeck(2, 18) = "2H"
StackedDeck(1, 19) = 19
StackedDeck(2, 19) = "7S"
StackedDeck(1, 20) = 20
StackedDeck(2, 20) = "9D"
StackedDeck(1, 21) = 21
StackedDeck(2, 21) = "5C"
StackedDeck(1, 22) = 22
StackedDeck(2, 22) = "QH"
StackedDeck(1, 23) = 23
StackedDeck(2, 23) = "4S"
StackedDeck(1, 24) = 24
StackedDeck(2, 24) = "AD"
StackedDeck(1, 25) = 25
StackedDeck(2, 25) = "6C"
StackedDeck(1, 26) = 26
StackedDeck(2, 26) = "JH"
StackedDeck(1, 27) = 27
StackedDeck(2, 27) = "8S"
StackedDeck(1, 28) = 28
StackedDeck(2, 28) = "KD"
StackedDeck(1, 29) = 29
StackedDeck(2, 29) = "3C"
StackedDeck(1, 30) = 30
StackedDeck(2, 30) = "10H"
StackedDeck(1, 31) = 31
StackedDeck(2, 31) = "2S"
StackedDeck(1, 32) = 32
StackedDeck(2, 32) = "7D"
StackedDeck(1, 33) = 33
StackedDeck(2, 33) = "9C"
StackedDeck(1, 34) = 34
StackedDeck(2, 34) = "5H"
StackedDeck(1, 35) = 35
StackedDeck(2, 35) = "QS"
StackedDeck(1, 36) = 36
StackedDeck(2, 36) = "4D"
StackedDeck(1, 37) = 37
StackedDeck(2, 37) = "AC"
StackedDeck(1, 38) = 38
StackedDeck(2, 38) = "6H"
StackedDeck(1, 39) = 39
StackedDeck(2, 39) = "JS"
StackedDeck(1, 40) = 40
StackedDeck(2, 40) = "8D"
StackedDeck(1, 41) = 41
StackedDeck(2, 41) = "KC"
StackedDeck(1, 42) = 42
StackedDeck(2, 42) = "3H"
StackedDeck(1, 43) = 43
StackedDeck(2, 43) = "10S"
StackedDeck(1, 44) = 44
StackedDeck(2, 44) = "2D"
StackedDeck(1, 45) = 45
StackedDeck(2, 45) = "7C"
StackedDeck(1, 46) = 46
StackedDeck(2, 46) = "9H"
StackedDeck(1, 47) = 47
StackedDeck(2, 47) = "5S"
StackedDeck(1, 48) = 48
StackedDeck(2, 48) = "QD"
StackedDeck(1, 49) = 49
StackedDeck(2, 49) = "4C"
StackedDeck(1, 50) = 50
StackedDeck(2, 50) = "AH"
StackedDeck(1, 51) = 51
StackedDeck(2, 51) = "6S"
StackedDeck(1, 52) = 52
StackedDeck(2, 52) = "JD"
CutDepth = 0
Cls
For i% = 1 To 52
    Deck(1, i%) = StackedDeck(1, i%)
    Deck(2, i%) = StackedDeck(2, i%)
    Deck(6, i%) = False
Next i%
ShowCards
End Sub
Public Sub SetDeckSiStebbins3()
DeckCount = 52
StackedDeck(1, 1) = 1
StackedDeck(2, 1) = "6H"
StackedDeck(1, 2) = 2
StackedDeck(2, 2) = "9S"
StackedDeck(1, 3) = 3
StackedDeck(2, 3) = "QD"
StackedDeck(1, 4) = 4
StackedDeck(2, 4) = "2C"
StackedDeck(1, 5) = 5
StackedDeck(2, 5) = "5H"
StackedDeck(1, 6) = 6
StackedDeck(2, 6) = "8S"
StackedDeck(1, 7) = 7
StackedDeck(2, 7) = "JD"
StackedDeck(1, 8) = 8
StackedDeck(2, 8) = "AC"
StackedDeck(1, 9) = 9
StackedDeck(2, 9) = "4H"
StackedDeck(1, 10) = 10
StackedDeck(2, 10) = "7S"
StackedDeck(1, 11) = 11
StackedDeck(2, 11) = "10D"
StackedDeck(1, 12) = 12
StackedDeck(2, 12) = "KC"
StackedDeck(1, 13) = 13
StackedDeck(2, 13) = "3H"
StackedDeck(1, 14) = 14
StackedDeck(2, 14) = "6S"
StackedDeck(1, 15) = 15
StackedDeck(2, 15) = "9D"
StackedDeck(1, 16) = 16
StackedDeck(2, 16) = "QC"
StackedDeck(1, 17) = 17
StackedDeck(2, 17) = "2H"
StackedDeck(1, 18) = 18
StackedDeck(2, 18) = "5S"
StackedDeck(1, 19) = 19
StackedDeck(2, 19) = "8D"
StackedDeck(1, 20) = 20
StackedDeck(2, 20) = "JC"
StackedDeck(1, 21) = 21
StackedDeck(2, 21) = "AH"
StackedDeck(1, 22) = 22
StackedDeck(2, 22) = "4S"
StackedDeck(1, 23) = 23
StackedDeck(2, 23) = "7D"
StackedDeck(1, 24) = 24
StackedDeck(2, 24) = "10C"
StackedDeck(1, 25) = 25
StackedDeck(2, 25) = "KH"
StackedDeck(1, 26) = 26
StackedDeck(2, 26) = "3S"
StackedDeck(1, 27) = 27
StackedDeck(2, 27) = "6D"
StackedDeck(1, 28) = 28
StackedDeck(2, 28) = "9C"
StackedDeck(1, 29) = 29
StackedDeck(2, 29) = "QH"
StackedDeck(1, 30) = 30
StackedDeck(2, 30) = "2S"
StackedDeck(1, 31) = 31
StackedDeck(2, 31) = "5D"
StackedDeck(1, 32) = 32
StackedDeck(2, 32) = "8C"
StackedDeck(1, 33) = 33
StackedDeck(2, 33) = "JH"
StackedDeck(1, 34) = 34
StackedDeck(2, 34) = "AS"
StackedDeck(1, 35) = 35
StackedDeck(2, 35) = "4D"
StackedDeck(1, 36) = 36
StackedDeck(2, 36) = "7C"
StackedDeck(1, 37) = 37
StackedDeck(2, 37) = "10H"
StackedDeck(1, 38) = 38
StackedDeck(2, 38) = "KS"
StackedDeck(1, 39) = 39
StackedDeck(2, 39) = "3D"
StackedDeck(1, 40) = 40
StackedDeck(2, 40) = "6C"
StackedDeck(1, 41) = 41
StackedDeck(2, 41) = "9H"
StackedDeck(1, 42) = 42
StackedDeck(2, 42) = "QS"
StackedDeck(1, 43) = 43
StackedDeck(2, 43) = "2D"
StackedDeck(1, 44) = 44
StackedDeck(2, 44) = "5C"
StackedDeck(1, 45) = 45
StackedDeck(2, 45) = "8H"
StackedDeck(1, 46) = 46
StackedDeck(2, 46) = "JS"
StackedDeck(1, 47) = 47
StackedDeck(2, 47) = "AD"
StackedDeck(1, 48) = 48
StackedDeck(2, 48) = "4C"
StackedDeck(1, 49) = 49
StackedDeck(2, 49) = "7H"
StackedDeck(1, 50) = 50
StackedDeck(2, 50) = "10S"
StackedDeck(1, 51) = 51
StackedDeck(2, 51) = "KD"
StackedDeck(1, 52) = 52
StackedDeck(2, 52) = "3C"
CutDepth = 0
Cls
For i% = 1 To 52
    Deck(1, i%) = StackedDeck(1, i%)
    Deck(2, i%) = StackedDeck(2, i%)
    Deck(6, i%) = False
Next i%
ShowCards
End Sub
Public Sub SetDeckSiStebbins4()
DeckCount = 52
StackedDeck(1, 1) = 1
StackedDeck(2, 1) = "6H"
StackedDeck(1, 2) = 2
StackedDeck(2, 2) = "10S"
StackedDeck(1, 3) = 3
StackedDeck(2, 3) = "AD"
StackedDeck(1, 4) = 4
StackedDeck(2, 4) = "5C"
StackedDeck(1, 5) = 5
StackedDeck(2, 5) = "9H"
StackedDeck(1, 6) = 6
StackedDeck(2, 6) = "KS"
StackedDeck(1, 7) = 7
StackedDeck(2, 7) = "4D"
StackedDeck(1, 8) = 8
StackedDeck(2, 8) = "8C"
StackedDeck(1, 9) = 9
StackedDeck(2, 9) = "QH"
StackedDeck(1, 10) = 10
StackedDeck(2, 10) = "3S"
StackedDeck(1, 11) = 11
StackedDeck(2, 11) = "7D"
StackedDeck(1, 12) = 12
StackedDeck(2, 12) = "JC"
StackedDeck(1, 13) = 13
StackedDeck(2, 13) = "2H"
StackedDeck(1, 14) = 14
StackedDeck(2, 14) = "6S"
StackedDeck(1, 15) = 15
StackedDeck(2, 15) = "10D"
StackedDeck(1, 16) = 16
StackedDeck(2, 16) = "AC"
StackedDeck(1, 17) = 17
StackedDeck(2, 17) = "5H"
StackedDeck(1, 18) = 18
StackedDeck(2, 18) = "9S"
StackedDeck(1, 19) = 19
StackedDeck(2, 19) = "KD"
StackedDeck(1, 20) = 20
StackedDeck(2, 20) = "4C"
StackedDeck(1, 21) = 21
StackedDeck(2, 21) = "8H"
StackedDeck(1, 22) = 22
StackedDeck(2, 22) = "QS"
StackedDeck(1, 23) = 23
StackedDeck(2, 23) = "3D"
StackedDeck(1, 24) = 24
StackedDeck(2, 24) = "7C"
StackedDeck(1, 25) = 25
StackedDeck(2, 25) = "JH"
StackedDeck(1, 26) = 26
StackedDeck(2, 26) = "2S"
StackedDeck(1, 27) = 27
StackedDeck(2, 27) = "6D"
StackedDeck(1, 28) = 28
StackedDeck(2, 28) = "10C"
StackedDeck(1, 29) = 29
StackedDeck(2, 29) = "AH"
StackedDeck(1, 30) = 30
StackedDeck(2, 30) = "5S"
StackedDeck(1, 31) = 31
StackedDeck(2, 31) = "9D"
StackedDeck(1, 32) = 32
StackedDeck(2, 32) = "KC"
StackedDeck(1, 33) = 33
StackedDeck(2, 33) = "4H"
StackedDeck(1, 34) = 34
StackedDeck(2, 34) = "8S"
StackedDeck(1, 35) = 35
StackedDeck(2, 35) = "QD"
StackedDeck(1, 36) = 36
StackedDeck(2, 36) = "3C"
StackedDeck(1, 37) = 37
StackedDeck(2, 37) = "7H"
StackedDeck(1, 38) = 38
StackedDeck(2, 38) = "JS"
StackedDeck(1, 39) = 39
StackedDeck(2, 39) = "2D"
StackedDeck(1, 40) = 40
StackedDeck(2, 40) = "6C"
StackedDeck(1, 41) = 41
StackedDeck(2, 41) = "10H"
StackedDeck(1, 42) = 42
StackedDeck(2, 42) = "AS"
StackedDeck(1, 43) = 43
StackedDeck(2, 43) = "5D"
StackedDeck(1, 44) = 44
StackedDeck(2, 44) = "9C"
StackedDeck(1, 45) = 45
StackedDeck(2, 45) = "KH"
StackedDeck(1, 46) = 46
StackedDeck(2, 46) = "4S"
StackedDeck(1, 47) = 47
StackedDeck(2, 47) = "8D"
StackedDeck(1, 48) = 48
StackedDeck(2, 48) = "QC"
StackedDeck(1, 49) = 49
StackedDeck(2, 49) = "3H"
StackedDeck(1, 50) = 50
StackedDeck(2, 50) = "7S"
StackedDeck(1, 51) = 51
StackedDeck(2, 51) = "JD"
StackedDeck(1, 52) = 52
StackedDeck(2, 52) = "2C"
CutDepth = 0
Cls
For i% = 1 To 52
    Deck(1, i%) = StackedDeck(1, i%)
    Deck(2, i%) = StackedDeck(2, i%)
    Deck(6, i%) = False
Next i%
ShowCards
End Sub
Public Sub SetDeckStanyon()
DeckCount = 52
StackedDeck(1, 1) = 1
StackedDeck(2, 1) = "AD"
StackedDeck(1, 2) = 2
StackedDeck(2, 2) = "3C"
StackedDeck(1, 3) = 3
StackedDeck(2, 3) = "6H"
StackedDeck(1, 4) = 4
StackedDeck(2, 4) = "10S"
StackedDeck(1, 5) = 5
StackedDeck(2, 5) = "2D"
StackedDeck(1, 6) = 6
StackedDeck(2, 6) = "4C"
StackedDeck(1, 7) = 7
StackedDeck(2, 7) = "7H"
StackedDeck(1, 8) = 8
StackedDeck(2, 8) = "JS"
StackedDeck(1, 9) = 9
StackedDeck(2, 9) = "3D"
StackedDeck(1, 10) = 10
StackedDeck(2, 10) = "5C"
StackedDeck(1, 11) = 11
StackedDeck(2, 11) = "8H"
StackedDeck(1, 12) = 12
StackedDeck(2, 12) = "QS"
StackedDeck(1, 13) = 13
StackedDeck(2, 13) = "4D"
StackedDeck(1, 14) = 14
StackedDeck(2, 14) = "6C"
StackedDeck(1, 15) = 15
StackedDeck(2, 15) = "9H"
StackedDeck(1, 16) = 16
StackedDeck(2, 16) = "KS"
StackedDeck(1, 17) = 17
StackedDeck(2, 17) = "5D"
StackedDeck(1, 18) = 18
StackedDeck(2, 18) = "7C"
StackedDeck(1, 19) = 19
StackedDeck(2, 19) = "10H"
StackedDeck(1, 20) = 20
StackedDeck(2, 20) = "AS"
StackedDeck(1, 21) = 21
StackedDeck(2, 21) = "6D"
StackedDeck(1, 22) = 22
StackedDeck(2, 22) = "8C"
StackedDeck(1, 23) = 23
StackedDeck(2, 23) = "JH"
StackedDeck(1, 24) = 24
StackedDeck(2, 24) = "2S"
StackedDeck(1, 25) = 25
StackedDeck(2, 25) = "7D"
StackedDeck(1, 26) = 26
StackedDeck(2, 26) = "9C"
StackedDeck(1, 27) = 27
StackedDeck(2, 27) = "QH"
StackedDeck(1, 28) = 28
StackedDeck(2, 28) = "3S"
StackedDeck(1, 29) = 29
StackedDeck(2, 29) = "8D"
StackedDeck(1, 30) = 30
StackedDeck(2, 30) = "10C"
StackedDeck(1, 31) = 31
StackedDeck(2, 31) = "KH"
StackedDeck(1, 32) = 32
StackedDeck(2, 32) = "4S"
StackedDeck(1, 33) = 33
StackedDeck(2, 33) = "9D"
StackedDeck(1, 34) = 34
StackedDeck(2, 34) = "JC"
StackedDeck(1, 35) = 35
StackedDeck(2, 35) = "AH"
StackedDeck(1, 36) = 36
StackedDeck(2, 36) = "5S"
StackedDeck(1, 37) = 37
StackedDeck(2, 37) = "10D"
StackedDeck(1, 38) = 38
StackedDeck(2, 38) = "QC"
StackedDeck(1, 39) = 39
StackedDeck(2, 39) = "2H"
StackedDeck(1, 40) = 40
StackedDeck(2, 40) = "6S"
StackedDeck(1, 41) = 41
StackedDeck(2, 41) = "JD"
StackedDeck(1, 42) = 42
StackedDeck(2, 42) = "KC"
StackedDeck(1, 43) = 43
StackedDeck(2, 43) = "3H"
StackedDeck(1, 44) = 44
StackedDeck(2, 44) = "7S"
StackedDeck(1, 45) = 45
StackedDeck(2, 45) = "QD"
StackedDeck(1, 46) = 46
StackedDeck(2, 46) = "AC"
StackedDeck(1, 47) = 47
StackedDeck(2, 47) = "4H"
StackedDeck(1, 48) = 48
StackedDeck(2, 48) = "8S"
StackedDeck(1, 49) = 49
StackedDeck(2, 49) = "KD"
StackedDeck(1, 50) = 50
StackedDeck(2, 50) = "2C"
StackedDeck(1, 51) = 51
StackedDeck(2, 51) = "5H"
StackedDeck(1, 52) = 52
StackedDeck(2, 52) = "9S"
CutDepth = 0
Cls
For i% = 1 To 52
    Deck(1, i%) = StackedDeck(1, i%)
    Deck(2, i%) = StackedDeck(2, i%)
    Deck(6, i%) = False
Next i%
ShowCards
End Sub
Public Sub SetDeckTamariz()
DeckCount = 52
StackedDeck(1, 1) = 1
StackedDeck(2, 1) = "4C"
StackedDeck(1, 2) = 2
StackedDeck(2, 2) = "2H"
StackedDeck(1, 3) = 3
StackedDeck(2, 3) = "7D"
StackedDeck(1, 4) = 4
StackedDeck(2, 4) = "3C"
StackedDeck(1, 5) = 5
StackedDeck(2, 5) = "4H"
StackedDeck(1, 6) = 6
StackedDeck(2, 6) = "6D"
StackedDeck(1, 7) = 7
StackedDeck(2, 7) = "AS"
StackedDeck(1, 8) = 8
StackedDeck(2, 8) = "5H"
StackedDeck(1, 9) = 9
StackedDeck(2, 9) = "9S"
StackedDeck(1, 10) = 10
StackedDeck(2, 10) = "2S"
StackedDeck(1, 11) = 11
StackedDeck(2, 11) = "QH"
StackedDeck(1, 12) = 12
StackedDeck(2, 12) = "3D"
StackedDeck(1, 13) = 13
StackedDeck(2, 13) = "QC"
StackedDeck(1, 14) = 14
StackedDeck(2, 14) = "8H"
StackedDeck(1, 15) = 15
StackedDeck(2, 15) = "6S"
StackedDeck(1, 16) = 16
StackedDeck(2, 16) = "5S"
StackedDeck(1, 17) = 17
StackedDeck(2, 17) = "9H"
StackedDeck(1, 18) = 18
StackedDeck(2, 18) = "KC"
StackedDeck(1, 19) = 19
StackedDeck(2, 19) = "2D"
StackedDeck(1, 20) = 20
StackedDeck(2, 20) = "JH"
StackedDeck(1, 21) = 21
StackedDeck(2, 21) = "3S"
StackedDeck(1, 22) = 22
StackedDeck(2, 22) = "8S"
StackedDeck(1, 23) = 23
StackedDeck(2, 23) = "6H"
StackedDeck(1, 24) = 24
StackedDeck(2, 24) = "10C"
StackedDeck(1, 25) = 25
StackedDeck(2, 25) = "5D"
StackedDeck(1, 26) = 26
StackedDeck(2, 26) = "KD"
StackedDeck(1, 27) = 27
StackedDeck(2, 27) = "2C"
StackedDeck(1, 28) = 28
StackedDeck(2, 28) = "3H"
StackedDeck(1, 29) = 29
StackedDeck(2, 29) = "8D"
StackedDeck(1, 30) = 30
StackedDeck(2, 30) = "5C"
StackedDeck(1, 31) = 31
StackedDeck(2, 31) = "KS"
StackedDeck(1, 32) = 32
StackedDeck(2, 32) = "JD"
StackedDeck(1, 33) = 33
StackedDeck(2, 33) = "8C"
StackedDeck(1, 34) = 34
StackedDeck(2, 34) = "10S"
StackedDeck(1, 35) = 35
StackedDeck(2, 35) = "KH"
StackedDeck(1, 36) = 36
StackedDeck(2, 36) = "JC"
StackedDeck(1, 37) = 37
StackedDeck(2, 37) = "7S"
StackedDeck(1, 38) = 38
StackedDeck(2, 38) = "10H"
StackedDeck(1, 39) = 39
StackedDeck(2, 39) = "AD"
StackedDeck(1, 40) = 40
StackedDeck(2, 40) = "4S"
StackedDeck(1, 41) = 41
StackedDeck(2, 41) = "7H"
StackedDeck(1, 42) = 42
StackedDeck(2, 42) = "4D"
StackedDeck(1, 43) = 43
StackedDeck(2, 43) = "AC"
StackedDeck(1, 44) = 44
StackedDeck(2, 44) = "9C"
StackedDeck(1, 45) = 45
StackedDeck(2, 45) = "JS"
StackedDeck(1, 46) = 46
StackedDeck(2, 46) = "QD"
StackedDeck(1, 47) = 47
StackedDeck(2, 47) = "7C"
StackedDeck(1, 48) = 48
StackedDeck(2, 48) = "QS"
StackedDeck(1, 49) = 49
StackedDeck(2, 49) = "10D"
StackedDeck(1, 50) = 50
StackedDeck(2, 50) = "6C"
StackedDeck(1, 51) = 51
StackedDeck(2, 51) = "AH"
StackedDeck(1, 52) = 52
StackedDeck(2, 52) = "9D"
CutDepth = 0
Cls
For i% = 1 To 52
    Deck(1, i%) = StackedDeck(1, i%)
    Deck(2, i%) = StackedDeck(2, i%)
    Deck(6, i%) = False
Next i%
ShowCards
End Sub

Public Sub SetDeckOsterlind()
DeckCount = 52
StackedDeck(1, 1) = 1
StackedDeck(2, 1) = "AS"
StackedDeck(1, 2) = 2
StackedDeck(2, 2) = "3S"
StackedDeck(1, 3) = 3
StackedDeck(2, 3) = "7D"
StackedDeck(1, 4) = 4
StackedDeck(2, 4) = "5H"
StackedDeck(1, 5) = 5
StackedDeck(2, 5) = "QC"
StackedDeck(1, 6) = 6
StackedDeck(2, 6) = "AC"
StackedDeck(1, 7) = 7
StackedDeck(2, 7) = "5S"
StackedDeck(1, 8) = 8
StackedDeck(2, 8) = "JH"
StackedDeck(1, 9) = 9
StackedDeck(2, 9) = "JC"
StackedDeck(1, 10) = 10
StackedDeck(2, 10) = "QD"
StackedDeck(1, 11) = 11
StackedDeck(2, 11) = "2D"
StackedDeck(1, 12) = 12
StackedDeck(2, 12) = "8C"
StackedDeck(1, 13) = 13
StackedDeck(2, 13) = "6S"
StackedDeck(1, 14) = 14
StackedDeck(2, 14) = "KH"
StackedDeck(1, 15) = 15
StackedDeck(2, 15) = "2H"
StackedDeck(1, 16) = 16
StackedDeck(2, 16) = "6D"
StackedDeck(1, 17) = 17
StackedDeck(2, 17) = "3D"
StackedDeck(1, 18) = 18
StackedDeck(2, 18) = "10S"
StackedDeck(1, 19) = 19
StackedDeck(2, 19) = "8D"
StackedDeck(1, 20) = 20
StackedDeck(2, 20) = "7C"
StackedDeck(1, 21) = 21
StackedDeck(2, 21) = "4S"
StackedDeck(1, 22) = 22
StackedDeck(2, 22) = "9D"
StackedDeck(1, 23) = 23
StackedDeck(2, 23) = "9C"
StackedDeck(1, 24) = 24
StackedDeck(2, 24) = "8H"
StackedDeck(1, 25) = 25
StackedDeck(2, 25) = "5D"
StackedDeck(1, 26) = 26
StackedDeck(2, 26) = "AD"
StackedDeck(1, 27) = 27
StackedDeck(2, 27) = "6H"
StackedDeck(1, 28) = 28
StackedDeck(2, 28) = "AH"
StackedDeck(1, 29) = 29
StackedDeck(2, 29) = "4D"
StackedDeck(1, 30) = 30
StackedDeck(2, 30) = "QS"
StackedDeck(1, 31) = 31
StackedDeck(2, 31) = "QH"
StackedDeck(1, 32) = 32
StackedDeck(2, 32) = "KC"
StackedDeck(1, 33) = 33
StackedDeck(2, 33) = "3C"
StackedDeck(1, 34) = 34
StackedDeck(2, 34) = "9H"
StackedDeck(1, 35) = 35
StackedDeck(2, 35) = "7S"
StackedDeck(1, 36) = 36
StackedDeck(2, 36) = "2S"
StackedDeck(1, 37) = 37
StackedDeck(2, 37) = "5C"
StackedDeck(1, 38) = 38
StackedDeck(2, 38) = "KD"
StackedDeck(1, 39) = 39
StackedDeck(2, 39) = "4H"
StackedDeck(1, 40) = 40
StackedDeck(2, 40) = "10C"
StackedDeck(1, 41) = 41
StackedDeck(2, 41) = "10D"
StackedDeck(1, 42) = 42
StackedDeck(2, 42) = "JS"
StackedDeck(1, 43) = 43
StackedDeck(2, 43) = "10H"
StackedDeck(1, 44) = 44
StackedDeck(2, 44) = "9S"
StackedDeck(1, 45) = 45
StackedDeck(2, 45) = "6C"
StackedDeck(1, 46) = 46
StackedDeck(2, 46) = "2C"
StackedDeck(1, 47) = 47
StackedDeck(2, 47) = "7H"
StackedDeck(1, 48) = 48
StackedDeck(2, 48) = "3H"
StackedDeck(1, 49) = 49
StackedDeck(2, 49) = "8S"
StackedDeck(1, 50) = 50
StackedDeck(2, 50) = "4C"
StackedDeck(1, 51) = 51
StackedDeck(2, 51) = "JD"
StackedDeck(1, 52) = 52
StackedDeck(2, 52) = "KS"
CutDepth = 0
Cls
For i% = 1 To 52
    Deck(1, i%) = StackedDeck(1, i%)
    Deck(2, i%) = StackedDeck(2, i%)
    Deck(6, i%) = False
Next i%
ShowCards
End Sub

Public Sub CutDeckRandom(paramreverse)
    Dim pReverse As String
    Dim pCutStartCard As Integer
    Dim pCutEndCard As Integer
    pReverse = paramreverse
    CutDepth = Int(Rnd * DeckCount) + 1
    'reverse a block if required
    If pReverse = "T" Then
        pCutStartCard = 1
        pCutEndCard = CutDepth
    ElseIf pReverse = "B" Then
        pCutStartCard = CutDepth + 1
        pCutEndCard = 52
    End If
    If Not pReverse = "X" Then
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
    'resume cut procedure
    i = 1
    For j% = CutDepth + 1 To DeckCount
        For z% = 1 To DeckProperties
            ChangedDeck(z%, i) = Deck(z%, j%)
        Next z%
        i = i + 1
    Next j%
    For k% = 1 To CutDepth
        For z% = 1 To DeckProperties
            ChangedDeck(z%, i) = Deck(z%, k%)
        Next z%
        i = i + 1
    Next k%
    For m% = 1 To DeckCount
        For z% = 1 To DeckProperties
            Deck(z%, m%) = ChangedDeck(z%, m%)
        Next z%
    Next m%
ShowCards
End Sub

Public Sub RiffleShuffle(paramreverse)
    Dim pReverse As String
    Dim pCutStartCard As Integer
    Dim pCutEndCard As Integer
    pReverse = paramreverse
    Cls
    CutDepth = Int(Rnd * 7) + 20
      'approximately half the deck for a riffle shuffle
    RemainingCut = DeckCount - CutDepth
    'reverse a block if required
    If pReverse = "T" Then
        pCutStartCard = 1
        pCutEndCard = CutDepth
    ElseIf pReverse = "B" Then
        pCutStartCard = CutDepth + 1
        pCutEndCard = 52
    End If
    If Not pReverse = "X" Then
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
    'resume riffle shuffle procedure
    ReDim TopCut(DeckProperties, CutDepth)
    ReDim BottomCut(DeckProperties, RemainingCut)
    For i% = 1 To CutDepth
        For z% = 1 To DeckProperties
            TopCut(z%, i%) = Deck(z%, i%)
        Next z%
    Next i%
    For j% = 1 To RemainingCut
        For z% = 1 To DeckProperties
            BottomCut(z%, j%) = Deck(z%, CutDepth + j%)
        Next z%
    Next j%
    TopIndex = 1
    BottomIndex = 1
    For k% = 1 To DeckCount
        side = Rnd
        'when low, shuffle from TopCut
        'when high, shuffle from BottomCut
        If side < CutDepth / DeckCount Then
          'compare Rnd with ratio of cuts to change odds
          'to a more even mixing of the talons
            If TopIndex <= CutDepth Then
                For z% = 1 To DeckProperties
                    Deck(z%, k%) = TopCut(z%, TopIndex)
                Next z%
                TopIndex = TopIndex + 1
            Else
                For z% = 1 To DeckProperties
                    Deck(z%, k%) = BottomCut(z%, BottomIndex)
                Next z%
                BottomIndex = BottomIndex + 1
            End If
        Else
            If BottomIndex <= RemainingCut Then
                For z% = 1 To DeckProperties
                    Deck(z%, k%) = BottomCut(z%, BottomIndex)
                Next z%
                BottomIndex = BottomIndex + 1
            Else
                For z% = 1 To DeckProperties
                    Deck(z%, k%) = TopCut(z%, TopIndex)
                Next z%
                TopIndex = TopIndex + 1
            End If
        End If
    Next k%
ShowCards
End Sub
Public Sub SelectCardsCutSelectNext1(param1)
    SelectedCutDepth = Int(Rnd * 10) + 21
    i = 1
    For j% = SelectedCutDepth + 2 To DeckCount
        For z% = 1 To DeckProperties
            ChangedDeck(z%, i) = Deck(z%, j%)
        Next z%
        i = i + 1
    Next j%
    For z% = 1 To DeckProperties
        ChangedDeck(z%, i) = Deck(z%, SelectedCutDepth + 1)
    Next z%
    ChangedDeck(4, i) = "Selected"
    If param1 = "R" Then
        ChangedDeck(6, i) = Not ChangedDeck(6, i)
    End If
    'SelectedCards(1, NumberOfSelectedCards + 1) _
        = Deck(2, SelectedCutDepth + 1)
    'SelectedCards(2, NumberOfSelectedCards + 1) _
        = "Card" & Deck(2, SelectedCutDepth + 1)
    If SelectionsTextBox.Text = Empty Then
        SelectionsTextBox.Text = Deck(2, SelectedCutDepth + 1)
    Else
        SelectionsTextBox.Text = SelectionsTextBox.Text & " " & _
            Deck(2, SelectedCutDepth + 1)
    End If
    'NumberOfSelectedCards = NumberOfSelectedCards + 1
    i = i + 1
    For k% = 1 To SelectedCutDepth
        For z% = 1 To DeckProperties
            ChangedDeck(z%, i) = Deck(z%, k%)
        Next z%
        i = i + 1
    Next k%
    For m% = 1 To DeckCount
        For z% = 1 To DeckProperties
            Deck(z%, m%) = ChangedDeck(z%, m%)
        Next z%
    Next m%
    ShowCards
End Sub
Public Sub SelectCardsCutSelectNext2(param1, param2)
    SelectedCutDepth = Int(Rnd * 8) + 13
    For j% = 1 To SelectedCutDepth
        For z% = 1 To DeckProperties
            ChangedDeck(z%, 52 - SelectedCutDepth + j%) = Deck(z%, j%)
        Next z%
    Next j%
    For z% = 1 To DeckProperties
        ChangedDeck(z%, 52 - SelectedCutDepth) = Deck(z%, SelectedCutDepth + 1)
    Next z%
    ChangedDeck(4, 52 - SelectedCutDepth) = "Selected"
    If param1 = "R" Then
        ChangedDeck(6, 52 - SelectedCutDepth) = Not ChangedDeck(6, 52 - SelectedCutDepth)
    End If
    If SelectionsTextBox.Text = Empty Then
        SelectionsTextBox.Text = Deck(2, SelectedCutDepth + 1)
    Else
        SelectionsTextBox.Text = SelectionsTextBox.Text & " " & _
            Deck(2, SelectedCutDepth + 1)
    End If
    i = SelectedCutDepth
    SelectedCutDepth = Int(Rnd * (52 - i) / 4 + 3 * (52 - i) / 8)
    For k% = 1 To SelectedCutDepth
        For z% = 1 To DeckProperties
            ChangedDeck(z%, 52 - i - SelectedCutDepth - 1 + k%) = Deck(z%, i + k% + 1)
        Next z%
    Next k%
    For z% = 1 To DeckProperties
        ChangedDeck(z%, 52 - i - SelectedCutDepth - 1) = Deck(z%, i + SelectedCutDepth + 2)
    Next z%
    ChangedDeck(4, 52 - i - SelectedCutDepth - 1) = "Selected"
    If param2 = "R" Then
        ChangedDeck(6, 52 - i - SelectedCutDepth - 1) = Not ChangedDeck(6, 52 - i - SelectedCutDepth - 1)
    End If
    If SelectionsTextBox.Text = Empty Then
        SelectionsTextBox.Text = Deck(2, i + SelectedCutDepth + 2)
    Else
        SelectionsTextBox.Text = SelectionsTextBox.Text & " " & _
            Deck(2, i + SelectedCutDepth + 2)
    End If
    For p% = 1 To 52 - i - SelectedCutDepth - 2
        For z% = 1 To DeckProperties
            ChangedDeck(z%, p%) = Deck(z%, i + SelectedCutDepth + 2 + p%)
        Next z%
    Next p%
    For m% = 1 To DeckCount
        For z% = 1 To DeckProperties
            Deck(z%, m%) = ChangedDeck(z%, m%)
        Next z%
    Next m%
    ShowCards
End Sub
Public Sub SelectCardsCutSelectNext3(param1, param2, param3)
    FirstCutStart = DeckCount
    FirstCutDepth = Int(Rnd * FirstCutStart / 8 + 3 * FirstCutStart / 16)
    SecondCutStart = DeckCount - FirstCutDepth - 1
    'the extra "-1" is due to the selected card being missing
    SecondCutDepth = Int(Rnd * SecondCutStart / 6 + 3 * SecondCutStart / 12)
    ThirdCutStart = SecondCutStart - SecondCutDepth - 1
    'the extra "-1" is due to the selected card being missing
    ThirdCutDepth = Int(Rnd * ThirdCutStart / 4 + 3 * ThirdCutStart / 8)
    'First Cut
    For j% = 1 To FirstCutDepth
        For z% = 1 To DeckProperties
            ChangedDeck(z%, 52 - FirstCutDepth + j%) = Deck(z%, j%)
        Next z%
    Next j%
    'First selection
    For z% = 1 To DeckProperties
        ChangedDeck(z%, 52 - FirstCutDepth) = Deck(z%, FirstCutDepth + 1)
    Next z%
    ChangedDeck(4, 52 - FirstCutDepth) = "Selected"
    If param1 = "R" Then
        ChangedDeck(6, 52 - FirstCutDepth) = Not ChangedDeck(6, 52 - FirstCutDepth)
    End If
    If SelectionsTextBox.Text = Empty Then
        SelectionsTextBox.Text = Deck(2, FirstCutDepth + 1)
    Else
        SelectionsTextBox.Text = SelectionsTextBox.Text & " " & _
            Deck(2, FirstCutDepth + 1)
    End If
    'Second cut
    For k% = 1 To SecondCutDepth
        For z% = 1 To DeckProperties
            ChangedDeck(z%, 52 - FirstCutDepth - SecondCutDepth - 1 + k%) _
            = Deck(z%, FirstCutDepth + k% + 1)
        Next z%
    Next k%
    'Second selection
    For z% = 1 To DeckProperties
        ChangedDeck(z%, 52 - FirstCutDepth - SecondCutDepth - 1) _
        = Deck(z%, FirstCutDepth + SecondCutDepth + 2)
    Next z%
    ChangedDeck(4, 52 - FirstCutDepth - SecondCutDepth - 1) = "Selected"
    If param2 = "R" Then
        ChangedDeck(6, 52 - FirstCutDepth - SecondCutDepth - 1) = _
            Not ChangedDeck(6, 52 - FirstCutDepth - SecondCutDepth - 1)
    End If
    If SelectionsTextBox.Text = Empty Then
        SelectionsTextBox.Text = Deck(2, FirstCutDepth + SecondCutDepth + 2)
    Else
        SelectionsTextBox.Text = SelectionsTextBox.Text & " " & _
            Deck(2, FirstCutDepth + SecondCutDepth + 2)
    End If
    'Third cut
    For k% = 1 To ThirdCutDepth
        For z% = 1 To DeckProperties
            ChangedDeck(z%, 52 - FirstCutDepth - SecondCutDepth _
            - ThirdCutDepth - 2 + k%) _
            = Deck(z%, FirstCutDepth + SecondCutDepth + k% + 2)
        Next z%
    Next k%
    'Third selection
    For z% = 1 To DeckProperties
        ChangedDeck(z%, 52 - FirstCutDepth - SecondCutDepth - ThirdCutDepth - 2) _
        = Deck(z%, FirstCutDepth + SecondCutDepth + ThirdCutDepth + 3)
    Next z%
    ChangedDeck(4, 52 - FirstCutDepth - SecondCutDepth - ThirdCutDepth - 2) _
    = "Selected"
    If param3 = "R" Then
        ChangedDeck(6, 52 - FirstCutDepth - SecondCutDepth - ThirdCutDepth - 2) = _
            Not ChangedDeck(6, 52 - FirstCutDepth - SecondCutDepth - ThirdCutDepth - 2)
    End If
    If SelectionsTextBox.Text = Empty Then
        SelectionsTextBox.Text = Deck(2, FirstCutDepth + SecondCutDepth + ThirdCutDepth + 3)
    Else
        SelectionsTextBox.Text = SelectionsTextBox.Text & " " & _
            Deck(2, FirstCutDepth + SecondCutDepth + ThirdCutDepth + 3)
    End If
    'Last block replacement
    For p% = 1 To DeckCount - FirstCutDepth - SecondCutDepth - ThirdCutDepth - 3
        For z% = 1 To DeckProperties
            ChangedDeck(z%, p%) = Deck(z%, FirstCutDepth + SecondCutDepth _
            + ThirdCutDepth + 3 + p%)
        Next z%
    Next p%
    For m% = 1 To DeckCount
        For z% = 1 To DeckProperties
            Deck(z%, m%) = ChangedDeck(z%, m%)
        Next z%
    Next m%
    ShowCards
End Sub
Public Sub SelectCardsCutSelectNextRepeat2(param1, param2, param3)
    FirstCutStart = DeckCount
    FirstCutDepth = Int(Rnd * FirstCutStart / 8 + 3 * FirstCutStart / 16)
    SecondCutDepth = Int(Rnd * FirstCutStart / 8 + 3 * FirstCutStart / 16)
    ThirdCutDepth = Int(Rnd * FirstCutStart / 8 + 3 * FirstCutStart / 16)
    FourthCutDepth = DeckCount - FirstCutDepth - SecondCutDepth - ThirdCutDepth - 3
    'First Cut
    For j% = 1 To FirstCutDepth
        For z% = 1 To DeckProperties
            ChangedDeck(z%, FourthCutDepth + 1 + j%) = Deck(z%, j%)
        Next z%
    Next j%
    'First selection
    For z% = 1 To DeckProperties
        ChangedDeck(z%, FourthCutDepth + 1) = Deck(z%, FirstCutDepth + 1)
    Next z%
    ChangedDeck(4, FourthCutDepth + 1) = "Selected"
    If param1 = "R" Then
        ChangedDeck(6, FourthCutDepth + 1) = Not ChangedDeck(6, FourthCutDepth + 1)
    End If
    If SelectionsTextBox.Text = Empty Then
        SelectionsTextBox.Text = Deck(2, FirstCutDepth + 1)
    Else
        SelectionsTextBox.Text = SelectionsTextBox.Text & " " & _
            Deck(2, FirstCutDepth + 1)
    End If
    'Second cut
    For k% = 1 To SecondCutDepth
        For z% = 1 To DeckProperties
            ChangedDeck(z%, 52 - ThirdCutDepth - SecondCutDepth - 1 + k%) _
            = Deck(z%, FirstCutDepth + k% + 1)
        Next z%
    Next k%
    'Second selection
    For z% = 1 To DeckProperties
        ChangedDeck(z%, 52 - ThirdCutDepth - SecondCutDepth - 1) _
        = Deck(z%, FirstCutDepth + SecondCutDepth + 2)
    Next z%
    ChangedDeck(4, 52 - ThirdCutDepth - SecondCutDepth - 1) = "Selected"
    If param2 = "R" Then
        ChangedDeck(6, 52 - ThirdCutDepth - SecondCutDepth - 1) = _
            Not ChangedDeck(6, 52 - ThirdCutDepth - SecondCutDepth - 1)
    End If
    If SelectionsTextBox.Text = Empty Then
        SelectionsTextBox.Text = Deck(2, FirstCutDepth + SecondCutDepth + 2)
    Else
        SelectionsTextBox.Text = SelectionsTextBox.Text & " " & _
            Deck(2, FirstCutDepth + SecondCutDepth + 2)
    End If
    'Third cut
    For k% = 1 To ThirdCutDepth
        For z% = 1 To DeckProperties
            ChangedDeck(z%, 52 - ThirdCutDepth + k%) _
            = Deck(z%, FirstCutDepth + SecondCutDepth + k% + 2)
        Next z%
    Next k%
    'Third selection
    For z% = 1 To DeckProperties
        ChangedDeck(z%, 52 - ThirdCutDepth) _
        = Deck(z%, FirstCutDepth + SecondCutDepth + ThirdCutDepth + 3)
    Next z%
    ChangedDeck(4, 52 - ThirdCutDepth) = "Selected"
    If param3 = "R" Then
        ChangedDeck(6, 52 - ThirdCutDepth) = Not ChangedDeck(6, 52 - ThirdCutDepth)
    End If
    If SelectionsTextBox.Text = Empty Then
        SelectionsTextBox.Text = Deck(2, FirstCutDepth + SecondCutDepth + ThirdCutDepth + 3)
    Else
        SelectionsTextBox.Text = SelectionsTextBox.Text & " " & _
            Deck(2, FirstCutDepth + SecondCutDepth + ThirdCutDepth + 3)
    End If
    'Last block replacement
    For p% = 1 To FourthCutDepth
        For z% = 1 To DeckProperties
            ChangedDeck(z%, p%) = Deck(z%, FirstCutDepth + SecondCutDepth _
            + ThirdCutDepth + 3 + p%)
        Next z%
    Next p%
    For m% = 1 To DeckCount
        For z% = 1 To DeckProperties
            Deck(z%, m%) = ChangedDeck(z%, m%)
        Next z%
    Next m%
    ShowCards
End Sub
Public Sub SelectCardsCutSelectFace1(param1)
    FirstCutStart = DeckCount
    FirstCutDepth = Int(Rnd * FirstCutStart / 6 + 3 * FirstCutStart / 12)
    SecondCutStart = DeckCount - FirstCutDepth
    SecondCutDepth = Int(Rnd * SecondCutStart / 4 + 3 * SecondCutStart / 8)
    ThirdCutDepth = SecondCutStart - SecondCutDepth
    'Preliminary Cut
    For j% = 1 To FirstCutDepth
        For z% = 1 To DeckProperties
            ChangedDeck(z%, 52 - FirstCutDepth + j%) = Deck(z%, j%)
        Next z%
    Next j%
    'First cut
    For j% = 1 To SecondCutDepth
        For z% = 1 To DeckProperties
            ChangedDeck(z%, 52 - FirstCutDepth - SecondCutDepth + j%) = _
            Deck(z%, FirstCutDepth + j%)
        Next z%
    Next j%
    'First selection
    For z% = 1 To DeckProperties
        ChangedDeck(z%, 52 - FirstCutDepth) = Deck(z%, FirstCutDepth + SecondCutDepth)
    Next z%
    ChangedDeck(4, 52 - FirstCutDepth) = "Selected"
    If param1 = "R" Then
        ChangedDeck(6, 52 - FirstCutDepth) = Not ChangedDeck(6, 52 - FirstCutDepth)
    End If
    If SelectionsTextBox.Text = Empty Then
        SelectionsTextBox.Text = Deck(2, FirstCutDepth + SecondCutDepth)
    Else
        SelectionsTextBox.Text = SelectionsTextBox.Text & " " & _
            Deck(2, FirstCutDepth + SecondCutDepth)
    End If
    'Last block replacement
    For p% = 1 To ThirdCutDepth
        For z% = 1 To DeckProperties
            ChangedDeck(z%, p%) = Deck(z%, FirstCutDepth + SecondCutDepth + p%)
        Next z%
    Next p%
    For m% = 1 To DeckCount
        For z% = 1 To DeckProperties
            Deck(z%, m%) = ChangedDeck(z%, m%)
        Next z%
    Next m%
    ShowCards
End Sub
Public Sub SelectCardsCutSelectFace2(param1, param2)
    FirstCutStart = DeckCount
    FirstCutDepth = Int(Rnd * FirstCutStart / 8 + 3 * FirstCutStart / 16)
    SecondCutStart = DeckCount - FirstCutDepth
    SecondCutDepth = Int(Rnd * SecondCutStart / 6 + 3 * SecondCutStart / 12)
    ThirdCutStart = SecondCutStart - SecondCutDepth
    ThirdCutDepth = Int(Rnd * ThirdCutStart / 4 + 3 * ThirdCutStart / 8)
    FourthCutDepth = ThirdCutStart - ThirdCutDepth
    'Preliminary Cut
    For j% = 1 To FirstCutDepth
        For z% = 1 To DeckProperties
            ChangedDeck(z%, 52 - FirstCutDepth + j%) = Deck(z%, j%)
        Next z%
    Next j%
    'First cut
    For j% = 1 To SecondCutDepth
        For z% = 1 To DeckProperties
            ChangedDeck(z%, 52 - FirstCutDepth - SecondCutDepth + j%) = _
            Deck(z%, FirstCutDepth + j%)
        Next z%
    Next j%
    'First selection
    For z% = 1 To DeckProperties
        ChangedDeck(z%, 52 - FirstCutDepth) = Deck(z%, FirstCutDepth + SecondCutDepth)
    Next z%
    ChangedDeck(4, 52 - FirstCutDepth) = "Selected"
    If param1 = "R" Then
        ChangedDeck(6, 52 - FirstCutDepth) = Not ChangedDeck(6, 52 - FirstCutDepth)
    End If
    If SelectionsTextBox.Text = Empty Then
        SelectionsTextBox.Text = Deck(2, FirstCutDepth + SecondCutDepth)
    Else
        SelectionsTextBox.Text = SelectionsTextBox.Text & " " & _
            Deck(2, FirstCutDepth + SecondCutDepth)
    End If
    'Second cut
    For k% = 1 To ThirdCutDepth
        For z% = 1 To DeckProperties
            ChangedDeck(z%, 52 - FirstCutDepth - SecondCutDepth - ThirdCutDepth + k%) _
            = Deck(z%, FirstCutDepth + SecondCutDepth + k%)
        Next z%
    Next k%
    'Second selection
    For z% = 1 To DeckProperties
        ChangedDeck(z%, 52 - FirstCutDepth - SecondCutDepth) _
        = Deck(z%, FirstCutDepth + SecondCutDepth + ThirdCutDepth)
    Next z%
    ChangedDeck(4, 52 - FirstCutDepth - SecondCutDepth) = "Selected"
    If param2 = "R" Then
        ChangedDeck(6, 52 - FirstCutDepth - SecondCutDepth) = _
            Not ChangedDeck(6, 52 - FirstCutDepth - SecondCutDepth)
    End If
    If SelectionsTextBox.Text = Empty Then
        SelectionsTextBox.Text = Deck(2, FirstCutDepth + SecondCutDepth + ThirdCutDepth)
    Else
        SelectionsTextBox.Text = SelectionsTextBox.Text & " " & _
            Deck(2, FirstCutDepth + SecondCutDepth + ThirdCutDepth)
    End If
    'Last block replacement
    For p% = 1 To FourthCutDepth
        For z% = 1 To DeckProperties
            ChangedDeck(z%, p%) = Deck(z%, FirstCutDepth + SecondCutDepth _
            + ThirdCutDepth + p%)
        Next z%
    Next p%
    For m% = 1 To DeckCount
        For z% = 1 To DeckProperties
            Deck(z%, m%) = ChangedDeck(z%, m%)
        Next z%
    Next m%
    ShowCards
End Sub
Public Sub SelectCardsCutSelectFace3(param1, param2, param3)
    FirstCutStart = DeckCount
    FirstCutDepth = Int(Rnd * FirstCutStart / 10 + 3 * FirstCutStart / 20)
    SecondCutStart = DeckCount - FirstCutDepth
    SecondCutDepth = Int(Rnd * SecondCutStart / 8 + 3 * SecondCutStart / 16)
    ThirdCutStart = SecondCutStart - SecondCutDepth
    ThirdCutDepth = Int(Rnd * ThirdCutStart / 6 + 3 * ThirdCutStart / 12)
    FourthCutStart = ThirdCutStart - ThirdCutDepth
    FourthCutDepth = Int(Rnd * FourthCutStart / 4 + 3 * FourthCutStart / 8)
    FifthCutDepth = FourthCutStart - FourthCutDepth
    'Preliminary Cut
    For j% = 1 To FirstCutDepth
        For z% = 1 To DeckProperties
            ChangedDeck(z%, 52 - FirstCutDepth + j%) = Deck(z%, j%)
        Next z%
    Next j%
    'First cut
    For j% = 1 To SecondCutDepth
        For z% = 1 To DeckProperties
            ChangedDeck(z%, 52 - FirstCutDepth - SecondCutDepth + j%) = _
            Deck(z%, FirstCutDepth + j%)
        Next z%
    Next j%
    'First selection
    For z% = 1 To DeckProperties
        ChangedDeck(z%, 52 - FirstCutDepth) = Deck(z%, FirstCutDepth + SecondCutDepth)
    Next z%
    ChangedDeck(4, 52 - FirstCutDepth) = "Selected"
    If param1 = "R" Then
        ChangedDeck(6, 52 - FirstCutDepth) = Not ChangedDeck(6, 52 - FirstCutDepth)
    End If
    If SelectionsTextBox.Text = Empty Then
        SelectionsTextBox.Text = Deck(2, FirstCutDepth + SecondCutDepth)
    Else
        SelectionsTextBox.Text = SelectionsTextBox.Text & " " & _
            Deck(2, FirstCutDepth + SecondCutDepth)
    End If
    'Second cut
    For k% = 1 To ThirdCutDepth
        For z% = 1 To DeckProperties
            ChangedDeck(z%, 52 - FirstCutDepth - SecondCutDepth - ThirdCutDepth + k%) _
            = Deck(z%, FirstCutDepth + SecondCutDepth + k%)
        Next z%
    Next k%
    'Second selection
    For z% = 1 To DeckProperties
        ChangedDeck(z%, 52 - FirstCutDepth - SecondCutDepth) _
        = Deck(z%, FirstCutDepth + SecondCutDepth + ThirdCutDepth)
    Next z%
    ChangedDeck(4, 52 - FirstCutDepth - SecondCutDepth) = "Selected"
    If param2 = "R" Then
        ChangedDeck(6, 52 - FirstCutDepth - SecondCutDepth) = _
            Not ChangedDeck(6, 52 - FirstCutDepth - SecondCutDepth)
    End If
    If SelectionsTextBox.Text = Empty Then
        SelectionsTextBox.Text = Deck(2, FirstCutDepth + SecondCutDepth + ThirdCutDepth)
    Else
        SelectionsTextBox.Text = SelectionsTextBox.Text & " " & _
            Deck(2, FirstCutDepth + SecondCutDepth + ThirdCutDepth)
    End If
    'Third cut
    For k% = 1 To FourthCutDepth
        For z% = 1 To DeckProperties
            ChangedDeck(z%, 52 - FirstCutDepth - SecondCutDepth _
            - ThirdCutDepth - FourthCutDepth + k%) _
            = Deck(z%, FirstCutDepth + SecondCutDepth + ThirdCutDepth + k%)
        Next z%
    Next k%
    'Third selection
    For z% = 1 To DeckProperties
        ChangedDeck(z%, 52 - FirstCutDepth - SecondCutDepth - ThirdCutDepth) _
        = Deck(z%, FirstCutDepth + SecondCutDepth + ThirdCutDepth + FourthCutDepth)
    Next z%
    ChangedDeck(4, 52 - FirstCutDepth - SecondCutDepth - ThirdCutDepth) _
        = "Selected"
    If param3 = "R" Then
        ChangedDeck(6, 52 - FirstCutDepth - SecondCutDepth - ThirdCutDepth) = _
            Not ChangedDeck(6, 52 - FirstCutDepth - SecondCutDepth - ThirdCutDepth)
    End If
    If SelectionsTextBox.Text = Empty Then
        SelectionsTextBox.Text = Deck(2, FirstCutDepth + SecondCutDepth + ThirdCutDepth + FourthCutDepth)
    Else
        SelectionsTextBox.Text = SelectionsTextBox.Text & " " & _
            Deck(2, FirstCutDepth + SecondCutDepth + ThirdCutDepth + FourthCutDepth)
    End If
    'Last block replacement
    For p% = 1 To FifthCutDepth
        For z% = 1 To DeckProperties
            ChangedDeck(z%, p%) = Deck(z%, FirstCutDepth + SecondCutDepth _
            + ThirdCutDepth + FourthCutDepth + p%)
        Next z%
    Next p%
    For m% = 1 To DeckCount
        For z% = 1 To DeckProperties
            Deck(z%, m%) = ChangedDeck(z%, m%)
        Next z%
    Next m%
    ShowCards
End Sub
Public Sub SelectCardsCutSelectNextRepeat(param1, param2)
    SelectedCutDepth = Int(Rnd * 8) + 13
    i = SelectedCutDepth
    SelectedCutDepth = Int(Rnd * (52 - i) / 4 + 3 * (52 - i) / 8)
    For j% = 1 To i
        For z% = 1 To DeckProperties
            ChangedDeck(z%, 52 - i - SelectedCutDepth - 1 + j%) = Deck(z%, j%)
        Next z%
    Next j%
    For z% = 1 To DeckProperties
        ChangedDeck(z%, 52 - i - SelectedCutDepth - 1) = Deck(z%, i + 1)
    Next z%
    ChangedDeck(4, 52 - i - SelectedCutDepth - 1) = "Selected"
    If param1 = "R" Then
        ChangedDeck(6, 52 - i - SelectedCutDepth - 1) = _
            Not ChangedDeck(6, 52 - i - SelectedCutDepth - 1)
    End If
    If SelectionsTextBox.Text = Empty Then
        SelectionsTextBox.Text = Deck(2, i + 1)
    Else
        SelectionsTextBox.Text = SelectionsTextBox.Text & " " & _
            Deck(2, i + 1)
    End If
    For k% = 1 To SelectedCutDepth
        For z% = 1 To DeckProperties
            ChangedDeck(z%, 52 - SelectedCutDepth + k%) = Deck(z%, i + k% + 1)
        Next z%
    Next k%
    For z% = 1 To DeckProperties
        ChangedDeck(z%, 52 - SelectedCutDepth) = Deck(z%, i + SelectedCutDepth + 2)
    Next z%
    ChangedDeck(4, 52 - SelectedCutDepth) = "Selected"
    If param2 = "R" Then
        ChangedDeck(6, 52 - SelectedCutDepth) = Not ChangedDeck(6, 52 - SelectedCutDepth)
    End If
    If SelectionsTextBox.Text = Empty Then
        SelectionsTextBox.Text = Deck(2, i + SelectedCutDepth + 2)
    Else
        SelectionsTextBox.Text = SelectionsTextBox.Text & " " & _
            Deck(2, i + SelectedCutDepth + 2)
    End If
    For p% = 1 To 52 - i - SelectedCutDepth - 2
        For z% = 1 To DeckProperties
            ChangedDeck(z%, p%) = Deck(z%, i + SelectedCutDepth + 2 + p%)
        Next z%
    Next p%
    For m% = 1 To DeckCount
        For z% = 1 To DeckProperties
            Deck(z%, m%) = ChangedDeck(z%, m%)
        Next z%
    Next m%
    ShowCards
End Sub

Public Sub Form_Load()
If SessionAlreadyOpen = 1 Then
    Exit Sub
    'if the session was already open, do not load in the default stack.
    'this is important when reopening the Control form with the View menu.
End If
Randomize
LoadDefaultMnemonic
SessionRecordButtons
On Error GoTo DefaultLoad
Dim fso As New FileSystemObject, deckfile As File, ts As TextStream
Set deckfile = fso.GetFile(App.Path & "\stackview.svf")
Set ts = deckfile.OpenAsTextStream(ForReading)
For i% = 1 To 52
    Deck(1, i%) = Val(ts.ReadLine)
    Deck(2, i%) = ts.ReadLine
Next i%
'XXX
Dim sProperty As String
If Not ts.AtEndOfStream Then
    sProperty = ts.ReadLine
    If Left(sProperty, 10) = "BackDesign" Then
        BackDesignCurrent = Right(sProperty, Len(sProperty) - 11)
    End If
End If
Call frmBackDesignDialog.LoadBackDesign(BackDesignCurrent)
'XXX
ts.Close
For k% = 1 To 52
    For m% = 1 To 52
        If Deck(1, m%) = k% Then
            TestOriginalDeck(1, k%) = Deck(1, m%)
            TestOriginalDeck(2, k%) = Deck(2, m%)
        End If
    Next m%
Next k%
'the above For/Next sets the current deck to its
'original position for testing
DeckProperties = 6
DeckCount = 52
ClearSelections_Click
ClearReversedCards_Click
NumberOfSelectedCards = 0
SelectionsTextBox.Text = Empty
ShowCards
SessionAlreadyOpen = 1
SessionRecordMode = False
SessionRecursionLimit = 10
SessionRecursionLevel = 0
SessionRecursing = False
ImportedCustomDeck = 0
'the next reset deck call eliminates the specified order when the stackview.svf
'file was originally saved.  I don't think it need to be here any more
'Call ResetCurrentDeck
Exit Sub
DefaultLoad:
SetDeckNewBicycle
For k% = 1 To 52
    For m% = 1 To 52
        If Deck(1, m%) = k% Then
            TestOriginalDeck(1, k%) = Deck(1, m%)
            TestOriginalDeck(2, k%) = Deck(2, m%)
        End If
    Next m%
Next k%
DeckProperties = 6
DeckCount = 52
ClearSelections_Click
ClearReversedCards_Click
NumberOfSelectedCards = 0
SelectionsTextBox.Text = Empty
ShowCards
SessionAlreadyOpen = 1
SessionRecordMode = False
SessionRecursionLimit = 10
SessionRecursionLevel = 0
SessionRecursing = False
ImportedCustomDeck = 0
Call ResetCurrentDeck
End Sub

Public Sub LoadDefaultMnemonic()
On Error GoTo DefaultMnemonicLoad
Dim fso As New FileSystemObject, mnemonicfile As File, ts As TextStream
Set mnemonicfile = fso.GetFile(App.Path & "\stackview.svm")
Set ts = mnemonicfile.OpenAsTextStream(ForReading)
For i% = 1 To 52
    MnemonicCards(i%) = ts.ReadLine
Next i%
For i% = 1 To 52
    MnemonicPositions(i%) = ts.ReadLine
Next i%
ts.Close
Call frmMnemonic.LoadMnemonicTable
MnemonicSaved = 1
Exit Sub
DefaultMnemonicLoad:
Call frmMnemonic.RestoreAronsonDefault_Click
MnemonicSaved = 1
End Sub



Public Sub ShowCards()
PokerCardsDealt = 0
PilesShown = 0
Call frmDeck.DisplayCards
End Sub

Public Sub ShowPiles()
PokerCardsDealt = 0
PilesShown = 1
Call frmDeck.DisplayPiles
End Sub


Public Sub ReturnCard(rp)
If rp = Chr(34) & "Anywhere" & Chr(34) Or _
    rp = "Anywhere" Then
    ReturnPosition = Int(Rnd * DeckCount) + 1
ElseIf rp = "Top Third" Then
    ReturnPosition = Int(Rnd * 17) + 1
ElseIf rp = "Middle Third" Then
    ReturnPosition = Int(Rnd * 18) + 18
ElseIf rp = "Bottom Third" Then
    ReturnPosition = Int(Rnd * 17) + 36
ElseIf rp = Chr(34) & "Original Position" & Chr(34) Or _
    rp = "Original Position" Then
    ReturnPosition = SelectedCard
ElseIf IsNumeric(rp) Then
    ReturnPosition = Val(rp)
End If
If ReturnPosition = SelectedCard Then
    ShowCards
    Exit Sub
End If
'move selected card to the return position
For z% = 1 To DeckProperties
    ChangedDeck(z%, ReturnPosition) = Deck(z%, SelectedCard)
Next z%
If ReturnPosition < SelectedCard Then
    For j% = 1 To ReturnPosition - 1
        For z% = 1 To DeckProperties
            ChangedDeck(z%, j%) = Deck(z%, j%)
        Next z%
    Next j%
    For k% = 1 To SelectedCard - ReturnPosition
        For z% = 1 To DeckProperties
            ChangedDeck(z%, ReturnPosition + k%) = _
            Deck(z%, ReturnPosition - 1 + k%)
        Next z%
    Next k%
    For n% = 1 To DeckCount - SelectedCard
        For z% = 1 To DeckProperties
            ChangedDeck(z%, SelectedCard + n%) = _
            Deck(z%, SelectedCard + n%)
        Next z%
    Next n%
    For m% = 1 To DeckCount
        For z% = 1 To DeckProperties
            Deck(z%, m%) = ChangedDeck(z%, m%)
        Next z%
    Next m%
ElseIf ReturnPosition > SelectedCard Then
    For k% = 1 To SelectedCard - 1
        For z% = 1 To DeckProperties
            ChangedDeck(z%, k%) = _
            Deck(z%, k%)
        Next z%
    Next k%
    For j% = 1 To ReturnPosition - SelectedCard
        For z% = 1 To DeckProperties
            ChangedDeck(z%, SelectedCard - 1 + j%) = _
            Deck(z%, SelectedCard + j%)
        Next z%
    Next j%
    For n% = 1 To DeckCount - ReturnPosition
        For z% = 1 To DeckProperties
            ChangedDeck(z%, ReturnPosition + n%) = _
            Deck(z%, ReturnPosition + n%)
        Next z%
    Next n%
    For m% = 1 To DeckCount
        For z% = 1 To DeckProperties
            Deck(z%, m%) = ChangedDeck(z%, m%)
        Next z%
    Next m%
End If
ShowCards
End Sub



Public Sub RunSingleButton_Click()
If (Not IsNumeric(RunSingleTextBox.Text) Or _
    Val(RunSingleTextBox.Text) < 1 Or _
    Val(RunSingleTextBox.Text) > 52) Then
    RunSingleTextBox.Text = Empty
    MsgBox "Please enter a valid card number (1 to 52)" & Chr(13) _
        & "in the 'Run Single Cards' Input Box"
    Exit Sub
End If
If RunSingleInverseCheck Then
    If RunSingleCardsReverseCheck Then
        InverseRunSingleCardsReverse (Val(RunSingleTextBox.Text))
        'SessionRecord
        If SessionRecordMode Then
            SessionCommand = "InverseRunSingleCardsReverse(" & RunSingleTextBox.Text & ")"
            SessionListBox.AddItem SessionCommand
            SessionStatusUpdate (0)
        End If
    Else
        InverseRunSingleCards (Val(RunSingleTextBox.Text))
        'SessionRecord
        If SessionRecordMode Then
            SessionCommand = "InverseRunSingleCards(" & RunSingleTextBox.Text & ")"
            SessionListBox.AddItem SessionCommand
            SessionStatusUpdate (0)
        End If
    End If
Else
    If RunSingleCardsReverseCheck Then
        RunSingleCardsReverse (Val(RunSingleTextBox.Text))
        'SessionRecord
        If SessionRecordMode Then
            SessionCommand = "RunSingleCardsReverse(" & RunSingleTextBox.Text & ")"
            SessionListBox.AddItem SessionCommand
            SessionStatusUpdate (0)
        End If
    Else
        RunSingleCards (Val(RunSingleTextBox.Text))
        'SessionRecord
        If SessionRecordMode Then
            SessionCommand = "RunSingleCards(" & RunSingleTextBox.Text & ")"
            SessionListBox.AddItem SessionCommand
            SessionStatusUpdate (0)
        End If
    End If
End If
End Sub

Private Sub InverseRunSingleCards(runcards)
i = 1
tmp = Val(runcards)
For j% = DeckCount - tmp + 1 To DeckCount
    For z% = 1 To DeckProperties
        ChangedDeck(z%, tmp - i + 1) = Deck(z%, j%)
    Next z%
    i = i + 1
Next j%
For k% = 1 To DeckCount - tmp
    For z% = 1 To DeckProperties
        ChangedDeck(z%, i) = Deck(z%, k%)
    Next z%
    i = i + 1
Next k%
For m% = 1 To DeckCount
    For z% = 1 To DeckProperties
        Deck(z%, m%) = ChangedDeck(z%, m%)
    Next z%
Next m%
ShowCards
End Sub

Private Sub InverseRunSingleCardsReverse(runcards)
i = 1
tmp = Val(runcards)
For j% = DeckCount - tmp + 1 To DeckCount
    For z% = 1 To DeckProperties
        ChangedDeck(z%, tmp - i + 1) = Deck(z%, j%)
    Next z%
    i = i + 1
Next j%
For k% = 1 To DeckCount - tmp
    For z% = 1 To DeckProperties
        ChangedDeck(z%, i) = Deck(z%, k%)
    Next z%
    i = i + 1
Next k%
For m% = 1 To DeckCount
    For z% = 1 To DeckProperties
        Deck(z%, m%) = ChangedDeck(z%, m%)
    Next z%
Next m%
ReverseBlock (tmp)
ShowCards
End Sub


Private Sub RunSingleCards(runcards)
i = 1
tmp = Val(runcards)
For j% = DeckCount - tmp + 1 To DeckCount
    For z% = 1 To DeckProperties
        ChangedDeck(z%, j%) = Deck(z%, tmp - i + 1)
    Next z%
    i = i + 1
Next j%
For k% = 1 To DeckCount - tmp
    For z% = 1 To DeckProperties
        ChangedDeck(z%, k%) = Deck(z%, i)
    Next z%
    i = i + 1
Next k%
For m% = 1 To DeckCount
    For z% = 1 To DeckProperties
        Deck(z%, m%) = ChangedDeck(z%, m%)
    Next z%
Next m%
ShowCards
End Sub

Private Sub RunSingleCardsReverse(runcards)
i = 1
tmp = Val(runcards)
For j% = DeckCount - tmp + 1 To DeckCount
    For z% = 1 To DeckProperties
        ChangedDeck(z%, j%) = Deck(z%, tmp - i + 1)
    Next z%
    i = i + 1
Next j%
For k% = 1 To DeckCount - tmp
    For z% = 1 To DeckProperties
        ChangedDeck(z%, k%) = Deck(z%, i)
    Next z%
    i = i + 1
Next k%
For m% = 1 To DeckCount
    For z% = 1 To DeckProperties
        Deck(z%, m%) = ChangedDeck(z%, m%)
    Next z%
Next m%
Call CutDeckPrecise((DeckCount - tmp), "X")
ReverseBlock (tmp)
Call CutDeckPrecise(tmp, "X")
ShowCards
End Sub


Public Sub SessionClearALL_Click()
If SessionListBox.ListCount = 0 Then
    MsgBox "There are no entries in this session to clear."
    Exit Sub
End If
Dim Msg, Style, Title, Help, Ctxt, Response, MyString
Msg = "This will CLEAR ALL Session Events." & Chr(13) & _
    "Do you want to continue?"   ' Define message.
Style = vbYesNo + vbQuestion + vbDefaultButton2   ' Define buttons.
Title = "Clear ALL Session Events"   ' Define title.
Response = MsgBox(Msg, Style, Title)
If Response = vbYes Then   ' User chose Yes.
    SessionListBox.Clear
    SessionStatusUpdate (1)
End If
End Sub

Public Sub SessionEventDelete_Click()
sessionindex = SessionListBox.ListIndex
If sessionindex = -1 Then
    MsgBox "Please select an Event first"
    Exit Sub
End If
If Left(SessionListBox.List(sessionindex), 4) = "Free" Or _
    Left(SessionListBox.List(sessionindex), 5) = "Force" Then
    SessionListBox.RemoveItem sessionindex
    SessionListBox.RemoveItem sessionindex
    SessionStatusUpdate (0)
ElseIf Left(SessionListBox.List(sessionindex), 6) = "Return" Then
    sessionindex = sessionindex - 1
    SessionListBox.ListIndex = sessionindex
    SessionListBox.RemoveItem sessionindex
    SessionListBox.RemoveItem sessionindex
    SessionStatusUpdate (0)
Else
    SessionListBox.RemoveItem sessionindex
    SessionStatusUpdate (0)
End If
If SessionListBox.ListCount = 0 Then
    Exit Sub
ElseIf sessionindex > SessionListBox.ListCount - 1 Then
    SessionListBox.ListIndex = SessionListBox.ListCount - 1
Else
    SessionListBox.ListIndex = sessionindex
End If
End Sub

Public Sub SessionEventMoveDown_Click()
'Events with "Free", "Force", and "Return" as starting text
'need to stay together as pairs
'there is also a check to make sure that there is not another pair
'next to it that need to be kept together
sessionindex = SessionListBox.ListIndex
sessioncount = SessionListBox.ListCount
If sessionindex = -1 Then
    MsgBox "Please select an Event first"
    Exit Sub
End If
If sessionindex = SessionListBox.ListCount - 1 Then
    Exit Sub
End If
If (Left(SessionListBox.List(sessionindex), 4) = "Free" Or _
    Left(SessionListBox.List(sessionindex), 5) = "Force") And _
    sessionindex = sessioncount - 2 Then
    Exit Sub
End If
If Left(SessionListBox.List(sessionindex), 4) = "Free" Or _
    Left(SessionListBox.List(sessionindex), 5) = "Force" Then
    If Left(SessionListBox.List(sessionindex + 2), 4) = "Free" Or _
        Left(SessionListBox.List(sessionindex + 2), 5) = "Force" Then
        movetext = SessionListBox.List(sessionindex)
        SessionListBox.AddItem movetext, sessionindex + 4
        SessionListBox.RemoveItem sessionindex
        movetext = SessionListBox.List(sessionindex)
        SessionListBox.AddItem movetext, sessionindex + 4
        SessionListBox.RemoveItem sessionindex
        SessionListBox.ListIndex = sessionindex + 2
        SessionStatusUpdate (0)
    Else
        movetext = SessionListBox.List(sessionindex)
        SessionListBox.AddItem movetext, sessionindex + 3
        SessionListBox.RemoveItem sessionindex
        movetext = SessionListBox.List(sessionindex)
        SessionListBox.AddItem movetext, sessionindex + 3
        SessionListBox.RemoveItem sessionindex
        SessionListBox.ListIndex = sessionindex + 1
        SessionStatusUpdate (0)
    End If
ElseIf Left(SessionListBox.List(sessionindex), 6) = "Return" Then
    If Left(SessionListBox.List(sessionindex + 1), 4) = "Free" Or _
        Left(SessionListBox.List(sessionindex + 1), 5) = "Force" Then
        sessionindex = SessionListBox.ListIndex - 1
        movetext = SessionListBox.List(sessionindex)
        SessionListBox.AddItem movetext, sessionindex + 4
        SessionListBox.RemoveItem sessionindex
        movetext = SessionListBox.List(sessionindex)
        SessionListBox.AddItem movetext, sessionindex + 4
        SessionListBox.RemoveItem sessionindex
        SessionListBox.ListIndex = sessionindex + 3
        SessionStatusUpdate (0)
    Else
        sessionindex = SessionListBox.ListIndex - 1
        movetext = SessionListBox.List(sessionindex)
        SessionListBox.AddItem movetext, sessionindex + 3
        SessionListBox.RemoveItem sessionindex
        movetext = SessionListBox.List(sessionindex)
        SessionListBox.AddItem movetext, sessionindex + 3
        SessionListBox.RemoveItem sessionindex
        SessionListBox.ListIndex = sessionindex + 2
        SessionStatusUpdate (0)
    End If
ElseIf Left(SessionListBox.List(sessionindex + 1), 4) = "Free" Or _
    Left(SessionListBox.List(sessionindex + 1), 5) = "Force" Then
    movetext = SessionListBox.List(sessionindex)
    SessionListBox.AddItem movetext, sessionindex + 3
    SessionListBox.RemoveItem sessionindex
    SessionListBox.ListIndex = sessionindex + 2
    SessionStatusUpdate (0)
Else
    movetext = SessionListBox.List(sessionindex)
    SessionListBox.AddItem movetext, sessionindex + 2
    SessionListBox.RemoveItem sessionindex
    SessionListBox.ListIndex = sessionindex + 1
    SessionStatusUpdate (0)
End If
End Sub

Public Sub SessionEventMoveUp_Click()
'Events with "Free", "Force", and "Return" as starting text
'need to stay together as pairs
'there is also a check to make sure that there is not another pair
'next to it that need to be kept together
sessionindex = SessionListBox.ListIndex
sessioncount = SessionListBox.ListCount
If sessionindex = -1 Then
    MsgBox "Please select an Event first"
    Exit Sub
End If
If sessionindex = 0 Then
    Exit Sub
End If
If Left(SessionListBox.List(sessionindex), 6) = "Return" And _
    sessionindex = 1 Then
    Exit Sub
End If
If Left(SessionListBox.List(sessionindex), 4) = "Free" Or _
    Left(SessionListBox.List(sessionindex), 5) = "Force" Then
    If Left(SessionListBox.List(sessionindex - 1), 6) = "Return" Then
        movetext = SessionListBox.List(sessionindex)
        SessionListBox.AddItem movetext, sessionindex - 2
        SessionListBox.RemoveItem sessionindex + 1
        sessionindex = sessionindex + 1
        movetext = SessionListBox.List(sessionindex)
        SessionListBox.AddItem movetext, sessionindex - 2
        SessionListBox.RemoveItem sessionindex + 1
        SessionListBox.ListIndex = sessionindex - 3
        SessionStatusUpdate (0)
    Else
        movetext = SessionListBox.List(sessionindex)
        SessionListBox.AddItem movetext, sessionindex - 1
        SessionListBox.RemoveItem sessionindex + 1
        sessionindex = sessionindex + 1
        movetext = SessionListBox.List(sessionindex)
        SessionListBox.AddItem movetext, sessionindex - 1
        SessionListBox.RemoveItem sessionindex + 1
        SessionListBox.ListIndex = sessionindex - 2
        SessionStatusUpdate (0)
    End If
ElseIf Left(SessionListBox.List(sessionindex), 6) = "Return" Then
    If Left(SessionListBox.List(sessionindex - 2), 6) = "Return" Then
        sessionindex = SessionListBox.ListIndex - 1
        movetext = SessionListBox.List(sessionindex)
        SessionListBox.AddItem movetext, sessionindex - 2
        SessionListBox.RemoveItem sessionindex + 1
        sessionindex = sessionindex + 1
        movetext = SessionListBox.List(sessionindex)
        SessionListBox.AddItem movetext, sessionindex - 2
        SessionListBox.RemoveItem sessionindex + 1
        SessionListBox.ListIndex = sessionindex - 2
        SessionStatusUpdate (0)
    Else
        sessionindex = SessionListBox.ListIndex - 1
        movetext = SessionListBox.List(sessionindex)
        SessionListBox.AddItem movetext, sessionindex - 1
        SessionListBox.RemoveItem sessionindex + 1
        sessionindex = sessionindex + 1
        movetext = SessionListBox.List(sessionindex)
        SessionListBox.AddItem movetext, sessionindex - 1
        SessionListBox.RemoveItem sessionindex + 1
        SessionListBox.ListIndex = sessionindex - 1
        SessionStatusUpdate (0)
    End If
ElseIf Left(SessionListBox.List(sessionindex - 1), 6) = "Return" Then
    movetext = SessionListBox.List(sessionindex)
    SessionListBox.AddItem movetext, sessionindex - 2
    SessionListBox.RemoveItem sessionindex + 1
    SessionListBox.ListIndex = sessionindex - 2
    SessionStatusUpdate (0)
Else
    movetext = SessionListBox.List(sessionindex)
    SessionListBox.AddItem movetext, sessionindex - 1
    SessionListBox.RemoveItem sessionindex + 1
    SessionListBox.ListIndex = sessionindex - 1
    SessionStatusUpdate (0)
End If
End Sub

Public Sub SessionPlayALL_Click()
If SessionListBox.ListCount = 0 Then
    MsgBox "There are no entries in this session to play."
    Exit Sub
End If
SessionRecursionLevel = 0
For i% = 0 To SessionListBox.ListCount - 1
    SessionListBox.ListIndex = i%
    Call SessionParse(i%)
    If SessionParseError Then
        SessionParseError = False
        Exit Sub
    End If
Next i%
End Sub


Public Sub SessionParse(sessionparameter)
Dim pCommaCount As Integer
Dim pParameterCounter As Integer
Dim pCommaPosition As Integer
Dim pLengthOfEvent As Integer
Dim pCurrentSearchPosition As Integer
Dim sessionevent As String
Dim stringpointer As Integer
Dim commapointer As Integer
Dim pTextPlug As String
'pSEP() is the new arrayed sessioneventparameter
'set pSep() to Empty
For i% = 1 To 11
    pSEP(i%) = Empty
Next i%
If SessionRecursing Then
    sessionevent = SessionRecursionList(SessionRecursionLevel).List(sessionparameter)
Else
    sessionevent = SessionListBox.List(sessionparameter)
End If
stringpointer = InStr(sessionevent, "(")
If stringpointer = 0 Then
    myCall = sessionevent
    'this for when there are no parameters passed - no "(" in sessionevent
Else
    myCall = Left(sessionevent, stringpointer - 1)
    pCurrentSearchPosition = stringpointer + 1
    pCommaPosition = 99
    'needs to be set to an arbitrary non-0 value
    pCommaCount = 0
    pParameterCounter = 1
    While pCommaPosition <> 0
        pCommaPosition = InStr(pCurrentSearchPosition, sessionevent, ",")
        If pCommaPosition > 0 Then
            pCommaCount = pCommaCount + 1
            pSEP(pCommaCount) = Mid(sessionevent, pCurrentSearchPosition, _
                pCommaPosition - pCurrentSearchPosition)
            pCurrentSearchPosition = pCommaPosition + 2
        End If
    Wend
    pLengthOfEvent = Len(sessionevent)
    pSEP(pCommaCount + 1) = Mid(sessionevent, pCurrentSearchPosition, _
        pLengthOfEvent - pCurrentSearchPosition)
End If
'now process the event and parameters
Select Case myCall
    Case "CutDeckPrecise"
        If Not (pSEP(2) = "X" Or pSEP(2) = "T" Or pSEP(2) = "B") Then
            MsgBox ("CutDeckPrecise Event second parameter is not valid." & Chr(13) & _
                "It must be a " & Chr(34) & "T" & Chr(34) & " or a " & _
                Chr(34) & "B" & Chr(34) & " or an " & Chr(34) & "X" & Chr(34) & ".")
            SessionParseError = True
            Exit Sub
        End If
        If IsNumeric(pSEP(1)) And _
            (Val(pSEP(1)) < 1 Or _
            Val(pSEP(1)) > 52) Then
                MsgBox ("CutDeckPrecise Event parameter is out of range." & Chr(13) & _
                "It must be between 1 and 52.")
                SessionParseError = True
                Exit Sub
        ElseIf Not IsNumeric(pSEP(1)) Then
            MsgBox ("CutDeckPrecise Event parameter is not valid." & Chr(13) & _
            "It must be a number between 1 and 52.")
            SessionParseError = True
            Exit Sub
        End If
        Call CutDeckPrecise(pSEP(1), pSEP(2))
    Case "ShiftTopBlock"
        If IsNumeric(pSEP(1)) And _
            (Val(pSEP(1)) < 1 Or _
            Val(pSEP(1)) > 52) Then
                MsgBox ("First event parameter is out of range." & Chr(13) & _
                "It must be between 1 and 52.")
                SessionParseError = True
                Exit Sub
        ElseIf Not IsNumeric(pSEP(1)) Then
            MsgBox ("First event parameter is not valid." & Chr(13) & _
            "It must be a number between 1 and 52.")
            SessionParseError = True
            Exit Sub
        End If
        If IsNumeric(pSEP(2)) And _
            (Val(pSEP(2)) < 1 Or _
            Val(pSEP(2)) > 52) Then
                MsgBox ("Second event parameter is out of range." & Chr(13) & _
                "It must be between 1 and 52.")
                SessionParseError = True
                Exit Sub
        ElseIf Not IsNumeric(pSEP(2)) Then
            MsgBox ("Second event parameter is not valid." & Chr(13) & _
            "It must be a number between 1 and 52.")
            SessionParseError = True
            Exit Sub
        End If
        Call ShiftTopBlock(pSEP(1), pSEP(2))
    Case "InverseShiftTopBlock"
        If IsNumeric(pSEP(1)) And _
            (Val(pSEP(1)) < 1 Or _
            Val(pSEP(1)) > 52) Then
                MsgBox ("First event parameter is out of range." & Chr(13) & _
                "It must be between 1 and 52.")
                SessionParseError = True
                Exit Sub
        ElseIf Not IsNumeric(pSEP(1)) Then
            MsgBox ("First event parameter is not valid." & Chr(13) & _
            "It must be a number between 1 and 52.")
            SessionParseError = True
            Exit Sub
        End If
        If IsNumeric(pSEP(2)) And _
            (Val(pSEP(2)) < 1 Or _
            Val(pSEP(2)) > 52) Then
                MsgBox ("Second event parameter is out of range." & Chr(13) & _
                "It must be between 1 and 52.")
                SessionParseError = True
                Exit Sub
        ElseIf Not IsNumeric(pSEP(2)) Then
            MsgBox ("Second event parameter is not valid." & Chr(13) & _
            "It must be a number between 1 and 52.")
            SessionParseError = True
            Exit Sub
        End If
        Call InverseShiftTopBlock(pSEP(1), pSEP(2))
    Case "ShiftTopBlockReverse"
        If IsNumeric(pSEP(1)) And _
            (Val(pSEP(1)) < 1 Or _
            Val(pSEP(1)) > 52) Then
                MsgBox ("First event parameter is out of range." & Chr(13) & _
                "It must be between 1 and 52.")
                SessionParseError = True
                Exit Sub
        ElseIf Not IsNumeric(pSEP(1)) Then
            MsgBox ("First event parameter is not valid." & Chr(13) & _
            "It must be a number between 1 and 52.")
            SessionParseError = True
            Exit Sub
        End If
        If IsNumeric(pSEP(2)) And _
            (Val(pSEP(2)) < 1 Or _
            Val(pSEP(2)) > 52) Then
                MsgBox ("Second event parameter is out of range." & Chr(13) & _
                "It must be between 1 and 52.")
                SessionParseError = True
                Exit Sub
        ElseIf Not IsNumeric(pSEP(2)) Then
            MsgBox ("Second event parameter is not valid." & Chr(13) & _
            "It must be a number between 1 and 52.")
            SessionParseError = True
            Exit Sub
        End If
        Call ShiftTopBlockReverse(pSEP(1), pSEP(2))
    Case "InverseShiftTopBlockReverse"
        If IsNumeric(pSEP(1)) And _
            (Val(pSEP(1)) < 1 Or _
            Val(pSEP(1)) > 52) Then
                MsgBox ("First event parameter is out of range." & Chr(13) & _
                "It must be between 1 and 52.")
                SessionParseError = True
                Exit Sub
        ElseIf Not IsNumeric(pSEP(1)) Then
            MsgBox ("First event parameter is not valid." & Chr(13) & _
            "It must be a number between 1 and 52.")
            SessionParseError = True
            Exit Sub
        End If
        If IsNumeric(pSEP(2)) And _
            (Val(pSEP(2)) < 1 Or _
            Val(pSEP(2)) > 52) Then
                MsgBox ("Second event parameter is out of range." & Chr(13) & _
                "It must be between 1 and 52.")
                SessionParseError = True
                Exit Sub
        ElseIf Not IsNumeric(pSEP(2)) Then
            MsgBox ("Second event parameter is not valid." & Chr(13) & _
            "It must be a number between 1 and 52.")
            SessionParseError = True
            Exit Sub
        End If
        Call InverseShiftTopBlockReverse(pSEP(1), pSEP(2))
    Case "ReverseCard"
        Dim pSafeMatchFound As Boolean
        pSafeMatchFound = False
        If IsNumeric(pSEP(1)) Then
            For i% = 1 To 52
                If i% = Val(pSEP(1)) Then
                    pSafeMatchFound = True
                End If
            Next i%
            If Not pSafeMatchFound Then
                MsgBox ("ReverseCard Event Error:" & Chr(13) & _
                "The single parameter can only be an interger from 1 to 52," & Chr(13) & _
                "or a card value such as AC, 2C, 3C, etc.")
                Exit Sub
            End If
        Else
            For i% = 1 To 52
                If Deck(2, i%) = pSEP(1) Then
                    pSafeMatchFound = True
                End If
            Next i%
            If Not pSafeMatchFound Then
                MsgBox ("ReverseCard Event Error:" & Chr(13) & _
                "The single parameter can only be an interger from 1 to 52," & Chr(13) & _
                "or a card value such as AC, 2C, 3C, etc.")
                Exit Sub
            End If
        End If
        Call frmDeck.ReverseCard(pSEP(1))
    Case "MoveCard"
        If IsNumeric(pSEP(1)) And _
            (Val(pSEP(1)) < 1 Or _
            Val(pSEP(1)) > 52) Then
                MsgBox ("First event parameter is out of range." & Chr(13) & _
                "It must be between 1 and 52.")
                SessionParseError = True
                Exit Sub
        ElseIf Not IsNumeric(pSEP(1)) Then
            MsgBox ("First event parameter is not valid." & Chr(13) & _
            "It must be a number between 1 and 52.")
            SessionParseError = True
            Exit Sub
        End If
        If IsNumeric(pSEP(2)) And _
            (Val(pSEP(2)) < 1 Or _
            Val(pSEP(2)) > 52) Then
                MsgBox ("Second event parameter is out of range." & Chr(13) & _
                "It must be between 1 and 52.")
                SessionParseError = True
                Exit Sub
        ElseIf Not IsNumeric(pSEP(2)) Then
            MsgBox ("Second event parameter is not valid." & Chr(13) & _
            "It must be a number between 1 and 52.")
            SessionParseError = True
            Exit Sub
        End If
        Call MoveCard(pSEP(1), pSEP(2))
    Case "MoveCardReverse"
        If IsNumeric(pSEP(1)) And _
            (Val(pSEP(1)) < 1 Or _
            Val(pSEP(1)) > 52) Then
                MsgBox ("First event parameter is out of range." & Chr(13) & _
                "It must be between 1 and 52.")
                SessionParseError = True
                Exit Sub
        ElseIf Not IsNumeric(pSEP(1)) Then
            MsgBox ("First event parameter is not valid." & Chr(13) & _
            "It must be a number between 1 and 52.")
            SessionParseError = True
            Exit Sub
        End If
        If IsNumeric(pSEP(2)) And _
            (Val(pSEP(2)) < 1 Or _
            Val(pSEP(2)) > 52) Then
                MsgBox ("Second event parameter is out of range." & Chr(13) & _
                "It must be between 1 and 52.")
                SessionParseError = True
                Exit Sub
        ElseIf Not IsNumeric(pSEP(2)) Then
            MsgBox ("Second event parameter is not valid." & Chr(13) & _
            "It must be a number between 1 and 52.")
            SessionParseError = True
            Exit Sub
        End If
        Call MoveCardReverse(pSEP(1), pSEP(2))
    Case "InverseMoveCard"
        If IsNumeric(pSEP(1)) And _
            (Val(pSEP(1)) < 1 Or _
            Val(pSEP(1)) > 52) Then
                MsgBox ("First event parameter is out of range." & Chr(13) & _
                "It must be between 1 and 52.")
                SessionParseError = True
                Exit Sub
        ElseIf Not IsNumeric(pSEP(1)) Then
            MsgBox ("First event parameter is not valid." & Chr(13) & _
            "It must be a number between 1 and 52.")
            SessionParseError = True
            Exit Sub
        End If
        If IsNumeric(pSEP(2)) And _
            (Val(pSEP(2)) < 1 Or _
            Val(pSEP(2)) > 52) Then
                MsgBox ("Second event parameter is out of range." & Chr(13) & _
                "It must be between 1 and 52.")
                SessionParseError = True
                Exit Sub
        ElseIf Not IsNumeric(pSEP(2)) Then
            MsgBox ("Second event parameter is not valid." & Chr(13) & _
            "It must be a number between 1 and 52.")
            SessionParseError = True
            Exit Sub
        End If
        Call InverseMoveCard(pSEP(1), pSEP(2))
    Case "InverseMoveCardReverse"
        If IsNumeric(pSEP(1)) And _
            (Val(pSEP(1)) < 1 Or _
            Val(pSEP(1)) > 52) Then
                MsgBox ("First event parameter is out of range." & Chr(13) & _
                "It must be between 1 and 52.")
                SessionParseError = True
                Exit Sub
        ElseIf Not IsNumeric(pSEP(1)) Then
            MsgBox ("First event parameter is not valid." & Chr(13) & _
            "It must be a number between 1 and 52.")
            SessionParseError = True
            Exit Sub
        End If
        If IsNumeric(pSEP(2)) And _
            (Val(pSEP(2)) < 1 Or _
            Val(pSEP(2)) > 52) Then
                MsgBox ("Second event parameter is out of range." & Chr(13) & _
                "It must be between 1 and 52.")
                SessionParseError = True
                Exit Sub
        ElseIf Not IsNumeric(pSEP(2)) Then
            MsgBox ("Second event parameter is not valid." & Chr(13) & _
            "It must be a number between 1 and 52.")
            SessionParseError = True
            Exit Sub
        End If
        Call InverseMoveCardReverse(pSEP(1), pSEP(2))
    Case "CutSpecialRandom"
        If Not (pSEP(2) = "X" Or pSEP(2) = "T" Or pSEP(2) = "B") Then
            MsgBox ("CutSpecialRandom Event second parameter is not valid." & Chr(13) & _
                "It must be a " & Chr(34) & "T" & Chr(34) & " or a " & _
                Chr(34) & "B" & Chr(34) & " or an " & Chr(34) & "X" & Chr(34) & ".")
            SessionParseError = True
            Exit Sub
        End If
        SessionParseError = True
        'set the error condition first,
        'then remove it if there is a matched parameter
        If _
            pSEP(1) = "Quarter" Or _
            pSEP(1) = "Third" Or _
            pSEP(1) = "Half" Or _
            pSEP(1) = "Two Thirds" Or _
            pSEP(1) = "Three Quarters" Or _
            pSEP(1) = "Shallow" Or _
            pSEP(1) = "Deep" Then
            SessionParseError = False
        End If
        If SessionParseError = True Then
            MsgBox ("CutSpecialRandom Event parameter is invalid." & Chr(13) & _
                "It can only be the text from the drop down list box.")
            Exit Sub
        End If
        Call CutSpecialRandom(pSEP(1), pSEP(2))
    Case "CutDeckRandom"
        If Not (pSEP(1) = "X" Or pSEP(1) = "T" Or pSEP(1) = "B") Then
            MsgBox ("CutDeckRandom Event parameter is not valid." & Chr(13) & _
                "It must be a " & Chr(34) & "T" & Chr(34) & " or a " & _
                Chr(34) & "B" & Chr(34) & " or an " & Chr(34) & "X" & Chr(34) & ".")
            SessionParseError = True
            Exit Sub
        End If
        Call CutDeckRandom(pSEP(1))
    Case "ForceCard"
        If IsNumeric(pSEP(1)) And _
            (Val(pSEP(1)) < 1 Or _
            Val(pSEP(1)) > 52) Then
                MsgBox ("ForceCard Event first parameter is out of range." & Chr(13) & _
                "It must be between 1 and 52.")
                SessionParseError = True
                Exit Sub
        ElseIf Not IsNumeric(pSEP(1)) Then
            MsgBox ("ForceCard Event parameter is not valid." & Chr(13) & _
            "It must be a number between 1 and 52.")
            SessionParseError = True
            Exit Sub
        End If
        If Not (pSEP(2) = "X" Or pSEP(2) = "R") Then
            MsgBox ("ForceCard Event second parameter is not valid." & Chr(13) & _
                "It must be an " & Chr(34) & "R" & Chr(34) & " or an " & _
                Chr(34) & "X" & Chr(34) & ".")
            SessionParseError = True
            Exit Sub
        End If
        Call ForceCard(pSEP(1), pSEP(2))
    Case "ReturnCard"
        If IsNumeric(pSEP(1)) And _
            (Val(pSEP(1)) < 1 Or _
            Val(pSEP(1)) > 52) Then
                MsgBox ("ReturnCard Event parameter is out of range." & Chr(13) & _
                "It must be between 1 and 52.")
                SessionParseError = True
                Exit Sub
        ElseIf Not IsNumeric(pSEP(1)) Then
            SessionParseError = True
            'set the error condition first,
            'then remove it if there is a matched parameter
            If pSEP(1) = Chr(34) & "Anywhere" & Chr(34) Or _
                pSEP(1) = Chr(34) & "Original Position" & Chr(34) Or _
                pSEP(1) = "Top Third" Or _
                pSEP(1) = "Middle Third" Or _
                pSEP(1) = "Bottom Third" Then
                SessionParseError = False
            End If
            If SessionParseError = True Then
                MsgBox ("ReturnCard Event parameter is not valid." & Chr(13) & _
                "It must be a number between 1 and 52," & Chr(13) & _
                "or the text " & Chr(34) & "Anywhere" & Chr(34) & Chr(13) & _
                "or the text from the drop down list.")
                SessionParseError = True
                Exit Sub
            End If
        End If
        Call ReturnCard(pSEP(1))
    Case "SelectCardsCutSelectNext1"
        If Not (pSEP(1) = "X" Or pSEP(1) = "R") Then
            MsgBox ("SelectCardsCutSelectNext1 Event:" & Chr(13) & _
                "First parameter is not valid." & Chr(13) & _
                "It must be an " & Chr(34) & "R" & Chr(34) & " or an " & _
                Chr(34) & "X" & Chr(34) & ".")
            SessionParseError = True
            Exit Sub
        End If
        Call SelectCardsCutSelectNext1(pSEP(1))
    Case "SelectCardsCutSelectNext2"
        If Not (pSEP(1) = "X" Or pSEP(1) = "R") Then
            MsgBox ("SelectCardsCutSelectNext2 Event:" & Chr(13) & _
                "First parameter is not valid." & Chr(13) & _
                "It must be an " & Chr(34) & "R" & Chr(34) & " or an " & _
                Chr(34) & "X" & Chr(34) & ".")
            SessionParseError = True
            Exit Sub
        End If
        If Not (pSEP(2) = "X" Or pSEP(1) = "R") Then
            MsgBox ("SelectCardsCutSelectNext2 Event:" & Chr(13) & _
                "Second parameter is not valid." & Chr(13) & _
                "It must be an " & Chr(34) & "R" & Chr(34) & " or an " & _
                Chr(34) & "X" & Chr(34) & ".")
            SessionParseError = True
            Exit Sub
        End If
        Call SelectCardsCutSelectNext2(pSEP(1), pSEP(2))
    Case "SelectCardsCutSelectNext3"
        If Not (pSEP(1) = "X" Or pSEP(1) = "R") Then
            MsgBox ("SelectCardsCutSelectNext3 Event:" & Chr(13) & _
                "First parameter is not valid." & Chr(13) & _
                "It must be an " & Chr(34) & "R" & Chr(34) & " or an " & _
                Chr(34) & "X" & Chr(34) & ".")
            SessionParseError = True
            Exit Sub
        End If
        If Not (pSEP(1) = "X" Or pSEP(1) = "R") Then
            MsgBox ("SelectCardsCutSelectNext3 Event:" & Chr(13) & _
                "Second parameter is not valid." & Chr(13) & _
                "It must be an " & Chr(34) & "R" & Chr(34) & " or an " & _
                Chr(34) & "X" & Chr(34) & ".")
            SessionParseError = True
            Exit Sub
        End If
        If Not (pSEP(1) = "X" Or pSEP(1) = "R") Then
            MsgBox ("SelectCardsCutSelectNext3 Event:" & Chr(13) & _
                "Third parameter is not valid." & Chr(13) & _
                "It must be an " & Chr(34) & "R" & Chr(34) & " or an " & _
                Chr(34) & "X" & Chr(34) & ".")
            SessionParseError = True
            Exit Sub
        End If
        Call SelectCardsCutSelectNext3(pSEP(1), pSEP(2), pSEP(3))
    Case "SelectCardsCutSelectFace1"
        If Not (pSEP(1) = "X" Or pSEP(1) = "R") Then
            MsgBox ("SelectCardsCutSelectFace1 Event:" & Chr(13) & _
                "First parameter is not valid." & Chr(13) & _
                "It must be an " & Chr(34) & "R" & Chr(34) & " or an " & _
                Chr(34) & "X" & Chr(34) & ".")
            SessionParseError = True
            Exit Sub
        End If
        Call SelectCardsCutSelectFace1(pSEP(1))
    Case "SelectCardsCutSelectFace2"
        If Not (pSEP(1) = "X" Or pSEP(1) = "R") Then
            MsgBox ("SelectCardsCutSelectFace2 Event:" & Chr(13) & _
                "First parameter is not valid." & Chr(13) & _
                "It must be an " & Chr(34) & "R" & Chr(34) & " or an " & _
                Chr(34) & "X" & Chr(34) & ".")
            SessionParseError = True
            Exit Sub
        End If
        If Not (pSEP(2) = "X" Or pSEP(1) = "R") Then
            MsgBox ("SelectCardsCutSelectFace2 Event:" & Chr(13) & _
                "Second parameter is not valid." & Chr(13) & _
                "It must be an " & Chr(34) & "R" & Chr(34) & " or an " & _
                Chr(34) & "X" & Chr(34) & ".")
            SessionParseError = True
            Exit Sub
        End If
        Call SelectCardsCutSelectFace2(pSEP(1), pSEP(2))
    Case "SelectCardsCutSelectFace3"
        If Not (pSEP(1) = "X" Or pSEP(1) = "R") Then
            MsgBox ("SelectCardsCutSelectFace3 Event:" & Chr(13) & _
                "First parameter is not valid." & Chr(13) & _
                "It must be an " & Chr(34) & "R" & Chr(34) & " or an " & _
                Chr(34) & "X" & Chr(34) & ".")
            SessionParseError = True
            Exit Sub
        End If
        If Not (pSEP(1) = "X" Or pSEP(1) = "R") Then
            MsgBox ("SelectCardsCutSelectFace3 Event:" & Chr(13) & _
                "Second parameter is not valid." & Chr(13) & _
                "It must be an " & Chr(34) & "R" & Chr(34) & " or an " & _
                Chr(34) & "X" & Chr(34) & ".")
            SessionParseError = True
            Exit Sub
        End If
        If Not (pSEP(1) = "X" Or pSEP(1) = "R") Then
            MsgBox ("SelectCardsCutSelectFace3 Event:" & Chr(13) & _
                "Third parameter is not valid." & Chr(13) & _
                "It must be an " & Chr(34) & "R" & Chr(34) & " or an " & _
                Chr(34) & "X" & Chr(34) & ".")
            SessionParseError = True
            Exit Sub
        End If
        Call SelectCardsCutSelectFace3(pSEP(1), pSEP(2), pSEP(3))
    Case "SelectCardsCutSelectNextRepeat"
        If Not (pSEP(1) = "X" Or pSEP(1) = "R") Then
            MsgBox ("SelectCardsCutSelectNextRepeat Event:" & Chr(13) & _
                "First parameter is not valid." & Chr(13) & _
                "It must be an " & Chr(34) & "R" & Chr(34) & " or an " & _
                Chr(34) & "X" & Chr(34) & ".")
            SessionParseError = True
            Exit Sub
        End If
        If Not (pSEP(2) = "X" Or pSEP(1) = "R") Then
            MsgBox ("SelectCardsCutSelectNextRepeat Event:" & Chr(13) & _
                "Second parameter is not valid." & Chr(13) & _
                "It must be an " & Chr(34) & "R" & Chr(34) & " or an " & _
                Chr(34) & "X" & Chr(34) & ".")
            SessionParseError = True
            Exit Sub
        End If
        Call SelectCardsCutSelectNextRepeat(pSEP(1), pSEP(2))
    Case "SelectCardsCutSelectNextRepeat2"
        If Not (pSEP(1) = "X" Or pSEP(1) = "R") Then
            MsgBox ("SelectCardsCutSelectNextRepeat2 Event:" & Chr(13) & _
                "First parameter is not valid." & Chr(13) & _
                "It must be an " & Chr(34) & "R" & Chr(34) & " or an " & _
                Chr(34) & "X" & Chr(34) & ".")
            SessionParseError = True
            Exit Sub
        End If
        If Not (pSEP(1) = "X" Or pSEP(1) = "R") Then
            MsgBox ("SelectCardsCutSelectNextRepeat2 Event:" & Chr(13) & _
                "Second parameter is not valid." & Chr(13) & _
                "It must be an " & Chr(34) & "R" & Chr(34) & " or an " & _
                Chr(34) & "X" & Chr(34) & ".")
            SessionParseError = True
            Exit Sub
        End If
        If Not (pSEP(1) = "X" Or pSEP(1) = "R") Then
            MsgBox ("SelectCardsCutSelectNextRepeat2 Event:" & Chr(13) & _
                "Third parameter is not valid." & Chr(13) & _
                "It must be an " & Chr(34) & "R" & Chr(34) & " or an " & _
                Chr(34) & "X" & Chr(34) & ".")
            SessionParseError = True
            Exit Sub
        End If
        Call SelectCardsCutSelectNextRepeat2(pSEP(1), pSEP(2), pSEP(3))
    Case "FreeChoiceSpreadSelect"
        SessionParseError = True
        'set the error condition first,
        'then remove it if there is a matched parameter
        If _
            pSEP(1) = Chr(34) & "Any Card" & Chr(34) Or _
            pSEP(1) = "Top Third" Or _
            pSEP(1) = "Middle Third" Or _
            pSEP(1) = "Bottom Third" Then
            SessionParseError = False
        End If
        If SessionParseError = True Then
            MsgBox ("FreeChoiceSpreadSelect Event first parameter is invalid.")
            Exit Sub
        End If
        If Not (pSEP(2) = "X" Or pSEP(2) = "R") Then
            MsgBox ("FreeChoiceSpreadSelect Event second parameter is not valid." & Chr(13) & _
                "It must be an " & Chr(34) & "R" & Chr(34) & " or an " & _
                Chr(34) & "X" & Chr(34) & ".")
            SessionParseError = True
            Exit Sub
        End If
        Call FreeChoiceSpreadSelect(pSEP(1), pSEP(2))
    Case "InFaro"
        Call InFaro
    Case "InFaroReverse"
        Call InFaroReverse
    Case "InFaroSpecialTop"
        If IsNumeric(pSEP(1)) And _
            (Val(pSEP(1)) < 1 Or _
            Val(pSEP(1)) > 52) Then
                MsgBox ("First event parameter is out of range." & Chr(13) & _
                "It must be between 1 and 52.")
                SessionParseError = True
                Exit Sub
        ElseIf Not IsNumeric(pSEP(1)) Then
            MsgBox ("First event parameter is not valid." & Chr(13) & _
            "It must be a number between 1 and 52.")
            SessionParseError = True
            Exit Sub
        End If
        If IsNumeric(pSEP(2)) And _
            (Val(pSEP(2)) < 1 Or _
            Val(pSEP(2)) > 52) Then
                MsgBox ("Second event parameter is out of range." & Chr(13) & _
                "It must be between 1 and 52.")
                SessionParseError = True
                Exit Sub
        ElseIf Not IsNumeric(pSEP(2)) Then
            MsgBox ("Second event parameter is not valid." & Chr(13) & _
            "It must be a number between 1 and 52.")
            SessionParseError = True
            Exit Sub
        End If
        If (Val(pSEP(1)) + Val(pSEP(2))) > 52 Then
            MsgBox "Please enter valid card positions" & Chr(13) _
                & "for the event parameters." & Chr(13) & Chr(13) _
                & "The sum of the two event parameters" & Chr(13) _
                & "must not be greater than 52."
            SessionParseError = True
            Exit Sub
        End If
        Call InFaroSpecialTop(pSEP(1), pSEP(2))
    Case "InFaroSpecialTopReverse"
        If IsNumeric(pSEP(1)) And _
            (Val(pSEP(1)) < 1 Or _
            Val(pSEP(1)) > 52) Then
                MsgBox ("First event parameter is out of range." & Chr(13) & _
                "It must be between 1 and 52.")
                SessionParseError = True
                Exit Sub
        ElseIf Not IsNumeric(pSEP(1)) Then
            MsgBox ("First event parameter is not valid." & Chr(13) & _
            "It must be a number between 1 and 52.")
            SessionParseError = True
            Exit Sub
        End If
        If IsNumeric(pSEP(2)) And _
            (Val(pSEP(2)) < 1 Or _
            Val(pSEP(2)) > 52) Then
                MsgBox ("Second event parameter is out of range." & Chr(13) & _
                "It must be between 1 and 52.")
                SessionParseError = True
                Exit Sub
        ElseIf Not IsNumeric(pSEP(2)) Then
            MsgBox ("Second event parameter is not valid." & Chr(13) & _
            "It must be a number between 1 and 52.")
            SessionParseError = True
            Exit Sub
        End If
        If (Val(pSEP(1)) + Val(pSEP(2))) > 52 Then
            MsgBox "Please enter valid card positions" & Chr(13) _
                & "for the event parameters." & Chr(13) & Chr(13) _
                & "The sum of the two event parameters" & Chr(13) _
                & "must not be greater than 52."
            SessionParseError = True
            Exit Sub
        End If
        Call InFaroSpecialTopReverse(pSEP(1), pSEP(2))
    Case "InFaroSpecialBottom"
        If IsNumeric(pSEP(1)) And _
            (Val(pSEP(1)) < 1 Or _
            Val(pSEP(1)) > 52) Then
                MsgBox ("First event parameter is out of range." & Chr(13) & _
                "It must be between 1 and 52.")
                SessionParseError = True
                Exit Sub
        ElseIf Not IsNumeric(pSEP(1)) Then
            MsgBox ("First event parameter is not valid." & Chr(13) & _
            "It must be a number between 1 and 52.")
            SessionParseError = True
            Exit Sub
        End If
        If IsNumeric(pSEP(2)) And _
            (Val(pSEP(2)) < 1 Or _
            Val(pSEP(2)) > 52) Then
                MsgBox ("Second event parameter is out of range." & Chr(13) & _
                "It must be between 1 and 52.")
                SessionParseError = True
                Exit Sub
        ElseIf Not IsNumeric(pSEP(2)) Then
            MsgBox ("Second event parameter is not valid." & Chr(13) & _
            "It must be a number between 1 and 52.")
            SessionParseError = True
            Exit Sub
        End If
        If (Val(pSEP(1)) + Val(pSEP(2))) > 52 Then
            MsgBox "Please enter valid card positions" & Chr(13) _
                & "for the event parameters." & Chr(13) & Chr(13) _
                & "The sum of the two event parameters" & Chr(13) _
                & "must not be greater than 52."
            SessionParseError = True
            Exit Sub
        End If
        Call InFaroSpecialBottom(pSEP(1), pSEP(2))
    Case "InFaroSpecialBottomReverse"
        If IsNumeric(pSEP(1)) And _
            (Val(pSEP(1)) < 1 Or _
            Val(pSEP(1)) > 52) Then
                MsgBox ("First event parameter is out of range." & Chr(13) & _
                "It must be between 1 and 52.")
                SessionParseError = True
                Exit Sub
        ElseIf Not IsNumeric(pSEP(1)) Then
            MsgBox ("First event parameter is not valid." & Chr(13) & _
            "It must be a number between 1 and 52.")
            SessionParseError = True
            Exit Sub
        End If
        If IsNumeric(pSEP(2)) And _
            (Val(pSEP(2)) < 1 Or _
            Val(pSEP(2)) > 52) Then
                MsgBox ("Second event parameter is out of range." & Chr(13) & _
                "It must be between 1 and 52.")
                SessionParseError = True
                Exit Sub
        ElseIf Not IsNumeric(pSEP(2)) Then
            MsgBox ("Second event parameter is not valid." & Chr(13) & _
            "It must be a number between 1 and 52.")
            SessionParseError = True
            Exit Sub
        End If
        If (Val(pSEP(1)) + Val(pSEP(2))) > 52 Then
            MsgBox "Please enter valid card positions" & Chr(13) _
                & "for the event parameters." & Chr(13) & Chr(13) _
                & "The sum of the two event parameters" & Chr(13) _
                & "must not be greater than 52."
            SessionParseError = True
            Exit Sub
        End If
        Call InFaroSpecialBottomReverse(pSEP(1), pSEP(2))
    Case "OutFaro"
        Call OutFaro
    Case "OutFaroReverse"
        Call OutFaroReverse
    Case "OutFaroSpecialTop"
        If IsNumeric(pSEP(1)) And _
            (Val(pSEP(1)) < 1 Or _
            Val(pSEP(1)) > 52) Then
                MsgBox ("First event parameter is out of range." & Chr(13) & _
                "It must be between 1 and 52.")
                SessionParseError = True
                Exit Sub
        ElseIf Not IsNumeric(pSEP(1)) Then
            MsgBox ("First event parameter is not valid." & Chr(13) & _
            "It must be a number between 1 and 52.")
            SessionParseError = True
            Exit Sub
        End If
        If IsNumeric(pSEP(2)) And _
            (Val(pSEP(2)) < 1 Or _
            Val(pSEP(2)) > 52) Then
                MsgBox ("Second event parameter is out of range." & Chr(13) & _
                "It must be between 1 and 52.")
                SessionParseError = True
                Exit Sub
        ElseIf Not IsNumeric(pSEP(2)) Then
            MsgBox ("Second event parameter is not valid." & Chr(13) & _
            "It must be a number between 1 and 52.")
            SessionParseError = True
            Exit Sub
        End If
        If (Val(pSEP(1)) + Val(pSEP(2))) > 52 Then
            MsgBox "Please enter valid card positions" & Chr(13) _
                & "for the event parameters." & Chr(13) & Chr(13) _
                & "The sum of the two event parameters" & Chr(13) _
                & "must not be greater than 52."
            SessionParseError = True
            Exit Sub
        End If
        Call OutFaroSpecialTop(pSEP(1), pSEP(2))
    Case "OutFaroSpecialTopReverse"
        If IsNumeric(pSEP(1)) And _
            (Val(pSEP(1)) < 1 Or _
            Val(pSEP(1)) > 52) Then
                MsgBox ("First event parameter is out of range." & Chr(13) & _
                "It must be between 1 and 52.")
                SessionParseError = True
                Exit Sub
        ElseIf Not IsNumeric(pSEP(1)) Then
            MsgBox ("First event parameter is not valid." & Chr(13) & _
            "It must be a number between 1 and 52.")
            SessionParseError = True
            Exit Sub
        End If
        If IsNumeric(pSEP(2)) And _
            (Val(pSEP(2)) < 1 Or _
            Val(pSEP(2)) > 52) Then
                MsgBox ("Second event parameter is out of range." & Chr(13) & _
                "It must be between 1 and 52.")
                SessionParseError = True
                Exit Sub
        ElseIf Not IsNumeric(pSEP(2)) Then
            MsgBox ("Second event parameter is not valid." & Chr(13) & _
            "It must be a number between 1 and 52.")
            SessionParseError = True
            Exit Sub
        End If
        If (Val(pSEP(1)) + Val(pSEP(2))) > 52 Then
            MsgBox "Please enter valid card positions" & Chr(13) _
                & "for the event parameters." & Chr(13) & Chr(13) _
                & "The sum of the two event parameters" & Chr(13) _
                & "must not be greater than 52."
            SessionParseError = True
            Exit Sub
        End If
        Call OutFaroSpecialTopReverse(pSEP(1), pSEP(2))
    Case "OutFaroSpecialBottom"
        If IsNumeric(pSEP(1)) And _
            (Val(pSEP(1)) < 1 Or _
            Val(pSEP(1)) > 52) Then
                MsgBox ("First event parameter is out of range." & Chr(13) & _
                "It must be between 1 and 52.")
                SessionParseError = True
                Exit Sub
        ElseIf Not IsNumeric(pSEP(1)) Then
            MsgBox ("First event parameter is not valid." & Chr(13) & _
            "It must be a number between 1 and 52.")
            SessionParseError = True
            Exit Sub
        End If
        If IsNumeric(pSEP(2)) And _
            (Val(pSEP(2)) < 1 Or _
            Val(pSEP(2)) > 52) Then
                MsgBox ("Second event parameter is out of range." & Chr(13) & _
                "It must be between 1 and 52.")
                SessionParseError = True
                Exit Sub
        ElseIf Not IsNumeric(pSEP(2)) Then
            MsgBox ("Second event parameter is not valid." & Chr(13) & _
            "It must be a number between 1 and 52.")
            SessionParseError = True
            Exit Sub
        End If
        If (Val(pSEP(1)) + Val(pSEP(2))) > 52 Then
            MsgBox "Please enter valid card positions" & Chr(13) _
                & "for the event parameters." & Chr(13) & Chr(13) _
                & "The sum of the two event parameters" & Chr(13) _
                & "must not be greater than 52."
            SessionParseError = True
            Exit Sub
        End If
        Call OutFaroSpecialBottom(pSEP(1), pSEP(2))
    Case "OutFaroSpecialBottomReverse"
        If IsNumeric(pSEP(1)) And _
            (Val(pSEP(1)) < 1 Or _
            Val(pSEP(1)) > 52) Then
                MsgBox ("First event parameter is out of range." & Chr(13) & _
                "It must be between 1 and 52.")
                SessionParseError = True
                Exit Sub
        ElseIf Not IsNumeric(pSEP(1)) Then
            MsgBox ("First event parameter is not valid." & Chr(13) & _
            "It must be a number between 1 and 52.")
            SessionParseError = True
            Exit Sub
        End If
        If IsNumeric(pSEP(2)) And _
            (Val(pSEP(2)) < 1 Or _
            Val(pSEP(2)) > 52) Then
                MsgBox ("Second event parameter is out of range." & Chr(13) & _
                "It must be between 1 and 52.")
                SessionParseError = True
                Exit Sub
        ElseIf Not IsNumeric(pSEP(2)) Then
            MsgBox ("Second event parameter is not valid." & Chr(13) & _
            "It must be a number between 1 and 52.")
            SessionParseError = True
            Exit Sub
        End If
        If (Val(pSEP(1)) + Val(pSEP(2))) > 52 Then
            MsgBox "Please enter valid card positions" & Chr(13) _
                & "for the event parameters." & Chr(13) & Chr(13) _
                & "The sum of the two event parameters" & Chr(13) _
                & "must not be greater than 52."
            SessionParseError = True
            Exit Sub
        End If
        Call OutFaroSpecialBottomReverse(pSEP(1), pSEP(2))
    Case "OHShuffle"
        Call OHShuffle
    Case "OHShuffleTop"
        If IsNumeric(pSEP(1)) And _
            (Val(pSEP(1)) < 1 Or _
            Val(pSEP(1)) > 52) Then
                MsgBox ("Overhand Shuffle Event parameter is out of range." & Chr(13) & _
                "It must be between 1 and 52.")
                SessionParseError = True
                Exit Sub
        ElseIf Not IsNumeric(pSEP(1)) Then
            MsgBox ("Overhand Shuffle Event parameter is not valid." & Chr(13) & _
            "It must be a number between 1 and 52.")
            SessionParseError = True
            Exit Sub
        End If
        Call OHShuffleTop(pSEP(1))
    Case "OHShuffleBottom"
        If IsNumeric(pSEP(1)) And _
            (Val(pSEP(1)) < 1 Or _
            Val(pSEP(1)) > 52) Then
                MsgBox ("Overhand Shuffle Event parameter is out of range." & Chr(13) & _
                "It must be between 1 and 52.")
                SessionParseError = True
                Exit Sub
        ElseIf Not IsNumeric(pSEP(1)) Then
            MsgBox ("Overhand Shuffle Event parameter is not valid." & Chr(13) & _
            "It must be a number between 1 and 52.")
            SessionParseError = True
            Exit Sub
        End If
        Call OHShuffleBottom(pSEP(1))
    Case "PokerDeal"
        If IsNumeric(pSEP(1)) And _
            (Val(pSEP(1)) < 2 Or _
            Val(pSEP(1)) > 10) Then
                MsgBox ("Poker Deal Event parameter is out of range." & Chr(13) & _
                "It must be between 2 and 10.")
                SessionParseError = True
                Exit Sub
        ElseIf Not IsNumeric(pSEP(1)) Then
            MsgBox ("Poker Deal Event parameter is not valid." & Chr(13) & _
            "It must be a number between 1 and 52.")
            SessionParseError = True
            Exit Sub
        End If
        Call PokerDeal(pSEP(1))
    Case "AssemblePokerDeal"
        SessionParseError = True
        'set the error condition first,
        'then remove it if there is a matched parameter
        If _
            pSEP(1) = "Backwards" Or _
            pSEP(1) = "Forwards" Or _
            pSEP(1) = "Unwind" Then
            SessionParseError = False
        End If
        If SessionParseError = True Then
            MsgBox ("Assemble Poker Deal Event parameter is invalid.")
            Exit Sub
        End If
        Call AssemblePokerDeal(pSEP(1))
    Case "SetStack"
        SessionParseError = True
        'set the error condition first,
        'then remove it if there is a matched parameter
        If _
            pSEP(1) = Chr(34) & "Default" & Chr(34) Or _
            pSEP(1) = Chr(34) & "Aronson" & Chr(34) Or _
            pSEP(1) = Chr(34) & "Ireland" & Chr(34) Or _
            pSEP(1) = Chr(34) & "Eight Kings" & Chr(34) Or _
            pSEP(1) = Chr(34) & "Joyal (CHaSeD)" & Chr(34) Or _
            pSEP(1) = Chr(34) & "Joyal (SHoCkeD)" & Chr(34) Or _
            pSEP(1) = Chr(34) & "New Deck (Bicycle)" & Chr(34) Or _
            pSEP(1) = Chr(34) & "New Deck (Fournier)" & Chr(34) Or _
            pSEP(1) = Chr(34) & "Nikola" & Chr(34) Or _
            pSEP(1) = Chr(34) & "Osterlind" & Chr(34) Or _
            pSEP(1) = Chr(34) & "Si Stebbins (3)" & Chr(34) Or _
            pSEP(1) = Chr(34) & "Si Stebbins (4)" & Chr(34) Or _
            pSEP(1) = Chr(34) & "Stanyon" & Chr(34) Or _
            pSEP(1) = Chr(34) & "Tamariz" & Chr(34) Then
            SessionParseError = False
        End If
        If SessionParseError = True Then
            MsgBox ("Set Stack Event parameter is invalid.")
            Exit Sub
        End If
        Call SetStack(pSEP(1))
    Case "ResetCurrentDeck"
        Call ResetCurrentDeck
    Case "InverseInFaro"
        Call InverseInFaro
    Case "InverseInFaroReverse"
        Call InverseInFaroReverse
    Case "InverseInFaroSpecialTop"
        If IsNumeric(pSEP(1)) And _
            (Val(pSEP(1)) < 1 Or _
            Val(pSEP(1)) > 52) Then
                MsgBox ("First event parameter is out of range." & Chr(13) & _
                "It must be between 1 and 52.")
                SessionParseError = True
                Exit Sub
        ElseIf Not IsNumeric(pSEP(1)) Then
            MsgBox ("First event parameter is not valid." & Chr(13) & _
            "It must be a number between 1 and 52.")
            SessionParseError = True
            Exit Sub
        End If
        If IsNumeric(pSEP(2)) And _
            (Val(pSEP(2)) < 1 Or _
            Val(pSEP(2)) > 52) Then
                MsgBox ("Second event parameter is out of range." & Chr(13) & _
                "It must be between 1 and 52.")
                SessionParseError = True
                Exit Sub
        ElseIf Not IsNumeric(pSEP(2)) Then
            MsgBox ("Second event parameter is not valid." & Chr(13) & _
            "It must be a number between 1 and 52.")
            SessionParseError = True
            Exit Sub
        End If
        If (Val(pSEP(1)) + Val(pSEP(2))) > 52 Then
            MsgBox "Please enter valid card positions" & Chr(13) _
                & "for the event parameters." & Chr(13) & Chr(13) _
                & "The sum of the two event parameters" & Chr(13) _
                & "must not be greater than 52."
            SessionParseError = True
            Exit Sub
        End If
        Call InverseInFaroSpecialTop(pSEP(1), pSEP(2))
    Case "InverseInFaroSpecialTopReverse"
        If IsNumeric(pSEP(1)) And _
            (Val(pSEP(1)) < 1 Or _
            Val(pSEP(1)) > 52) Then
                MsgBox ("First event parameter is out of range." & Chr(13) & _
                "It must be between 1 and 52.")
                SessionParseError = True
                Exit Sub
        ElseIf Not IsNumeric(pSEP(1)) Then
            MsgBox ("First event parameter is not valid." & Chr(13) & _
            "It must be a number between 1 and 52.")
            SessionParseError = True
            Exit Sub
        End If
        If IsNumeric(pSEP(2)) And _
            (Val(pSEP(2)) < 1 Or _
            Val(pSEP(2)) > 52) Then
                MsgBox ("Second event parameter is out of range." & Chr(13) & _
                "It must be between 1 and 52.")
                SessionParseError = True
                Exit Sub
        ElseIf Not IsNumeric(pSEP(2)) Then
            MsgBox ("Second event parameter is not valid." & Chr(13) & _
            "It must be a number between 1 and 52.")
            SessionParseError = True
            Exit Sub
        End If
        If (Val(pSEP(1)) + Val(pSEP(2))) > 52 Then
            MsgBox "Please enter valid card positions" & Chr(13) _
                & "for the event parameters." & Chr(13) & Chr(13) _
                & "The sum of the two event parameters" & Chr(13) _
                & "must not be greater than 52."
            SessionParseError = True
            Exit Sub
        End If
        Call InverseInFaroSpecialTopReverse(pSEP(1), pSEP(2))
    Case "InverseInFaroSpecialBottom"
        If IsNumeric(pSEP(1)) And _
            (Val(pSEP(1)) < 1 Or _
            Val(pSEP(1)) > 52) Then
                MsgBox ("First event parameter is out of range." & Chr(13) & _
                "It must be between 1 and 52.")
                SessionParseError = True
                Exit Sub
        ElseIf Not IsNumeric(pSEP(1)) Then
            MsgBox ("First event parameter is not valid." & Chr(13) & _
            "It must be a number between 1 and 52.")
            SessionParseError = True
            Exit Sub
        End If
        If IsNumeric(pSEP(2)) And _
            (Val(pSEP(2)) < 1 Or _
            Val(pSEP(2)) > 52) Then
                MsgBox ("Second event parameter is out of range." & Chr(13) & _
                "It must be between 1 and 52.")
                SessionParseError = True
                Exit Sub
        ElseIf Not IsNumeric(pSEP(2)) Then
            MsgBox ("Second event parameter is not valid." & Chr(13) & _
            "It must be a number between 1 and 52.")
            SessionParseError = True
            Exit Sub
        End If
        If (Val(pSEP(1)) + Val(pSEP(2))) > 52 Then
            MsgBox "Please enter valid card positions" & Chr(13) _
                & "for the event parameters." & Chr(13) & Chr(13) _
                & "The sum of the two event parameters" & Chr(13) _
                & "must not be greater than 52."
            SessionParseError = True
            Exit Sub
        End If
        Call InverseInFaroSpecialBottom(pSEP(1), pSEP(2))
    Case "InverseInFaroSpecialBottomReverse"
        If IsNumeric(pSEP(1)) And _
            (Val(pSEP(1)) < 1 Or _
            Val(pSEP(1)) > 52) Then
                MsgBox ("First event parameter is out of range." & Chr(13) & _
                "It must be between 1 and 52.")
                SessionParseError = True
                Exit Sub
        ElseIf Not IsNumeric(pSEP(1)) Then
            MsgBox ("First event parameter is not valid." & Chr(13) & _
            "It must be a number between 1 and 52.")
            SessionParseError = True
            Exit Sub
        End If
        If IsNumeric(pSEP(2)) And _
            (Val(pSEP(2)) < 1 Or _
            Val(pSEP(2)) > 52) Then
                MsgBox ("Second event parameter is out of range." & Chr(13) & _
                "It must be between 1 and 52.")
                SessionParseError = True
                Exit Sub
        ElseIf Not IsNumeric(pSEP(2)) Then
            MsgBox ("Second event parameter is not valid." & Chr(13) & _
            "It must be a number between 1 and 52.")
            SessionParseError = True
            Exit Sub
        End If
        If (Val(pSEP(1)) + Val(pSEP(2))) > 52 Then
            MsgBox "Please enter valid card positions" & Chr(13) _
                & "for the event parameters." & Chr(13) & Chr(13) _
                & "The sum of the two event parameters" & Chr(13) _
                & "must not be greater than 52."
            SessionParseError = True
            Exit Sub
        End If
        Call InverseInFaroSpecialBottomReverse(pSEP(1), pSEP(2))
    Case "InverseOutFaro"
        Call InverseOutFaro
    Case "InverseOutFaroReverse"
        Call InverseOutFaroReverse
    Case "InverseOutFaroSpecialTop"
        If IsNumeric(pSEP(1)) And _
            (Val(pSEP(1)) < 1 Or _
            Val(pSEP(1)) > 52) Then
                MsgBox ("First event parameter is out of range." & Chr(13) & _
                "It must be between 1 and 52.")
                SessionParseError = True
                Exit Sub
        ElseIf Not IsNumeric(pSEP(1)) Then
            MsgBox ("First event parameter is not valid." & Chr(13) & _
            "It must be a number between 1 and 52.")
            SessionParseError = True
            Exit Sub
        End If
        If IsNumeric(pSEP(2)) And _
            (Val(pSEP(2)) < 1 Or _
            Val(pSEP(2)) > 52) Then
                MsgBox ("Second event parameter is out of range." & Chr(13) & _
                "It must be between 1 and 52.")
                SessionParseError = True
                Exit Sub
        ElseIf Not IsNumeric(pSEP(2)) Then
            MsgBox ("Second event parameter is not valid." & Chr(13) & _
            "It must be a number between 1 and 52.")
            SessionParseError = True
            Exit Sub
        End If
        If (Val(pSEP(1)) + Val(pSEP(2))) > 52 Then
            MsgBox "Please enter valid card positions" & Chr(13) _
                & "for the event parameters." & Chr(13) & Chr(13) _
                & "The sum of the two event parameters" & Chr(13) _
                & "must not be greater than 52."
            SessionParseError = True
            Exit Sub
        End If
        Call InverseOutFaroSpecialTop(pSEP(1), pSEP(2))
    Case "InverseOutFaroSpecialTopReverse"
        If IsNumeric(pSEP(1)) And _
            (Val(pSEP(1)) < 1 Or _
            Val(pSEP(1)) > 52) Then
                MsgBox ("First event parameter is out of range." & Chr(13) & _
                "It must be between 1 and 52.")
                SessionParseError = True
                Exit Sub
        ElseIf Not IsNumeric(pSEP(1)) Then
            MsgBox ("First event parameter is not valid." & Chr(13) & _
            "It must be a number between 1 and 52.")
            SessionParseError = True
            Exit Sub
        End If
        If IsNumeric(pSEP(2)) And _
            (Val(pSEP(2)) < 1 Or _
            Val(pSEP(2)) > 52) Then
                MsgBox ("Second event parameter is out of range." & Chr(13) & _
                "It must be between 1 and 52.")
                SessionParseError = True
                Exit Sub
        ElseIf Not IsNumeric(pSEP(2)) Then
            MsgBox ("Second event parameter is not valid." & Chr(13) & _
            "It must be a number between 1 and 52.")
            SessionParseError = True
            Exit Sub
        End If
        If (Val(pSEP(1)) + Val(pSEP(2))) > 52 Then
            MsgBox "Please enter valid card positions" & Chr(13) _
                & "for the event parameters." & Chr(13) & Chr(13) _
                & "The sum of the two event parameters" & Chr(13) _
                & "must not be greater than 52."
            SessionParseError = True
            Exit Sub
        End If
        Call InverseOutFaroSpecialTopReverse(pSEP(1), pSEP(2))
    Case "InverseOutFaroSpecialBottom"
        If IsNumeric(pSEP(1)) And _
            (Val(pSEP(1)) < 1 Or _
            Val(pSEP(1)) > 52) Then
                MsgBox ("First event parameter is out of range." & Chr(13) & _
                "It must be between 1 and 52.")
                SessionParseError = True
                Exit Sub
        ElseIf Not IsNumeric(pSEP(1)) Then
            MsgBox ("First event parameter is not valid." & Chr(13) & _
            "It must be a number between 1 and 52.")
            SessionParseError = True
            Exit Sub
        End If
        If IsNumeric(pSEP(2)) And _
            (Val(pSEP(2)) < 1 Or _
            Val(pSEP(2)) > 52) Then
                MsgBox ("Second event parameter is out of range." & Chr(13) & _
                "It must be between 1 and 52.")
                SessionParseError = True
                Exit Sub
        ElseIf Not IsNumeric(pSEP(2)) Then
            MsgBox ("Second event parameter is not valid." & Chr(13) & _
            "It must be a number between 1 and 52.")
            SessionParseError = True
            Exit Sub
        End If
        If (Val(pSEP(1)) + Val(pSEP(2))) > 52 Then
            MsgBox "Please enter valid card positions" & Chr(13) _
                & "for the event parameters." & Chr(13) & Chr(13) _
                & "The sum of the two event parameters" & Chr(13) _
                & "must not be greater than 52."
            SessionParseError = True
            Exit Sub
        End If
        Call InverseOutFaroSpecialBottom(pSEP(1), pSEP(2))
    Case "InverseOutFaroSpecialBottomReverse"
        If IsNumeric(pSEP(1)) And _
            (Val(pSEP(1)) < 1 Or _
            Val(pSEP(1)) > 52) Then
                MsgBox ("First event parameter is out of range." & Chr(13) & _
                "It must be between 1 and 52.")
                SessionParseError = True
                Exit Sub
        ElseIf Not IsNumeric(pSEP(1)) Then
            MsgBox ("First event parameter is not valid." & Chr(13) & _
            "It must be a number between 1 and 52.")
            SessionParseError = True
            Exit Sub
        End If
        If IsNumeric(pSEP(2)) And _
            (Val(pSEP(2)) < 1 Or _
            Val(pSEP(2)) > 52) Then
                MsgBox ("Second event parameter is out of range." & Chr(13) & _
                "It must be between 1 and 52.")
                SessionParseError = True
                Exit Sub
        ElseIf Not IsNumeric(pSEP(2)) Then
            MsgBox ("Second event parameter is not valid." & Chr(13) & _
            "It must be a number between 1 and 52.")
            SessionParseError = True
            Exit Sub
        End If
        If (Val(pSEP(1)) + Val(pSEP(2))) > 52 Then
            MsgBox "Please enter valid card positions" & Chr(13) _
                & "for the event parameters." & Chr(13) & Chr(13) _
                & "The sum of the two event parameters" & Chr(13) _
                & "must not be greater than 52."
            SessionParseError = True
            Exit Sub
        End If
        Call InverseOutFaroSpecialBottomReverse(pSEP(1), pSEP(2))
    Case "RiffleShuffle"
        If Not (pSEP(1) = "X" Or pSEP(1) = "T" Or pSEP(1) = "B") Then
            MsgBox ("RiffleShuffle Event parameter is not valid." & Chr(13) & _
                "It must be a " & Chr(34) & "T" & Chr(34) & " or a " & _
                Chr(34) & "B" & Chr(34) & " or an " & Chr(34) & "X" & Chr(34) & ".")
            SessionParseError = True
            Exit Sub
        End If
        Call RiffleShuffle(pSEP(1))
    Case "RiffleShuffleTop"
        If IsNumeric(pSEP(1)) And _
            (Val(pSEP(1)) < 1 Or _
            Val(pSEP(1)) > 52) Then
                MsgBox ("Riffle Shuffle Event parameter is out of range." & Chr(13) & _
                "It must be between 1 and 52.")
                SessionParseError = True
                Exit Sub
        ElseIf Not IsNumeric(pSEP(1)) Then
            MsgBox ("Riffle Shuffle Event parameter is not valid." & Chr(13) & _
            "It must be a number between 1 and 52.")
            SessionParseError = True
            Exit Sub
        End If
        If Not (pSEP(2) = "X" Or pSEP(2) = "T" Or pSEP(2) = "B") Then
            MsgBox ("RiffleShuffleTop second Event parameter is not valid." & Chr(13) & _
                "It must be a " & Chr(34) & "T" & Chr(34) & " or a " & _
                Chr(34) & "B" & Chr(34) & " or an " & Chr(34) & "X" & Chr(34) & ".")
            SessionParseError = True
            Exit Sub
        End If
        Call RiffleShuffleTop(pSEP(1), pSEP(2))
    Case "RiffleShuffleBottom"
        If IsNumeric(pSEP(1)) And _
            (Val(pSEP(1)) < 1 Or _
            Val(pSEP(1)) > 52) Then
                MsgBox ("Riffle Shuffle Event parameter is out of range." & Chr(13) & _
                "It must be between 1 and 52.")
                SessionParseError = True
                Exit Sub
        ElseIf Not IsNumeric(pSEP(1)) Then
            MsgBox ("Riffle Shuffle Event parameter is not valid." & Chr(13) & _
            "It must be a number between 1 and 52.")
            SessionParseError = True
            Exit Sub
        End If
        If Not (pSEP(2) = "X" Or pSEP(2) = "T" Or pSEP(2) = "B") Then
            MsgBox ("RiffleShuffleBottom second Event parameter is not valid." & Chr(13) & _
                "It must be a " & Chr(34) & "T" & Chr(34) & " or a " & _
                Chr(34) & "B" & Chr(34) & " or an " & Chr(34) & "X" & Chr(34) & ".")
            SessionParseError = True
            Exit Sub
        End If
        Call RiffleShuffleBottom(pSEP(1), pSEP(2))
    Case "RunSingleCards"
        If IsNumeric(pSEP(1)) And _
            (Val(pSEP(1)) < 1 Or _
            Val(pSEP(1)) > 52) Then
                MsgBox ("Run Single Cards Event parameter is out of range." & Chr(13) & _
                "It must be between 1 and 52.")
                SessionParseError = True
                Exit Sub
        ElseIf Not IsNumeric(pSEP(1)) Then
            MsgBox ("Run Single Cards Event parameter is not valid." & Chr(13) & _
            "It must be a number between 1 and 52.")
            SessionParseError = True
            Exit Sub
        End If
        Call RunSingleCards(pSEP(1))
    Case "InverseRunSingleCards"
        If IsNumeric(pSEP(1)) And _
            (Val(pSEP(1)) < 1 Or _
            Val(pSEP(1)) > 52) Then
                MsgBox ("Run Single Cards Event parameter is out of range." & Chr(13) & _
                "It must be between 1 and 52.")
                SessionParseError = True
                Exit Sub
        ElseIf Not IsNumeric(pSEP(1)) Then
            MsgBox ("Run Single Cards Event parameter is not valid." & Chr(13) & _
            "It must be a number between 1 and 52.")
            SessionParseError = True
            Exit Sub
        End If
        Call InverseRunSingleCards(pSEP(1))
    Case "RunSingleCardsReverse"
        If IsNumeric(pSEP(1)) And _
            (Val(pSEP(1)) < 1 Or _
            Val(pSEP(1)) > 52) Then
                MsgBox ("Run Single Cards Event parameter is out of range." & Chr(13) & _
                "It must be between 1 and 52.")
                SessionParseError = True
                Exit Sub
        ElseIf Not IsNumeric(pSEP(1)) Then
            MsgBox ("Run Single Cards Event parameter is not valid." & Chr(13) & _
            "It must be a number between 1 and 52.")
            SessionParseError = True
            Exit Sub
        End If
        Call RunSingleCardsReverse(pSEP(1))
    Case "InverseRunSingleCardsReverse"
        If IsNumeric(pSEP(1)) And _
            (Val(pSEP(1)) < 1 Or _
            Val(pSEP(1)) > 52) Then
                MsgBox ("Run Single Cards Event parameter is out of range." & Chr(13) & _
                "It must be between 1 and 52.")
                SessionParseError = True
                Exit Sub
        ElseIf Not IsNumeric(pSEP(1)) Then
            MsgBox ("Run Single Cards Event parameter is not valid." & Chr(13) & _
            "It must be a number between 1 and 52.")
            SessionParseError = True
            Exit Sub
        End If
        Call InverseRunSingleCardsReverse(pSEP(1))
    Case "CreatePilesDealAlternatingRandom"
        If IsNumeric(pSEP(1)) And _
            (Val(pSEP(1)) < 1 Or _
            Val(pSEP(1)) > 8) Then
                MsgBox ("Create Piles Deal Alternating Random" & Chr(13) & _
                "first Event parameter is out of range." & Chr(13) & _
                "It must be between 1 and 8.")
                SessionParseError = True
                Exit Sub
        ElseIf Not IsNumeric(pSEP(1)) Then
            MsgBox ("Create Piles Deal Alternating Random" & Chr(13) & _
            "first Event parameter is not valid." & Chr(13) & _
            "It must be a number between 1 and 8.")
            SessionParseError = True
            Exit Sub
        End If
        If IsNumeric(pSEP(2)) And _
            (Val(pSEP(2)) < Val(pSEP(1)) Or _
            Val(pSEP(2)) > 52) Then
                MsgBox ("Create Piles Deal Alternating Random" & Chr(13) & _
                "second Event parameter is out of range." & Chr(13) & _
                "It must be between " & pSEP(1) & "(first parameter) and 52.")
                SessionParseError = True
                Exit Sub
        ElseIf Not IsNumeric(pSEP(2)) Then
            MsgBox ("Create Piles Deal Alternating Random" & Chr(13) & _
            "second Event parameter is not valid." & Chr(13) & _
            "It must be a number between " & pSEP(1) & "(first parameter) and 52.")
            SessionParseError = True
            Exit Sub
        End If
        If Val(pSEP(1)) = 8 And Val(pSEP(2)) < 52 Then
            MsgBox ("If you are dealing less than the full deck," & Chr(13) & _
            "you can only deal up to 7 piles.  (The 8th pile" & Chr(13) & _
            "is reserved for the remaining cards.)")
            SessionParseError = True
            Exit Sub
        End If
        Call frmPiles.CreatePilesDealAlternatingRandom(pSEP(1), pSEP(2))
    Case "CreatePilesDealAlternatingRegular"
        If IsNumeric(pSEP(1)) And _
            (Val(pSEP(1)) < 1 Or _
            Val(pSEP(1)) > 8) Then
                MsgBox ("Create Piles Deal Alternating Regular" & Chr(13) & _
                "first Event parameter is out of range." & Chr(13) & _
                "It must be between 1 and 8.")
                SessionParseError = True
                Exit Sub
        ElseIf Not IsNumeric(pSEP(1)) Then
            MsgBox ("Create Piles Deal Alternating Regular" & Chr(13) & _
            "first Event parameter is not valid." & Chr(13) & _
            "It must be a number between 1 and 8.")
            SessionParseError = True
            Exit Sub
        End If
        If IsNumeric(pSEP(2)) And _
            (Val(pSEP(2)) < Val(pSEP(1)) Or _
            Val(pSEP(2)) > 52) Then
                MsgBox ("Create Piles Deal Alternating Regular" & Chr(13) & _
                "second Event parameter is out of range." & Chr(13) & _
                "It must be between " & pSEP(1) & "(first parameter) and 52.")
                SessionParseError = True
                Exit Sub
        ElseIf Not IsNumeric(pSEP(2)) Then
            MsgBox ("Create Piles Deal Alternating Regular" & Chr(13) & _
            "second Event parameter is not valid." & Chr(13) & _
            "It must be a number between " & pSEP(1) & "(first parameter) and 52.")
            SessionParseError = True
            Exit Sub
        End If
        If Val(pSEP(1)) = 8 And Val(pSEP(2)) < 52 Then
            MsgBox ("If you are dealing less than the full deck," & Chr(13) & _
            "you can only deal up to 7 piles.  (The 8th pile" & Chr(13) & _
            "is reserved for the remaining cards.)")
            SessionParseError = True
            Exit Sub
        End If
        Call frmPiles.CreatePilesDealAlternatingRegular(pSEP(1), pSEP(2))
    Case "CreatePilesDealCompleteRandom"
        If IsNumeric(pSEP(1)) And _
            (Val(pSEP(1)) < 1 Or _
            Val(pSEP(1)) > 8) Then
                MsgBox ("Create Piles Deal Complete Random" & Chr(13) & _
                "first Event parameter is out of range." & Chr(13) & _
                "It must be between 1 and 8.")
                SessionParseError = True
                Exit Sub
        ElseIf Not IsNumeric(pSEP(1)) Then
            MsgBox ("Create Piles Deal Complete Random" & Chr(13) & _
            "first Event parameter is not valid." & Chr(13) & _
            "It must be a number between 1 and 8.")
            SessionParseError = True
            Exit Sub
        End If
        If IsNumeric(pSEP(2)) And _
            (Val(pSEP(2)) < Val(pSEP(1)) Or _
            Val(pSEP(2)) > 52) Then
                MsgBox ("Create Piles Deal Complete Random" & Chr(13) & _
                "second Event parameter is out of range." & Chr(13) & _
                "It must be between " & pSEP(1) & "(first parameter) and 52.")
                SessionParseError = True
                Exit Sub
        ElseIf Not IsNumeric(pSEP(2)) Then
            MsgBox ("Create Piles Deal Complete Random" & Chr(13) & _
            "second Event parameter is not valid." & Chr(13) & _
            "It must be a number between " & pSEP(1) & "(first parameter) and 52.")
            SessionParseError = True
            Exit Sub
        End If
        If Val(pSEP(1)) = 8 And Val(pSEP(2)) < 52 Then
            MsgBox ("If you are dealing less than the full deck," & Chr(13) & _
            "you can only deal up to 7 piles.  (The 8th pile" & Chr(13) & _
            "is reserved for the remaining cards.)")
            SessionParseError = True
            Exit Sub
        End If
        Call frmPiles.CreatePilesDealCompleteRandom(pSEP(1), pSEP(2))
    Case "CreatePilesDealComplete"
        'pPiles
        If IsNumeric(pSEP(1)) And _
            (Val(pSEP(1)) < 1 Or _
            Val(pSEP(1)) > 8) Then
                MsgBox ("Create Piles Deal Complete" & Chr(13) & _
                "first Event parameter is out of range." & Chr(13) & _
                "It must be between 1 and 8.")
                SessionParseError = True
                Exit Sub
        ElseIf Not IsNumeric(pSEP(1)) Then
            MsgBox ("Create Piles Deal Complete" & Chr(13) & _
            "first Event parameter is not valid." & Chr(13) & _
            "It must be a number between 1 and 8.")
            SessionParseError = True
            Exit Sub
        End If
        'pCode1-8 (pSEP(2 - 9))
        For i% = 2 To 9 'pCode1-8
            SessionParseError = True
            'set the error condition first,
            'then remove it if there is a matched parameter
            'check for "R"
            If pSEP(i%) = "R" Then
                SessionParseError = False
            End If
            'check for "X"
            If pSEP(i%) = "X" Then
                SessionParseError = False
            End If
            'check for "Rxx" - where xx must be greater than 0
            If Len(pSEP(i%)) > 1 And Left(pSEP(i%), 1) = "R" Then
                If IsNumeric(Right(pSEP(i%), Len(pSEP(i%)) - 1)) And _
                    Val(Right(pSEP(i%), Len(pSEP(i%)) - 1)) > 0 Then
                    SessionParseError = False
                End If
            End If
            'check for "Sxx" - where xx must be from within 1 and 52
            If Len(pSEP(i%)) > 1 And Left(pSEP(i%), 1) = "S" Then
                If IsNumeric(Right(pSEP(i%), Len(pSEP(i%)) - 1)) And _
                    Val(Right(pSEP(i%), Len(pSEP(i%)) - 1)) >= 1 And _
                    Val(Right(pSEP(i%), Len(pSEP(i%)) - 1)) <= 52 Then
                    SessionParseError = False
                End If
            End If
            If SessionParseError = True Then
                If i% = 2 Then
                    pTextPlug = "second"
                ElseIf i% = 3 Then
                    pTextPlug = "third"
                ElseIf i% = 4 Then
                    pTextPlug = "fourth"
                ElseIf i% = 5 Then
                    pTextPlug = "fifth"
                ElseIf i% = 6 Then
                    pTextPlug = "sixth"
                ElseIf i% = 7 Then
                    pTextPlug = "seventh"
                ElseIf i% = 8 Then
                    pTextPlug = "eighth"
                ElseIf i% = 9 Then
                    pTextPlug = "ninth"
                End If
                MsgBox ("Create Piles Deal Complete" & Chr(13) & _
                    "The " & pTextPlug & " Event parameter is invalid." & Chr(13) & Chr(13) & _
                    "Allowable parameter entries are:" & Chr(13) & _
                    Chr(34) & "R" & Chr(34) & " (for Random)" & Chr(13) & _
                    Chr(34) & "Rxx" & Chr(34) & " (xx > 0)" & Chr(13) & _
                    Chr(34) & "Syy" & Chr(34) & " (Specified yy from 1 to 52)" & Chr(13) & _
                    Chr(34) & "X" & Chr(34) & " (for excluded pile)" & Chr(13))
                Exit Sub
            End If
        Next i%
        Call frmPiles.CreatePilesDealComplete(pSEP(1), pSEP(2), pSEP(3), pSEP(4), pSEP(5), pSEP(6), pSEP(7), pSEP(8), pSEP(9))
    Case "CreatePilesCut"
        'pPiles
        If IsNumeric(pSEP(1)) And _
            (Val(pSEP(1)) < 1 Or _
            Val(pSEP(1)) > 8) Then
                MsgBox ("Create Piles Deal Cut" & Chr(13) & _
                "first Event parameter is out of range." & Chr(13) & _
                "It must be between 1 and 8.")
                SessionParseError = True
                Exit Sub
        ElseIf Not IsNumeric(pSEP(1)) Then
            MsgBox ("Create Piles Cut" & Chr(13) & _
            "first Event parameter is not valid." & Chr(13) & _
            "It must be a number between 1 and 8.")
            SessionParseError = True
            Exit Sub
        End If
        'pCode1-8 (pSEP(2 - 9))
        SessionParseError = True
        'set the error condition first,
        'then remove it if there is a matched parameter
        For i% = 2 To 9 'pCode1-8
            'check for "R"
            If pSEP(i%) = "R" Then
                SessionParseError = False
            End If
            'check for "X"
            If pSEP(i%) = "X" Then
                SessionParseError = False
            End If
            'check for "Rxx" - where xx must be greater than 0
            If Len(pSEP(i%)) > 1 And Left(pSEP(i%), 1) = "R" Then
                If IsNumeric(Right(pSEP(i%), Len(pSEP(i%)) - 1)) And _
                    Val(Right(pSEP(i%), Len(pSEP(i%)) - 1)) > 0 Then
                    SessionParseError = False
                End If
            End If
            'check for "Sxx" - where xx must be from within 1 and 52
            If Len(pSEP(i%)) > 1 And Left(pSEP(i%), 1) = "S" Then
                If IsNumeric(Right(pSEP(i%), Len(pSEP(i%)) - 1)) And _
                    Val(Right(pSEP(i%), Len(pSEP(i%)) - 1)) >= 1 And _
                    Val(Right(pSEP(i%), Len(pSEP(i%)) - 1)) <= 52 Then
                    SessionParseError = False
                End If
            End If
            If SessionParseError = True Then
                If i% = 2 Then
                    pTextPlug = "second"
                ElseIf i% = 3 Then
                    pTextPlug = "third"
                ElseIf i% = 4 Then
                    pTextPlug = "fourth"
                ElseIf i% = 5 Then
                    pTextPlug = "fifth"
                ElseIf i% = 6 Then
                    pTextPlug = "sixth"
                ElseIf i% = 7 Then
                    pTextPlug = "seventh"
                ElseIf i% = 8 Then
                    pTextPlug = "eighth"
                ElseIf i% = 9 Then
                    pTextPlug = "ninth"
                End If
                MsgBox ("Create Piles Cut" & Chr(13) & _
                    "The " & pTextPlug & " Event parameter is invalid." & Chr(13) & Chr(13) & _
                    "Allowable parameter entries are:" & Chr(13) & _
                    Chr(34) & "R" & Chr(34) & " (for Random)" & Chr(13) & _
                    Chr(34) & "Rxx" & Chr(34) & " (xx > 0)" & Chr(13) & _
                    Chr(34) & "Syy" & Chr(34) & " (Specified yy from 1 to 52)" & Chr(13) & _
                    Chr(34) & "X" & Chr(34) & " (for excluded pile)" & Chr(13))
                Exit Sub
            End If
        Next i%
        Call frmPiles.CreatePilesCut(pSEP(1), pSEP(2), pSEP(3), pSEP(4), pSEP(5), pSEP(6), pSEP(7), pSEP(8), pSEP(9))
    Case "CutPiles"
        'check for existing piles
        If PilesShown = 0 Then
            MsgBox ("This event requires the presence of piles.")
            Exit Sub
        End If
        For i% = 1 To 4
            'set error condition to True
            SessionParseError = True
            'p1
            If i% = 1 Then
                'check for "R"
                If pSEP(i%) = "R" Then
                    SessionParseError = False
                End If
                'check for "Pxx" - where xx must be greater than 0
                If Len(pSEP(i%)) > 1 And Left(pSEP(i%), 1) = "P" Then
                    If IsNumeric(Right(pSEP(i%), Len(pSEP(i%)) - 1)) And _
                        Val(Right(pSEP(i%), Len(pSEP(i%)) - 1)) >= 1 Then
                        SessionParseError = False
                    End If
                End If
            'p2
            ElseIf i% = 2 Then
                'check for "R"
                If pSEP(i%) = "R" Then
                    SessionParseError = False
                End If
                'check for "C"
                If pSEP(i%) = "C" Then
                    SessionParseError = False
                End If
                'check for "Sxx" - where xx must be from within 1 and 52
                If Len(pSEP(i%)) > 1 And Left(pSEP(i%), 1) = "S" Then
                    If IsNumeric(Right(pSEP(i%), Len(pSEP(i%)) - 1)) And _
                        Val(Right(pSEP(i%), Len(pSEP(i%)) - 1)) >= 1 And _
                        Val(Right(pSEP(i%), Len(pSEP(i%)) - 1)) <= 52 Then
                        SessionParseError = False
                    End If
                End If
            'p3
            ElseIf i% = 3 Then
                'check for "P"
                If pSEP(i%) = "P" Then
                    SessionParseError = False
                End If
                'check for "E"
                If pSEP(i%) = "E" Then
                    SessionParseError = False
                End If
                'check for "M"
                If pSEP(i%) = "M" Then
                    SessionParseError = False
                End If
                'check for "R"
                If pSEP(i%) = "R" Then
                    SessionParseError = False
                End If
                'check for "D"
                If pSEP(i%) = "D" Then
                    SessionParseError = False
                End If
                'check for "L"
                If pSEP(i%) = "L" Then
                    SessionParseError = False
                End If
                'check for "Sxx" - where xx must be from within 1 and 52
                If Len(pSEP(i%)) > 1 And Left(pSEP(i%), 1) = "S" Then
                    If IsNumeric(Right(pSEP(i%), Len(pSEP(i%)) - 1)) And _
                        Val(Right(pSEP(i%), Len(pSEP(i%)) - 1)) >= 1 Then
                        SessionParseError = False
                    End If
                End If
                'check for "Nxx" - where xx must be greater than 0
                If Len(pSEP(i%)) > 1 And Left(pSEP(i%), 1) = "N" Then
                    If IsNumeric(Right(pSEP(i%), Len(pSEP(i%)) - 1)) And _
                        Val(Right(pSEP(i%), Len(pSEP(i%)) - 1)) >= 1 Then
                        SessionParseError = False
                    End If
                End If
            'p4
            ElseIf i% = 4 Then
                'check for "R"
                If pSEP(i%) = "R" Then
                    SessionParseError = False
                End If
                'check for "X"
                If pSEP(i%) = "X" Then
                    SessionParseError = False
                End If
            End If
            If SessionParseError = True Then
                If i% = 1 Then
                    pTextPlug = "first"
                ElseIf i% = 2 Then
                    pTextPlug = "second"
                ElseIf i% = 3 Then
                    pTextPlug = "third"
                ElseIf i% = 4 Then
                    pTextPlug = "fourth"
                End If
                If i% = 1 Then
                    MsgBox ("Cut Piles" & Chr(13) & _
                        "The " & pTextPlug & " Event parameter is invalid." & Chr(13) & Chr(13) & _
                        "Allowable parameter entries are:" & Chr(13) & _
                        Chr(34) & "R" & Chr(34) & " (for Random)" & Chr(13) & _
                        Chr(34) & "Pxx" & Chr(34) & " (Primary, xx > 0)" & Chr(13))
                    Exit Sub
                ElseIf i% = 2 Then
                    MsgBox ("Cut Piles" & Chr(13) & _
                        "The " & pTextPlug & " Event parameter is invalid." & Chr(13) & Chr(13) & _
                        "Allowable parameter entries are:" & Chr(13) & _
                        Chr(34) & "R" & Chr(34) & " (for Random)" & Chr(13) & _
                        Chr(34) & "C" & Chr(34) & " (for Complete)" & Chr(13) & _
                        Chr(34) & "Syy" & Chr(34) & " (Specified yy from 1 to 52)" & Chr(13))
                    Exit Sub
                ElseIf i% = 3 Then
                    MsgBox ("Cut Piles" & Chr(13) & _
                        "The " & pTextPlug & " Event parameter is invalid." & Chr(13) & Chr(13) & _
                        "Allowable parameter entries are:" & Chr(13) & _
                        Chr(34) & "P" & Chr(34) & " (for Primary)" & Chr(13) & _
                        Chr(34) & "E" & Chr(34) & " (for Equivalent Random)" & Chr(13) & _
                        Chr(34) & "M" & Chr(34) & " (for Top Same)" & Chr(13) & _
                        Chr(34) & "R" & Chr(34) & " (for Random)" & Chr(13) & _
                        Chr(34) & "D" & Chr(34) & " (for Top Random Not Same)" & Chr(13) & _
                        Chr(34) & "L" & Chr(34) & " (for Random Pile)" & Chr(13) & _
                        Chr(34) & "Nxx" & Chr(34) & " (xx > 0)" & Chr(13) & _
                        Chr(34) & "Syy" & Chr(34) & " (Secondary, yy > 0)" & Chr(13))
                    Exit Sub
                ElseIf i% = 4 Then
                    MsgBox ("Cut Piles" & Chr(13) & _
                        "The " & pTextPlug & " Event parameter is invalid." & Chr(13) & Chr(13) & _
                        "Allowable parameter entries are:" & Chr(13) & _
                        Chr(34) & "R" & Chr(34) & " (for Reverse)" & Chr(13) & _
                        Chr(34) & "X" & Chr(34) & " (for not Reversed)" & Chr(13))
                    Exit Sub
                End If
            End If
        Next i%
        Call frmPiles.CutPiles(pSEP(1), pSEP(2), pSEP(3), pSEP(4))
    Case "SelectReturn"
        'run special long case that saved space in this procedure
        Call SelectReturnCaseExtension
    Case "CombinePiles"
        'run special long case that saved space in this procedure
        Call CombinePilesCaseExtension
    Case "RiffleShufflePile"
        'run special long case that saved space in this procedure
        Call RiffleShufflePilesCaseExtension
    Case "AustralianDeal"
        'run special long case that saved space in this procedure
        Call AustralianDealCaseExtension
    Case "SwapPiles"
        'run special long case that saved space in this procedure
        Call SwapPilesCaseExtension
    Case "ElmsleyCount"
        'check for existing piles
        If PilesShown = 0 Then
            MsgBox ("This event requires the presence of piles.")
            Exit Sub
        End If
        'pPiles
        If IsNumeric(pSEP(1)) And _
            (Val(pSEP(1)) < 1 Or _
            Val(pSEP(1)) > NumPiles) Then
                MsgBox ("Elmsley Count" & Chr(13) & _
                "Event parameter is out of range." & Chr(13) & _
                "It must be between 1 and the number of piles.")
                SessionParseError = True
                Exit Sub
        ElseIf Not IsNumeric(pSEP(1)) Then
            MsgBox ("Elmsley Count" & Chr(13) & _
            "Event parameter is not valid." & Chr(13) & _
            "It must be a number between 1 and the number of piles.")
            SessionParseError = True
            Exit Sub
        End If
        Call frmPiles.ElmsleyCount(pSEP(1))
    Case "InverseElmsleyCount"
        'check for existing piles
        If PilesShown = 0 Then
            MsgBox ("This event requires the presence of piles.")
            Exit Sub
        End If
        'pPiles
        If IsNumeric(pSEP(1)) And _
            (Val(pSEP(1)) < 1 Or _
            Val(pSEP(1)) > NumPiles) Then
                MsgBox ("Inverse Elmsley Count" & Chr(13) & _
                "Event parameter is out of range." & Chr(13) & _
                "It must be between 1 and the number of piles.")
                SessionParseError = True
                Exit Sub
        ElseIf Not IsNumeric(pSEP(1)) Then
            MsgBox ("Inverse Elmsley Count" & Chr(13) & _
            "Event parameter is not valid." & Chr(13) & _
            "It must be a number between 1 and the number of piles.")
            SessionParseError = True
            Exit Sub
        End If
        Call frmPiles.InverseElmsleyCount(pSEP(1))
    Case "JordanCount"
        'check for existing piles
        If PilesShown = 0 Then
            MsgBox ("This event requires the presence of piles.")
            Exit Sub
        End If
        'pPiles
        If IsNumeric(pSEP(1)) And _
            (Val(pSEP(1)) < 1 Or _
            Val(pSEP(1)) > NumPiles) Then
                MsgBox ("Jordan Count" & Chr(13) & _
                "Event parameter is out of range." & Chr(13) & _
                "It must be between 1 and the number of piles.")
                SessionParseError = True
                Exit Sub
        ElseIf Not IsNumeric(pSEP(1)) Then
            MsgBox ("Jordan Count" & Chr(13) & _
            "Event parameter is not valid." & Chr(13) & _
            "It must be a number between 1 and the number of piles.")
            SessionParseError = True
            Exit Sub
        End If
        Call frmPiles.JordanCount(pSEP(1))
    Case "InverseJordanCount"
        'check for existing piles
        If PilesShown = 0 Then
            MsgBox ("This event requires the presence of piles.")
            Exit Sub
        End If
        'pPiles
        If IsNumeric(pSEP(1)) And _
            (Val(pSEP(1)) < 1 Or _
            Val(pSEP(1)) > NumPiles) Then
                MsgBox ("Inverse Jordan Count" & Chr(13) & _
                "Event parameter is out of range." & Chr(13) & _
                "It must be between 1 and the number of piles.")
                SessionParseError = True
                Exit Sub
        ElseIf Not IsNumeric(pSEP(1)) Then
            MsgBox ("Inverse Jordan Count" & Chr(13) & _
            "Event parameter is not valid." & Chr(13) & _
            "It must be a number between 1 and the number of piles.")
            SessionParseError = True
            Exit Sub
        End If
        Call frmPiles.InverseJordanCount(pSEP(1))
    Case "TurnOver"
        'check for existing piles
        If PilesShown = 0 Then
            MsgBox ("This event requires the presence of piles.")
            Exit Sub
        End If
        'pPiles
        If IsNumeric(pSEP(1)) And _
            (Val(pSEP(1)) < 1 Or _
            Val(pSEP(1)) > NumPiles) Then
                MsgBox ("Turn Over" & Chr(13) & _
                "Event parameter is out of range." & Chr(13) & _
                "It must be between 1 and the number of piles.")
                SessionParseError = True
                Exit Sub
        ElseIf Not IsNumeric(pSEP(1)) Then
            MsgBox ("Turn Over" & Chr(13) & _
            "Event parameter is not valid." & Chr(13) & _
            "It must be a number between 1 and the number of piles.")
            SessionParseError = True
            Exit Sub
        End If
        Call frmPiles.TurnOver(pSEP(1))
    Case "SwapCardsRandom"
        'check for valid number 0 - 1 for "No Selection" status
        If IsNumeric(pSEP(1)) Then
            If Not (Val(pSEP(1)) = 0 Or Val(pSEP(1)) = 1) Then
                MsgBox ("SwapCardsRandom" & Chr(13) & _
                "Event parameter is not valid." & Chr(13) & _
                "It can only be a 0 or 1.")
                SessionParseError = True
                Exit Sub
            End If
        Else
            MsgBox ("SwapCardsRandom" & Chr(13) & _
            "Event parameter is not valid." & Chr(13) & _
            "It can only be a 0 or 1.")
            SessionParseError = True
            Exit Sub
        End If
        Call SwapCardsRandom(pSEP(1))
    Case "SwapDifferentColors"
        'check for valid number 0 - 1 for "No Selection" status
        If IsNumeric(pSEP(1)) Then
            If Not (Val(pSEP(1)) = 0 Or Val(pSEP(1)) = 1) Then
                MsgBox ("SwapDifferentColors" & Chr(13) & _
                "Event parameter is not valid." & Chr(13) & _
                "It can only be a 0 or 1.")
                SessionParseError = True
                Exit Sub
            End If
        Else
            MsgBox ("SwapDifferentColors" & Chr(13) & _
            "Event parameter is not valid." & Chr(13) & _
            "It can only be a 0 or 1.")
            SessionParseError = True
            Exit Sub
        End If
        Call SwapDifferentColors(pSEP(1))
    Case "SwapSpecifiedCards"
        'check for valid number 0 - 1 for "No Selection" status
        If IsNumeric(pSEP(3)) Then
            If Not (Val(pSEP(3)) = 0 Or Val(pSEP(3)) = 1) Then
                MsgBox ("SwapSpecifiedCards" & Chr(13) & _
                "Third Event parameter is not valid." & Chr(13) & _
                "It can only be a 0 or 1.")
                SessionParseError = True
                Exit Sub
            End If
        Else
            MsgBox ("SwapSpecifiedCards" & Chr(13) & _
            "Third Event parameter is not valid." & Chr(13) & _
            "It can only be a 0 or 1.")
            SessionParseError = True
            Exit Sub
        End If
        If IsNumeric(pSEP(1)) And _
            (Val(pSEP(1)) < 1 Or _
            Val(pSEP(1)) > 52) Then
                MsgBox ("SwapSpecifiedCards" & Chr(13) & _
                "First event parameter is out of range." & Chr(13) & _
                "It must be between 1 and 52.")
                SessionParseError = True
                Exit Sub
        ElseIf Not IsNumeric(pSEP(1)) Then
            MsgBox ("SwapSpecifiedCards" & Chr(13) & _
            "First event parameter is not valid." & Chr(13) & _
            "It must be a number between 1 and 52.")
            SessionParseError = True
            Exit Sub
        End If
        If IsNumeric(pSEP(2)) And _
            (Val(pSEP(2)) < 1 Or _
            Val(pSEP(2)) > 52) Then
                MsgBox ("SwapSpecifiedCards" & Chr(13) & _
                "Second event parameter is out of range." & Chr(13) & _
                "It must be between 1 and 52.")
                SessionParseError = True
                Exit Sub
        ElseIf Not IsNumeric(pSEP(2)) Then
            MsgBox ("SwapSpecifiedCards" & Chr(13) & _
            "Second event parameter is not valid." & Chr(13) & _
            "It must be a number between 1 and 52.")
            SessionParseError = True
            Exit Sub
        End If
        If pSEP(1) = pSEP(2) Then
            MsgBox ("SwapSpecifiedCards" & Chr(13) & _
            "Both event parameters can not be the same value.")
            SessionParseError = True
            Exit Sub
        End If
        Call SwapSpecifiedCards(pSEP(1), pSEP(2), pSEP(3))
    Case "SwapSpecifiedPositions"
        'check for valid number 0 - 1 for "No Selection" status
        If IsNumeric(pSEP(3)) Then
            If Not (Val(pSEP(3)) = 0 Or Val(pSEP(3)) = 1) Then
                MsgBox ("SwapSpecifiedPositions" & Chr(13) & _
                "Third Event parameter is not valid." & Chr(13) & _
                "It can only be a 0 or 1.")
                SessionParseError = True
                Exit Sub
            End If
        Else
            MsgBox ("SwapSpecifiedPositions" & Chr(13) & _
            "Third Event parameter is not valid." & Chr(13) & _
            "It can only be a 0 or 1.")
            SessionParseError = True
            Exit Sub
        End If
        If IsNumeric(pSEP(1)) And _
            (Val(pSEP(1)) < 1 Or _
            Val(pSEP(1)) > 52) Then
                MsgBox ("SwapSpecifiedPositions" & Chr(13) & _
                "First event parameter is out of range." & Chr(13) & _
                "It must be between 1 and 52.")
                SessionParseError = True
                Exit Sub
        ElseIf Not IsNumeric(pSEP(1)) Then
            MsgBox ("SwapSpecifiedPositions" & Chr(13) & _
            "First event parameter is not valid." & Chr(13) & _
            "It must be a number between 1 and 52.")
            SessionParseError = True
            Exit Sub
        End If
        If IsNumeric(pSEP(2)) And _
            (Val(pSEP(2)) < 1 Or _
            Val(pSEP(2)) > 52) Then
                MsgBox ("SwapSpecifiedPositions" & Chr(13) & _
                "Second event parameter is out of range." & Chr(13) & _
                "It must be between 1 and 52.")
                SessionParseError = True
                Exit Sub
        ElseIf Not IsNumeric(pSEP(2)) Then
            MsgBox ("SwapSpecifiedPositions" & Chr(13) & _
            "Second event parameter is not valid." & Chr(13) & _
            "It must be a number between 1 and 52.")
            SessionParseError = True
            Exit Sub
        End If
        If pSEP(1) = pSEP(2) Then
            MsgBox ("SwapSpecifiedPositions" & Chr(13) & _
            "Both event parameters can not be the same value.")
            SessionParseError = True
            Exit Sub
        End If
        Call SwapSpecifiedPositions(pSEP(1), pSEP(2), pSEP(3))
    Case "SwapSameColor"
        'check for valid number 0 - 1 for "No Selection" status
        If IsNumeric(pSEP(2)) Then
            If Not (Val(pSEP(2)) = 0 Or Val(pSEP(2)) = 1) Then
                MsgBox ("SwapSameColor" & Chr(13) & _
                "Second Event parameter is not valid." & Chr(13) & _
                "It can only be a 0 or 1.")
                SessionParseError = True
                Exit Sub
            End If
        Else
            MsgBox ("SwapSameColor" & Chr(13) & _
            "Second Event parameter is not valid." & Chr(13) & _
            "It can only be a 0 or 1.")
            SessionParseError = True
            Exit Sub
        End If
        If Not pSEP(1) = "X" And _
            Not pSEP(1) = "R" And _
            Not pSEP(1) = "B" Then
                MsgBox ("SwapSameColor" & Chr(13) & _
                    "Event parameter is invalid." & Chr(13) & _
                    "Allowable parameter entries are:" & Chr(13) & _
                    Chr(34) & "X" & Chr(34) & " (Random color)" & Chr(13) & _
                    Chr(34) & "R" & Chr(34) & " (Red)" & Chr(13) & _
                    Chr(34) & "B" & Chr(34) & " (Black)" & Chr(13))
            SessionParseError = True
            Exit Sub
        End If
        Call SwapSameColor(pSEP(1), pSEP(2))
    Case "SwapSameSuit"
        'check for valid number 0 - 1 for "No Selection" status
        If IsNumeric(pSEP(2)) Then
            If Not (Val(pSEP(2)) = 0 Or Val(pSEP(2)) = 1) Then
                MsgBox ("SwapSameSuit" & Chr(13) & _
                "Second Event parameter is not valid." & Chr(13) & _
                "It can only be a 0 or 1.")
                SessionParseError = True
                Exit Sub
            End If
        Else
            MsgBox ("SwapSameSuit" & Chr(13) & _
            "Second Event parameter is not valid." & Chr(13) & _
            "It can only be a 0 or 1.")
            SessionParseError = True
            Exit Sub
        End If
        If Not pSEP(1) = "X" And _
            Not pSEP(1) = "C" And _
            Not pSEP(1) = "H" And _
            Not pSEP(1) = "S" And _
            Not pSEP(1) = "D" Then
                MsgBox ("SwapSameSuit" & Chr(13) & _
                    "Event parameter is invalid." & Chr(13) & _
                    "Allowable parameter entries are:" & Chr(13) & _
                    Chr(34) & "X" & Chr(34) & " (Random suit)" & Chr(13) & _
                    Chr(34) & "C" & Chr(34) & " (Clubs)" & Chr(13) & _
                    Chr(34) & "H" & Chr(34) & " (Hearts)" & Chr(13) & _
                    Chr(34) & "S" & Chr(34) & " (Spades)" & Chr(13) & _
                    Chr(34) & "D" & Chr(34) & " (Diamonds)" & Chr(13))
            SessionParseError = True
            Exit Sub
        End If
        Call SwapSameSuit(pSEP(1), pSEP(2))
    Case "SwapDifferentSuits"
        'check for valid number 0 - 1 for "No Selection" status
        If IsNumeric(pSEP(3)) Then
            If Not (Val(pSEP(3)) = 0 Or Val(pSEP(3)) = 1) Then
                MsgBox ("SwapDifferentSuits" & Chr(13) & _
                "Third Event parameter is not valid." & Chr(13) & _
                "It can only be a 0 or 1.")
                SessionParseError = True
                Exit Sub
            End If
        Else
            MsgBox ("SwapDifferentSuits" & Chr(13) & _
            "Third Event parameter is not valid." & Chr(13) & _
            "It can only be a 0 or 1.")
            SessionParseError = True
            Exit Sub
        End If
        If Not pSEP(1) = "X" And _
            Not pSEP(1) = "C" And _
            Not pSEP(1) = "H" And _
            Not pSEP(1) = "S" And _
            Not pSEP(1) = "D" Then
                MsgBox ("SwapDifferentSuits" & Chr(13) & _
                    "First event parameter is invalid." & Chr(13) & _
                    "Allowable parameter entries are:" & Chr(13) & _
                    Chr(34) & "X" & Chr(34) & " (Random suit)" & Chr(13) & _
                    Chr(34) & "C" & Chr(34) & " (Clubs)" & Chr(13) & _
                    Chr(34) & "H" & Chr(34) & " (Hearts)" & Chr(13) & _
                    Chr(34) & "S" & Chr(34) & " (Spades)" & Chr(13) & _
                    Chr(34) & "D" & Chr(34) & " (Diamonds)" & Chr(13))
            SessionParseError = True
            Exit Sub
        End If
        If Not pSEP(2) = "X" And _
            Not pSEP(2) = "C" And _
            Not pSEP(2) = "H" And _
            Not pSEP(2) = "S" And _
            Not pSEP(2) = "D" Then
                MsgBox ("SwapDifferentSuits" & Chr(13) & _
                    "Second event parameter is invalid." & Chr(13) & _
                    "Allowable parameter entries are:" & Chr(13) & _
                    Chr(34) & "X" & Chr(34) & " (Random suit)" & Chr(13) & _
                    Chr(34) & "C" & Chr(34) & " (Clubs)" & Chr(13) & _
                    Chr(34) & "H" & Chr(34) & " (Hearts)" & Chr(13) & _
                    Chr(34) & "S" & Chr(34) & " (Spades)" & Chr(13) & _
                    Chr(34) & "D" & Chr(34) & " (Diamonds)" & Chr(13))
            SessionParseError = True
            Exit Sub
        End If
        If pSEP(1) = pSEP(2) And pSEP(1) <> "X" Then
            MsgBox ("SwapDifferentSuits" & Chr(13) & _
            "Both event parameters can not be the same." & Chr(13) & _
            "(Unless they are both 'X'.)")
            SessionParseError = True
            Exit Sub
        End If
        If pSEP(1) <> "X" And pSEP(2) = "X" Then
            MsgBox ("SwapDifferentSuits" & Chr(13) & _
            "The second event parameter can only be an 'X'" & Chr(13) & _
            "if the first event parameter is also an 'X'")
            SessionParseError = True
            Exit Sub
        End If
        Call SwapDifferentSuits(pSEP(1), pSEP(2), pSEP(3))
    Case "PokerDiscard"
        If PokerCardsDealt = 1 Then
            Dim pSafeMatch As Boolean
            pSafeMatch = False
            If IsNumeric(pSEP(1)) Then
                For i% = 1 To 50
                    If i% = Val(pSEP(1)) Then
                        pSafeMatch = True
                    End If
                Next i%
                If Not pSafeMatch Then
                    MsgBox ("PokerDiscard Event Error:" & Chr(13) & _
                    "The single parameter can only be an integer from 1 to 50," & Chr(13) & _
                    "or a card value such as AC, 2C, 3C, etc.")
                    Exit Sub
                End If
            Else
                For i% = 1 To 52
                    If Deck(2, i%) = pSEP(1) Then
                        pSafeMatch = True
                    End If
                Next i%
                If Not pSafeMatch Then
                    MsgBox ("Poker Discard Event Error:" & Chr(13) & _
                    "The single parameter can only be an integer from 1 to 50," & Chr(13) & _
                    "or a card value such as AC, 2C, 3C, etc.")
                    Exit Sub
                End If
            End If
            Call frmDeck.PokerDiscard(pSEP(1))
        Else
            MsgBox ("Poker Discard Event Error:" & Chr(13) & _
                    "This event may only be called when there are" & Chr(13) & _
                    "active poker hands dealt.")
            Exit Sub
        End If
    Case "Macro"
        Call SessionFilePlay(pSEP(1))
    Case Else
        MsgBox ("frmStackView: This is an unknown Event entry." & _
        Chr(13) & myCall & " error")
        SessionParseError = True
End Select
End Sub

Private Sub SwapPilesCaseExtension()
'check for existing piles
If PilesShown = 0 Then
    MsgBox ("This event requires the presence of piles.")
    Exit Sub
End If
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
pStringPointer = 0
'decode first (left) parameter
'first section checks if there are three characters which can only be "PxR"
If Len(pSEP(1)) = 3 Then
    If Left(pSEP(1), 1) = "P" Then
        If IsNumeric(Mid(pSEP(1), 2, 1)) Then
            pLeftPile = Val(Mid(pSEP(1), 2, 1))
            If pLeftPile > NumPiles Then
                MsgBox ("Error: Swap Piles" & Chr(13) & _
                    "The first parameter (" & Chr(34) & pSEP(1) & Chr(34) & _
                    ") is referencing a pile number" & Chr(13) & _
                    "that is greater than the current number of piles.")
                Exit Sub
            End If
            If pLeftPile < 1 Then
                MsgBox ("Error: Swap Piles" & Chr(13) & _
                    "The first parameter (" & Chr(34) & pSEP(1) & Chr(34) & _
                    ") is referencing a pile number" & Chr(13) & _
                    "that is less than 1.  You must reference an actual pile number.")
                Exit Sub
            End If
        Else
            MsgBox ("Error: Swap Piles" & Chr(13) & _
                "The first parameter (" & Chr(34) & pSEP(1) & Chr(34) & _
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
            "The first parameter (" & Chr(34) & pSEP(1) & Chr(34) & _
            ") can only be:" & Chr(13) & _
            Chr(34) & "Px" & Chr(34) & " (specified Pile with number)" & Chr(13) & _
            Chr(34) & "R" & Chr(34) & " (Random pile)" & Chr(13) & _
            Chr(34) & "S" & Chr(34) & " (Pile with a selected card)" & Chr(13) & _
            Chr(34) & "N" & Chr(34) & " (pile with No selected card)" & Chr(13) & _
            Chr(13) & "The above parameters may have an " & Chr(34) & "R" & Chr(34) & Chr(13) & _
            "suffix which indicates the pile should be Reversed.")
        Exit Sub
    End If
    If Right(pSEP(1), 1) = "R" Then
        pLeftReverse = True
    Else
        MsgBox ("Error: Swap Piles" & Chr(13) & _
            "The first parameter (" & Chr(34) & pSEP(1) & Chr(34) & _
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
ElseIf Len(pSEP(1)) = 2 Then
    'if the first character is a "P" then the second must be a pile number
    If Left(pSEP(1), 1) = "P" Then
        If IsNumeric(Right(pSEP(1), 1)) Then
            pLeftPile = Val(Right(pSEP(1), 1))
            If pLeftPile > NumPiles Then
                MsgBox ("Error: Swap Piles" & Chr(13) & _
                    "The first parameter (" & Chr(34) & pSEP(1) & Chr(34) & _
                    ") is referencing a pile number" & Chr(13) & _
                    "that is greater than the current number of piles.")
                Exit Sub
            End If
            If pLeftPile < 1 Then
                MsgBox ("Error: Swap Piles" & Chr(13) & _
                    "The first parameter (" & Chr(34) & pSEP(1) & Chr(34) & _
                    ") is referencing a pile number" & Chr(13) & _
                    "that is less than 1.  You must reference an actual pile number.")
                Exit Sub
            End If
        Else
            MsgBox ("Error: Swap Piles" & Chr(13) & _
                "The first parameter (" & Chr(34) & pSEP(1) & Chr(34) & _
                ") can only be:" & Chr(13) & _
                Chr(34) & "Px" & Chr(34) & " (specified Pile with number)" & Chr(13) & _
                Chr(34) & "R" & Chr(34) & " (Random pile)" & Chr(13) & _
                Chr(34) & "S" & Chr(34) & " (Pile with a selected card)" & Chr(13) & _
                Chr(34) & "N" & Chr(34) & " (pile with No selected card)" & Chr(13) & _
                Chr(13) & "The above parameters may have an " & Chr(34) & "R" & Chr(34) & Chr(13) & _
                "suffix which indicates the pile should be Reversed.")
            Exit Sub
        End If
    ElseIf Left(pSEP(1), 1) = "R" Then
        'set the first pile to a random number
        pLeftPile = Int(Rnd * NumPiles + 1)
        'check for a valid second parameter (can only be an "R")
        If Right(pSEP(1), 1) = "R" Then
            pLeftReverse = True
        Else
            MsgBox ("Error: Swap Piles" & Chr(13) & _
                "The first parameter (" & Chr(34) & pSEP(1) & Chr(34) & _
                ") can only be:" & Chr(13) & _
                Chr(34) & "Px" & Chr(34) & " (specified Pile with number)" & Chr(13) & _
                Chr(34) & "R" & Chr(34) & " (Random pile)" & Chr(13) & _
                Chr(34) & "S" & Chr(34) & " (Pile with a selected card)" & Chr(13) & _
                Chr(34) & "N" & Chr(34) & " (pile with No selected card)" & Chr(13) & _
                Chr(13) & "The above parameters may have an " & Chr(34) & "R" & Chr(34) & Chr(13) & _
                "suffix which indicates the pile should be Reversed.")
            Exit Sub
        End If
    ElseIf Left(pSEP(1), 1) = "S" Then
            If Right(pSEP(1), 1) = "R" Then
                pLeftReverse = True
            Else
                MsgBox ("Error: Swap Piles" & Chr(13) & _
                    "The first parameter (" & Chr(34) & pSEP(1) & Chr(34) & _
                    ") can only be:" & Chr(13) & _
                    Chr(34) & "Px" & Chr(34) & " (specified Pile with number)" & Chr(13) & _
                    Chr(34) & "R" & Chr(34) & " (Random pile)" & Chr(13) & _
                    Chr(34) & "S" & Chr(34) & " (Pile with a selected card)" & Chr(13) & _
                    Chr(34) & "N" & Chr(34) & " (pile with No selected card)" & Chr(13) & _
                    Chr(13) & "The above parameters may have an " & Chr(34) & "R" & Chr(34) & Chr(13) & _
                    "suffix which indicates the pile should be Reversed.")
                Exit Sub
            End If
    ElseIf Left(pSEP(1), 1) = "N" Then
        'check for a valid second parameter
        If Right(pSEP(1), 1) = "R" Then
            pLeftReverse = True
        Else
            MsgBox ("Error: Swap Piles" & Chr(13) & _
                "The first parameter (" & Chr(34) & pSEP(1) & Chr(34) & _
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
            "The first parameter (" & Chr(34) & pSEP(1) & Chr(34) & _
            ") can only be:" & Chr(13) & _
            Chr(34) & "Px" & Chr(34) & " (specified Pile with number)" & Chr(13) & _
            Chr(34) & "R" & Chr(34) & " (Random pile)" & Chr(13) & _
            Chr(34) & "S" & Chr(34) & " (Pile with a selected card)" & Chr(13) & _
            Chr(34) & "N" & Chr(34) & " (pile with No selected card)" & Chr(13) & _
            Chr(13) & "The above parameters may have an " & Chr(34) & "R" & Chr(34) & Chr(13) & _
            "suffix which indicates the pile should be Reversed.")
        Exit Sub
    End If
ElseIf Len(pSEP(1)) = 1 Then
    If pSEP(1) = "R" Then
        'set the first pile to a random number
        pLeftPile = Int(Rnd * NumPiles + 1)
    ElseIf pSEP(1) = "S" Then
        pLeftContainsSelected = True
    ElseIf pSEP(1) = "N" Then
        'nothing to do
    Else
        MsgBox ("Error: Swap Piles" & Chr(13) & _
            "The first parameter (" & Chr(34) & pSEP(1) & Chr(34) & _
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
'need to identify the errors for the first parameter wherever pSwapPilesError=True
'initialize appropriate variables
'decode second (right) parameter
'first section checks if there are three characters which can only be "PxR"
If Len(pSEP(2)) = 3 Then
    If Left(pSEP(2), 1) = "P" Then
        If IsNumeric(Mid(pSEP(2), 2, 1)) Then
            pRightPile = Val(Mid(pSEP(2), 2, 1))
            If pRightPile > NumPiles Then
                MsgBox ("Error: Swap Piles" & Chr(13) & _
                    "The second parameter (" & Chr(34) & pSEP(2) & Chr(34) & _
                    ") is referencing a pile number" & Chr(13) & _
                    "that is greater than the current number of piles.")
                Exit Sub
            End If
            If pRightPile < 1 Then
                MsgBox ("Error: Swap Piles" & Chr(13) & _
                    "The first parameter (" & Chr(34) & pSEP(2) & Chr(34) & _
                    ") is referencing a pile number" & Chr(13) & _
                    "that is less than 1.  You must reference an actual pile number.")
                Exit Sub
            End If
        Else
            MsgBox ("Error: Swap Piles" & Chr(13) & _
                "The second parameter (" & Chr(34) & pSEP(2) & Chr(34) & _
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
            "The second parameter (" & Chr(34) & pSEP(2) & Chr(34) & _
            ") can only be:" & Chr(13) & _
            Chr(34) & "Px" & Chr(34) & " (specified Pile with number)" & Chr(13) & _
            Chr(34) & "R" & Chr(34) & " (Random pile)" & Chr(13) & _
            Chr(34) & "S" & Chr(34) & " (Pile with a selected card)" & Chr(13) & _
            Chr(34) & "N" & Chr(34) & " (pile with No selected card)" & Chr(13) & _
            Chr(13) & "The above parameters may have an " & Chr(34) & "R" & Chr(34) & Chr(13) & _
            "suffix which indicates the pile should be Reversed.")
        Exit Sub
    End If
    If Right(pSEP(2), 1) = "R" Then
        pRightReverse = True
    Else
        MsgBox ("Error: Swap Piles" & Chr(13) & _
            "The second parameter (" & Chr(34) & pSEP(2) & Chr(34) & _
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
ElseIf Len(pSEP(2)) = 2 Then
    'if the first character is a "P" then the second must be a pile number
    If Left(pSEP(2), 1) = "P" Then
        If IsNumeric(Right(pSEP(2), 1)) Then
            pRightPile = Val(Right(pSEP(2), 1))
            If pRightPile > NumPiles Then
                MsgBox ("Error: Swap Piles" & Chr(13) & _
                    "The second parameter (" & Chr(34) & pSEP(2) & Chr(34) & _
                    ") is referencing a pile number" & Chr(13) & _
                    "that is greater than the current number of piles.")
                Exit Sub
            End If
            If pRightPile < 1 Then
                MsgBox ("Error: Swap Piles" & Chr(13) & _
                    "The first parameter (" & Chr(34) & pSEP(2) & Chr(34) & _
                    ") is referencing a pile number" & Chr(13) & _
                    "that is less than 1.  You must reference an actual pile number.")
                Exit Sub
            End If
        Else
            MsgBox ("Error: Swap Piles" & Chr(13) & _
                "The second parameter (" & Chr(34) & pSEP(2) & Chr(34) & _
                ") can only be:" & Chr(13) & _
                Chr(34) & "Px" & Chr(34) & " (specified Pile with number)" & Chr(13) & _
                Chr(34) & "R" & Chr(34) & " (Random pile)" & Chr(13) & _
                Chr(34) & "S" & Chr(34) & " (Pile with a selected card)" & Chr(13) & _
                Chr(34) & "N" & Chr(34) & " (pile with No selected card)" & Chr(13) & _
                Chr(13) & "The above parameters may have an " & Chr(34) & "R" & Chr(34) & Chr(13) & _
                "suffix which indicates the pile should be Reversed.")
            Exit Sub
        End If
    ElseIf Left(pSEP(2), 1) = "R" Then
        'check for a valid second parameter (can only be an "R")
        If Right(pSEP(2), 1) = "R" Then
            pRightReverse = True
        Else
            MsgBox ("Error: Swap Piles" & Chr(13) & _
                "The second parameter (" & Chr(34) & pSEP(2) & Chr(34) & _
                ") can only be:" & Chr(13) & _
                Chr(34) & "Px" & Chr(34) & " (specified Pile with number)" & Chr(13) & _
                Chr(34) & "R" & Chr(34) & " (Random pile)" & Chr(13) & _
                Chr(34) & "S" & Chr(34) & " (Pile with a selected card)" & Chr(13) & _
                Chr(34) & "N" & Chr(34) & " (pile with No selected card)" & Chr(13) & _
                Chr(13) & "The above parameters may have an " & Chr(34) & "R" & Chr(34) & Chr(13) & _
                "suffix which indicates the pile should be Reversed.")
            Exit Sub
        End If
    ElseIf Left(pSEP(2), 1) = "S" Then
        'check for a valid second parameter
        If Right(pSEP(2), 1) = "R" Then
            pRightReverse = True
        Else
            MsgBox ("Error: Swap Piles" & Chr(13) & _
                "The second parameter (" & Chr(34) & pSEP(2) & Chr(34) & _
                ") can only be:" & Chr(13) & _
                Chr(34) & "Px" & Chr(34) & " (specified Pile with number)" & Chr(13) & _
                Chr(34) & "R" & Chr(34) & " (Random pile)" & Chr(13) & _
                Chr(34) & "S" & Chr(34) & " (Pile with a selected card)" & Chr(13) & _
                Chr(34) & "N" & Chr(34) & " (pile with No selected card)" & Chr(13) & _
                Chr(13) & "The above parameters may have an " & Chr(34) & "R" & Chr(34) & Chr(13) & _
                "suffix which indicates the pile should be Reversed.")
            Exit Sub
        End If
    ElseIf Left(pSEP(2), 1) = "N" Then
        'check for a valid second parameter
        If Right(pSEP(2), 1) = "R" Then
            pRightReverse = True
        Else
            MsgBox ("Error: Swap Piles" & Chr(13) & _
                "The second parameter (" & Chr(34) & pSEP(2) & Chr(34) & _
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
            "The second parameter (" & Chr(34) & pSEP(2) & Chr(34) & _
            ") can only be:" & Chr(13) & _
            Chr(34) & "Px" & Chr(34) & " (specified Pile with number)" & Chr(13) & _
            Chr(34) & "R" & Chr(34) & " (Random pile)" & Chr(13) & _
            Chr(34) & "S" & Chr(34) & " (Pile with a selected card)" & Chr(13) & _
            Chr(34) & "N" & Chr(34) & " (pile with No selected card)" & Chr(13) & _
            Chr(13) & "The above parameters may have an " & Chr(34) & "R" & Chr(34) & Chr(13) & _
            "suffix which indicates the pile should be Reversed.")
        Exit Sub
    End If
ElseIf Len(pSEP(2)) = 1 Then
    If pSEP(2) = "R" Then
        'set the second pile to a random number
    ElseIf pSEP(2) = "S" Then
        'do nothing
    ElseIf pSEP(2) = "N" Then
        'do nothing
    Else
        MsgBox ("Error: Swap Piles" & Chr(13) & _
            "The second parameter (" & Chr(34) & pSEP(2) & Chr(34) & _
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
    Call frmPiles.SwapPiles(pSEP(1), pSEP(2))
End Sub


Public Sub AustralianDealCaseExtension()
        'check for existing piles
        If PilesShown = 0 Then
            MsgBox ("This event requires the presence of piles.")
            Exit Sub
        End If
        Dim specialError(8) As Boolean
        'initially set all specialErrors to True
        For m% = 1 To 8
            specialError(m%) = True
        Next m%
        For i% = 1 To 11
            'set error condition to True
            SessionParseError = True
            'p1
            If i% = 1 Then
                'check for "Ax" - where x must be greater than 0 and less than NumPiles
                'and A can be equal to "D" or "U"
                If Len(pSEP(i%)) = 2 And (Left(pSEP(i%), 1) = "D" Or _
                    Left(pSEP(i%), 1) = "U") Then
                    If IsNumeric(Right(pSEP(i%), 1)) And _
                        Val(Right(pSEP(i%), 1)) > 0 And _
                        Val(Right(pSEP(i%), 1)) <= NumPiles Then
                        SessionParseError = False
                    End If
                End If
            'p2
            ElseIf i% = 2 Then
                'check for valid number 0 - 52
                If IsNumeric(pSEP(i%)) Then
                    If Val(pSEP(i%)) >= 0 And Val(pSEP(i%)) <= 52 Then
                        SessionParseError = False
                    End If
                End If
            'p3
            ElseIf i% = 3 Then
                'check for valid number 0 - 52
                If IsNumeric(pSEP(i%)) Then
                    If Val(pSEP(i%)) >= 0 And Val(pSEP(i%)) <= 52 Then
                        SessionParseError = False
                    End If
                End If
            'p4
            ElseIf i% = 4 Then
                'check for valid number 0 - 52
                If IsNumeric(pSEP(i%)) Then
                    If Val(pSEP(i%)) >= 0 And Val(pSEP(i%)) <= 52 Then
                        SessionParseError = False
                    End If
                End If
            'p5
            ElseIf i% = 5 Then
                'check for valid number 0 - 52
                If IsNumeric(pSEP(i%)) Then
                    If Val(pSEP(i%)) >= 0 And Val(pSEP(i%)) <= 52 Then
                        SessionParseError = False
                    End If
                End If
            'p6
            ElseIf i% = 6 Then
                'check for specialErrors 1 to 4
                If pSEP(2) = 0 Or pSEP(4) = 0 Then
                    specialError(1) = False
                End If
                If pSEP(3) = 0 Or pSEP(5) = 0 Then
                    specialError(2) = False
                End If
                If pSEP(2) > 0 Or pSEP(4) > 0 Then
                    specialError(3) = False
                End If
                If pSEP(3) > 0 Or pSEP(5) > 0 Then
                    specialError(4) = False
                End If
                'check for "S"
                If pSEP(i%) = "S" Then
                    SessionParseError = False
                'check for valid number 0 - 1
                ElseIf IsNumeric(pSEP(i%)) Then
                    If Val(pSEP(i%)) = 0 Or Val(pSEP(i%)) = 1 Then
                        SessionParseError = False
                    End If
                End If
            'p7
            ElseIf i% = 7 Then
                'check for valid number 0 - 1
                If IsNumeric(pSEP(i%)) Then
                    If Val(pSEP(i%)) = 0 Or Val(pSEP(i%)) = 1 Then
                        SessionParseError = False
                    End If
                End If
            'p8
            ElseIf i% = 8 Then
                'check for valid number 0 - 1
                If IsNumeric(pSEP(i%)) Then
                    If Val(pSEP(i%)) = 0 Or Val(pSEP(i%)) = 1 Then
                        SessionParseError = False
                    End If
                End If
            'p9
            ElseIf i% = 9 Then
                'check for valid number 0 - 1
                If IsNumeric(pSEP(i%)) Then
                    If Val(pSEP(i%)) = 0 Or Val(pSEP(i%)) = 1 Then
                        SessionParseError = False
                    End If
                End If
            'p10
            ElseIf i% = 10 Then
                'check for specialErrors 5 to 6
                If pSEP(6) = "S" And Val(pSEP(8)) = 0 Then
                    specialError(5) = False
                ElseIf Val(pSEP(6)) = 0 Or Val(pSEP(8)) = 0 Then
                    specialError(5) = False
                End If
                If Val(pSEP(7)) = 0 Or Val(pSEP(9)) = 0 Then
                    specialError(6) = False
                End If
                'check for valid entries (0, "F", "R")
                If pSEP(i%) = 0 Or pSEP(i%) = "F" Or pSEP(i%) = "R" Then
                    SessionParseError = False
                End If
                'check for consistency between Reverse Selected card (P6)
                '  and Selected Card (P10)
                If pSEP(6) = "S" And pSEP(10) <> 0 Then
                        specialError(8) = False
                End If
            'p11
            ElseIf i% = 11 Then
                'check for valid number 0 - 1
                If IsNumeric(pSEP(i%)) Then
                    If Val(pSEP(i%)) = 0 Or Val(pSEP(i%)) = 1 Then
                        SessionParseError = False
                    End If
                End If
                'check for specialError 7
                If Val(pSEP(i%)) = 0 Or _
                    (Val(pSEP(i%)) = 1 And _
                    Val(pSEP(4)) = 0 And _
                    Val(pSEP(5)) = 0 And _
                    Val(pSEP(8)) = 0 And _
                    Val(pSEP(9)) = 0 And _
                    pSEP(10) <> "R") Then
                        specialError(7) = False
                End If
            End If
            If SessionParseError = True Then
                If i% = 1 Then
                    pTextPlug = "first"
                ElseIf i% = 2 Then
                    pTextPlug = "second"
                ElseIf i% = 3 Then
                    pTextPlug = "third"
                ElseIf i% = 4 Then
                    pTextPlug = "fourth"
                ElseIf i% = 5 Then
                    pTextPlug = "fifth"
                ElseIf i% = 6 Then
                    pTextPlug = "sixth"
                ElseIf i% = 7 Then
                    pTextPlug = "seventh"
                ElseIf i% = 8 Then
                    pTextPlug = "eighth"
                ElseIf i% = 9 Then
                    pTextPlug = "ninth"
                ElseIf i% = 10 Then
                    pTextPlug = "tenth"
                ElseIf i% = 11 Then
                    pTextPlug = "eleventh"
                End If
                If i% = 1 Then
                    MsgBox ("Australian Deal" & Chr(13) & _
                        "The " & pTextPlug & " Event parameter is invalid." & Chr(13) & Chr(13) & _
                        "Allowable parameter entries are:" & Chr(13) & _
                        Chr(34) & "Dx" & Chr(34) & " (Down first, x = valid pile number)" & Chr(13) & _
                        Chr(34) & "Ux" & Chr(34) & " (Under first, x = valid pile number)" & Chr(13))
                    Exit Sub
                ElseIf i% = 2 Then
                    MsgBox ("Australian Deal" & Chr(13) & _
                        "The " & pTextPlug & " Event parameter is invalid." & Chr(13) & Chr(13) & _
                        "Allowable parameter entries are:" & Chr(13) & _
                        Chr(34) & "x" & Chr(34) & " (0 to 52)" & Chr(13))
                    Exit Sub
                ElseIf i% = 3 Then
                    MsgBox ("Australian Deal" & Chr(13) & _
                        "The " & pTextPlug & " Event parameter is invalid." & Chr(13) & Chr(13) & _
                        "Allowable parameter entries are:" & Chr(13) & _
                        Chr(34) & "x" & Chr(34) & " (0 to 52)" & Chr(13))
                    Exit Sub
                ElseIf i% = 4 Then
                    MsgBox ("Australian Deal" & Chr(13) & _
                        "The " & pTextPlug & " Event parameter is invalid." & Chr(13) & Chr(13) & _
                        "Allowable parameter entries are:" & Chr(13) & _
                        Chr(34) & "x" & Chr(34) & " (0 to 52)" & Chr(13))
                    Exit Sub
                ElseIf i% = 5 Then
                    MsgBox ("Australian Deal" & Chr(13) & _
                        "The " & pTextPlug & " Event parameter is invalid." & Chr(13) & Chr(13) & _
                        "Allowable parameter entries are:" & Chr(13) & _
                        Chr(34) & "x" & Chr(34) & " (0 to 52)" & Chr(13))
                    Exit Sub
                ElseIf i% = 6 Then
                    MsgBox ("Australian Deal" & Chr(13) & _
                        "The " & pTextPlug & " Event parameter is invalid." & Chr(13) & Chr(13) & _
                        "Allowable parameter entries are:" & Chr(13) & _
                        Chr(34) & "S" & Chr(34) & " (reverse Selected card)" & Chr(13) & _
                        Chr(34) & "x" & Chr(34) & " (Reverse All Down = 0 or 1)" & Chr(13))
                    Exit Sub
                ElseIf i% = 7 Then
                    MsgBox ("Australian Deal" & Chr(13) & _
                        "The " & pTextPlug & " Event parameter is invalid." & Chr(13) & Chr(13) & _
                        "Allowable parameter entries are:" & Chr(13) & _
                        Chr(34) & "x" & Chr(34) & " (Reverse All Under = 0 or 1)" & Chr(13))
                    Exit Sub
                ElseIf i% = 8 Then
                    MsgBox ("Australian Deal" & Chr(13) & _
                        "The " & pTextPlug & " Event parameter is invalid." & Chr(13) & Chr(13) & _
                        "Allowable parameter entries are:" & Chr(13) & _
                        Chr(34) & "x" & Chr(34) & " (Reverse Random Down = 0 or 1)" & Chr(13))
                    Exit Sub
                ElseIf i% = 9 Then
                    MsgBox ("Australian Deal" & Chr(13) & _
                        "The " & pTextPlug & " Event parameter is invalid." & Chr(13) & Chr(13) & _
                        "Allowable parameter entries are:" & Chr(13) & _
                        Chr(34) & "x" & Chr(34) & " (Reverse Random Under = 0 or 1)" & Chr(13))
                    Exit Sub
                ElseIf i% = 10 Then
                    MsgBox ("Australian Deal" & Chr(13) & _
                        "The " & pTextPlug & " Event parameter is invalid." & Chr(13) & Chr(13) & _
                        "Allowable parameter entries are:" & Chr(13) & _
                        Chr(34) & "x" & Chr(34) & " (x = 0, for no selection)" & Chr(13) & _
                        Chr(34) & "F" & Chr(34) & " (select Final card)" & Chr(13) & _
                        Chr(34) & "R" & Chr(34) & " (select Random card)" & Chr(13))
                    Exit Sub
                ElseIf i% = 11 Then
                    MsgBox ("Australian Deal" & Chr(13) & _
                        "The " & pTextPlug & " Event parameter is invalid." & Chr(13) & Chr(13) & _
                        "Allowable parameter entries are:" & Chr(13) & _
                        Chr(34) & "x" & Chr(34) & " (Inverse = 0 or 1)" & Chr(13))
                    Exit Sub
                End If
            End If
        Next i%
        If specialError(1) Then
            MsgBox ("Australian Deal: Either of parameters 2 or 4 must be set to 0." & Chr(13) & _
            "The Down setting can't be 'Specified' and 'Random'.")
            SessionParseError = True
            Exit Sub
        End If
        If specialError(2) Then
            MsgBox ("Australian Deal: Either of parameters 3 or 5 must be set to 0." & Chr(13) & _
            "The Under setting can't be 'Specified' and 'Random'.")
            SessionParseError = True
            Exit Sub
        End If
        If specialError(3) Then
            MsgBox ("Australian Deal: Either of parameters 2 or 4 must be greater than 0." & Chr(13) & _
            "A number of cards for the Down setting must be 'Specified' or 'Random'.")
            SessionParseError = True
            Exit Sub
        End If
        If specialError(4) Then
            MsgBox ("Australian Deal: Either of parameters 3 or 5 must be greater than 0." & Chr(13) & _
            "A number of cards for the Under setting must be 'Specified' or 'Random'.")
            SessionParseError = True
            Exit Sub
        End If
        If specialError(5) Then
            MsgBox ("Australian Deal: Either of parameters 6 or 8 must be set to 0." & Chr(13) & _
            "The Reverse card Down setting can't be 'All' and 'Random'.")
            SessionParseError = True
            Exit Sub
        End If
        If specialError(6) Then
            MsgBox ("Australian Deal: Either of parameters 7 or 9 must be set to 0." & Chr(13) & _
            "The Reverse card Down setting can't be 'All' and 'Random'.")
            SessionParseError = True
            Exit Sub
        End If
        If specialError(7) Then
            MsgBox ("Australian Deal: If the Inverse parameter (11) is set, " & Chr(13) & _
            "there can not be any Random settings." & Chr(13) & _
            "Parameters 4, 5, 8, and 9 must all be set to 0," & Chr(13) & _
            " parameter 10 can not be set to 'R'.")
            SessionParseError = True
            Exit Sub
        End If
        If specialError(8) Then
            MsgBox ("Australian Deal:" & Chr(13) & Chr(13) & _
            "You have indicated that the selected card is to be reversed, (P6 = S)" & Chr(13) & _
            "but you have not indicated which card will be the selected one." & Chr(13) & Chr(13) & _
            "P10 must be an " & Chr(34) & Chr(34) & "F (for Final)" & Chr(34) & Chr(34) & " or an " & _
            Chr(34) & Chr(34) & "R (for Random)" & Chr(34) & Chr(34))
            SessionParseError = True
            Exit Sub
        End If
        Call frmPiles.AustralianDeal(pSEP(1), pSEP(2), pSEP(3), pSEP(4), pSEP(5), pSEP(6), _
            pSEP(7), pSEP(8), pSEP(9), pSEP(10), pSEP(11))
End Sub

Public Sub RiffleShufflePilesCaseExtension()
        'check for existing piles
        If PilesShown = 0 Then
            MsgBox ("This event requires the presence of piles.")
            Exit Sub
        End If
        Dim protectedBlockLimit As Integer
        Dim protectedBlockIndex As Integer
        For i% = 1 To 2
        If Len(pSEP(i%)) = 1 Then
            If IsNumeric(pSEP(i%)) And _
                (Val(pSEP(i%)) < 1 Or _
                Val(pSEP(i%)) > NumPiles) Then
                    MsgBox ("Riffle Shuffle Piles Event parameter is out of range." & Chr(13) & _
                    "It must be between 1 and the number of piles," & Chr(13) & _
                    "with an optional prefix of " & Chr(34) & "T" & Chr(34) & _
                    " or " & Chr(34) & "B" & Chr(34) & "," & _
                    "and an optional suffix of " & Chr(34) & "R" & Chr(34) & ".")
                    SessionParseError = True
                    Exit Sub
            ElseIf Not IsNumeric(pSEP(i%)) Then
                    MsgBox ("Riffle Shuffle Piles Event parameter is out of range." & Chr(13) & _
                    "It must be between 1 and the number of piles," & Chr(13) & _
                    "with an optional prefix of " & Chr(34) & "T" & Chr(34) & _
                    " or " & Chr(34) & "B" & Chr(34) & "," & _
                    "and an optional suffix of " & Chr(34) & "R" & Chr(34) & ".")
                SessionParseError = True
                Exit Sub
            End If
        ElseIf Len(pSEP(i%)) = 2 Then
            If IsNumeric(Left(pSEP(i%), 1)) Then
                If Val(Left(pSEP(i%), 1)) < 1 Or _
                    Val(Left(pSEP(i%), 1)) > NumPiles Then
                    MsgBox ("Riffle Shuffle Piles Event parameter is out of range." & Chr(13) & _
                    "It must be between 1 and the number of piles," & Chr(13) & _
                    "with an optional prefix of " & Chr(34) & "T" & Chr(34) & _
                    " or " & Chr(34) & "B" & Chr(34) & "," & _
                    "and an optional suffix of " & Chr(34) & "R" & Chr(34) & ".")
                    SessionParseError = True
                    Exit Sub
                End If
                If Right(pSEP(i%), 1) <> "R" Then
                    MsgBox ("Riffle Shuffle Piles Event parameter is out of range." & Chr(13) & _
                    "It must be between 1 and the number of piles," & Chr(13) & _
                    "with an optional prefix of " & Chr(34) & "T" & Chr(34) & _
                    " or " & Chr(34) & "B" & Chr(34) & "," & _
                    "and an optional suffix of " & Chr(34) & "R" & Chr(34) & ".")
                    SessionParseError = True
                    Exit Sub
                End If
            ElseIf IsNumeric(Right(pSEP(i%), 1)) Then
                If Val(Right(pSEP(i%), 1)) < 1 Or _
                    Val(Right(pSEP(i%), 1)) > NumPiles Then
                    MsgBox ("Riffle Shuffle Piles Event parameter is out of range." & Chr(13) & _
                    "It must be between 1 and the number of piles," & Chr(13) & _
                    "with an optional prefix of " & Chr(34) & "T" & Chr(34) & _
                    " or " & Chr(34) & "B" & Chr(34) & "," & _
                    "and an optional suffix of " & Chr(34) & "R" & Chr(34) & ".")
                    SessionParseError = True
                    Exit Sub
                End If
                If Left(pSEP(i%), 1) <> "T" And _
                    Left(pSEP(i%), 1) <> "B" Then
                    MsgBox ("Riffle Shuffle Piles Event parameter is out of range." & Chr(13) & _
                    "It must be between 1 and the number of piles," & Chr(13) & _
                    "with an optional prefix of " & Chr(34) & "T" & Chr(34) & _
                    " or " & Chr(34) & "B" & Chr(34) & "," & _
                    "and an optional suffix of " & Chr(34) & "R" & Chr(34) & ".")
                    SessionParseError = True
                    Exit Sub
                End If
                protectedBlockIndex = Val(Right(pSEP(i%), 1))
                protectedBlockLimit = PileTable(protectedBlockIndex, 2) - _
                    PileTable(protectedBlockIndex, 1) + 1
            Else
                    MsgBox ("Riffle Shuffle Piles Event parameter is out of range." & Chr(13) & _
                    "It must be between 1 and the number of piles," & Chr(13) & _
                    "with an optional prefix of " & Chr(34) & "T" & Chr(34) & _
                    " or " & Chr(34) & "B" & Chr(34) & "," & _
                    "and an optional suffix of " & Chr(34) & "R" & Chr(34) & ".")
                SessionParseError = True
                Exit Sub
            End If
        ElseIf Len(pSEP(i%)) = 3 Then
            If IsNumeric(Mid(pSEP(i%), 2, 1)) Then
                If Val(Mid(pSEP(i%), 2, 1)) < 1 Or _
                    Val(Mid(pSEP(i%), 2, 1)) > NumPiles Then
                    MsgBox ("Riffle Shuffle Piles Event parameter is out of range." & Chr(13) & _
                    "It must be between 1 and the number of piles," & Chr(13) & _
                    "with an optional prefix of " & Chr(34) & "T" & Chr(34) & _
                    " or " & Chr(34) & "B" & Chr(34) & "," & _
                    "and an optional suffix of " & Chr(34) & "R" & Chr(34) & ".")
                    SessionParseError = True
                    Exit Sub
                End If
            Else
                    MsgBox ("Riffle Shuffle Piles Event parameter is out of range." & Chr(13) & _
                    "It must be between 1 and the number of piles," & Chr(13) & _
                    "with an optional prefix of " & Chr(34) & "T" & Chr(34) & _
                    " or " & Chr(34) & "B" & Chr(34) & "," & _
                    "and an optional suffix of " & Chr(34) & "R" & Chr(34) & ".")
                SessionParseError = True
                Exit Sub
            End If
            If Left(pSEP(i%), 1) <> "T" And _
                Left(pSEP(i%), 1) <> "B" Then
                    MsgBox ("Riffle Shuffle Piles Event parameter is out of range." & Chr(13) & _
                    "It must be between 1 and the number of piles," & Chr(13) & _
                    "with an optional prefix of " & Chr(34) & "T" & Chr(34) & _
                    " or " & Chr(34) & "B" & Chr(34) & "," & _
                    "and an optional suffix of " & Chr(34) & "R" & Chr(34) & ".")
                SessionParseError = True
                Exit Sub
            End If
            protectedBlockIndex = Val(Mid(pSEP(i%), 2, 1))
            protectedBlockLimit = PileTable(protectedBlockIndex, 2) - _
                PileTable(protectedBlockIndex, 1) + 1
            If Right(pSEP(i%), 1) <> "R" Then
                    MsgBox ("Riffle Shuffle Piles Event parameter is out of range." & Chr(13) & _
                    "It must be between 1 and the number of piles," & Chr(13) & _
                    "with an optional prefix of " & Chr(34) & "T" & Chr(34) & _
                    " or " & Chr(34) & "B" & Chr(34) & "," & _
                    "and an optional suffix of " & Chr(34) & "R" & Chr(34) & ".")
                SessionParseError = True
                Exit Sub
            End If
        End If
        Next i%
        If Not IsNumeric(Left(pSEP(1), 1)) And Not IsNumeric(Left(pSEP(2), 1)) Then
            MsgBox ("Riffle Shuffle Piles Event parameters can not " & _
            "both start with a " & Chr(34) & "T" & Chr(34) & " or " & _
            Chr(34) & "B" & Chr(34) & ".")
            SessionParseError = True
            Exit Sub
        End If
        If IsNumeric(pSEP(3)) And _
            (Val(pSEP(3)) < 0 Or _
            Val(pSEP(3)) > protectedBlockLimit) Then
                MsgBox ("Riffle Shuffle Piles Event parameter is out of range." & Chr(13) & _
                "The third parameter must be between 0 and the number of" & Chr(13) & _
                "cards in the protected pile.")
                SessionParseError = True
                Exit Sub
        ElseIf Not IsNumeric(pSEP(3)) Then
            MsgBox ("Riffle Shuffle Piles Event parameter is out of range." & Chr(13) & _
            "The third parameter must be a number between 0 and the number of" & Chr(13) & _
                "cards in the protected pile.")
            SessionParseError = True
            Exit Sub
        End If
        If Not pSEP(4) = "G" And _
            Not pSEP(4) = "X" Then
            MsgBox ("Riffle Shuffle Piles Event parameter is not valid." & Chr(13) & _
            "The fourth parameter must be either a 'G' for Gilbreath View enabled," & Chr(13) & _
                "or an 'X' for Gilbreath View disabled.")
            SessionParseError = True
            Exit Sub
        End If
        Call frmPiles.RiffleShufflePile(pSEP(1), pSEP(2), pSEP(3), pSEP(4))
End Sub
Public Sub CombinePilesCaseExtension()
        'check for existing piles
        If PilesShown = 0 Then
            MsgBox ("This event requires the presence of piles.")
            Exit Sub
        End If
        Dim leftParam As Integer
        Dim rightParam As Integer
        For i% = 1 To 2
        If Len(pSEP(i%)) = 1 Then
            If IsNumeric(pSEP(i%)) And _
                (Val(pSEP(i%)) < 1 Or _
                Val(pSEP(i%)) > 8) Then
                    MsgBox ("Combine Piles Event parameter is out of range." & Chr(13) & _
                    "It must be between 1 and 8," & Chr(13) & _
                    "with an optional prefix of " & Chr(34) & "T" & Chr(34) & "," & _
                    "and an optional suffix of " & Chr(34) & "R" & Chr(34) & ".")
                    SessionParseError = True
                    Exit Sub
            ElseIf Not IsNumeric(pSEP(i%)) Then
                    MsgBox ("Combine Piles Event parameter is out of range." & Chr(13) & _
                    "It must be between 1 and 8," & Chr(13) & _
                    "with an optional prefix of " & Chr(34) & "T" & Chr(34) & "," & _
                    "and an optional suffix of " & Chr(34) & "R" & Chr(34) & ".")
                SessionParseError = True
                Exit Sub
            End If
        ElseIf Len(pSEP(i%)) = 2 Then
            If IsNumeric(Left(pSEP(i%), 1)) Then
                If Val(Left(pSEP(i%), 1)) < 1 Or _
                    Val(Left(pSEP(i%), 1)) > 8 Then
                    MsgBox ("Combine Piles Event parameter is out of range." & Chr(13) & _
                    "It must be between 1 and 8," & Chr(13) & _
                    "with an optional prefix of " & Chr(34) & "T" & Chr(34) & "," & _
                    "and an optional suffix of " & Chr(34) & "R" & Chr(34) & ".")
                    SessionParseError = True
                    Exit Sub
                End If
                If Right(pSEP(i%), 1) <> "R" Then
                    MsgBox ("Combine Piles Event parameter is out of range." & Chr(13) & _
                    "It must be between 1 and 8," & Chr(13) & _
                    "with an optional prefix of " & Chr(34) & "T" & Chr(34) & "," & _
                    "and an optional suffix of " & Chr(34) & "R" & Chr(34) & ".")
                    SessionParseError = True
                    Exit Sub
                End If
            ElseIf IsNumeric(Right(pSEP(i%), 1)) Then
                If Val(Right(pSEP(i%), 1)) < 1 Or _
                    Val(Right(pSEP(i%), 1)) > 8 Then
                    MsgBox ("Combine Piles Event parameter is out of range." & Chr(13) & _
                    "It must be between 1 and 8," & Chr(13) & _
                    "with an optional prefix of " & Chr(34) & "T" & Chr(34) & "," & _
                    "and an optional suffix of " & Chr(34) & "R" & Chr(34) & ".")
                    SessionParseError = True
                    Exit Sub
                End If
                If Left(pSEP(i%), 1) <> "T" Then
                    MsgBox ("Combine Piles Event parameter is out of range." & Chr(13) & _
                    "It must be between 1 and 8," & Chr(13) & _
                    "with an optional prefix of " & Chr(34) & "T" & Chr(34) & "," & _
                    "and an optional suffix of " & Chr(34) & "R" & Chr(34) & ".")
                    SessionParseError = True
                    Exit Sub
                End If
            Else
                MsgBox ("Combine Piles Event parameter is out of range." & Chr(13) & _
                "It must be between 1 and 8," & Chr(13) & _
                "with an optional prefix of " & Chr(34) & "T" & Chr(34) & "," & _
                "and an optional suffix of " & Chr(34) & "R" & Chr(34) & ".")
                SessionParseError = True
                Exit Sub
            End If
        ElseIf Len(pSEP(i%)) = 3 Then
            If IsNumeric(Mid(pSEP(i%), 2, 1)) Then
                If Val(Mid(pSEP(i%), 2, 1)) < 1 Or _
                    Val(Mid(pSEP(i%), 2, 1)) > 8 Then
                    MsgBox ("Combine Piles Event parameter is out of range." & Chr(13) & _
                    "It must be between 1 and 8," & Chr(13) & _
                    "with an optional prefix of " & Chr(34) & "T" & Chr(34) & "," & _
                    "and an optional suffix of " & Chr(34) & "R" & Chr(34) & ".")
                    SessionParseError = True
                    Exit Sub
                End If
            Else
                MsgBox ("Combine Piles Event parameter is out of range." & Chr(13) & _
                "It must be between 1 and 8," & Chr(13) & _
                "with an optional prefix of " & Chr(34) & "T" & Chr(34) & "," & _
                "and an optional suffix of " & Chr(34) & "R" & Chr(34) & ".")
                SessionParseError = True
                Exit Sub
            End If
            If Left(pSEP(i%), 1) <> "T" Then
                MsgBox ("Combine Piles Event parameter is out of range." & Chr(13) & _
                "It must be between 1 and 8," & Chr(13) & _
                "with an optional prefix of " & Chr(34) & "T" & Chr(34) & "," & _
                "and an optional suffix of " & Chr(34) & "R" & Chr(34) & ".")
                SessionParseError = True
                Exit Sub
            End If
            If Right(pSEP(i%), 1) <> "R" Then
                MsgBox ("Combine Piles Event parameter is out of range." & Chr(13) & _
                "It must be between 1 and 8," & Chr(13) & _
                "with an optional prefix of " & Chr(34) & "T" & Chr(34) & "," & _
                "and an optional suffix of " & Chr(34) & "R" & Chr(34) & ".")
                SessionParseError = True
                Exit Sub
            End If
        End If
        Next i%
        If Len(pSEP(1)) = 3 And Len(pSEP(2)) = 3 Then
            MsgBox ("Combine Piles Event parameters can not " & _
            "both start with a " & Chr(34) & "T" & ".")
            SessionParseError = True
            Exit Sub
        End If
        'check if both piles are the same
        'establish first pile
        If Len(pSEP(1)) = 1 Then
            leftParam = Val(pSEP(1))
        ElseIf Len(pSEP(1)) = 2 Then
            If IsNumeric(Left(pSEP(1), 1)) Then
                leftParam = Val(Left(pSEP(1), 1))
            ElseIf IsNumeric(Right(pSEP(1), 1)) Then
                leftParam = Val(Right(pSEP(1), 1))
            End If
        ElseIf Len(pSEP(1)) = 3 Then
            leftParam = Val(Mid(pSEP(1), 2, 1))
        End If
        'establish second pile
        If Len(pSEP(2)) = 1 Then
            rightParam = Val(pSEP(1))
        ElseIf Len(pSEP(2)) = 2 Then
            If IsNumeric(Left(pSEP(2), 1)) Then
                rightParam = Val(Left(pSEP(2), 1))
            ElseIf IsNumeric(Right(pSEP(2), 1)) Then
                rightParam = Val(Right(pSEP(2), 1))
            End If
        ElseIf Len(pSEP(2)) = 3 Then
            rightParam = Val(Mid(pSEP(2), 2, 1))
        End If
        If leftParam = rightParam Then
            MsgBox ("Combine Piles Event parameters can not " & _
            "both indicate the same pile.")
            SessionParseError = True
            Exit Sub
        End If
        Call frmPiles.CombinePiles(pSEP(1), pSEP(2))
End Sub
Public Sub SelectReturnCaseExtension()
        'check for existing piles
        If PilesShown = 0 Then
            MsgBox ("This event requires the presence of piles.")
            Exit Sub
        End If
        For i% = 1 To 5
            'set error condition to True
            SessionParseError = True
            'p1
            If i% = 1 Then
                'check for "R"
                If pSEP(i%) = "R" Then
                    SessionParseError = False
                End If
                'check for "Px" - where x must be greater than 0
                If Len(pSEP(i%)) > 1 And Left(pSEP(i%), 1) = "P" Then
                    If IsNumeric(Right(pSEP(i%), Len(pSEP(i%)) - 1)) And _
                        Val(Right(pSEP(i%), Len(pSEP(i%)) - 1)) >= 1 Then
                        SessionParseError = False
                    End If
                End If
            'p2
            ElseIf i% = 2 Then
                'check for "T"
                If pSEP(i%) = "T" Then
                    SessionParseError = False
                End If
                'check for "B"
                If pSEP(i%) = "B" Then
                    SessionParseError = False
                End If
                'check for "R"
                If pSEP(i%) = "R" Then
                    SessionParseError = False
                End If
                'check for "TR"
                If pSEP(i%) = "TR" Then
                    SessionParseError = False
                End If
                'check for "BR"
                If pSEP(i%) = "BR" Then
                    SessionParseError = False
                End If
                'check for "RR"
                If pSEP(i%) = "RR" Then
                    SessionParseError = False
                End If
                'check for "Sxx" - where xx must be from within 1 and 52
                If Len(pSEP(i%)) > 1 And Left(pSEP(i%), 1) = "S" Then
                    If IsNumeric(Right(pSEP(i%), Len(pSEP(i%)) - 1)) And _
                        Val(Right(pSEP(i%), Len(pSEP(i%)) - 1)) >= 1 And _
                        Val(Right(pSEP(i%), Len(pSEP(i%)) - 1)) <= 52 Then
                        SessionParseError = False
                    End If
                End If
                'check for "SxxR" - where xx must be from within 1 and 52
                If Len(pSEP(i%)) > 1 And Left(pSEP(i%), 1) = "S" And _
                    Right(pSEP(i%), 1) = "R" Then
                    If IsNumeric(Mid(pSEP(i%), 2, Len(pSEP(i%)) - 2)) And _
                        Val(Mid(pSEP(i%), 2, Len(pSEP(i%)) - 2)) >= 1 And _
                        Val(Mid(pSEP(i%), 2, Len(pSEP(i%)) - 2)) <= 52 Then
                        SessionParseError = False
                    End If
                End If
            'p3
            ElseIf i% = 3 Then
                'check for "Sx" - where x must be greater than 0
                If Len(pSEP(i%)) > 1 And Left(pSEP(i%), 1) = "S" Then
                    If IsNumeric(Right(pSEP(i%), Len(pSEP(i%)) - 1)) And _
                        Val(Right(pSEP(i%), Len(pSEP(i%)) - 1)) >= 1 Then
                        SessionParseError = False
                    End If
                End If
                'check for "Px" - where x must be greater than 0
                If Len(pSEP(i%)) > 1 And Left(pSEP(i%), 1) = "P" Then
                    If IsNumeric(Right(pSEP(i%), Len(pSEP(i%)) - 1)) And _
                        Val(Right(pSEP(i%), Len(pSEP(i%)) - 1)) >= 1 Then
                        SessionParseError = False
                    End If
                End If
                'check for "E"
                If pSEP(i%) = "E" Then
                    SessionParseError = False
                End If
                'check for "R"
                If pSEP(i%) = "R" Then
                    SessionParseError = False
                End If
                'check for "D"
                If pSEP(i%) = "D" Then
                    SessionParseError = False
                End If
                'check for "L"
                If pSEP(i%) = "L" Then
                    SessionParseError = False
                End If
                'check for "Nxx" - where xx must be greater than 0
                If Len(pSEP(i%)) > 1 And Left(pSEP(i%), 1) = "N" Then
                    If IsNumeric(Right(pSEP(i%), Len(pSEP(i%)) - 1)) And _
                        Val(Right(pSEP(i%), Len(pSEP(i%)) - 1)) >= 1 Then
                        SessionParseError = False
                    End If
                End If
            'p4
            ElseIf i% = 4 Then
                'check for "N"
                If pSEP(i%) = "N" Then
                    SessionParseError = False
                End If
                'check for "T"
                If pSEP(i%) = "T" Then
                    SessionParseError = False
                End If
                'check for "B"
                If pSEP(i%) = "B" Then
                    SessionParseError = False
                End If
                'check for "E"
                If pSEP(i%) = "E" Then
                    SessionParseError = False
                End If
                'check for "R"
                If pSEP(i%) = "R" Then
                    SessionParseError = False
                End If
                'check for "Sx" - where x must be greater than 0
                If Len(pSEP(i%)) > 1 And Left(pSEP(i%), 1) = "S" Then
                    If IsNumeric(Right(pSEP(i%), Len(pSEP(i%)) - 1)) And _
                        Val(Right(pSEP(i%), Len(pSEP(i%)) - 1)) >= 1 Then
                        SessionParseError = False
                    End If
                End If
            'p5
            ElseIf i% = 5 Then
                'check for "M"
                If pSEP(i%) = "M" Then
                    SessionParseError = False
                End If
                'check for "S"
                If pSEP(i%) = "S" Then
                    SessionParseError = False
                End If
            End If
            If SessionParseError = True Then
                If i% = 1 Then
                    pTextPlug = "first"
                ElseIf i% = 2 Then
                    pTextPlug = "second"
                ElseIf i% = 3 Then
                    pTextPlug = "third"
                ElseIf i% = 4 Then
                    pTextPlug = "fourth"
                ElseIf i% = 5 Then
                    pTextPlug = "fifth"
                End If
                If i% = 1 Then
                    MsgBox ("Select Return" & Chr(13) & _
                        "The " & pTextPlug & " Event parameter is invalid." & Chr(13) & Chr(13) & _
                        "Allowable parameter entries are:" & Chr(13) & _
                        Chr(34) & "R" & Chr(34) & " (for Random)" & Chr(13) & _
                        Chr(34) & "Px" & Chr(34) & " (Primary, x > 0)" & Chr(13))
                    Exit Sub
                ElseIf i% = 2 Then
                    MsgBox ("Select Return" & Chr(13) & _
                        "The " & pTextPlug & " Event parameter is invalid." & Chr(13) & Chr(13) & _
                        "Allowable parameter entries are:" & Chr(13) & _
                        Chr(34) & "T" & Chr(34) & " (Top Card)" & Chr(13) & _
                        Chr(34) & "B" & Chr(34) & " (Bottom Card)" & Chr(13) & _
                        Chr(34) & "R" & Chr(34) & " (Random)" & Chr(13) & _
                        Chr(34) & "TR" & Chr(34) & " (Top Card - Reverse)" & Chr(13) & _
                        Chr(34) & "BR" & Chr(34) & " (Bottom Card - Reverse)" & Chr(13) & _
                        Chr(34) & "RR" & Chr(34) & " (Random - Reverse)" & Chr(13) & _
                        Chr(34) & "Syy" & Chr(34) & " (Specified yy from 1 to 52)" & Chr(13) & _
                        Chr(34) & "SyyR" & Chr(34) & " (Specified yy  - Reverse from 1 to 52)" & Chr(13))
                    Exit Sub
                ElseIf i% = 3 Then
                    MsgBox ("Select Return" & Chr(13) & _
                        "The " & pTextPlug & " Event parameter is invalid." & Chr(13) & Chr(13) & _
                        "Allowable parameter entries are:" & Chr(13) & _
                        Chr(34) & "Sx" & Chr(34) & " (Secondary, x > 0)" & Chr(13) & _
                        Chr(34) & "Px" & Chr(34) & " (Primary, x > 0)" & Chr(13) & _
                        Chr(34) & "E" & Chr(34) & " (for Equivalent Random)" & Chr(13) & _
                        Chr(34) & "R" & Chr(34) & " (Random)" & Chr(13) & _
                        Chr(34) & "D" & Chr(34) & " (Random not same)" & Chr(13) & _
                        Chr(34) & "Nx" & Chr(34) & " (New Pile, x > 0)" & Chr(13) & _
                        Chr(34) & "L" & Chr(34) & " (Random Pile)" & Chr(13))
                    Exit Sub
                ElseIf i% = 4 Then
                    MsgBox ("Select Return" & Chr(13) & _
                        "The " & pTextPlug & " Event parameter is invalid." & Chr(13) & Chr(13) & _
                        "Allowable parameter entries are:" & Chr(13) & _
                        Chr(34) & "N" & Chr(34) & " (New Pile)" & Chr(13) & _
                        Chr(34) & "T" & Chr(34) & " (Top)" & Chr(13) & _
                        Chr(34) & "B" & Chr(34) & " (Bottom)" & Chr(13) & _
                        Chr(34) & "E" & Chr(34) & " (Same Position)" & Chr(13) & _
                        Chr(34) & "R" & Chr(34) & " (Random)" & Chr(13) & _
                        Chr(34) & "Sx" & Chr(34) & " (Specified, x > 0)" & Chr(13))
                    Exit Sub
                ElseIf i% = 5 Then
                    MsgBox ("Select Return" & Chr(13) & _
                        "The " & pTextPlug & " Event parameter is invalid." & Chr(13) & Chr(13) & _
                        "Allowable parameter entries are:" & Chr(13) & _
                        Chr(34) & "M" & Chr(34) & " (Move Only)" & Chr(13) & _
                        Chr(34) & "S" & Chr(34) & " (Select Card)" & Chr(13))
                    Exit Sub
                End If
            End If
        Next i%
        Call frmPiles.SelectReturn(pSEP(1), pSEP(2), pSEP(3), pSEP(4), pSEP(5))
End Sub

Public Sub SessionFilePlay(fileParam)
If SessionRecursionLevel = 10 Then
    MsgBox ("You may not go more than 10 levels deep" & Chr(13) & _
    "with Session Macro files." & Chr(13) & Chr(13) & _
    "Macro file: " & fileParam & Chr(13) & _
    "was called from Level 10 and will not play.")
    Exit Sub
End If
If Not ((Right(fileParam, 3) = "svs") Or (Right(fileParam, 3) = "SVS")) Then
    MsgBox ("Requested file: " & fileParam & "  Invalid file type." & Chr(13) & _
    "Macro Session files must have a .svs extension.")
    Exit Sub
End If

'set the recusion level
SessionRecursionLevel = SessionRecursionLevel + 1
SessionRecursing = True
'get the file for recursion level
On Error GoTo SessionOpenError
SessionRecursionList(SessionRecursionLevel).Clear
Dim fso As New FileSystemObject, sessionfile As File, ts As TextStream
Set sessionfile = fso.GetFile(App.Path & "\" & fileParam)
Set ts = sessionfile.OpenAsTextStream(ForReading)
Dim listCounter As Integer
listCounter = 0 'set file index counter to top line of SessionListBox
Do While ts.AtEndOfStream <> True
    SessionRecursionList(SessionRecursionLevel).List(listCounter) = ts.ReadLine
    listCounter = listCounter + 1
Loop
ts.Close
'run the macro file
If SessionRecursionList(SessionRecursionLevel).ListCount = 0 Then
    MsgBox ("There are no entries in macro session:" & Chr(13) & _
    fileParam & " to play.")
    Exit Sub
End If
For i% = 0 To SessionRecursionList(SessionRecursionLevel).ListCount - 1
    SessionRecursionList(SessionRecursionLevel).ListIndex = i%
    Call SessionParse(i%)
    If SessionParseError Then
        SessionParseError = False
        Exit Sub
    End If
Next i%
SessionRecursionLevel = SessionRecursionLevel - 1
If SessionRecursionLevel = 0 Then
    SessionRecursing = False
End If
Exit Sub
SessionOpenError:
'first, reset the recursion counter
SessionRecursionLevel = SessionRecursionLevel - 1
If SessionRecursionLevel = 0 Then
    SessionRecursing = False
End If
MsgBox ("Error reading macro file.  " & Chr(13) & _
    "File: " & fileParam & "  may be corrupt or missing.")
Exit Sub
End Sub


Public Sub SessionPlayEvent_Click()
sessionindex = SessionListBox.ListIndex
If sessionindex = -1 Then
    MsgBox "Please select an Event to play."
    Exit Sub
End If
SessionRecursionLevel = 0
If Left(SessionListBox.List(sessionindex), 4) = "Free" Or _
    Left(SessionListBox.List(sessionindex), 5) = "Force" Then
    Call SessionParse(sessionindex)
    If SessionParseError Then
        SessionParseError = False
        Exit Sub
    End If
    Call SessionParse(sessionindex + 1)
    If SessionParseError Then
        SessionParseError = False
        SessionListBox.ListIndex = sessionindex + 1
        Exit Sub
    End If
    SessionListBox.ListIndex = sessionindex + 1
    sessionindex = sessionindex + 1
ElseIf Left(SessionListBox.List(sessionindex), 6) = "Return" Then
    Call SessionParse(sessionindex - 1)
    If SessionParseError Then
        SessionParseError = False
        SessionListBox.ListIndex = sessionindex - 1
        Exit Sub
    End If
    Call SessionParse(sessionindex)
    If SessionParseError Then
        SessionParseError = False
        Exit Sub
    End If
Else
    Call SessionParse(sessionindex)
    If SessionParseError Then
        SessionParseError = False
        Exit Sub
    End If
End If
If sessionindex = SessionListBox.ListCount - 1 Then
    MsgBox "That was the last Event in this Session"
    Exit Sub
Else
    SessionListBox.ListIndex = sessionindex + 1
End If
End Sub
Public Sub SessionStatusUpdate(sessionparameter)
'a value of 0 means that the session has changed and is not saved
'a value of 1 means that there is no current session
'a text filename means that the file has been just opened or saved
Select Case sessionparameter
    Case 0
        SessionFileName.Caption = "Current Session NOT saved"
        SessionSaved = 0
    Case 1
        SessionFileName.Caption = "No current session"
        SessionSaved = 1
    Case Else
        SessionFileName.Caption = sessionparameter
        SessionSaved = 1
End Select
End Sub



Public Sub ShiftTopBlockButton_Click()
If ShiftTopBlockTextBox.Text = Empty Or _
    ShiftTopDepthTextBox.Text = Empty Or _
    Not IsNumeric(ShiftTopBlockTextBox.Text) Or _
    Not IsNumeric(ShiftTopDepthTextBox.Text) Or _
    (Val(ShiftTopBlockTextBox.Text) + _
        Val(ShiftTopDepthTextBox.Text)) > 52 Then
    ShiftTopBlockTextBox.Text = Empty
    ShiftTopDepthTextBox.Text = Empty
    MsgBox "Please enter a valid card position in the" & Chr(13) _
        & "'Block' and 'Depth' Input Boxes" & Chr(13) & Chr(13) _
        & "The sum of 'Block' and 'Depth'" & Chr(13) _
        & "must not be greater than 52."
    Exit Sub
End If
If ShiftTopBlockInverseCheck Then
    If ShiftTopBlockReverseCheck Then
        Call InverseShiftTopBlockReverse(Val(ShiftTopBlockTextBox.Text), Val(ShiftTopDepthTextBox.Text))
        'SessionRecord
        If SessionRecordMode Then
            SessionCommand = "InverseShiftTopBlockReverse(" & Val(ShiftTopBlockTextBox.Text) _
            & ", " & Val(ShiftTopDepthTextBox.Text) & ")"
            SessionListBox.AddItem SessionCommand
            SessionStatusUpdate (0)
        End If
    Else
        Call InverseShiftTopBlock(Val(ShiftTopBlockTextBox.Text), Val(ShiftTopDepthTextBox.Text))
        'SessionRecord
        If SessionRecordMode Then
            SessionCommand = "InverseShiftTopBlock(" & Val(ShiftTopBlockTextBox.Text) _
            & ", " & Val(ShiftTopDepthTextBox.Text) & ")"
            SessionListBox.AddItem SessionCommand
            SessionStatusUpdate (0)
        End If
    End If
Else
    If ShiftTopBlockReverseCheck Then
        Call ShiftTopBlockReverse(Val(ShiftTopBlockTextBox.Text), Val(ShiftTopDepthTextBox.Text))
        'SessionRecord
        If SessionRecordMode Then
            SessionCommand = "ShiftTopBlockReverse(" & Val(ShiftTopBlockTextBox.Text) _
            & ", " & Val(ShiftTopDepthTextBox.Text) & ")"
            SessionListBox.AddItem SessionCommand
            SessionStatusUpdate (0)
        End If
    Else
        Call ShiftTopBlock(Val(ShiftTopBlockTextBox.Text), Val(ShiftTopDepthTextBox.Text))
        'SessionRecord
        If SessionRecordMode Then
            SessionCommand = "ShiftTopBlock(" & Val(ShiftTopBlockTextBox.Text) _
            & ", " & Val(ShiftTopDepthTextBox.Text) & ")"
            SessionListBox.AddItem SessionCommand
            SessionStatusUpdate (0)
        End If
    End If
End If

End Sub

Public Sub ShiftTopBlockTextBox_LostFocus()
If ShiftTopBlockTextBox.Text <> Empty And _
    (Not IsNumeric(ShiftTopBlockTextBox.Text) Or _
    Val(ShiftTopBlockTextBox.Text) < 1 Or _
    Val(ShiftTopBlockTextBox.Text) > 52) Then
    ShiftTopBlockTextBox.Text = Empty
    MsgBox "Please enter a valid card position (1 to 52)" & Chr(13) _
        & "in the 'Block' Input Box"
    ShiftTopBlockTextBox.SetFocus
    Exit Sub
End If
End Sub

Public Sub ShiftTopDepthTextBox_LostFocus()
If ShiftTopDepthTextBox.Text <> Empty And _
    (Not IsNumeric(ShiftTopDepthTextBox.Text) Or _
    Val(ShiftTopDepthTextBox.Text) < 1 Or _
    Val(ShiftTopDepthTextBox.Text) > 52) Then
    ShiftTopDepthTextBox.Text = Empty
    MsgBox "Please enter a valid card position (1 to 52)" & Chr(13) _
        & "in the 'Depth' Input Box"
    ShiftTopDepthTextBox.SetFocus
    Exit Sub
End If
End Sub





Public Sub ShowIndexValues_Click()
If PilesShown = 1 Then
    ShowPiles
ElseIf PokerCardsDealt = 1 Then
    ShowDeal
Else
    ShowCards
End If
End Sub




Public Sub Form_Unload(Cancel As Integer)
    Dim Msg   ' Declare variable.
    ' Set the message text.
    Msg = ""
    If SessionSaved = 0 Then
        Msg = Msg & "You have unsaved Session activity still present." & Chr(13)
        Msg = Msg & "Closing this window will clear your Session progress." & Chr(13)
        Msg = Msg & Chr(13) & "Do you really want to close this window?" & Chr(13)
        ' If user clicks the No button, stop QueryUnload.
        If MsgBox(Msg, vbQuestion + vbYesNo, Me.Caption) = vbNo Then
            Cancel = True
        Else
            SessionSaved = 1
            SessionRecordMode = False
            frmMain.mnuControl.Checked = False
        End If
    Else
        frmMain.mnuControl.Checked = False
    End If
End Sub

Private Sub ShowPositionValues_Click()
If ShowPositionValues = 1 Then
    CountFromBack.Enabled = True
    CountFromFace.Enabled = True
Else
    CountFromBack.Enabled = False
    CountFromFace.Enabled = False
End If
If PilesShown = 1 Then
    ShowPiles
ElseIf PokerCardsDealt = 1 Then
    ShowDeal
Else
    ShowCards
End If
End Sub

Private Sub SwapCardsButton_Click()
Dim pGeneral1 As String
Dim pGeneral2 As String
Dim pGeneral3 As Integer
Dim pGeneral4 As Integer
'set card selection status
pGeneral4 = SwapCardsNoSelection.Value
If SwapRandomOption.Value = True Then
    SwapCardsRandom (pGeneral4)
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "SwapCardsRandom(" & pGeneral4 & ")"
        SessionListBox.AddItem SessionCommand
        SessionStatusUpdate (0)
    End If
ElseIf SwapSpecifiedCardsOption.Value = True Then
    If Not IsNumeric(SwapValue1.Text) Or _
        Val(SwapValue1.Text) < 1 Or _
        Val(SwapValue1.Text) > 52 Then
        SwapValue1.Text = Empty
        MsgBox "Please enter a valid card stack value (1 to 52)" & Chr(13) _
            & "in the 'Specified Card Stack Values' left Input Box"
        Exit Sub
    End If
    If Not IsNumeric(SwapValue2.Text) Or _
        Val(SwapValue2.Text) < 1 Or _
        Val(SwapValue2.Text) > 52 Then
        SwapValue2.Text = Empty
        MsgBox "Please enter a valid card stack value (1 to 52)" & Chr(13) _
            & "in the 'Specified Card Stack Values' right Input Box"
        Exit Sub
    End If
    If Val(SwapValue1.Text) = Val(SwapValue2.Text) Then
        MsgBox "You must enter different values" & Chr(13) _
            & "in the 'Specified Card Stack Values' Input Boxes."
        Exit Sub
    End If
    Call SwapSpecifiedCards(Val(SwapValue1.Text), Val(SwapValue2.Text), pGeneral4)
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "SwapSpecifiedCards(" & SwapValue1.Text & ", " & _
            SwapValue2.Text & ", " & pGeneral4 & ")"
        SessionListBox.AddItem SessionCommand
        SessionStatusUpdate (0)
    End If
ElseIf SwapSpecifiedPositionsOption.Value = True Then
    If Not IsNumeric(SwapPosition1.Text) Or _
        Val(SwapPosition1.Text) < 1 Or _
        Val(SwapPosition1.Text) > 52 Then
        SwapPosition1.Text = Empty
        MsgBox "Please enter a valid card stack value (1 to 52)" & Chr(13) _
            & "in the 'Specified Card Positions' left Input Box"
        Exit Sub
    End If
    If Not IsNumeric(SwapPosition2.Text) Or _
        Val(SwapPosition2.Text) < 1 Or _
        Val(SwapPosition2.Text) > 52 Then
        SwapPosition2.Text = Empty
        MsgBox "Please enter a valid card stack value (1 to 52)" & Chr(13) _
            & "in the 'Specified Card Positions' right Input Box"
        Exit Sub
    End If
    If Val(SwapPosition1.Text) = Val(SwapPosition2.Text) Then
        MsgBox "You must enter different values" & Chr(13) _
            & "in the 'Specified Position Values' Input Boxes."
        Exit Sub
    End If
    Call SwapSpecifiedPositions(Val(SwapPosition1.Text), Val(SwapPosition2.Text), pGeneral4)
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "SwapSpecifiedPositions(" & SwapPosition1.Text & ", " & _
            SwapPosition2.Text & ", " & pGeneral4 & ")"
        SessionListBox.AddItem SessionCommand
        SessionStatusUpdate (0)
    End If
ElseIf SwapDifferentColorsOption.Value = True Then
    SwapDifferentColors (pGeneral4)
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "SwapDifferentColors(" & pGeneral4 & ")"
        SessionListBox.AddItem SessionCommand
        SessionStatusUpdate (0)
    End If
ElseIf SwapSameColorOption.Value = True Then
    If SwapSameColorRandom.Value + _
        SwapSameColorRed.Value + _
        SwapSameColorBlack.Value = 0 Then
        MsgBox "You must select one of the" & Chr(13) _
            & "'Same Color' check boxes."
        Exit Sub
    End If
    If SwapSameColorRandom.Value + _
        SwapSameColorRed.Value + _
        SwapSameColorBlack.Value > 1 Then
        MsgBox "You may only select one of the" & Chr(13) _
            & "'Same Color' check boxes."
        Exit Sub
    End If
    If SwapSameColorRandom.Value = 1 Then
        pGeneral1 = "X"
    ElseIf SwapSameColorRed.Value = 1 Then
        pGeneral1 = "R"
    ElseIf SwapSameColorBlack.Value = 1 Then
        pGeneral1 = "B"
    End If
    Call SwapSameColor(pGeneral1, pGeneral4)
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "SwapSameColor(" & pGeneral1 & ", " & pGeneral4 & ")"
        SessionListBox.AddItem SessionCommand
        SessionStatusUpdate (0)
    End If
ElseIf SwapDifferentSuitsOption.Value = True Then
    If SwapDifferentSuitsRandom.Value + _
        SwapDifferentSuitsClub.Value + _
        SwapDifferentSuitsHeart.Value + _
        SwapDifferentSuitsSpade.Value + _
        SwapDifferentSuitsDiamond.Value = 0 Then
        MsgBox "You have not selected any of the check boxes." & Chr(13) & Chr(13) _
            & "You must select either the 'Random' check box alone" & Chr(13) _
            & "so that both suits are randomly selected, or" & Chr(13) _
            & "any two of the five 'Different Suits' check boxes."
        Exit Sub
    End If
    If SwapDifferentSuitsRandom.Value = 0 And _
        (SwapDifferentSuitsClub.Value + _
        SwapDifferentSuitsHeart.Value + _
        SwapDifferentSuitsSpade.Value + _
        SwapDifferentSuitsDiamond.Value) = 1 Then
        MsgBox "You have specified only one of the suits." & Chr(13) & Chr(13) _
            & "You must select either the 'Random' check box alone" & Chr(13) _
            & "so that both suits are randomly selected, or" & Chr(13) _
            & "any two of the five 'Different Suits' check boxes."
        Exit Sub
    End If
    If SwapDifferentSuitsRandom.Value + _
        SwapDifferentSuitsClub.Value + _
        SwapDifferentSuitsHeart.Value + _
        SwapDifferentSuitsSpade.Value + _
        SwapDifferentSuitsDiamond.Value > 2 Then
        MsgBox "You have selected too many of the check boxes." & Chr(13) & Chr(13) _
            & "You must select either the 'Random' check box alone" & Chr(13) _
            & "so that both suits are randomly selected, or" & Chr(13) _
            & "any two of the five 'Different Suits' check boxes."
        Exit Sub
    End If
    pGeneral3 = 0
    'this sets a marker that the pGeneral1 variable has not been set
    If SwapDifferentSuitsRandom.Value = 1 Then
        pGeneral1 = "X"
        pGeneral3 = 1
    End If
    If SwapDifferentSuitsClub.Value = 1 Then
        If pGeneral3 = 0 Then
            pGeneral1 = "C"
            pGeneral3 = 1
        ElseIf pGeneral3 = 1 Then
            pGeneral2 = "C"
            pGeneral3 = 2
        End If
    End If
    If SwapDifferentSuitsHeart.Value = 1 Then
        If pGeneral3 = 0 Then
            pGeneral1 = "H"
            pGeneral3 = 1
        ElseIf pGeneral3 = 1 Then
            pGeneral2 = "H"
            pGeneral3 = 2
        End If
    End If
    If SwapDifferentSuitsSpade.Value = 1 Then
        If pGeneral3 = 0 Then
            pGeneral1 = "S"
            pGeneral3 = 1
        ElseIf pGeneral3 = 1 Then
            pGeneral2 = "S"
            pGeneral3 = 2
        End If
    End If
    If SwapDifferentSuitsDiamond.Value = 1 Then
        If pGeneral3 = 0 Then
            pGeneral1 = "D"
            pGeneral3 = 1
        ElseIf pGeneral3 = 1 Then
            pGeneral2 = "D"
            pGeneral3 = 2
        End If
    End If
    If pGeneral3 = 1 Then
        pGeneral2 = "X"
        'this is for the condition where only the Random box was checked
    End If
    Call SwapDifferentSuits(pGeneral1, pGeneral2, pGeneral4)
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "SwapDifferentSuits(" & pGeneral1 & ", " & _
            pGeneral2 & ", " & pGeneral4 & ")"
        SessionListBox.AddItem SessionCommand
        SessionStatusUpdate (0)
    End If
ElseIf SwapSameSuitOption.Value = True Then
    If SwapSameSuitRandom.Value + _
        SwapSameSuitClub.Value + _
        SwapSameSuitHeart.Value + _
        SwapSameSuitSpade.Value + _
        SwapSameSuitDiamond.Value = 0 Then
        MsgBox "You have not selected any of the check boxes." & Chr(13) & Chr(13) _
            & "You must select one of the 'Same Suit' check boxes."
        Exit Sub
    End If
    If SwapSameSuitRandom.Value + _
        SwapSameSuitClub.Value + _
        SwapSameSuitHeart.Value + _
        SwapSameSuitSpade.Value + _
        SwapSameSuitDiamond.Value > 1 Then
        MsgBox "You have selected too many of the check boxes." & Chr(13) & Chr(13) _
            & "You must select only one of the 'Same Suit' check boxes."
        Exit Sub
    End If
    If SwapSameSuitRandom.Value = 1 Then
        pGeneral1 = "X"
    ElseIf SwapSameSuitClub.Value = 1 Then
        pGeneral1 = "C"
    ElseIf SwapSameSuitHeart.Value = 1 Then
        pGeneral1 = "H"
    ElseIf SwapSameSuitSpade.Value = 1 Then
        pGeneral1 = "S"
    ElseIf SwapSameSuitDiamond.Value = 1 Then
        pGeneral1 = "D"
    End If
    Call SwapSameSuit(pGeneral1, pGeneral4)
    'SessionRecord
    If SessionRecordMode Then
        SessionCommand = "SwapSameSuit(" & pGeneral1 & ", " & pGeneral4 & ")"
        SessionListBox.AddItem SessionCommand
        SessionStatusUpdate (0)
    End If
End If
End Sub

Private Sub SwapCardsRandom(param1)
Dim pCard1 As Integer
Dim pCard2 As Integer
Dim pCard3 As Integer
Dim pControl As Integer
'set the "No Selection" check box status passed parameter
pCard3 = param1
pControl = 0
'when pControl is set to 1, pCard2 has been successfully set
pCard1 = Int(Rnd * 52) + 1
While pControl = 0
    pCard2 = Int(Rnd * 52) + 1
    If pCard2 <> pCard1 Then
        pControl = 1
    End If
Wend
Call SwapCards(pCard1, pCard2, pCard3)
End Sub

Private Sub SwapSpecifiedCards(param1, param2, param3)
Dim pCard1 As Integer
Dim pCard2 As Integer
Dim pCard3 As Integer
'set the "No Selection" check box status passed parameter
pCard3 = param3
For i% = 1 To DeckCount
    If Val(Deck(1, i%)) = param1 Then
        pCard1 = i%
    End If
    If Val(Deck(1, i%)) = param2 Then
        pCard2 = i%
    End If
Next i%
Call SwapCards(pCard1, pCard2, pCard3)
End Sub

Private Sub SwapSpecifiedPositions(param1, param2, param3)
'set the "No Selection" check box status passed parameter
Call SwapCards(param1, param2, param3)
End Sub

Private Sub SwapDifferentColors(param1)
Dim pCard1 As Integer
Dim pCard2 As Integer
Dim pCard3 As Integer
Dim pControl As Integer
'set the "No Selection" check box status passed parameter
pCard3 = param1
'when pControl is set to 1, pCardX has been successfully set
'-----------------
'set Red Card
pControl = 0
While pControl = 0
    pCard1 = Int(Rnd * 52) + 1
    If Right(Deck(3, pCard1), 1) = "H" Or _
        Right(Deck(3, pCard1), 1) = "D" Then
        pControl = 1
    End If
Wend
'-----------------
'set Black card
pControl = 0
While pControl = 0
    pCard2 = Int(Rnd * 52) + 1
    If Right(Deck(3, pCard2), 1) = "C" Or _
        Right(Deck(3, pCard2), 1) = "S" Then
        pControl = 1
    End If
Wend
Call SwapCards(pCard1, pCard2, pCard3)
End Sub

Private Sub SwapSameColor(param1, param2)
Dim pCard1 As Integer
Dim pCard2 As Integer
Dim pCard3 As Integer
Dim pControl As Integer
Dim pColor As Integer
'set the "No Selection" check box status passed parameter
pCard3 = param2
'pColor=1=Red
'pColor=2=Black
pColor = Int(Rnd * 2) + 1
If (param1 = "X" And pColor = 1) Or param1 = "R" Then
    'set first Red Card
    pControl = 0
    While pControl = 0
        pCard1 = Int(Rnd * 52) + 1
        If Right(Deck(3, pCard1), 1) = "H" Or _
            Right(Deck(3, pCard1), 1) = "D" Then
            pControl = 1
        End If
    Wend
    'set second Red card
    pControl = 0
    While pControl = 0
        pCard2 = Int(Rnd * 52) + 1
        If Right(Deck(3, pCard2), 1) = "H" Or _
            Right(Deck(3, pCard2), 1) = "D" Then
            If pCard2 <> pCard1 Then
                pControl = 1
            End If
        End If
    Wend
ElseIf (param1 = "X" And pColor = 2) Or param1 = "B" Then
    'set first Black Card
    pControl = 0
    While pControl = 0
        pCard1 = Int(Rnd * 52) + 1
        If Right(Deck(3, pCard1), 1) = "C" Or _
            Right(Deck(3, pCard1), 1) = "S" Then
            pControl = 1
        End If
    Wend
    'set second Black card
    pControl = 0
    While pControl = 0
        pCard2 = Int(Rnd * 52) + 1
        If Right(Deck(3, pCard2), 1) = "C" Or _
            Right(Deck(3, pCard2), 1) = "S" Then
            If pCard2 <> pCard1 Then
                pControl = 1
            End If
        End If
    Wend
End If
Call SwapCards(pCard1, pCard2, pCard3)
End Sub

Private Sub SwapDifferentSuits(param1, param2, param3)
Dim pCard1 As Integer
Dim pCard1Suit As Integer
Dim pCard2 As Integer
Dim pCard2Suit As Integer
Dim pCard3 As Integer
Dim pControl As Integer
Dim pSuit As Integer
'pSuit=1=Club
'pSuit=2=Heart
'pCSuit=3=Spade
'pSuit=4=Diamond
'initilly set to 0, and will be set to 1 when the first card is established
'-------------------
'set the "No Selection" check box status passed parameter
pCard3 = param3
'set second card
'(  the second card must be set before the first card...
'   the call to this sub can be SwapDifferentSuits(X,C).  If the first random suit is Clubs,
'   then the second parameter will also select Clubs.  The only way there can be an X in the
'   second parameter is if there is also one in the first parameter.)
pSuit = Int(Rnd * 4) + 1
If (param2 = "X" And pSuit = 1) Or param2 = "C" Then
    'set second card as Club
    pControl = 0
    While pControl = 0
        pCard2 = Int(Rnd * 52) + 1
        If Right(Deck(3, pCard2), 1) = "C" Then
            pControl = 1
            pCard2Suit = 1
        End If
    Wend
ElseIf (param2 = "X" And pSuit = 2) Or param2 = "H" Then
    'set second card as Heart
    pControl = 0
    While pControl = 0
        pCard2 = Int(Rnd * 52) + 1
        If Right(Deck(3, pCard2), 1) = "H" Then
            pControl = 1
            pCard2Suit = 2
        End If
    Wend
ElseIf (param2 = "X" And pSuit = 3) Or param2 = "S" Then
    'set second card as Spade
    pControl = 0
    While pControl = 0
        pCard2 = Int(Rnd * 52) + 1
        If Right(Deck(3, pCard2), 1) = "S" Then
            pControl = 1
            pCard2Suit = 3
        End If
    Wend
ElseIf (param2 = "X" And pSuit = 4) Or param2 = "D" Then
    'set second card as Diamond
    pControl = 0
    While pControl = 0
        pCard2 = Int(Rnd * 52) + 1
        If Right(Deck(3, pCard2), 1) = "D" Then
            pControl = 1
            pCard2Suit = 4
        End If
    Wend
End If
'-------------
'set first card
pControl = 0
While pControl = 0
    pSuit = Int(Rnd * 4) + 1
    If pCard2Suit <> pSuit Then
        pControl = 1
    End If
Wend
If (param1 = "X" And pSuit = 1) Or param1 = "C" Then
    'set first card as Club
    pControl = 0
    While pControl = 0
        pCard1 = Int(Rnd * 52) + 1
        If Right(Deck(3, pCard1), 1) = "C" Then
            pControl = 1
            pCard1Suit = 1
        End If
    Wend
ElseIf (param1 = "X" And pSuit = 2) Or param1 = "H" Then
    'set first card as Heart
    pControl = 0
    While pControl = 0
        pCard1 = Int(Rnd * 52) + 1
        If Right(Deck(3, pCard1), 1) = "H" Then
            pControl = 1
            pCard1Suit = 2
        End If
    Wend
ElseIf (param1 = "X" And pSuit = 3) Or param1 = "S" Then
    'set first card as Spade
    pControl = 0
    While pControl = 0
        pCard1 = Int(Rnd * 52) + 1
        If Right(Deck(3, pCard1), 1) = "S" Then
            pControl = 1
            pCard1Suit = 3
        End If
    Wend
ElseIf (param1 = "X" And pSuit = 4) Or param1 = "D" Then
    'set first card as Diamond
    pControl = 0
    While pControl = 0
        pCard1 = Int(Rnd * 52) + 1
        If Right(Deck(3, pCard1), 1) = "D" Then
            pControl = 1
            pCard1Suit = 4
        End If
    Wend
End If
Call SwapCards(pCard1, pCard2, pCard3)
End Sub

Private Sub SwapSameSuit(param1, param2)
Dim pCard1 As Integer
Dim pCard2 As Integer
Dim pCard3 As Integer
Dim pControl As Integer
Dim pSuit As Integer
'pSuit=1=Club
'pSuit=2=Heart
'pSuit=3=Spade
'pSuit=4=Diamond
'set the "No Selection" check box status passed parameter
pCard3 = param2
pSuit = Int(Rnd * 4) + 1
If (param1 = "X" And pSuit = 1) Or param1 = "C" Then
    'set first Club
    pControl = 0
    While pControl = 0
        pCard1 = Int(Rnd * 52) + 1
        If Right(Deck(3, pCard1), 1) = "C" Then
            pControl = 1
        End If
    Wend
    'set second Club
    pControl = 0
    While pControl = 0
        pCard2 = Int(Rnd * 52) + 1
        If Right(Deck(3, pCard2), 1) = "C" Then
            If pCard2 <> pCard1 Then
                pControl = 1
            End If
        End If
    Wend
ElseIf (param1 = "X" And pSuit = 2) Or param1 = "H" Then
    'set first Club
    pControl = 0
    While pControl = 0
        pCard1 = Int(Rnd * 52) + 1
        If Right(Deck(3, pCard1), 1) = "H" Then
            pControl = 1
        End If
    Wend
    'set second Club
    pControl = 0
    While pControl = 0
        pCard2 = Int(Rnd * 52) + 1
        If Right(Deck(3, pCard2), 1) = "H" Then
            If pCard2 <> pCard1 Then
                pControl = 1
            End If
        End If
    Wend
ElseIf (param1 = "X" And pSuit = 3) Or param1 = "S" Then
    'set first Club
    pControl = 0
    While pControl = 0
        pCard1 = Int(Rnd * 52) + 1
        If Right(Deck(3, pCard1), 1) = "S" Then
            pControl = 1
        End If
    Wend
    'set second Club
    pControl = 0
    While pControl = 0
        pCard2 = Int(Rnd * 52) + 1
        If Right(Deck(3, pCard2), 1) = "S" Then
            If pCard2 <> pCard1 Then
                pControl = 1
            End If
        End If
    Wend
ElseIf (param1 = "X" And pSuit = 4) Or param1 = "D" Then
    'set first Club
    pControl = 0
    While pControl = 0
        pCard1 = Int(Rnd * 52) + 1
        If Right(Deck(3, pCard1), 1) = "D" Then
            pControl = 1
        End If
    Wend
    'set second Club
    pControl = 0
    While pControl = 0
        pCard2 = Int(Rnd * 52) + 1
        If Right(Deck(3, pCard2), 1) = "D" Then
            If pCard2 <> pCard1 Then
                pControl = 1
            End If
        End If
    Wend
End If
Call SwapCards(pCard1, pCard2, pCard3)
End Sub

Private Sub SwapCards(pCard1, pCard2, pCard3)
'in this subroutine, pCard1 and pCard2 refer to the two card positions that will swap
'pCard3 represents the "No Selection" setting:
'0=swapped cards are selected
'1=swapped cards are not selected
Dim placeCard(6)
For i% = 1 To DeckProperties
    placeCard(i%) = Deck(i%, pCard1)
Next i%
For i% = 1 To DeckProperties
    Deck(i%, pCard1) = Deck(i%, pCard2)
Next i%
For i% = 1 To DeckProperties
    Deck(i%, pCard2) = placeCard(i%)
Next i%
If pCard3 = 0 Then
    Deck(4, pCard2) = "Selected"
    If SelectionsTextBox.Text = Empty Then
        SelectionsTextBox.Text = Deck(2, pCard2)
    Else
        SelectionsTextBox.Text = SelectionsTextBox.Text & " " & Deck(2, pCard2)
    End If
    Deck(4, pCard1) = "Selected"
    If SelectionsTextBox.Text = Empty Then
        SelectionsTextBox.Text = Deck(2, pCard1)
    Else
        SelectionsTextBox.Text = SelectionsTextBox.Text & " " & Deck(2, pCard1)
    End If
End If
ShowCards
End Sub

Private Sub SwapDifferentSuitsRandom_Click()
SwapDifferentSuitsOption.Value = True
End Sub
Private Sub SwapDifferentSuitsClub_Click()
SwapDifferentSuitsOption.Value = True
End Sub
Private Sub SwapDifferentSuitsHeart_Click()
SwapDifferentSuitsOption.Value = True
End Sub
Private Sub SwapDifferentSuitsSpade_Click()
SwapDifferentSuitsOption.Value = True
End Sub
Private Sub SwapDifferentSuitsDiamond_Click()
SwapDifferentSuitsOption.Value = True
End Sub

Private Sub SwapPosition1_Click()
SwapSpecifiedPositionsOption.Value = True
End Sub
Private Sub SwapPosition2_Click()
SwapSpecifiedPositionsOption.Value = True
End Sub

Private Sub SwapSameColorRandom_Click()
SwapSameColorOption.Value = True
End Sub
Private Sub SwapSameColorRed_Click()
SwapSameColorOption.Value = True
End Sub
Private Sub SwapSameColorBlack_Click()
SwapSameColorOption.Value = True
End Sub

Private Sub SwapSameSuitRandom_Click()
SwapSameSuitOption.Value = True
End Sub
Private Sub SwapSameSuitClub_Click()
SwapSameSuitOption.Value = True
End Sub
Private Sub SwapSameSuitHeart_Click()
SwapSameSuitOption.Value = True
End Sub
Private Sub SwapSameSuitSpade_Click()
SwapSameSuitOption.Value = True
End Sub
Private Sub SwapSameSuitDiamond_Click()
SwapSameSuitOption.Value = True
End Sub

Private Sub SwapValue1_Click()
SwapSpecifiedCardsOption.Value = True
End Sub
Private Sub SwapValue2_Click()
SwapSpecifiedCardsOption.Value = True
End Sub

Private Sub ViewDeckAbove_Click()
ShowCards
End Sub

Private Sub ViewDeckBeneath_Click()
ShowCards
End Sub
