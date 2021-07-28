VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmSearch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "StackView Search"
   ClientHeight    =   6630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8760
   Icon            =   "frmSearch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   8760
   Begin VB.CommandButton TargetDeckButton 
      Height          =   315
      Left            =   1440
      TabIndex        =   1
      Top             =   480
      Width           =   585
   End
   Begin MSComctlLib.ProgressBar SearchProgress 
      Height          =   240
      Left            =   240
      TabIndex        =   54
      Top             =   3585
      Visible         =   0   'False
      Width           =   5700
      _ExtentX        =   10054
      _ExtentY        =   423
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComDlg.CommonDialog SearchCommonDialog 
      Left            =   5775
      Top             =   60
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame SearchItems 
      Height          =   2655
      Left            =   90
      TabIndex        =   49
      Top             =   3915
      Width           =   8595
      Begin VB.CheckBox SearchInFaroSpecialTopAll 
         Caption         =   "Check1"
         Height          =   240
         Left            =   5895
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   1455
         Width           =   210
      End
      Begin VB.CheckBox SearchInFaroSpecialTopInvAll 
         Caption         =   "Check1"
         Height          =   240
         Left            =   5895
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   1755
         Width           =   210
      End
      Begin VB.CheckBox SearchInFaroSpecialBottomAll 
         Caption         =   "Check1"
         Height          =   240
         Left            =   5895
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   2055
         Width           =   210
      End
      Begin VB.CheckBox SearchInFaroSpecialBottomInvAll 
         Caption         =   "Check1"
         Height          =   240
         Left            =   5895
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   2355
         Width           =   210
      End
      Begin VB.CheckBox SearchCutDeckPreciseAll 
         Caption         =   "Check1"
         Height          =   240
         Left            =   225
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   870
         Width           =   210
      End
      Begin VB.CheckBox SearchMoveCardAll 
         Caption         =   "Check1"
         Height          =   240
         Left            =   225
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1170
         Width           =   210
      End
      Begin VB.CheckBox SearchRunSingleCardsAll 
         Caption         =   "Check1"
         Height          =   240
         Left            =   225
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1470
         Width           =   210
      End
      Begin VB.CheckBox SearchRunSingleCardsInvAll 
         Caption         =   "Check1"
         Height          =   240
         Left            =   225
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   1770
         Width           =   210
      End
      Begin VB.CheckBox SearchShiftTopBlockAll 
         Caption         =   "Check1"
         Height          =   240
         Left            =   225
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   2070
         Width           =   210
      End
      Begin VB.CheckBox SearchShiftTopBlockInvAll 
         Caption         =   "Check1"
         Height          =   240
         Left            =   225
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   2370
         Width           =   210
      End
      Begin VB.CheckBox SearchOutFaroSpecialTopAll 
         Caption         =   "Check1"
         Height          =   240
         Left            =   3060
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   1455
         Width           =   210
      End
      Begin VB.CheckBox SearchOutFaroSpecialTopInvAll 
         Caption         =   "Check1"
         Height          =   240
         Left            =   3060
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   1755
         Width           =   210
      End
      Begin VB.CheckBox SearchOutFaroSpecialBottomAll 
         Caption         =   "Check1"
         Height          =   240
         Left            =   3060
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   2055
         Width           =   210
      End
      Begin VB.CheckBox SearchOutFaroSpecialBottomInvAll 
         Caption         =   "Check1"
         Height          =   240
         Left            =   3060
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   2355
         Width           =   210
      End
      Begin VB.CommandButton SearchUncheckAllButton 
         Caption         =   "Uncheck All"
         Height          =   315
         Left            =   7260
         TabIndex        =   37
         Top             =   225
         Width           =   1200
      End
      Begin VB.CheckBox SearchInFaroSpecialBottomInv 
         Caption         =   "In Faro Special Bottom Inv"
         Height          =   285
         Left            =   6120
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   2325
         Width           =   2415
      End
      Begin VB.CheckBox SearchInFaroSpecialBottom 
         Caption         =   "In Faro Special Bottom"
         Height          =   285
         Left            =   6120
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   2025
         Width           =   2190
      End
      Begin VB.CheckBox SearchInFaroSpecialTopInv 
         Caption         =   "In Faro Special Top Inv"
         Height          =   285
         Left            =   6120
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   1725
         Width           =   2190
      End
      Begin VB.CheckBox SearchInFaroSpecialTop 
         Caption         =   "In Faro Special Top"
         Height          =   285
         Left            =   6120
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   1425
         Width           =   2190
      End
      Begin VB.CheckBox SearchInFaroInverse 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Left            =   6120
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   1125
         Width           =   195
      End
      Begin VB.CheckBox SearchInFaro 
         Enabled         =   0   'False
         Height          =   285
         Left            =   6120
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   825
         Width           =   225
      End
      Begin VB.CheckBox SearchOutFaroSpecialBottomInv 
         Caption         =   "Out Faro Special Bottom Inv"
         Height          =   285
         Left            =   3285
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   2325
         Width           =   2445
      End
      Begin VB.CheckBox SearchOutFaroSpecialBottom 
         Caption         =   "Out Faro Special Bottom"
         Height          =   285
         Left            =   3285
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   2025
         Width           =   2190
      End
      Begin VB.CheckBox SearchOutFaroSpecialTopInv 
         Caption         =   "Out Faro Special Top Inv"
         Height          =   285
         Left            =   3285
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   1725
         Width           =   2190
      End
      Begin VB.CheckBox SearchOutFaroSpecialTop 
         Caption         =   "Out Faro Special Top"
         Height          =   285
         Left            =   3285
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   1425
         Width           =   2190
      End
      Begin VB.CheckBox SearchOutFaroInverse 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3285
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   1125
         Width           =   195
      End
      Begin VB.CheckBox SearchOutFaro 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3285
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   825
         Width           =   195
      End
      Begin VB.CheckBox SearchShiftTopBlockInv 
         Caption         =   "Shift Top Block Inv"
         Height          =   285
         Left            =   450
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   2340
         Width           =   2190
      End
      Begin VB.CheckBox SearchShiftTopBlock 
         Caption         =   "Shift Top Block"
         Height          =   285
         Left            =   450
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   2040
         Width           =   2190
      End
      Begin VB.CheckBox SearchRunSingleCardsInv 
         Caption         =   "Run Single Cards Inv"
         Height          =   285
         Left            =   450
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1740
         Width           =   2190
      End
      Begin VB.CheckBox SearchRunSingleCards 
         Caption         =   "Run Single Cards"
         Height          =   285
         Left            =   450
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1440
         Width           =   2190
      End
      Begin VB.CheckBox SearchMoveCard 
         Caption         =   "Move Card"
         Height          =   285
         Left            =   450
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1140
         Width           =   2190
      End
      Begin VB.CheckBox SearchCutDeckPrecise 
         Caption         =   "Cut Deck Precise"
         Height          =   285
         Left            =   450
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   840
         Width           =   2190
      End
      Begin VB.CheckBox SearchOutFaroAll 
         Caption         =   "     Out Faro"
         Height          =   240
         Left            =   3060
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   855
         Width           =   2085
      End
      Begin VB.CheckBox SearchOutFaroInverseAll 
         Caption         =   "     Out Faro Inverse"
         Height          =   240
         Left            =   3060
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   1155
         Width           =   2175
      End
      Begin VB.CheckBox SearchInFaroAll 
         Caption         =   "     In Faro"
         Height          =   240
         Left            =   5895
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   855
         Width           =   1965
      End
      Begin VB.CheckBox SearchInFaroInverseAll 
         Caption         =   "     In Faro Inverse"
         Height          =   240
         Left            =   5895
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   1155
         Width           =   1950
      End
      Begin VB.Label Label8 
         Caption         =   "All | Specific"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3075
         TabIndex        =   58
         Top             =   630
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "All | Specific"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5910
         TabIndex        =   57
         Top             =   630
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "All | Specific"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   56
         Top             =   630
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Identify Manipulations to Include in Search"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   135
         TabIndex        =   53
         Top             =   240
         Width           =   4395
      End
   End
   Begin VB.TextBox SearchLevelsTextBox 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   1455
      TabIndex        =   4
      Top             =   1170
      Width           =   585
   End
   Begin VB.CommandButton StartDeckButton 
      Height          =   315
      Left            =   1440
      TabIndex        =   0
      Top             =   105
      Width           =   585
   End
   Begin VB.CommandButton TransferToSession 
      Caption         =   "Transfer List to Sessions"
      Height          =   390
      Left            =   6210
      TabIndex        =   38
      Top             =   3480
      Width           =   2355
   End
   Begin VB.ListBox MatchListBox 
      Height          =   2790
      Left            =   6075
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   540
      Width           =   2535
   End
   Begin VB.Frame SearchMatchFrame 
      BorderStyle     =   0  'None
      Height          =   405
      Left            =   420
      TabIndex        =   59
      Top             =   765
      Width           =   4710
      Begin VB.OptionButton SearchMatchPartialOption 
         Caption         =   "Partial Deck Match"
         Height          =   240
         Left            =   1965
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   90
         Width           =   2400
      End
      Begin VB.OptionButton SearchMatchWholeOption 
         Caption         =   "Whole Deck Match"
         Height          =   240
         Left            =   135
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   90
         Value           =   -1  'True
         Width           =   1770
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   5820
      X2              =   165
      Y1              =   2355
      Y2              =   2355
   End
   Begin VB.Label ProgressLabel 
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
      Height          =   285
      Left            =   210
      TabIndex        =   63
      Top             =   2400
      Width           =   5580
   End
   Begin VB.Label ManipulationsLabel 
      Caption         =   "0"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   4140
      TabIndex        =   62
      Top             =   1245
      Width           =   1875
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      Caption         =   "Manipulations per Move"
      Height          =   285
      Left            =   2205
      TabIndex        =   61
      Top             =   1245
      Width           =   1785
   End
   Begin VB.Label TargetDeckLabel 
      Caption         =   "Set Start Deck"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   2130
      TabIndex        =   45
      Top             =   525
      Visible         =   0   'False
      Width           =   3840
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      Height          =   1125
      Left            =   165
      Top             =   1575
      Width           =   5670
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      Caption         =   "Elapsed Time"
      Height          =   285
      Left            =   240
      TabIndex        =   60
      Top             =   2100
      Width           =   1530
   End
   Begin VB.Image ContinueSearchToggle 
      Height          =   600
      Index           =   0
      Left            =   2055
      Picture         =   "frmSearch.frx":1CCA
      Top             =   2835
      Width           =   3750
   End
   Begin VB.Image ContinueSearchToggle 
      Height          =   600
      Index           =   1
      Left            =   2055
      Picture         =   "frmSearch.frx":2E3D
      Top             =   2835
      Width           =   3750
   End
   Begin VB.Label TimerResult 
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   1920
      TabIndex        =   55
      Top             =   2100
      Width           =   3855
   End
   Begin VB.Image SearchToggle 
      Height          =   600
      Index           =   3
      Left            =   225
      Picture         =   "frmSearch.frx":3E21
      Top             =   2835
      Visible         =   0   'False
      Width           =   1740
   End
   Begin VB.Image SearchToggle 
      Height          =   600
      Index           =   2
      Left            =   225
      Picture         =   "frmSearch.frx":5190
      Top             =   2835
      Visible         =   0   'False
      Width           =   1740
   End
   Begin VB.Image SearchToggle 
      Height          =   600
      Index           =   1
      Left            =   225
      Picture         =   "frmSearch.frx":64E2
      Top             =   2835
      Visible         =   0   'False
      Width           =   1740
   End
   Begin VB.Image SearchToggle 
      Height          =   600
      Index           =   0
      Left            =   225
      Picture         =   "frmSearch.frx":77FF
      Top             =   2835
      Width           =   1740
   End
   Begin VB.Label SearchTimeLabel 
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   1890
      TabIndex        =   48
      Top             =   1725
      Width           =   3855
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Estimated Possible Search Time "
      Height          =   465
      Left            =   270
      TabIndex        =   47
      Top             =   1620
      Width           =   1530
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Maximum Moves"
      Height          =   255
      Left            =   150
      TabIndex        =   46
      Top             =   1245
      Width           =   1230
   End
   Begin VB.Label StartDeckLabel 
      Caption         =   "Set Start Deck"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   2115
      TabIndex        =   44
      Top             =   150
      Visible         =   0   'False
      Width           =   3870
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Set Target Deck"
      Height          =   255
      Left            =   165
      TabIndex        =   43
      Top             =   525
      Width           =   1230
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Set Start Deck"
      Height          =   255
      Left            =   150
      TabIndex        =   42
      Top             =   135
      Width           =   1230
   End
   Begin VB.Label MatchFoundLabel 
      Alignment       =   2  'Center
      Caption         =   "Match Found!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   6045
      TabIndex        =   41
      Top             =   135
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      Height          =   690
      Left            =   165
      Top             =   2790
      Width           =   5685
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub AppendTrapFile()
Dim fso, txtfile
Set fso = CreateObject("Scripting.FileSystemObject")
Set txtfile = fso.OpenTextFile(TrapPathFinal, ForAppending)
txtfile.WriteLine ("Threshold: " & PartialMatchCounter)
For i% = 0 To MatchListBox.ListCount - 1
    txtfile.WriteLine (MatchListBox.List(i%))
Next i%
txtfile.WriteBlankLines (1)
txtfile.Close
End Sub

Private Sub ContinueSearchToggle_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
ContinueSearchToggle(0).Visible = False
ContinueSearchToggle(1).Visible = True
End Sub

Private Sub ContinueSearchToggle_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
ContinueSearchToggle(0).Visible = False
ContinueSearchToggle(1).Visible = False
SearchToggle(0).Visible = False
SearchToggle(1).Visible = False
SearchToggle(2).Visible = True
SearchToggle(3).Visible = False
SearchToggle(0).Refresh
SearchToggle(1).Refresh
SearchToggle(2).Refresh
SearchToggle(3).Refresh
SearchingMode = True
SearchContinueReady = 0
   
    'turn off search label
    MatchFoundLabel.Visible = False
  
    'disable the controls
    SearchItems.Enabled = False
    SearchLevelsTextBox.Enabled = False
    StartDeckButton.Enabled = False
    TargetDeckButton.Enabled = False
    TransferToSession.Enabled = False
    SearchCutDeckPrecise.Enabled = False
    SearchMoveCard.Enabled = False
    SearchRunSingleCards.Enabled = False
    SearchRunSingleCardsInv.Enabled = False
    SearchShiftTopBlock.Enabled = False
    SearchShiftTopBlockInv.Enabled = False
    SearchOutFaro.Enabled = False
    SearchOutFaroInverse.Enabled = False
    SearchOutFaroSpecialTop.Enabled = False
    SearchOutFaroSpecialTopInv.Enabled = False
    SearchOutFaroSpecialBottom.Enabled = False
    SearchOutFaroSpecialBottomInv.Enabled = False
    SearchInFaro.Enabled = False
    SearchInFaroInverse.Enabled = False
    SearchInFaroSpecialTop.Enabled = False
    SearchInFaroSpecialTopInv.Enabled = False
    SearchInFaroSpecialBottom.Enabled = False
    SearchInFaroSpecialBottomInv.Enabled = False
    SearchCutDeckPreciseAll.Enabled = False
    SearchMoveCardAll.Enabled = False
    SearchRunSingleCardsAll.Enabled = False
    SearchRunSingleCardsInvAll.Enabled = False
    SearchShiftTopBlockAll.Enabled = False
    SearchShiftTopBlockInvAll.Enabled = False
    SearchOutFaroAll.Enabled = False
    SearchOutFaroInverseAll.Enabled = False
    SearchOutFaroSpecialTopAll.Enabled = False
    SearchOutFaroSpecialTopInvAll.Enabled = False
    SearchOutFaroSpecialBottomAll.Enabled = False
    SearchOutFaroSpecialBottomInvAll.Enabled = False
    SearchInFaroAll.Enabled = False
    SearchInFaroInverseAll.Enabled = False
    SearchInFaroSpecialTopAll.Enabled = False
    SearchInFaroSpecialTopInvAll.Enabled = False
    SearchInFaroSpecialBottomAll.Enabled = False
    SearchInFaroSpecialBottomInvAll.Enabled = False
    SearchMatchFrame.Enabled = False
    'SearchMatchWholeOption.Enabled = False
    'SearchMatchPartialOption.Enabled = False
    
    'start the progress bar and timer
    SearchProgress.Value = 0
    SearchProgress.Max = 100
    SearchProgress.Visible = True
    
    'start timer -- this is for my debugging and procedure time comparisons only
    SearchStartTime = Timer
    
    'RUN THE SEARCH
    SearchCurrentLevel = SearchCurrentLevelRestart
    ShowProgress
    SearchTransactionRun
    
    'check for a threshold trap pause (allow continue)
    If PartialMatchFound = 1 Then
        SearchToggle(0).Visible = True
        SearchToggle(1).Visible = False
        SearchToggle(2).Visible = False
        SearchToggle(3).Visible = False
        ContinueSearchToggle(0).Visible = True
        ContinueSearchToggle(1).Visible = False
        SearchContinueReady = 1
        SearchSaved = 0
        SearchCurrentLevelRestart = SearchCurrentLevel
        PartialMatchFound = 0
    End If
    
    
    'stop timer
    If Timer < SearchStartTime Then
        SearchStartTime = SearchStartTime - 86400
        'error condition is test crosses midnight
    End If
    SearchElapsedTime = SearchElapsedTime + Timer - SearchStartTime
    'TimerResult.Caption = Int(SearchElapsedTime / 60) & " min. " & _
                Round((SearchElapsedTime Mod 60), 1) & " sec."
    ShowElapsedTime
    
    're-enable the controls
    SearchItems.Enabled = True
    SearchLevelsTextBox.Enabled = True
    StartDeckButton.Enabled = True
    TargetDeckButton.Enabled = True
    TransferToSession.Enabled = True
    SearchCutDeckPrecise.Enabled = True
    SearchMoveCard.Enabled = True
    SearchRunSingleCards.Enabled = True
    SearchRunSingleCardsInv.Enabled = True
    SearchShiftTopBlock.Enabled = True
    SearchShiftTopBlockInv.Enabled = True
    SearchOutFaro.Enabled = False
    SearchOutFaroInverse.Enabled = False
    SearchOutFaroSpecialTop.Enabled = True
    SearchOutFaroSpecialTopInv.Enabled = True
    SearchOutFaroSpecialBottom.Enabled = True
    SearchOutFaroSpecialBottomInv.Enabled = True
    SearchInFaro.Enabled = False
    SearchInFaroInverse.Enabled = False
    SearchInFaroSpecialTop.Enabled = True
    SearchInFaroSpecialTopInv.Enabled = True
    SearchInFaroSpecialBottom.Enabled = True
    SearchInFaroSpecialBottomInv.Enabled = True
    SearchCutDeckPreciseAll.Enabled = True
    SearchMoveCardAll.Enabled = True
    SearchRunSingleCardsAll.Enabled = True
    SearchRunSingleCardsInvAll.Enabled = True
    SearchShiftTopBlockAll.Enabled = True
    SearchShiftTopBlockInvAll.Enabled = True
    SearchOutFaroAll.Enabled = True
    SearchOutFaroInverseAll.Enabled = True
    SearchOutFaroSpecialTopAll.Enabled = True
    SearchOutFaroSpecialTopInvAll.Enabled = True
    SearchOutFaroSpecialBottomAll.Enabled = True
    SearchOutFaroSpecialBottomInvAll.Enabled = True
    SearchInFaroAll.Enabled = True
    SearchInFaroInverseAll.Enabled = True
    SearchInFaroSpecialTopAll.Enabled = True
    SearchInFaroSpecialTopInvAll.Enabled = True
    SearchInFaroSpecialBottomAll.Enabled = True
    SearchInFaroSpecialBottomInvAll.Enabled = True
    SearchingMode = False
    SearchToggle(0).Visible = True
    SearchToggle(1).Visible = False
    SearchToggle(2).Visible = False
    SearchToggle(3).Visible = False
    SearchProgress.Visible = False
    'SearchMatchWholeOption.Enabled = True
    'SearchMatchPartialOption.Enabled = True
    SearchMatchFrame.Enabled = True
End Sub

Private Sub Form_Load()
SuspendTrapFinal = True
ManipulationsPerSecond = 4000
SpeedMod = 750
MatchFound = 0
NoMatchFound = 0
ThresholdMatchCards = 0
TrapThreshold = False
WholeDeckMatchSet = True
PartialMatchFound = 0
PartialMatchCounter = 0
SearchContinueReady = 0
SearchSpecialCancel = False
SearchProgress.Visible = False
SearchingMode = False
'LoadManipulations
SearchToggle(0).Visible = True
SearchToggle(1).Visible = False
SearchToggle(2).Visible = False
SearchToggle(3).Visible = False
ContinueSearchToggle(0).Visible = False
ContinueSearchToggle(1).Visible = False
SearchMatchStartCard = 1
SearchMatchEndCard = 52
SearchLevelsTextBox.Text = 1
'set the default level depth to 1
SearchTimeLabel.Caption = Empty
ManipulationsLabel.Caption = "0"
SearchCounter = 1
'set the counter to the initial position for it's use in the Manipulations array
SearchCounterMax = 0
'set the max value for SearchCounter.  It will be increased based on what is checked
SearchCDPMin = 0
SearchCDPMax = 0
SearchCDP = 0
SearchRSC = 0
SearchRSCMin = 0
SearchRSCMax = 0
SearchRSCR = 0
SearchRSCRMin = 0
SearchRSCRMax = 0
SearchMC1Min = 0
SearchMC2Min = 0
SearchMC1Max = 0
SearchMC2Max = 0
SearchMC = 0
SearchSTB1Min = 0
SearchSTB2Min = 0
SearchSTB1Max = 0
SearchSTB2Max = 0
SearchSTB = 0
SearchSTBR1Min = 0
SearchSTBR2Min = 0
SearchSTBR1Max = 0
SearchSTBR2Max = 0
SearchSTBR = 0
SearchOFST1Min = 0
SearchOFST2Min = 0
SearchOFST1Max = 0
SearchOFST2Max = 0
SearchOFST = 0
SearchOFSTR1Min = 0
SearchOFSTR2Min = 0
SearchOFSTR1Max = 0
SearchOFSTR2Max = 0
SearchOFSTR = 0
SearchOFSB1Min = 0
SearchOFSB2Min = 0
SearchOFSB1Max = 0
SearchOFSB2Max = 0
SearchOFSB = 0
SearchOFSBR1Min = 0
SearchOFSBR2Min = 0
SearchOFSBR1Max = 0
SearchOFSBR2Max = 0
SearchOFSBR = 0
SearchIFST1Min = 0
SearchIFST2Min = 0
SearchIFST1Max = 0
SearchIFST2Max = 0
SearchIFST = 0
SearchIFSTR1Min = 0
SearchIFSTR2Min = 0
SearchIFSTR1Max = 0
SearchIFSTR2Max = 0
SearchIFSTR = 0
SearchIFSB1Min = 0
SearchIFSB2Min = 0
SearchIFSB1Max = 0
SearchIFSB2Max = 0
SearchIFSB = 0
SearchIFSBR1Min = 0
SearchIFSBR2Min = 0
SearchIFSBR1Max = 0
SearchIFSBR2Max = 0
SearchIFSBR = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim Msg   ' Declare variable.
    ' Set the message text.
    Msg = ""
    If SearchSaved = 0 Then
        Msg = Msg & "You have unsaved Search activity still present." & Chr(13)
        Msg = Msg & "Closing this window will clear your Search progress." & Chr(13)
        Msg = Msg & Chr(13) & "Do you really want to close this window?" & Chr(13)
        ' If user clicks the No button, stop QueryUnload.
        If MsgBox(Msg, vbQuestion + vbYesNo, Me.Caption) = vbNo Then
            Cancel = True
        Else
            SearchSaved = 1
            SearchingMode = False
            frmMain.mnuSearch.Checked = False
        End If
    Else
        frmMain.mnuSearch.Checked = False
    End If
End Sub

Public Sub LoadManipulations()
SearchCounter = 1
ReDim Manipulations(2, SearchCounterMax) As Variant
If SearchCutDeckPreciseAll.Value = 1 Then
    For i% = 1 To 51
        Manipulations(1, SearchCounter) = "CutDeckPrecise(" & i% & ", X)"
        Manipulations(2, SearchCounter) = "CutDeckPrecise(" & 52 - i% & ", X)"
        SearchCounter = SearchCounter + 1
    Next i%
End If
If SearchCutDeckPrecise.Value = 1 Then
    For i% = SearchCDPMin To SearchCDPMax
        Manipulations(1, SearchCounter) = "CutDeckPrecise(" & i% & ", X)"
        Manipulations(2, SearchCounter) = "CutDeckPrecise(" & 52 - i% & ", X)"
            'correct Manipulations(2... if SearchCDPMax=52
            If i% = 52 Then
                Manipulations(2, SearchCounter) = "CutDeckPrecise(" & 52 & ", X)"
            End If
        SearchCounter = SearchCounter + 1
    Next i%
End If
If SearchMoveCardAll.Value = 1 Then
    For i% = 1 To 52
        For j% = 1 To 52
            If i% <> j% Then
                Manipulations(1, SearchCounter) = "MoveCard(" & i% & ", " & j% & ")"
                Manipulations(2, SearchCounter) = "MoveCard(" & j% & ", " & i% & ")"
                SearchCounter = SearchCounter + 1
            End If
        Next j%
    Next i%
End If
If SearchMoveCard.Value = 1 Then
    For i% = SearchMC1Min To SearchMC1Max
        For j% = SearchMC2Min To SearchMC2Max
            'If i% <> j% And i% + j% < 53 Then
            If i% <> j% Then
                Manipulations(1, SearchCounter) = "MoveCard(" & i% & ", " & j% & ")"
                Manipulations(2, SearchCounter) = "MoveCard(" & j% & ", " & i% & ")"
                SearchCounter = SearchCounter + 1
            End If
        Next j%
    Next i%
End If
If SearchRunSingleCardsAll.Value = 1 Then
    For i% = 1 To 52
        Manipulations(1, SearchCounter) = "RunSingleCards(" & i% & ")"
        Manipulations(2, SearchCounter) = "RunSingleCardsInverse(" & i% & ")"
        SearchCounter = SearchCounter + 1
    Next i%
End If
If SearchRunSingleCards.Value = 1 Then
    For i% = SearchRSCMin To SearchRSCMax
        Manipulations(1, SearchCounter) = "RunSingleCards(" & i% & ")"
        Manipulations(2, SearchCounter) = "RunSingleCardsInverse(" & i% & ")"
        SearchCounter = SearchCounter + 1
    Next i%
End If
If SearchRunSingleCardsInvAll.Value = 1 Then
    For i% = 1 To 52
        Manipulations(1, SearchCounter) = "RunSingleCardsInverse(" & i% & ")"
        Manipulations(2, SearchCounter) = "RunSingleCards(" & i% & ")"
        SearchCounter = SearchCounter + 1
    Next i%
End If
If SearchRunSingleCardsInv.Value = 1 Then
    For i% = SearchRSCRMin To SearchRSCRMax
        Manipulations(1, SearchCounter) = "RunSingleCardsInverse(" & i% & ")"
        Manipulations(2, SearchCounter) = "RunSingleCards(" & i% & ")"
        SearchCounter = SearchCounter + 1
    Next i%
End If
If SearchShiftTopBlockAll.Value = 1 Then
    For i% = 1 To 51
        For j% = 1 To 52 - i%
                Manipulations(1, SearchCounter) = "ShiftTopBlock(" & i% & ", " & j% & ")"
                Manipulations(2, SearchCounter) = "ShiftTopBlockInverse(" & i% & ", " & j% & ")"
                SearchCounter = SearchCounter + 1
        Next j%
    Next i%
End If
If SearchShiftTopBlock.Value = 1 Then
    For i% = SearchSTB1Min To SearchSTB1Max
        For j% = SearchSTB2Min To SearchSTB2Max
                If i% + j% < 53 Then
                    Manipulations(1, SearchCounter) = "ShiftTopBlock(" & i% & ", " & j% & ")"
                    Manipulations(2, SearchCounter) = "ShiftTopBlockInverse(" & i% & ", " & j% & ")"
                    SearchCounter = SearchCounter + 1
                End If
        Next j%
    Next i%
End If
If SearchShiftTopBlockInvAll.Value = 1 Then
    For i% = 1 To 51
        For j% = 1 To 52 - i%
                Manipulations(1, SearchCounter) = "ShiftTopBlockInverse(" & i% & ", " & j% & ")"
                Manipulations(2, SearchCounter) = "ShiftTopBlock(" & i% & ", " & j% & ")"
                SearchCounter = SearchCounter + 1
        Next j%
    Next i%
End If
If SearchShiftTopBlockInv.Value = 1 Then
    For i% = SearchSTBR1Min To SearchSTBR1Max
        For j% = SearchSTBR2Min To SearchSTBR2Max
                If i% + j% < 53 Then
                    Manipulations(1, SearchCounter) = "ShiftTopBlockInverse(" & i% & ", " & j% & ")"
                    Manipulations(2, SearchCounter) = "ShiftTopBlock(" & i% & ", " & j% & ")"
                    SearchCounter = SearchCounter + 1
                End If
        Next j%
    Next i%
End If
If SearchOutFaroAll.Value = 1 Then
    Manipulations(1, SearchCounter) = "OutFaro"
    Manipulations(2, SearchCounter) = "InverseOutFaro"
    SearchCounter = SearchCounter + 1
End If
If SearchOutFaroInverseAll.Value = 1 Then
    Manipulations(1, SearchCounter) = "InverseOutFaro"
    Manipulations(2, SearchCounter) = "OutFaro"
    SearchCounter = SearchCounter + 1
End If
If SearchOutFaroSpecialTopAll.Value = 1 Then
    For i% = 1 To 51
        For j% = 1 To 52 - i%
                Manipulations(1, SearchCounter) = "OutFaroSpecialTop(" & i% & ", " & j% & ")"
                Manipulations(2, SearchCounter) = "InverseOutFaroSpecialTop(" & i% & ", " & j% & ")"
                SearchCounter = SearchCounter + 1
        Next j%
    Next i%
End If
If SearchOutFaroSpecialTop.Value = 1 Then
    For i% = SearchOFST1Min To SearchOFST1Max
        For j% = SearchOFST2Min To SearchOFST2Max
            If i% + j% < 53 Then
                Manipulations(1, SearchCounter) = "OutFaroSpecialTop(" & i% & ", " & j% & ")"
                Manipulations(2, SearchCounter) = "InverseOutFaroSpecialTop(" & i% & ", " & j% & ")"
                SearchCounter = SearchCounter + 1
            End If
        Next j%
    Next i%
End If
If SearchOutFaroSpecialTopInvAll.Value = 1 Then
    For i% = 1 To 51
        For j% = 1 To 52 - i%
                Manipulations(1, SearchCounter) = "InverseOutFaroSpecialTop(" & i% & ", " & j% & ")"
                Manipulations(2, SearchCounter) = "OutFaroSpecialTop(" & i% & ", " & j% & ")"
                SearchCounter = SearchCounter + 1
        Next j%
    Next i%
End If
If SearchOutFaroSpecialTopInv.Value = 1 Then
    For i% = SearchOFSTR1Min To SearchOFSTR1Max
        For j% = SearchOFSTR2Min To SearchOFSTR2Max
            If i% + j% < 53 Then
                Manipulations(1, SearchCounter) = "InverseOutFaroSpecialTop(" & i% & ", " & j% & ")"
                Manipulations(2, SearchCounter) = "OutFaroSpecialTop(" & i% & ", " & j% & ")"
                SearchCounter = SearchCounter + 1
            End If
        Next j%
    Next i%
End If
If SearchOutFaroSpecialBottomAll.Value = 1 Then
    For i% = 1 To 51
        For j% = 1 To 52 - i%
                Manipulations(1, SearchCounter) = "OutFaroSpecialBottom(" & i% & ", " & j% & ")"
                Manipulations(2, SearchCounter) = "InverseOutFaroSpecialBottom(" & i% & ", " & j% & ")"
                SearchCounter = SearchCounter + 1
        Next j%
    Next i%
End If
If SearchOutFaroSpecialBottom.Value = 1 Then
    For i% = SearchOFSB1Min To SearchOFSB1Max
        For j% = SearchOFSB2Min To SearchOFSB2Max
            If i% + j% < 53 Then
                Manipulations(1, SearchCounter) = "OutFaroSpecialBottom(" & i% & ", " & j% & ")"
                Manipulations(2, SearchCounter) = "InverseOutFaroSpecialBottom(" & i% & ", " & j% & ")"
                SearchCounter = SearchCounter + 1
            End If
        Next j%
    Next i%
End If
If SearchOutFaroSpecialBottomInvAll.Value = 1 Then
    For i% = 1 To 51
        For j% = 1 To 52 - i%
                Manipulations(1, SearchCounter) = "InverseOutFaroSpecialBottom(" & i% & ", " & j% & ")"
                Manipulations(2, SearchCounter) = "OutFaroSpecialBottom(" & i% & ", " & j% & ")"
                SearchCounter = SearchCounter + 1
        Next j%
    Next i%
End If
If SearchOutFaroSpecialBottomInv.Value = 1 Then
    For i% = SearchOFSBR1Min To SearchOFSBR1Max
        For j% = SearchOFSBR2Min To SearchOFSBR2Max
            If i% + j% < 53 Then
                Manipulations(1, SearchCounter) = "InverseOutFaroSpecialBottom(" & i% & ", " & j% & ")"
                Manipulations(2, SearchCounter) = "OutFaroSpecialBottom(" & i% & ", " & j% & ")"
                SearchCounter = SearchCounter + 1
            End If
        Next j%
    Next i%
End If
If SearchInFaroAll.Value = 1 Then
    Manipulations(1, SearchCounter) = "InFaro"
    Manipulations(2, SearchCounter) = "InverseInFaro"
    SearchCounter = SearchCounter + 1
End If
If SearchInFaroInverseAll.Value = 1 Then
    Manipulations(1, SearchCounter) = "InverseInFaro"
    Manipulations(2, SearchCounter) = "InFaro"
    SearchCounter = SearchCounter + 1
End If
If SearchInFaroSpecialTopAll.Value = 1 Then
    For i% = 1 To 51
        For j% = 1 To 52 - i%
                Manipulations(1, SearchCounter) = "InFaroSpecialTop(" & i% & ", " & j% & ")"
                Manipulations(2, SearchCounter) = "InverseInFaroSpecialTop(" & i% & ", " & j% & ")"
                SearchCounter = SearchCounter + 1
        Next j%
    Next i%
End If
If SearchInFaroSpecialTop.Value = 1 Then
    For i% = SearchIFST1Min To SearchIFST1Max
        For j% = SearchIFST2Min To SearchIFST2Max
            If i% + j% < 53 Then
                Manipulations(1, SearchCounter) = "InFaroSpecialTop(" & i% & ", " & j% & ")"
                Manipulations(2, SearchCounter) = "InverseInFaroSpecialTop(" & i% & ", " & j% & ")"
                SearchCounter = SearchCounter + 1
            End If
        Next j%
    Next i%
End If
If SearchInFaroSpecialTopInvAll.Value = 1 Then
    For i% = 1 To 51
        For j% = 1 To 52 - i%
                Manipulations(1, SearchCounter) = "InverseInFaroSpecialTop(" & i% & ", " & j% & ")"
                Manipulations(2, SearchCounter) = "InFaroSpecialTop(" & i% & ", " & j% & ")"
                SearchCounter = SearchCounter + 1
        Next j%
    Next i%
End If
If SearchInFaroSpecialTopInv.Value = 1 Then
    For i% = SearchIFSTR1Min To SearchIFSTR1Max
        For j% = SearchIFSTR2Min To SearchIFSTR2Max
            If i% + j% < 53 Then
                Manipulations(1, SearchCounter) = "InverseInFaroSpecialTop(" & i% & ", " & j% & ")"
                Manipulations(2, SearchCounter) = "InFaroSpecialTop(" & i% & ", " & j% & ")"
                SearchCounter = SearchCounter + 1
            End If
        Next j%
    Next i%
End If
If SearchInFaroSpecialBottomAll.Value = 1 Then
    For i% = 1 To 51
        For j% = 1 To 52 - i%
                Manipulations(1, SearchCounter) = "InFaroSpecialBottom(" & i% & ", " & j% & ")"
                Manipulations(2, SearchCounter) = "InverseInFaroSpecialBottom(" & i% & ", " & j% & ")"
                SearchCounter = SearchCounter + 1
        Next j%
    Next i%
End If
If SearchInFaroSpecialBottom.Value = 1 Then
    For i% = SearchIFSB1Min To SearchIFSB1Max
        For j% = SearchIFSB2Min To SearchIFSB2Max
            If i% + j% < 53 Then
                Manipulations(1, SearchCounter) = "InFaroSpecialBottom(" & i% & ", " & j% & ")"
                Manipulations(2, SearchCounter) = "InverseInFaroSpecialBottom(" & i% & ", " & j% & ")"
                SearchCounter = SearchCounter + 1
            End If
        Next j%
    Next i%
End If
If SearchInFaroSpecialBottomInvAll.Value = 1 Then
    For i% = 1 To 51
        For j% = 1 To 52 - i%
                Manipulations(1, SearchCounter) = "InverseInFaroSpecialBottom(" & i% & ", " & j% & ")"
                Manipulations(2, SearchCounter) = "InFaroSpecialBottom(" & i% & ", " & j% & ")"
                SearchCounter = SearchCounter + 1
        Next j%
    Next i%
End If
If SearchInFaroSpecialBottomInv.Value = 1 Then
    For i% = SearchIFSBR1Min To SearchIFSBR1Max
        For j% = SearchIFSBR2Min To SearchIFSBR2Max
            If i% + j% < 53 Then
                Manipulations(1, SearchCounter) = "InverseInFaroSpecialBottom(" & i% & ", " & j% & ")"
                Manipulations(2, SearchCounter) = "InFaroSpecialBottom(" & i% & ", " & j% & ")"
                SearchCounter = SearchCounter + 1
            End If
        Next j%
    Next i%
End If
'For i% = 1 To SearchCounterMax
'    Debug.Print Manipulations(1, i%)
'Next i%
End Sub



Private Sub SearchCutDeckPrecise_Click()
If SearchFileLoading Then
    If SearchCutDeckPrecise.Value = 1 Then
        SearchCounterMax = SearchCounterMax + SearchCDP
    End If
    Exit Sub
End If
If SearchCutDeckPrecise.Value = 1 Then
    SearchCutDeckPreciseAll.Value = 0
End If
If SearchCutDeckPrecise.Value = 0 Then
    If SearchSpecialCancel Then
        Exit Sub
    End If
    SearchCounterMax = SearchCounterMax - SearchCDP
Else
    SearchSpecialRangeOne ("CutDeckPrecise")
    If SearchSpecialCancel Then
        Exit Sub
    End If
    SearchCDP = SearchCDPMax - SearchCDPMin + 1
    SearchCounterMax = SearchCounterMax + SearchCDP
End If
ShowEstimatedTime
End Sub

Private Sub SearchCutDeckPrecise_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
SearchSpecialCancel = False
SearchDisableContinueCheck
If SearchContinueReady = 1 Then
    Exit Sub
End If
End Sub

Private Sub SearchCutDeckPreciseAll_Click()
If SearchFileLoading Then
    If SearchCutDeckPreciseAll.Value = 1 Then
        SearchCounterMax = SearchCounterMax + 51
    End If
    Exit Sub
End If
If SearchCutDeckPreciseAll.Value = 1 Then
    SearchCutDeckPrecise.Value = 0
End If
If SearchCutDeckPreciseAll.Value = 0 Then
    SearchCounterMax = SearchCounterMax - 51
Else
    SearchCounterMax = SearchCounterMax + 51
End If
ShowEstimatedTime
End Sub

Private Sub SearchCutDeckPreciseAll_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
SearchDisableContinueCheck
If SearchContinueReady = 1 Then
    Exit Sub
End If
End Sub

Private Sub SearchInFaroAll_Click()
If SearchFileLoading Then
    If SearchInFaroAll.Value = 1 Then
        SearchCounterMax = SearchCounterMax + 1
    End If
    Exit Sub
End If
If SearchInFaroAll.Value = 0 Then
    SearchCounterMax = SearchCounterMax - 1
Else
    SearchCounterMax = SearchCounterMax + 1
End If
ShowEstimatedTime
End Sub

Private Sub SearchInFaroAll_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
SearchDisableContinueCheck
If SearchContinueReady = 1 Then
    Exit Sub
End If

End Sub


Private Sub SearchInFaroInverseAll_Click()
If SearchFileLoading Then
    If SearchInFaroInverse.Value = 1 Then
        SearchCounterMax = SearchCounterMax + 1
    End If
    Exit Sub
End If
If SearchInFaroInverseAll.Value = 0 Then
    SearchCounterMax = SearchCounterMax - 1
Else
    SearchCounterMax = SearchCounterMax + 1
End If
ShowEstimatedTime
End Sub

Private Sub SearchInFaroInverseAll_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
SearchDisableContinueCheck
If SearchContinueReady = 1 Then
    Exit Sub
End If

End Sub

Private Sub SearchInFaroSpecialBottom_Click()
If SearchFileLoading Then
    If SearchInFaroSpecialBottom.Value = 1 Then
        SearchCounterMax = SearchCounterMax + SearchIFSB
    End If
    Exit Sub
End If
If SearchInFaroSpecialBottom.Value = 1 Then
    SearchInFaroSpecialBottomAll.Value = 0
End If
If SearchInFaroSpecialBottom.Value = 0 Then
    If SearchSpecialCancel Then
        Exit Sub
    End If
    SearchCounterMax = SearchCounterMax - SearchIFSB
Else
    SearchSpecialRangeTwo ("InFaroSpecialBottom")
    SearchInFaroSpecialBottomAll.Value = 0
    SearchIFSB = 0
    If SearchSpecialCancel Then
        Exit Sub
    End If
    For i% = SearchIFSB1Min To SearchIFSB1Max
        For j% = SearchIFSB2Min To SearchIFSB2Max
            If i% + j% < 53 Then
                SearchIFSB = SearchIFSB + 1
            End If
        Next j%
    Next i%
    SearchCounterMax = SearchCounterMax + SearchIFSB
    If SearchIFSB = 0 Then
        MsgBox "Your range values do not allow for any manipulations to be included." & Chr(13) _
            & "The sum of the values taken from the ranges can not be greater than 52."
    End If
    SearchInFaroSpecialBottom.Value = 0
End If
ShowEstimatedTime
End Sub

Private Sub SearchInFaroSpecialBottom_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
SearchSpecialCancel = False
SearchDisableContinueCheck
If SearchContinueReady = 1 Then
    Exit Sub
End If

End Sub

Private Sub SearchInFaroSpecialBottomAll_Click()
If SearchFileLoading Then
    If SearchInFaroSpecialBottomAll.Value = 1 Then
        SearchCounterMax = SearchCounterMax + 1326
    End If
    Exit Sub
End If
If SearchInFaroSpecialBottomAll.Value = 1 Then
    SearchInFaroSpecialBottom.Value = 0
End If
If SearchInFaroSpecialBottomAll.Value = 0 Then
    SearchCounterMax = SearchCounterMax - 1326
Else
    SearchCounterMax = SearchCounterMax + 1326
End If
ShowEstimatedTime
End Sub

Private Sub SearchInFaroSpecialBottomAll_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
SearchDisableContinueCheck
If SearchContinueReady = 1 Then
    Exit Sub
End If

End Sub

Private Sub SearchInFaroSpecialBottomInv_Click()
If SearchFileLoading Then
    If SearchInFaroSpecialBottomInv.Value = 1 Then
        SearchCounterMax = SearchCounterMax + SearchIFSBR
    End If
    Exit Sub
End If
If SearchInFaroSpecialBottomInv.Value = 1 Then
    SearchInFaroSpecialBottomInvAll.Value = 0
End If
If SearchInFaroSpecialBottomInv.Value = 0 Then
    If SearchSpecialCancel Then
        Exit Sub
    End If
    SearchCounterMax = SearchCounterMax - SearchIFSBR
Else
    SearchSpecialRangeTwo ("InFaroSpecialBottomInv")
    SearchInFaroSpecialBottomInvAll.Value = 0
    SearchIFSBR = 0
    If SearchSpecialCancel Then
        Exit Sub
    End If
    For i% = SearchIFSBR1Min To SearchIFSBR1Max
        For j% = SearchIFSBR2Min To SearchIFSBR2Max
            If i% + j% < 53 Then
                SearchIFSBR = SearchIFSBR + 1
            End If
        Next j%
    Next i%
    SearchCounterMax = SearchCounterMax + SearchIFSBR
    If SearchIFSBR = 0 Then
        MsgBox "Your range values do not allow for any manipulations to be included." & Chr(13) _
            & "The sum of the values taken from the ranges can not be greater than 52."
    End If
    SearchInFaroSpecialBottomInv.Value = 0
End If
ShowEstimatedTime
End Sub

Private Sub SearchInFaroSpecialBottomInv_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
SearchSpecialCancel = False
SearchDisableContinueCheck
If SearchContinueReady = 1 Then
    Exit Sub
End If

End Sub

Private Sub SearchInFaroSpecialBottomInvAll_Click()
If SearchFileLoading Then
    If SearchInFaroSpecialBottomInvAll.Value = 1 Then
        SearchCounterMax = SearchCounterMax + 1326
    End If
    Exit Sub
End If
If SearchInFaroSpecialBottomInvAll.Value = 1 Then
    SearchInFaroSpecialBottomInv.Value = 0
End If
If SearchInFaroSpecialBottomInvAll.Value = 0 Then
    SearchCounterMax = SearchCounterMax - 1326
Else
    SearchCounterMax = SearchCounterMax + 1326
End If
ShowEstimatedTime
End Sub

Private Sub SearchInFaroSpecialBottomInvAll_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
SearchDisableContinueCheck
If SearchContinueReady = 1 Then
    Exit Sub
End If

End Sub

Private Sub SearchInFaroSpecialTop_Click()
If SearchFileLoading Then
    If SearchInFaroSpecialTop.Value = 1 Then
        SearchCounterMax = SearchCounterMax + SearchIFST
    End If
    Exit Sub
End If
If SearchInFaroSpecialTop.Value = 1 Then
    SearchInFaroSpecialTopAll.Value = 0
End If
If SearchInFaroSpecialTop.Value = 0 Then
    If SearchSpecialCancel Then
        Exit Sub
    End If
    SearchCounterMax = SearchCounterMax - SearchIFST
Else
    SearchSpecialRangeTwo ("InFaroSpecialTop")
    SearchInFaroSpecialTopAll.Value = 0
    SearchIFST = 0
    If SearchSpecialCancel Then
        Exit Sub
    End If
    For i% = SearchIFST1Min To SearchIFST1Max
        For j% = SearchIFST2Min To SearchIFST2Max
            If i% + j% < 53 Then
                SearchIFST = SearchIFST + 1
            End If
        Next j%
    Next i%
    SearchCounterMax = SearchCounterMax + SearchIFST
    If SearchIFST = 0 Then
        MsgBox "Your range values do not allow for any manipulations to be included." & Chr(13) _
            & "The sum of the values taken from the ranges can not be greater than 52."
    End If
    SearchInFaroSpecialTop.Value = 0
End If
ShowEstimatedTime
End Sub

Private Sub SearchInFaroSpecialTop_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
SearchSpecialCancel = False
SearchDisableContinueCheck
If SearchContinueReady = 1 Then
    Exit Sub
End If

End Sub

Private Sub SearchInFaroSpecialTopAll_Click()
If SearchFileLoading Then
    If SearchInFaroSpecialTopAll.Value = 1 Then
        SearchCounterMax = SearchCounterMax + 1326
    End If
    Exit Sub
End If
If SearchInFaroSpecialTopAll.Value = 1 Then
    SearchInFaroSpecialTop.Value = 0
End If
If SearchInFaroSpecialTopAll.Value = 0 Then
    SearchCounterMax = SearchCounterMax - 1326
Else
    SearchCounterMax = SearchCounterMax + 1326
End If
ShowEstimatedTime
End Sub

Private Sub SearchInFaroSpecialTopAll_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
SearchDisableContinueCheck
If SearchContinueReady = 1 Then
    Exit Sub
End If

End Sub

Private Sub SearchInFaroSpecialTopInv_Click()
If SearchFileLoading Then
    If SearchInFaroSpecialTopInv.Value = 1 Then
        SearchCounterMax = SearchCounterMax + SearchIFSTR
    End If
    Exit Sub
End If
If SearchInFaroSpecialTopInv.Value = 1 Then
    SearchInFaroSpecialTopInvAll.Value = 0
End If
If SearchInFaroSpecialTopInv.Value = 0 Then
    If SearchSpecialCancel Then
        Exit Sub
    End If
    SearchCounterMax = SearchCounterMax - SearchIFSTR
Else
    SearchSpecialRangeTwo ("InFaroSpecialTopInv")
    SearchInFaroSpecialTopInvAll.Value = 0
    SearchIFSTR = 0
    If SearchSpecialCancel Then
        Exit Sub
    End If
    For i% = SearchIFSTR1Min To SearchIFSTR1Max
        For j% = SearchIFSTR2Min To SearchIFSTR2Max
            If i% + j% < 53 Then
                SearchIFSTR = SearchIFSTR + 1
            End If
        Next j%
    Next i%
    SearchCounterMax = SearchCounterMax + SearchIFSTR
    If SearchIFSTR = 0 Then
        MsgBox "Your range values do not allow for any manipulations to be included." & Chr(13) _
            & "The sum of the values taken from the ranges can not be greater than 52."
    End If
    SearchInFaroSpecialTopInv.Value = 0
End If
ShowEstimatedTime
End Sub

Private Sub SearchInFaroSpecialTopInv_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
SearchSpecialCancel = False
SearchDisableContinueCheck
If SearchContinueReady = 1 Then
    Exit Sub
End If

End Sub

Private Sub SearchInFaroSpecialTopInvAll_Click()
If SearchFileLoading Then
    If SearchInFaroSpecialTopInvAll.Value = 1 Then
        SearchCounterMax = SearchCounterMax + 1326
    End If
    Exit Sub
End If
If SearchInFaroSpecialTopInvAll.Value = 1 Then
    SearchInFaroSpecialTopInv.Value = 0
End If
If SearchInFaroSpecialTopInvAll.Value = 0 Then
    SearchCounterMax = SearchCounterMax - 1326
Else
    SearchCounterMax = SearchCounterMax + 1326
End If
ShowEstimatedTime
End Sub



Private Sub SearchInFaroSpecialTopInvAll_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
SearchDisableContinueCheck
If SearchContinueReady = 1 Then
    Exit Sub
End If

End Sub

Private Sub SearchLevelsTextBox_Change()
If SearchFileLoading Then
    Exit Sub
End If
'If SearchLevelsTextBox.Text <> Empty And
If (Not IsNumeric(SearchLevelsTextBox.Text) Or _
    Val(SearchLevelsTextBox.Text) < 1 Or _
    Val(SearchLevelsTextBox.Text) > 26) Then
    SearchLevelsTextBox.Text = Empty
    MsgBox "Please enter a valid search level (1 to 26)" & Chr(13) _
        & "in the 'Search Levels' Input Box"
    SearchLevelsTextBox.SetFocus
    Exit Sub
End If
If Val(SearchLevelsTextBox.Text) - Int(Val(SearchLevelsTextBox.Text)) > 0 Then
    SearchLevelsTextBox.Text = Empty
    MsgBox "Please enter a valid search level (1 to 26)" & Chr(13) _
        & "in the 'Search Levels' Input Box" & Chr(13) & Chr(13) & _
        "Your entry must be an integer (no fractional portion)."
    SearchLevelsTextBox.SetFocus
    Exit Sub
End If
ShowEstimatedTime
End Sub

Private Sub SearchLevelsTextBox_GotFocus()
If SearchFileLoading Then
    Exit Sub
End If
SearchDisableContinueCheck
If SearchContinueReady = 1 Then
    Exit Sub
End If
End Sub

Private Sub SearchLevelsTextBox_LostFocus()
If SearchFileLoading Then
    Exit Sub
End If
'If SearchLevelsTextBox.Text <> Empty And
If (Not IsNumeric(SearchLevelsTextBox.Text) Or _
    Val(SearchLevelsTextBox.Text) < 1 Or _
    Val(SearchLevelsTextBox.Text) > 26) Then
    SearchLevelsTextBox.Text = Empty
    MsgBox "Please enter a valid search level (1 to 26)" & Chr(13) _
        & "in the 'Search Levels' Input Box"
    SearchLevelsTextBox.SetFocus
    Exit Sub
End If
If Val(SearchLevelsTextBox.Text) - Int(Val(SearchLevelsTextBox.Text)) > 0 Then
    SearchLevelsTextBox.Text = Empty
    MsgBox "Please enter a valid search level (1 to 26)" & Chr(13) _
        & "in the 'Search Levels' Input Box" & Chr(13) & Chr(13) & _
        "Your entry must be an integer (no fractional portion)."
    SearchLevelsTextBox.SetFocus
    Exit Sub
End If
ShowEstimatedTime
End Sub

Private Sub SearchLevelsTextBox_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
SearchDisableContinueCheck
If SearchContinueReady = 1 Then
    Exit Sub
End If
End Sub

'Private Sub SearchMatchPartialOption_Click()
'If SearchFileLoading Then
'    Exit Sub
'End If
'    frmPartialDeckMatch.Show vbModal
'    'SearchSpecialRangeOne ("MatchPartial")
'End Sub

Private Sub SearchMatchPartialOption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
SearchDisableContinueCheck
If SearchContinueReady = 1 Then
    Exit Sub
End If
End Sub

Private Sub SearchMatchPartialOption_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    frmPartialDeckMatch.Show vbModal
End Sub

'Private Sub SearchMatchWholeOption_Click()
'If SearchFileLoading Then
'    Exit Sub
'End If
'    SearchMatchStartCard = 1
'    SearchMatchEndCard = 52
'End Sub

Private Sub SearchMatchWholeOption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
SearchDisableContinueCheck
If SearchContinueReady = 1 Then
    Exit Sub
End If
End Sub

Private Sub SearchMatchWholeOption_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    frmWholeDeckMatch.Show vbModal
End Sub

Private Sub SearchMoveCard_Click()
If SearchFileLoading Then
    If SearchMoveCard.Value = 1 Then
        SearchCounterMax = SearchCounterMax + SearchMC
    End If
    Exit Sub
End If
If SearchMoveCard.Value = 1 Then
    SearchMoveCardAll.Value = 0
End If
If SearchMoveCard.Value = 0 Then
    If SearchSpecialCancel Then
        Exit Sub
    End If
    SearchCounterMax = SearchCounterMax - SearchMC
Else
    SearchSpecialRangeTwo ("MoveCard")
    SearchMoveCardAll.Value = 0
    If SearchSpecialCancel Then
        Exit Sub
    End If
    SearchMC = (SearchMC1Max - SearchMC1Min + 1) * (SearchMC2Max - SearchMC2Min + 1)
    For i% = SearchMC1Min To SearchMC1Max
        For j% = SearchMC2Min To SearchMC2Max
            If i% = j% Then
                SearchMC = SearchMC - 1
            End If
        Next j%
    Next i%
    SearchCounterMax = SearchCounterMax + SearchMC
End If
ShowEstimatedTime
End Sub

Private Sub SearchMoveCard_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
SearchSpecialCancel = False
SearchDisableContinueCheck
If SearchContinueReady = 1 Then
    Exit Sub
End If

End Sub

Private Sub SearchMoveCardAll_Click()
If SearchFileLoading Then
    If SearchMoveCardAll.Value = 1 Then
        SearchCounterMax = SearchCounterMax + 2652
    End If
    Exit Sub
End If
If SearchMoveCardAll.Value = 1 Then
    SearchMoveCard.Value = 0
End If
If SearchMoveCardAll.Value = 0 Then
    SearchCounterMax = SearchCounterMax - 2652
Else
    SearchCounterMax = SearchCounterMax + 2652
End If
ShowEstimatedTime
End Sub

Private Sub SearchMoveCardAll_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
SearchDisableContinueCheck
If SearchContinueReady = 1 Then
    Exit Sub
End If

End Sub

Private Sub SearchOutFaroAll_Click()
If SearchFileLoading Then
    If SearchOutFaroAll.Value = 1 Then
        SearchCounterMax = SearchCounterMax + 1
    End If
    Exit Sub
End If
If SearchOutFaroAll.Value = 0 Then
    SearchCounterMax = SearchCounterMax - 1
Else
    SearchCounterMax = SearchCounterMax + 1
End If
ShowEstimatedTime
End Sub

Private Sub SearchOutFaroAll_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
SearchDisableContinueCheck
If SearchContinueReady = 1 Then
    Exit Sub
End If

End Sub


Private Sub SearchOutFaroInverseAll_Click()
If SearchFileLoading Then
    If SearchOutFaroInverseAll.Value = 1 Then
        SearchCounterMax = SearchCounterMax + 1
    End If
    Exit Sub
End If
If SearchOutFaroInverseAll.Value = 0 Then
    SearchCounterMax = SearchCounterMax - 1
Else
    SearchCounterMax = SearchCounterMax + 1
End If
ShowEstimatedTime
End Sub

Private Sub SearchOutFaroInverseAll_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
SearchDisableContinueCheck
If SearchContinueReady = 1 Then
    Exit Sub
End If

End Sub

Private Sub SearchOutFaroSpecialBottom_Click()
If SearchFileLoading Then
    If SearchOutFaroSpecialBottom.Value = 1 Then
        SearchCounterMax = SearchCounterMax + SearchOFSB
    End If
    Exit Sub
End If
If SearchOutFaroSpecialBottom.Value = 1 Then
    SearchOutFaroSpecialBottomAll.Value = 0
End If
If SearchOutFaroSpecialBottom.Value = 0 Then
    If SearchSpecialCancel Then
        Exit Sub
    End If
    SearchCounterMax = SearchCounterMax - SearchOFSB
Else
    SearchSpecialRangeTwo ("OutFaroSpecialBottom")
    SearchOutFaroSpecialBottomAll.Value = 0
    SearchOFSB = 0
    If SearchSpecialCancel Then
        Exit Sub
    End If
    For i% = SearchOFSB1Min To SearchOFSB1Max
        For j% = SearchOFSB2Min To SearchOFSB2Max
            If i% + j% < 53 Then
                SearchOFSB = SearchOFSB + 1
            End If
        Next j%
    Next i%
    SearchCounterMax = SearchCounterMax + SearchOFSB
    If SearchOFSB = 0 Then
        MsgBox "Your range values do not allow for any manipulations to be included." & Chr(13) _
            & "The sum of the values taken from the ranges can not be greater than 52."
    End If
    SearchOutFaroSpecialBottom.Value = 0
End If
ShowEstimatedTime
End Sub

Private Sub SearchOutFaroSpecialBottom_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
SearchSpecialCancel = False
SearchDisableContinueCheck
If SearchContinueReady = 1 Then
    Exit Sub
End If

End Sub

Private Sub SearchOutFaroSpecialBottomAll_Click()
If SearchFileLoading Then
    If SearchOutFaroSpecialBottomAll.Value = 1 Then
        SearchCounterMax = SearchCounterMax + 1326
    End If
    Exit Sub
End If
If SearchOutFaroSpecialBottomAll.Value = 1 Then
    SearchOutFaroSpecialBottom.Value = 0
End If
If SearchOutFaroSpecialBottomAll.Value = 0 Then
    SearchCounterMax = SearchCounterMax - 1326
Else
    SearchCounterMax = SearchCounterMax + 1326
End If
ShowEstimatedTime
End Sub

Private Sub SearchOutFaroSpecialBottomAll_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
SearchDisableContinueCheck
If SearchContinueReady = 1 Then
    Exit Sub
End If

End Sub

Private Sub SearchOutFaroSpecialBottomInv_Click()
If SearchFileLoading Then
    If SearchOutFaroSpecialBottomInv.Value = 1 Then
        SearchCounterMax = SearchCounterMax + SearchOFSBR
    End If
    Exit Sub
End If
If SearchOutFaroSpecialBottomInv.Value = 1 Then
    SearchOutFaroSpecialBottomInvAll.Value = 0
End If
If SearchOutFaroSpecialBottomInv.Value = 0 Then
    If SearchSpecialCancel Then
        Exit Sub
    End If
    SearchCounterMax = SearchCounterMax - SearchOFSBR
Else
    SearchSpecialRangeTwo ("OutFaroSpecialBottomInv")
    SearchOutFaroSpecialBottomInvAll.Value = 0
    SearchOFSBR = 0
    If SearchSpecialCancel Then
        Exit Sub
    End If
    For i% = SearchOFSBR1Min To SearchOFSBR1Max
        For j% = SearchOFSBR2Min To SearchOFSBR2Max
            If i% + j% < 53 Then
                SearchOFSBR = SearchOFSBR + 1
            End If
        Next j%
    Next i%
    SearchCounterMax = SearchCounterMax + SearchOFSBR
    If SearchOFSBR = 0 Then
        MsgBox "Your range values do not allow for any manipulations to be included." & Chr(13) _
            & "The sum of the values taken from the ranges can not be greater than 52."
    End If
    SearchOutFaroSpecialBottomInv.Value = 0
End If
ShowEstimatedTime
End Sub

Private Sub SearchOutFaroSpecialBottomInv_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
SearchSpecialCancel = False
SearchDisableContinueCheck
If SearchContinueReady = 1 Then
    Exit Sub
End If

End Sub

Private Sub SearchOutFaroSpecialBottomInvAll_Click()
If SearchFileLoading Then
    If SearchOutFaroSpecialBottomInvAll.Value = 1 Then
        SearchCounterMax = SearchCounterMax + 1326
    End If
    Exit Sub
End If
If SearchOutFaroSpecialBottomInvAll.Value = 1 Then
    SearchOutFaroSpecialBottomInv.Value = 0
End If
If SearchOutFaroSpecialBottomInvAll.Value = 0 Then
    SearchCounterMax = SearchCounterMax - 1326
Else
    SearchCounterMax = SearchCounterMax + 1326
End If
ShowEstimatedTime
End Sub

Private Sub SearchOutFaroSpecialBottomInvAll_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
SearchDisableContinueCheck
If SearchContinueReady = 1 Then
    Exit Sub
End If

End Sub

Private Sub SearchOutFaroSpecialTop_Click()
If SearchFileLoading Then
    If SearchOutFaroSpecialTop.Value = 1 Then
        SearchCounterMax = SearchCounterMax + SearchOFST
    End If
    Exit Sub
End If
If SearchOutFaroSpecialTop.Value = 1 Then
    SearchOutFaroSpecialTopAll.Value = 0
End If
If SearchOutFaroSpecialTop.Value = 0 Then
    If SearchSpecialCancel Then
        Exit Sub
    End If
    SearchCounterMax = SearchCounterMax - SearchOFST
Else
    SearchSpecialRangeTwo ("OutFaroSpecialTop")
    SearchOutFaroSpecialTopAll.Value = 0
    SearchOFST = 0
    If SearchSpecialCancel Then
        Exit Sub
    End If
    For i% = SearchOFST1Min To SearchOFST1Max
        For j% = SearchOFST2Min To SearchOFST2Max
            If i% + j% < 53 Then
                SearchOFST = SearchOFST + 1
            End If
        Next j%
    Next i%
    SearchCounterMax = SearchCounterMax + SearchOFST
    If SearchOFST = 0 Then
        MsgBox "Your range values do not allow for any manipulations to be included." & Chr(13) _
            & "The sum of the values taken from the ranges can not be greater than 52."
    End If
    SearchOutFaroSpecialTop.Value = 0
End If
ShowEstimatedTime
End Sub

Private Sub SearchOutFaroSpecialTop_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
SearchSpecialCancel = False
SearchDisableContinueCheck
If SearchContinueReady = 1 Then
    Exit Sub
End If

End Sub

Private Sub SearchOutFaroSpecialTopAll_Click()
If SearchFileLoading Then
    If SearchOutFaroSpecialTopAll.Value = 1 Then
        SearchCounterMax = SearchCounterMax + 1326
    End If
    Exit Sub
End If
If SearchOutFaroSpecialTopAll.Value = 1 Then
    SearchOutFaroSpecialTop.Value = 0
End If
If SearchOutFaroSpecialTopAll.Value = 0 Then
    SearchCounterMax = SearchCounterMax - 1326
Else
    SearchCounterMax = SearchCounterMax + 1326
End If
ShowEstimatedTime
End Sub

Private Sub SearchOutFaroSpecialTopAll_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
SearchDisableContinueCheck
If SearchContinueReady = 1 Then
    Exit Sub
End If

End Sub

Private Sub SearchOutFaroSpecialTopInv_Click()
If SearchFileLoading Then
    If SearchOutFaroSpecialTopInv.Value = 1 Then
        SearchCounterMax = SearchCounterMax + SearchOFSTR
    End If
    Exit Sub
End If
If SearchOutFaroSpecialTopInv.Value = 1 Then
    SearchOutFaroSpecialTopInvAll.Value = 0
End If
If SearchOutFaroSpecialTopInv.Value = 0 Then
    If SearchSpecialCancel Then
        Exit Sub
    End If
    SearchCounterMax = SearchCounterMax - SearchOFSTR
Else
    SearchSpecialRangeTwo ("OutFaroSpecialTopInv")
    SearchOutFaroSpecialTopInvAll.Value = 0
    SearchOFSTR = 0
    If SearchSpecialCancel Then
        Exit Sub
    End If
    For i% = SearchOFSTR1Min To SearchOFSTR1Max
        For j% = SearchOFSTR2Min To SearchOFSTR2Max
            If i% + j% < 53 Then
                SearchOFSTR = SearchOFSTR + 1
            End If
        Next j%
    Next i%
    SearchCounterMax = SearchCounterMax + SearchOFSTR
    If SearchOFSTR = 0 Then
        MsgBox "Your range values do not allow for any manipulations to be included." & Chr(13) _
            & "The sum of the values taken from the ranges can not be greater than 52."
    End If
    SearchOutFaroSpecialTopInv.Value = 0
End If
ShowEstimatedTime
End Sub

Private Sub SearchOutFaroSpecialTopInv_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
SearchSpecialCancel = False
SearchDisableContinueCheck
If SearchContinueReady = 1 Then
    Exit Sub
End If

End Sub

Private Sub SearchOutFaroSpecialTopInvAll_Click()
If SearchFileLoading Then
    If SearchOutFaroSpecialTopInvAll.Value = 1 Then
        SearchCounterMax = SearchCounterMax + 1326
    End If
    Exit Sub
End If
If SearchOutFaroSpecialTopInvAll.Value = 1 Then
    SearchOutFaroSpecialTopInv.Value = 0
End If
If SearchOutFaroSpecialTopInvAll.Value = 0 Then
    SearchCounterMax = SearchCounterMax - 1326
Else
    SearchCounterMax = SearchCounterMax + 1326
End If
ShowEstimatedTime
End Sub

Private Sub SearchOutFaroSpecialTopInvAll_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
SearchDisableContinueCheck
If SearchContinueReady = 1 Then
    Exit Sub
End If

End Sub

Private Sub SearchRunSingleCards_Click()
If SearchFileLoading Then
    If SearchRunSingleCards.Value = 1 Then
        SearchCounterMax = SearchCounterMax + (SearchRSCMax - SearchRSCMin + 1)
    End If
    Exit Sub
End If
If SearchRunSingleCards.Value = 1 Then
    SearchRunSingleCardsAll.Value = 0
End If
If SearchRunSingleCards.Value = 0 Then
    If SearchSpecialCancel Then
        Exit Sub
    End If
    SearchCounterMax = SearchCounterMax - (SearchRSCMax - SearchRSCMin + 1)
Else
    SearchSpecialRangeOne ("RunSingleCards")
    If SearchSpecialCancel Then
        Exit Sub
    End If
    SearchCounterMax = SearchCounterMax + (SearchRSCMax - SearchRSCMin + 1)
End If
ShowEstimatedTime
End Sub

Private Sub SearchRunSingleCards_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
SearchSpecialCancel = False
SearchDisableContinueCheck
If SearchContinueReady = 1 Then
    Exit Sub
End If

End Sub

Private Sub SearchRunSingleCardsAll_Click()
If SearchFileLoading Then
    If SearchRunSingleCardsAll.Value = 1 Then
        SearchCounterMax = SearchCounterMax + 52
    End If
    Exit Sub
End If
If SearchRunSingleCardsAll.Value = 1 Then
    SearchRunSingleCards.Value = 0
End If
If SearchRunSingleCardsAll.Value = 0 Then
    SearchCounterMax = SearchCounterMax - 52
Else
    SearchCounterMax = SearchCounterMax + 52
End If
ShowEstimatedTime
End Sub

Private Sub SearchRunSingleCardsAll_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
SearchDisableContinueCheck
If SearchContinueReady = 1 Then
    Exit Sub
End If

End Sub

Private Sub SearchRunSingleCardsInv_Click()
If SearchFileLoading Then
    If SearchRunSingleCardsInv.Value = 1 Then
        SearchCounterMax = SearchCounterMax + (SearchRSCRMax - SearchRSCRMin + 1)
    End If
    Exit Sub
End If
If SearchRunSingleCardsInv.Value = 1 Then
    SearchRunSingleCardsInvAll.Value = 0
End If
If SearchRunSingleCardsInv.Value = 0 Then
    If SearchSpecialCancel Then
        Exit Sub
    End If
    SearchCounterMax = SearchCounterMax - (SearchRSCRMax - SearchRSCRMin + 1)
Else
    SearchSpecialRangeOne ("RunSingleCardsInv")
    If SearchSpecialCancel Then
        Exit Sub
    End If
    SearchCounterMax = SearchCounterMax + (SearchRSCRMax - SearchRSCRMin + 1)
End If
ShowEstimatedTime
End Sub

Private Sub SearchRunSingleCardsInv_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
SearchSpecialCancel = False
SearchDisableContinueCheck
If SearchContinueReady = 1 Then
    Exit Sub
End If

End Sub

Private Sub SearchRunSingleCardsInvAll_Click()
If SearchFileLoading Then
    If SearchRunSingleCardsInvAll.Value = 1 Then
        SearchCounterMax = SearchCounterMax + 52
    End If
    Exit Sub
End If
If SearchRunSingleCardsInvAll.Value = 1 Then
    SearchRunSingleCardsInv.Value = 0
End If
If SearchRunSingleCardsInvAll.Value = 0 Then
    SearchCounterMax = SearchCounterMax - 52
Else
    SearchCounterMax = SearchCounterMax + 52
End If
ShowEstimatedTime
End Sub

Private Sub SearchRunSingleCardsInvAll_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
SearchDisableContinueCheck
If SearchContinueReady = 1 Then
    Exit Sub
End If

End Sub

Private Sub SearchShiftTopBlock_Click()
If SearchFileLoading Then
    If SearchShiftTopBlock.Value = 1 Then
        SearchCounterMax = SearchCounterMax + SearchSTB
    End If
    Exit Sub
End If
If SearchShiftTopBlock.Value = 1 Then
    SearchShiftTopBlockAll.Value = 0
End If
If SearchShiftTopBlock.Value = 0 Then
    If SearchSpecialCancel Then
        Exit Sub
    End If
    SearchCounterMax = SearchCounterMax - SearchSTB
Else
    SearchSpecialRangeTwo ("ShiftTopBlock")
    SearchShiftTopBlockAll.Value = 0
    SearchSTB = 0
    If SearchSpecialCancel Then
        Exit Sub
    End If
    For i% = SearchSTB1Min To SearchSTB1Max
        For j% = SearchSTB2Min To SearchSTB2Max
            If i% + j% < 53 Then
                SearchSTB = SearchSTB + 1
            End If
        Next j%
    Next i%
    SearchCounterMax = SearchCounterMax + SearchSTB
    If SearchSTB = 0 Then
        MsgBox "Your range values do not allow for any manipulations to be included." & Chr(13) _
            & "The sum of the values taken from the ranges can not be greater than 52."
    End If
    SearchShiftTopBlock.Value = 0
End If
ShowEstimatedTime
End Sub

Private Sub SearchShiftTopBlock_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
SearchSpecialCancel = False
SearchDisableContinueCheck
If SearchContinueReady = 1 Then
    Exit Sub
End If

End Sub

Private Sub SearchShiftTopBlockAll_Click()
If SearchFileLoading Then
    If SearchShiftTopBlockAll.Value = 1 Then
        SearchCounterMax = SearchCounterMax + 1326
    End If
    Exit Sub
End If
If SearchShiftTopBlockAll.Value = 1 Then
    SearchShiftTopBlock.Value = 0
End If
If SearchShiftTopBlockAll.Value = 0 Then
    SearchCounterMax = SearchCounterMax - 1326
Else
    SearchCounterMax = SearchCounterMax + 1326
End If
ShowEstimatedTime
End Sub

Private Sub SearchShiftTopBlockAll_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
SearchDisableContinueCheck
If SearchContinueReady = 1 Then
    Exit Sub
End If

End Sub

Private Sub SearchShiftTopBlockInv_Click()
If SearchFileLoading Then
    If SearchShiftTopBlockInv.Value = 1 Then
        SearchCounterMax = SearchCounterMax + SearchSTBR
    End If
    Exit Sub
End If
If SearchShiftTopBlockInv.Value = 1 Then
    SearchShiftTopBlockInvAll.Value = 0
End If
If SearchShiftTopBlockInv.Value = 0 Then
    If SearchSpecialCancel Then
        Exit Sub
    End If
    SearchCounterMax = SearchCounterMax - SearchSTBR
Else
    SearchSpecialRangeTwo ("ShiftTopBlockInv")
    SearchShiftTopBlockInvAll.Value = 0
    SearchSTBR = 0
    If SearchSpecialCancel Then
        Exit Sub
    End If
    For i% = SearchSTBR1Min To SearchSTBR1Max
        For j% = SearchSTBR2Min To SearchSTBR2Max
            If i% + j% < 53 Then
                SearchSTBR = SearchSTBR + 1
            End If
        Next j%
    Next i%
    SearchCounterMax = SearchCounterMax + SearchSTBR
    If SearchSTBR = 0 Then
        MsgBox "Your range values do not allow for any manipulations to be included." & Chr(13) _
            & "The sum of the values taken from the ranges can not be greater than 52."
    End If
    SearchShiftTopBlockInv.Value = 0
End If
ShowEstimatedTime
End Sub

Private Sub SearchShiftTopBlockInv_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
SearchSpecialCancel = False
SearchDisableContinueCheck
If SearchContinueReady = 1 Then
    Exit Sub
End If

End Sub

Private Sub SearchShiftTopBlockInvAll_Click()
If SearchFileLoading Then
    If SearchShiftTopBlockInvAll.Value = 1 Then
        SearchCounterMax = SearchCounterMax + 1326
    End If
    Exit Sub
End If
If SearchShiftTopBlockInvAll.Value = 1 Then
    SearchShiftTopBlockInv.Value = 0
End If
If SearchShiftTopBlockInvAll.Value = 0 Then
    SearchCounterMax = SearchCounterMax - 1326
Else
    SearchCounterMax = SearchCounterMax + 1326
End If
ShowEstimatedTime
End Sub

Private Sub SearchSpecialRangeOne(checkboxname)
SearchSpecialName = checkboxname
frmSearchSpecialOne.Show vbModal
End Sub

Private Sub SearchSpecialRangeTwo(checkboxname)
SearchSpecialName = checkboxname
frmSearchSpecialTwo.Show vbModal
End Sub



Private Sub SearchShiftTopBlockInvAll_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
SearchDisableContinueCheck
If SearchContinueReady = 1 Then
    Exit Sub
End If

End Sub

Private Sub SearchToggle_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If SearchingMode Then
    SearchToggle(0).Visible = False
    SearchToggle(1).Visible = False
    SearchToggle(2).Visible = False
    SearchToggle(3).Visible = True
Else
    SearchToggle(0).Visible = False
    SearchToggle(1).Visible = True
    SearchToggle(2).Visible = False
    SearchToggle(3).Visible = False
End If
End Sub

Private Sub SearchToggle_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If SearchingMode Then
    FirstPass = False
    SearchToggle(0).Visible = True
    SearchToggle(1).Visible = False
    SearchToggle(2).Visible = False
    SearchToggle(3).Visible = False
    ContinueSearchToggle(0).Visible = True
    ContinueSearchToggle(1).Visible = False
    SearchContinueReady = 1
    SearchSaved = 0
    SearchCurrentLevelRestart = SearchCurrentLevel
    ShowProgress
    ShowElapsedTime
    MatchListBox.Clear
    MatchFoundLabel.Caption = Empty
    MatchFoundLabel.Visible = False
Else
    RestartSearchCheck
    If SearchContinueReady = 1 And SearchSaved = 0 Then
        SearchToggle(0).Visible = True
        SearchToggle(1).Visible = False
        SearchToggle(2).Visible = False
        SearchToggle(3).Visible = False
        Exit Sub
    End If
    SearchToggle(0).Visible = False
    SearchToggle(1).Visible = False
    SearchToggle(2).Visible = True
    SearchToggle(3).Visible = False
    SearchContinueReady = 0
    ContinueSearchToggle(0).Visible = False
    ContinueSearchToggle(1).Visible = False
    ProgressLabel.Caption = ""
    FirstPass = True
End If
SearchToggle(0).Refresh
SearchToggle(1).Refresh
SearchToggle(2).Refresh
SearchToggle(3).Refresh
SearchingMode = Not SearchingMode
If SearchingMode Then
    'make sure there is no untransferred session remaining
    If SearchSessionTransferred = 0 And MatchListBox.ListCount > 0 Then
        Dim Msg, Style, Title, Help, Ctxt, Response, MyString
        Msg = "There are still untransferred Search results" & _
            Chr(13) & _
            "This will delete the Search results." & Chr(13) & _
            Chr(13) & "Do you want to continue?"   ' Define message.
        Style = vbYesNo + vbQuestion + vbDefaultButton2   ' Define buttons.
        Title = "Clear Search Results?"   ' Define title.
        Response = MsgBox(Msg, Style, Title)
        If Response = vbYes Then   ' User chose Yes.
            MatchListBox.Clear
        Else
            SearchToggle(0).Visible = True
            SearchToggle(1).Visible = False
            SearchToggle(2).Visible = False
            SearchToggle(3).Visible = False
            SearchToggle(0).Refresh
            SearchToggle(1).Refresh
            SearchToggle(2).Refresh
            SearchToggle(3).Refresh
            SearchingMode = False
            Exit Sub
        End If
    Else
        MatchListBox.Clear
    End If
    
    'turn off search label
    MatchFoundLabel.Visible = False
    
    'make sure the start and target decks are set
    If SearchStartDeckSet = 0 Then
        'fix the button state
        SearchToggle(0).Visible = True
        SearchToggle(1).Visible = False
        SearchToggle(2).Visible = False
        SearchToggle(3).Visible = False
        MsgBox "You need to set the Start Deck." & Chr(13)
        SearchingMode = False
        Exit Sub
    End If
    If SearchTargetDeckSet = 0 Then
        'fix the button state
        SearchToggle(0).Visible = True
        SearchToggle(1).Visible = False
        SearchToggle(2).Visible = False
        SearchToggle(3).Visible = False
        MsgBox "You need to set the Target Deck." & Chr(13)
        SearchingMode = False
        Exit Sub
    End If
    If SearchCounterMax = 0 Then
        'fix the button state
        SearchToggle(0).Visible = True
        SearchToggle(1).Visible = False
        SearchToggle(2).Visible = False
        SearchToggle(3).Visible = False
        MsgBox "You need to select one or more manipulations " & Chr(13) & _
            "in order to run a search."
        SearchingMode = False
        Exit Sub
    End If
    
    'transfer over the initial start deck to the deck used for search
    For m% = 1 To 52
        For z% = 1 To 2
            StartDeck(z%, m%) = StartDeckInitial(z%, m%)
        Next z%
    Next m%

    
    'disable the controls
    SearchItems.Enabled = False
    SearchLevelsTextBox.Enabled = False
    StartDeckButton.Enabled = False
    TargetDeckButton.Enabled = False
    TransferToSession.Enabled = False
    SearchCutDeckPrecise.Enabled = False
    SearchMoveCard.Enabled = False
    SearchRunSingleCards.Enabled = False
    SearchRunSingleCardsInv.Enabled = False
    SearchShiftTopBlock.Enabled = False
    SearchShiftTopBlockInv.Enabled = False
    SearchOutFaro.Enabled = False
    SearchOutFaroInverse.Enabled = False
    SearchOutFaroSpecialTop.Enabled = False
    SearchOutFaroSpecialTopInv.Enabled = False
    SearchOutFaroSpecialBottom.Enabled = False
    SearchOutFaroSpecialBottomInv.Enabled = False
    SearchInFaro.Enabled = False
    SearchInFaroInverse.Enabled = False
    SearchInFaroSpecialTop.Enabled = False
    SearchInFaroSpecialTopInv.Enabled = False
    SearchInFaroSpecialBottom.Enabled = False
    SearchInFaroSpecialBottomInv.Enabled = False
    SearchCutDeckPreciseAll.Enabled = False
    SearchMoveCardAll.Enabled = False
    SearchRunSingleCardsAll.Enabled = False
    SearchRunSingleCardsInvAll.Enabled = False
    SearchShiftTopBlockAll.Enabled = False
    SearchShiftTopBlockInvAll.Enabled = False
    SearchOutFaroAll.Enabled = False
    SearchOutFaroInverseAll.Enabled = False
    SearchOutFaroSpecialTopAll.Enabled = False
    SearchOutFaroSpecialTopInvAll.Enabled = False
    SearchOutFaroSpecialBottomAll.Enabled = False
    SearchOutFaroSpecialBottomInvAll.Enabled = False
    SearchInFaroAll.Enabled = False
    SearchInFaroInverseAll.Enabled = False
    SearchInFaroSpecialTopAll.Enabled = False
    SearchInFaroSpecialTopInvAll.Enabled = False
    SearchInFaroSpecialBottomAll.Enabled = False
    SearchInFaroSpecialBottomInvAll.Enabled = False
    'SearchMatchWholeOption.Enabled = False
    'SearchMatchPartialOption.Enabled = False
    SearchMatchFrame.Enabled = False
    
    'start the progress bar and timer
    SearchProgress.Value = 0
    SearchProgress.Max = 100
    SearchProgress.Visible = True
    
    'start timer -- this is for my debugging and procedure time comparisons only
    SearchElapsedTime = 0
    SearchStartTime = Timer
    ShowElapsedTime
        
    'RUN THE SEARCH
    SearchTransactionInitialize
    LoadManipulations
    SearchTransactionRun
    
    'check for a threshold trap pause (allow continue)
    If PartialMatchFound = 1 Then
        SearchToggle(0).Visible = True
        SearchToggle(1).Visible = False
        SearchToggle(2).Visible = False
        SearchToggle(3).Visible = False
        ContinueSearchToggle(0).Visible = True
        ContinueSearchToggle(1).Visible = False
        SearchContinueReady = 1
        SearchSaved = 0
        SearchCurrentLevelRestart = SearchCurrentLevel
        PartialMatchFound = 0
    End If

    
    
    'stop timer
    If Timer < SearchStartTime Then
        SearchStartTime = SearchStartTime - 86400
        'error condition is test crosses midnight
    End If
    SearchElapsedTime = SearchElapsedTime + Timer - SearchStartTime
    ShowElapsedTime
    
    're-enable the controls
    SearchItems.Enabled = True
    SearchLevelsTextBox.Enabled = True
    StartDeckButton.Enabled = True
    TargetDeckButton.Enabled = True
    TransferToSession.Enabled = True
    SearchCutDeckPrecise.Enabled = True
    SearchMoveCard.Enabled = True
    SearchRunSingleCards.Enabled = True
    SearchRunSingleCardsInv.Enabled = True
    SearchShiftTopBlock.Enabled = True
    SearchShiftTopBlockInv.Enabled = True
    SearchOutFaro.Enabled = False
    SearchOutFaroInverse.Enabled = False
    SearchOutFaroSpecialTop.Enabled = True
    SearchOutFaroSpecialTopInv.Enabled = True
    SearchOutFaroSpecialBottom.Enabled = True
    SearchOutFaroSpecialBottomInv.Enabled = True
    SearchInFaro.Enabled = False
    SearchInFaroInverse.Enabled = False
    SearchInFaroSpecialTop.Enabled = True
    SearchInFaroSpecialTopInv.Enabled = True
    SearchInFaroSpecialBottom.Enabled = True
    SearchInFaroSpecialBottomInv.Enabled = True
    SearchCutDeckPreciseAll.Enabled = True
    SearchMoveCardAll.Enabled = True
    SearchRunSingleCardsAll.Enabled = True
    SearchRunSingleCardsInvAll.Enabled = True
    SearchShiftTopBlockAll.Enabled = True
    SearchShiftTopBlockInvAll.Enabled = True
    SearchOutFaroAll.Enabled = True
    SearchOutFaroInverseAll.Enabled = True
    SearchOutFaroSpecialTopAll.Enabled = True
    SearchOutFaroSpecialTopInvAll.Enabled = True
    SearchOutFaroSpecialBottomAll.Enabled = True
    SearchOutFaroSpecialBottomInvAll.Enabled = True
    SearchInFaroAll.Enabled = True
    SearchInFaroInverseAll.Enabled = True
    SearchInFaroSpecialTopAll.Enabled = True
    SearchInFaroSpecialTopInvAll.Enabled = True
    SearchInFaroSpecialBottomAll.Enabled = True
    SearchInFaroSpecialBottomInvAll.Enabled = True
    SearchingMode = False
    SearchToggle(0).Visible = True
    SearchToggle(1).Visible = False
    SearchToggle(2).Visible = False
    SearchToggle(3).Visible = False
    SearchProgress.Visible = False
    'SearchMatchWholeOption.Enabled = True
    'SearchMatchPartialOption.Enabled = True
    SearchMatchFrame.Enabled = True

Else
    'THIS CODE IS FOR THE "MOUSE UP" WHEN THE SEARCHING IS FINISHED
    
    'enable the controls
    SearchItems.Enabled = True
    SearchLevelsTextBox.Enabled = True
    StartDeckButton.Enabled = True
    TargetDeckButton.Enabled = True
    TransferToSession.Enabled = True
    SearchCutDeckPrecise.Enabled = True
    SearchMoveCard.Enabled = True
    SearchRunSingleCards.Enabled = True
    SearchRunSingleCardsInv.Enabled = True
    SearchShiftTopBlock.Enabled = True
    SearchShiftTopBlockInv.Enabled = True
    SearchOutFaro.Enabled = False
    SearchOutFaroInverse.Enabled = False
    SearchOutFaroSpecialTop.Enabled = True
    SearchOutFaroSpecialTopInv.Enabled = True
    SearchOutFaroSpecialBottom.Enabled = True
    SearchOutFaroSpecialBottomInv.Enabled = True
    SearchInFaro.Enabled = False
    SearchInFaroInverse.Enabled = False
    SearchInFaroSpecialTop.Enabled = True
    SearchInFaroSpecialTopInv.Enabled = True
    SearchInFaroSpecialBottom.Enabled = True
    SearchInFaroSpecialBottomInv.Enabled = True
    SearchCutDeckPreciseAll.Enabled = True
    SearchMoveCardAll.Enabled = True
    SearchRunSingleCardsAll.Enabled = True
    SearchRunSingleCardsInvAll.Enabled = True
    SearchShiftTopBlockAll.Enabled = True
    SearchShiftTopBlockInvAll.Enabled = True
    SearchOutFaroAll.Enabled = True
    SearchOutFaroInverseAll.Enabled = True
    SearchOutFaroSpecialTopAll.Enabled = True
    SearchOutFaroSpecialTopInvAll.Enabled = True
    SearchOutFaroSpecialBottomAll.Enabled = True
    SearchOutFaroSpecialBottomInvAll.Enabled = True
    SearchInFaroAll.Enabled = True
    SearchInFaroInverseAll.Enabled = True
    SearchInFaroSpecialTopAll.Enabled = True
    SearchInFaroSpecialTopInvAll.Enabled = True
    SearchInFaroSpecialBottomAll.Enabled = True
    SearchInFaroSpecialBottomInvAll.Enabled = True
    'SearchMatchWholeOption.Enabled = True
    'SearchMatchPartialOption.Enabled = True
    SearchMatchFrame.Enabled = True
    
End If
End Sub


Private Sub SearchTransactionInitialize()
MatchFound = 0
NoMatchFound = 0
SearchSessionTransferred = 0
SearchCurrentLevel = 1
For i% = 1 To 26
    SearchLevelCounter(i%) = 1
Next i%
End Sub

Private Sub SearchTransactionCheckMatch()
For card% = SearchMatchStartCard To SearchMatchEndCard
    If StartDeck(2, card%) <> TargetDeck(2, card%) Then
        MatchFound = 0
        Exit Sub
    End If
Next card%
MatchFound = 1
NoMatchFound = 0
FirstPass = False
MatchListBox.Clear
For i% = 0 To SearchCurrentLevel - 1
    MatchListBox.List(i%) = Manipulations(1, SearchLevelCounter(i% + 1))
Next i%
SearchSessionTransferred = 0
End Sub

Private Sub SearchTransactionCheckPartialMatch()
PartialMatchCounter = 0
For card% = SearchMatchStartCard To SearchMatchEndCard
    If StartDeck(2, card%) = TargetDeck(2, card%) Then
        PartialMatchCounter = PartialMatchCounter + 1
    End If
Next card%
If PartialMatchCounter <= ThresholdMatchCards Then
    PartialMatchFound = 0
    Exit Sub
Else
    PartialMatchFound = 1
    FirstPass = False
    MatchListBox.Clear
    For i% = 0 To SearchCurrentLevel - 1
        MatchListBox.List(i%) = Manipulations(1, SearchLevelCounter(i% + 1))
    Next i%
    SearchSessionTransferred = 0
End If
End Sub


Private Sub SearchTransactionRun()
SearchProgressCounter = 0
'set the progress bar counter to 0

    SearchLevel1
        If MatchFound = 1 Or PartialMatchFound = 1 Or SearchContinueReady = 1 Then
            NoMatchFound = 0
            Exit Sub
        End If
    SearchLevel2
        If MatchFound = 1 Or PartialMatchFound = 1 Or SearchContinueReady = 1 Then
            NoMatchFound = 0
            Exit Sub
        End If
    SearchLevel3
        If MatchFound = 1 Or PartialMatchFound = 1 Or SearchContinueReady = 1 Then
            NoMatchFound = 0
            Exit Sub
        End If
    SearchLevel4
        If MatchFound = 1 Or PartialMatchFound = 1 Or SearchContinueReady = 1 Then
            NoMatchFound = 0
            Exit Sub
        End If
    SearchLevel5
        If MatchFound = 1 Or PartialMatchFound = 1 Or SearchContinueReady = 1 Then
            NoMatchFound = 0
            Exit Sub
        End If
    SearchLevel6
        If MatchFound = 1 Or PartialMatchFound = 1 Or SearchContinueReady = 1 Then
            NoMatchFound = 0
            Exit Sub
        End If
    SearchLevel7
        If MatchFound = 1 Or PartialMatchFound = 1 Or SearchContinueReady = 1 Then
            NoMatchFound = 0
            Exit Sub
        End If
    SearchLevel8
        If MatchFound = 1 Or PartialMatchFound = 1 Or SearchContinueReady = 1 Then
            NoMatchFound = 0
            Exit Sub
        End If
    SearchLevel9
        If MatchFound = 1 Or PartialMatchFound = 1 Or SearchContinueReady = 1 Then
            NoMatchFound = 0
            Exit Sub
        End If
    SearchLevel10
        If MatchFound = 1 Or PartialMatchFound = 1 Or SearchContinueReady = 1 Then
            NoMatchFound = 0
            Exit Sub
        End If
    SearchLevel11
        If MatchFound = 1 Or PartialMatchFound = 1 Or SearchContinueReady = 1 Then
            NoMatchFound = 0
            Exit Sub
        End If
    SearchLevel12
        If MatchFound = 1 Or PartialMatchFound = 1 Or SearchContinueReady = 1 Then
            NoMatchFound = 0
            Exit Sub
        End If
    SearchLevel13
        If MatchFound = 1 Or PartialMatchFound = 1 Or SearchContinueReady = 1 Then
            NoMatchFound = 0
            Exit Sub
        End If
    SearchLevel14
        If MatchFound = 1 Or PartialMatchFound = 1 Or SearchContinueReady = 1 Then
            NoMatchFound = 0
            Exit Sub
        End If
    SearchLevel15
        If MatchFound = 1 Or PartialMatchFound = 1 Or SearchContinueReady = 1 Then
            NoMatchFound = 0
            Exit Sub
        End If
    SearchLevel16
        If MatchFound = 1 Or PartialMatchFound = 1 Or SearchContinueReady = 1 Then
            NoMatchFound = 0
            Exit Sub
        End If
    SearchLevel17
        If MatchFound = 1 Or PartialMatchFound = 1 Or SearchContinueReady = 1 Then
            NoMatchFound = 0
            Exit Sub
        End If
    SearchLevel18
        If MatchFound = 1 Or PartialMatchFound = 1 Or SearchContinueReady = 1 Then
            NoMatchFound = 0
            Exit Sub
        End If
    SearchLevel19
        If MatchFound = 1 Or PartialMatchFound = 1 Or SearchContinueReady = 1 Then
            NoMatchFound = 0
            Exit Sub
        End If
    SearchLevel20
        If MatchFound = 1 Or PartialMatchFound = 1 Or SearchContinueReady = 1 Then
            NoMatchFound = 0
            Exit Sub
        End If
    SearchLevel21
        If MatchFound = 1 Or PartialMatchFound = 1 Or SearchContinueReady = 1 Then
            NoMatchFound = 0
            Exit Sub
        End If
    SearchLevel22
        If MatchFound = 1 Or PartialMatchFound = 1 Or SearchContinueReady = 1 Then
            NoMatchFound = 0
            Exit Sub
        End If
    SearchLevel23
        If MatchFound = 1 Or PartialMatchFound = 1 Or SearchContinueReady = 1 Then
            NoMatchFound = 0
            Exit Sub
        End If
    SearchLevel24
        If MatchFound = 1 Or PartialMatchFound = 1 Or SearchContinueReady = 1 Then
            NoMatchFound = 0
            Exit Sub
        End If
    SearchLevel25
        If MatchFound = 1 Or PartialMatchFound = 1 Or SearchContinueReady = 1 Then
            NoMatchFound = 0
            Exit Sub
        End If
    SearchLevel26
        If MatchFound = 1 Or PartialMatchFound = 1 Or SearchContinueReady = 1 Then
            NoMatchFound = 0
            Exit Sub
        End If
SearchingMode = False
ProgressLabel.Caption = Empty
MatchListBox.Clear
MatchFoundLabel.Caption = "No Match Found"
NoMatchFound = 1
MatchFoundLabel.Visible = True
ShowElapsedTime
End Sub


Private Sub SearchLevel1()
'SEARCH FIRST LEVEL (1)
If SearchCurrentLevel = 1 And _
  SearchCurrentLevel <= Val(SearchLevelsTextBox.Text) Then
    ResetCurrentDeck
    Do While SearchingMode And SearchLevelCounter(1) <= SearchCounterMax
        SearchParse (Manipulations(1, SearchLevelCounter(1)))
        SearchTransactionCheckMatch
        If MatchFound = 0 Then
            If TrapThreshold Then
                SearchTransactionCheckPartialMatch
                If PartialMatchFound = 1 Then
                    ShowProgress
                    If SuspendTrapFinal Then
                        SearchParse (Manipulations(2, SearchLevelCounter(1)))
                        SearchLevelCounter(1) = SearchLevelCounter(1) + 1
                        Exit Sub
                    Else
                        AppendTrapFile
                    End If
                End If
            End If
            SearchParse (Manipulations(2, SearchLevelCounter(1)))
            SearchLevelCounter(1) = SearchLevelCounter(1) + 1
        Else
            MatchFoundLabel.Caption = "Match Found!!!"
            MatchFoundLabel.Visible = True
                If Timer < SearchStartTime Then
                    SearchStartTime = SearchStartTime - 86400
                    'error condition is test crosses midnight
                End If
                SearchElapsedTime = SearchElapsedTime + Timer - SearchStartTime
                ShowElapsedTime
                ProgressLabel.Caption = Empty
            Exit Sub
        End If
        If SearchProgressCounter Mod SpeedMod = 0 Then
            SearchProgressCounter = 0
            If SearchProgress.Value < SearchProgress.Max Then
                SearchProgress.Value = SearchProgress.Value + 1
            Else
                FirstPass = False
                SearchProgress.Value = 0
                If Timer < SearchStartTime Then
                    SearchStartTime = SearchStartTime - 86400
                    'error condition is test crosses midnight
                End If
                SearchElapsedTime = SearchElapsedTime + Timer - SearchStartTime
                ShowElapsedTime
                ShowProgress
                SearchStartTime = Timer
            End If
            DoEvents
            If SearchContinueReady = 1 Then
                Exit Sub
            End If
        End If
        SearchProgressCounter = SearchProgressCounter + 1
    Loop
    SearchLevelCounter(1) = 1
    'this is needed to set the counter to the beginning for the next
    'level search
    SearchCurrentLevel = SearchCurrentLevel + 1
    'this increments the level counter
End If
End Sub


Private Sub SearchLevel2()
'SEARCH SECOND LEVEL (2)

If SearchCurrentLevel = 2 And _
  SearchCurrentLevel <= Val(SearchLevelsTextBox.Text) Then
    Do While SearchingMode And SearchLevelCounter(1) <= SearchCounterMax
        ResetCurrentDeck
        SearchParse (Manipulations(1, SearchLevelCounter(1)))
        'Debug.Print "Correctly in 2nd level"
        Do While SearchingMode And SearchLevelCounter(2) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(2)))
            SearchTransactionCheckMatch
            If MatchFound = 0 Then
                If TrapThreshold Then
                    SearchTransactionCheckPartialMatch
                    If PartialMatchFound = 1 Then
                        ShowProgress
                        If SuspendTrapFinal Then
                            SearchParse (Manipulations(2, SearchLevelCounter(2)))
                            SearchLevelCounter(2) = SearchLevelCounter(2) + 1
                            Exit Sub
                        Else
                            AppendTrapFile
                        End If
                    End If
                End If
                SearchParse (Manipulations(2, SearchLevelCounter(2)))
                SearchLevelCounter(2) = SearchLevelCounter(2) + 1
            Else
                MatchFoundLabel.Caption = "Match Found!!!"
                MatchFoundLabel.Visible = True
                If Timer < SearchStartTime Then
                    SearchStartTime = SearchStartTime - 86400
                    'error condition is test crosses midnight
                End If
                SearchElapsedTime = SearchElapsedTime + Timer - SearchStartTime
                ShowElapsedTime
                ProgressLabel.Caption = Empty
                Exit Sub
            End If
            If SearchProgressCounter Mod SpeedMod = 0 Then
                SearchProgressCounter = 0
                If SearchProgress.Value < SearchProgress.Max Then
                    SearchProgress.Value = SearchProgress.Value + 1
                Else
                    FirstPass = False
                    SearchProgress.Value = 0
                    If Timer < SearchStartTime Then
                        SearchStartTime = SearchStartTime - 86400
                        'error condition is test crosses midnight
                    End If
                    ManipulationsPerSecond = (SpeedMod * 100) / (Timer - SearchStartTime)
                    SpeedMod = Int(ManipulationsPerSecond / 8)
                    SearchElapsedTime = SearchElapsedTime + Timer - SearchStartTime
                    ShowElapsedTime
                    ShowProgress
                    ShowEstimatedTime
                    SearchStartTime = Timer
                End If
                DoEvents
                If SearchContinueReady = 1 Then
                    Exit Sub
                End If
            End If
            SearchProgressCounter = SearchProgressCounter + 1
        Loop
        SearchLevelCounter(1) = SearchLevelCounter(1) + 1
        SearchLevelCounter(2) = 1
    Loop
    SearchLevelCounter(1) = 1
    SearchCurrentLevel = SearchCurrentLevel + 1
    'this increments the level counter
End If
End Sub


Private Sub SearchLevel3()
'SEARCH THIRD LEVEL (3)
If SearchCurrentLevel = 3 And _
  SearchCurrentLevel <= Val(SearchLevelsTextBox.Text) Then
    Do While SearchingMode And SearchLevelCounter(1) <= SearchCounterMax
        ResetCurrentDeck
        SearchParse (Manipulations(1, SearchLevelCounter(1)))
        Do While SearchingMode And SearchLevelCounter(2) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(2)))
        Do While SearchingMode And SearchLevelCounter(3) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(3)))
            SearchTransactionCheckMatch
            If MatchFound = 0 Then
                If TrapThreshold Then
                    SearchTransactionCheckPartialMatch
                    If PartialMatchFound = 1 Then
                        ShowProgress
                        If SuspendTrapFinal Then
                            SearchParse (Manipulations(2, SearchLevelCounter(3)))
                            SearchLevelCounter(3) = SearchLevelCounter(3) + 1
                            Exit Sub
                        Else
                            AppendTrapFile
                        End If
                    End If
                End If
                SearchParse (Manipulations(2, SearchLevelCounter(3)))
                SearchLevelCounter(3) = SearchLevelCounter(3) + 1
            Else
                MatchFoundLabel.Caption = "Match Found!!!"
                MatchFoundLabel.Visible = True
                If Timer < SearchStartTime Then
                    SearchStartTime = SearchStartTime - 86400
                    'error condition is test crosses midnight
                End If
                SearchElapsedTime = SearchElapsedTime + Timer - SearchStartTime
                ShowElapsedTime
                ProgressLabel.Caption = Empty
                Exit Sub
            End If
            If SearchProgressCounter Mod SpeedMod = 0 Then
                SearchProgressCounter = 0
                If SearchProgress.Value < SearchProgress.Max Then
                    SearchProgress.Value = SearchProgress.Value + 1
                Else
                    FirstPass = False
                    SearchProgress.Value = 0
                    If Timer < SearchStartTime Then
                        SearchStartTime = SearchStartTime - 86400
                        'error condition is test crosses midnight
                    End If
                    ManipulationsPerSecond = (SpeedMod * 100) / (Timer - SearchStartTime)
                    SpeedMod = Int(ManipulationsPerSecond / 8)
                    SearchElapsedTime = SearchElapsedTime + Timer - SearchStartTime
                    ShowElapsedTime
                    ShowProgress
                    ShowEstimatedTime
                    SearchStartTime = Timer
                End If
                DoEvents
                If SearchContinueReady = 1 Then
                    Exit Sub
                End If
            End If
            SearchProgressCounter = SearchProgressCounter + 1
        Loop 'level 3
            SearchParse (Manipulations(2, SearchLevelCounter(2)))
            'undo the level 2 manipulation for the next plunge into level 3
            SearchLevelCounter(3) = 1
            'restart level 3 counter
            SearchLevelCounter(2) = SearchLevelCounter(2) + 1
            'increment level 2 counter
        Loop 'level 2
            SearchLevelCounter(2) = 1
            'restart level 2 counter
            SearchLevelCounter(1) = SearchLevelCounter(1) + 1
            'increment level 1 counter
    Loop 'level 1
    SearchLevelCounter(1) = 1
    SearchCurrentLevel = SearchCurrentLevel + 1
    'this increments the level counter
End If
End Sub


Private Sub SearchLevel4()
'SEARCH FOURTH LEVEL (4)
If SearchCurrentLevel = 4 And _
  SearchCurrentLevel <= Val(SearchLevelsTextBox.Text) Then
    Do While SearchingMode And SearchLevelCounter(1) <= SearchCounterMax
        ResetCurrentDeck
        SearchParse (Manipulations(1, SearchLevelCounter(1)))
        Do While SearchingMode And SearchLevelCounter(2) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(2)))
        Do While SearchingMode And SearchLevelCounter(3) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(3)))
        Do While SearchingMode And SearchLevelCounter(4) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(4)))
            SearchTransactionCheckMatch
            If MatchFound = 0 Then
                If TrapThreshold Then
                    SearchTransactionCheckPartialMatch
                    If PartialMatchFound = 1 Then
                        ShowProgress
                        If SuspendTrapFinal Then
                            SearchParse (Manipulations(2, SearchLevelCounter(4)))
                            SearchLevelCounter(4) = SearchLevelCounter(4) + 1
                            Exit Sub
                        Else
                            AppendTrapFile
                        End If
                    End If
                End If
                SearchParse (Manipulations(2, SearchLevelCounter(4)))
                SearchLevelCounter(4) = SearchLevelCounter(4) + 1
            Else
                MatchFoundLabel.Caption = "Match Found!!!"
                MatchFoundLabel.Visible = True
                If Timer < SearchStartTime Then
                    SearchStartTime = SearchStartTime - 86400
                    'error condition is test crosses midnight
                End If
                SearchElapsedTime = SearchElapsedTime + Timer - SearchStartTime
                ShowElapsedTime
                ProgressLabel.Caption = Empty
                Exit Sub
            End If
            If SearchProgressCounter Mod SpeedMod = 0 Then
                SearchProgressCounter = 0
                If SearchProgress.Value < SearchProgress.Max Then
                    SearchProgress.Value = SearchProgress.Value + 1
                Else
                    FirstPass = False
                    SearchProgress.Value = 0
                    If Timer < SearchStartTime Then
                        SearchStartTime = SearchStartTime - 86400
                        'error condition is test crosses midnight
                    End If
                    ManipulationsPerSecond = (SpeedMod * 100) / (Timer - SearchStartTime)
                    SpeedMod = Int(ManipulationsPerSecond / 8)
                    SearchElapsedTime = SearchElapsedTime + Timer - SearchStartTime
                    ShowElapsedTime
                    ShowProgress
                    ShowEstimatedTime
                    SearchStartTime = Timer
                End If
                DoEvents
                If SearchContinueReady = 1 Then
                    Exit Sub
                End If
            End If
            SearchProgressCounter = SearchProgressCounter + 1
        Loop 'level 4
            SearchParse (Manipulations(2, SearchLevelCounter(3)))
            'undo the level 3 manipulation for the next plunge into level 4
            SearchLevelCounter(4) = 1
            'restart level 4 counter
            SearchLevelCounter(3) = SearchLevelCounter(3) + 1
            'increment level 3 counter
        Loop 'level 3
            SearchParse (Manipulations(2, SearchLevelCounter(2)))
            'undo the level 2 manipulation for the next plunge into level 3
            SearchLevelCounter(3) = 1
            'restart level 3 counter
            SearchLevelCounter(2) = SearchLevelCounter(2) + 1
            'increment level 2 counter
        Loop 'level 2
            SearchLevelCounter(2) = 1
            'restart level 2 counter
            SearchLevelCounter(1) = SearchLevelCounter(1) + 1
            'increment level 1 counter
    Loop 'level 1
    SearchLevelCounter(1) = 1
    SearchCurrentLevel = SearchCurrentLevel + 1
    'this increments the level counter
End If
End Sub


Private Sub SearchLevel5()
'SEARCH FIFTH LEVEL (5)
If SearchCurrentLevel = 5 And _
  SearchCurrentLevel <= Val(SearchLevelsTextBox.Text) Then
    Do While SearchingMode And SearchLevelCounter(1) <= SearchCounterMax
        ResetCurrentDeck
        SearchParse (Manipulations(1, SearchLevelCounter(1)))
        Do While SearchingMode And SearchLevelCounter(2) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(2)))
        Do While SearchingMode And SearchLevelCounter(3) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(3)))
        Do While SearchingMode And SearchLevelCounter(4) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(4)))
        Do While SearchingMode And SearchLevelCounter(5) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(5)))
            SearchTransactionCheckMatch
            If MatchFound = 0 Then
                If TrapThreshold Then
                    SearchTransactionCheckPartialMatch
                    If PartialMatchFound = 1 Then
                        ShowProgress
                        If SuspendTrapFinal Then
                            SearchParse (Manipulations(2, SearchLevelCounter(5)))
                            SearchLevelCounter(5) = SearchLevelCounter(5) + 1
                            Exit Sub
                        Else
                            AppendTrapFile
                        End If
                    End If
                End If
                SearchParse (Manipulations(2, SearchLevelCounter(5)))
                SearchLevelCounter(5) = SearchLevelCounter(5) + 1
            Else
                MatchFoundLabel.Caption = "Match Found!!!"
                MatchFoundLabel.Visible = True
                If Timer < SearchStartTime Then
                    SearchStartTime = SearchStartTime - 86400
                    'error condition is test crosses midnight
                End If
                SearchElapsedTime = SearchElapsedTime + Timer - SearchStartTime
                ShowElapsedTime
                ProgressLabel.Caption = Empty
                Exit Sub
            End If
            If SearchProgressCounter Mod SpeedMod = 0 Then
                SearchProgressCounter = 0
                If SearchProgress.Value < SearchProgress.Max Then
                    SearchProgress.Value = SearchProgress.Value + 1
                Else
                    FirstPass = False
                    SearchProgress.Value = 0
                    If Timer < SearchStartTime Then
                        SearchStartTime = SearchStartTime - 86400
                        'error condition is test crosses midnight
                    End If
                    ManipulationsPerSecond = (SpeedMod * 100) / (Timer - SearchStartTime)
                    SpeedMod = Int(ManipulationsPerSecond / 8)
                    SearchElapsedTime = SearchElapsedTime + Timer - SearchStartTime
                    ShowElapsedTime
                    ShowProgress
                    ShowEstimatedTime
                    SearchStartTime = Timer
                End If
                DoEvents
                If SearchContinueReady = 1 Then
                    Exit Sub
                End If
            End If
            SearchProgressCounter = SearchProgressCounter + 1
        Loop 'level 5
            SearchParse (Manipulations(2, SearchLevelCounter(4)))
            SearchLevelCounter(5) = 1
            SearchLevelCounter(4) = SearchLevelCounter(4) + 1
        Loop 'level 4
            SearchParse (Manipulations(2, SearchLevelCounter(3)))
            SearchLevelCounter(4) = 1
            SearchLevelCounter(3) = SearchLevelCounter(3) + 1
        Loop 'level 3
            SearchParse (Manipulations(2, SearchLevelCounter(2)))
            SearchLevelCounter(3) = 1
            SearchLevelCounter(2) = SearchLevelCounter(2) + 1
        Loop 'level 2
            SearchLevelCounter(2) = 1
            SearchLevelCounter(1) = SearchLevelCounter(1) + 1
    Loop 'level 1
    SearchLevelCounter(1) = 1
    SearchCurrentLevel = SearchCurrentLevel + 1
End If
End Sub


Private Sub SearchLevel6()
'SEARCH SIXTH LEVEL (6)
If SearchCurrentLevel = 6 And _
  SearchCurrentLevel <= Val(SearchLevelsTextBox.Text) Then
    Do While SearchingMode And SearchLevelCounter(1) <= SearchCounterMax
        ResetCurrentDeck
        SearchParse (Manipulations(1, SearchLevelCounter(1)))
        Do While SearchingMode And SearchLevelCounter(2) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(2)))
        Do While SearchingMode And SearchLevelCounter(3) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(3)))
        Do While SearchingMode And SearchLevelCounter(4) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(4)))
        Do While SearchingMode And SearchLevelCounter(5) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(5)))
        Do While SearchingMode And SearchLevelCounter(6) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(6)))
            SearchTransactionCheckMatch
            If MatchFound = 0 Then
                If TrapThreshold Then
                    SearchTransactionCheckPartialMatch
                    If PartialMatchFound = 1 Then
                        ShowProgress
                        If SuspendTrapFinal Then
                            SearchParse (Manipulations(2, SearchLevelCounter(6)))
                            SearchLevelCounter(6) = SearchLevelCounter(6) + 1
                            Exit Sub
                        Else
                            AppendTrapFile
                        End If
                    End If
                End If
                SearchParse (Manipulations(2, SearchLevelCounter(6)))
                SearchLevelCounter(6) = SearchLevelCounter(6) + 1
            Else
                MatchFoundLabel.Caption = "Match Found!!!"
                MatchFoundLabel.Visible = True
                If Timer < SearchStartTime Then
                    SearchStartTime = SearchStartTime - 86400
                    'error condition is test crosses midnight
                End If
                SearchElapsedTime = SearchElapsedTime + Timer - SearchStartTime
                ShowElapsedTime
                ProgressLabel.Caption = Empty
                Exit Sub
            End If
            If SearchProgressCounter Mod SpeedMod = 0 Then
                SearchProgressCounter = 0
                If SearchProgress.Value < SearchProgress.Max Then
                    SearchProgress.Value = SearchProgress.Value + 1
                Else
                    FirstPass = False
                    SearchProgress.Value = 0
                    If Timer < SearchStartTime Then
                        SearchStartTime = SearchStartTime - 86400
                        'error condition is test crosses midnight
                    End If
                    ManipulationsPerSecond = (SpeedMod * 100) / (Timer - SearchStartTime)
                    SpeedMod = Int(ManipulationsPerSecond / 8)
                    SearchElapsedTime = SearchElapsedTime + Timer - SearchStartTime
                    ShowElapsedTime
                    ShowProgress
                    ShowEstimatedTime
                    SearchStartTime = Timer
                End If
                DoEvents
                If SearchContinueReady = 1 Then
                    Exit Sub
                End If
            End If
            SearchProgressCounter = SearchProgressCounter + 1
        Loop 'level 6
            SearchParse (Manipulations(2, SearchLevelCounter(5)))
            SearchLevelCounter(6) = 1
            SearchLevelCounter(5) = SearchLevelCounter(5) + 1
        Loop 'level 5
            SearchParse (Manipulations(2, SearchLevelCounter(4)))
            SearchLevelCounter(5) = 1
            SearchLevelCounter(4) = SearchLevelCounter(4) + 1
        Loop 'level 4
            SearchParse (Manipulations(2, SearchLevelCounter(3)))
            SearchLevelCounter(4) = 1
            SearchLevelCounter(3) = SearchLevelCounter(3) + 1
        Loop 'level 3
            SearchParse (Manipulations(2, SearchLevelCounter(2)))
            SearchLevelCounter(3) = 1
            SearchLevelCounter(2) = SearchLevelCounter(2) + 1
        Loop 'level 2
            SearchLevelCounter(2) = 1
            SearchLevelCounter(1) = SearchLevelCounter(1) + 1
    Loop 'level 1
    SearchLevelCounter(1) = 1
    SearchCurrentLevel = SearchCurrentLevel + 1
End If
End Sub


Private Sub SearchLevel7()
'SEARCH SEVENTH LEVEL (7)
If SearchCurrentLevel = 7 And _
  SearchCurrentLevel <= Val(SearchLevelsTextBox.Text) Then
    Do While SearchingMode And SearchLevelCounter(1) <= SearchCounterMax
        ResetCurrentDeck
        SearchParse (Manipulations(1, SearchLevelCounter(1)))
        Do While SearchingMode And SearchLevelCounter(2) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(2)))
        Do While SearchingMode And SearchLevelCounter(3) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(3)))
        Do While SearchingMode And SearchLevelCounter(4) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(4)))
        Do While SearchingMode And SearchLevelCounter(5) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(5)))
        Do While SearchingMode And SearchLevelCounter(6) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(6)))
        Do While SearchingMode And SearchLevelCounter(7) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(7)))
            SearchTransactionCheckMatch
            If MatchFound = 0 Then
                If TrapThreshold Then
                    SearchTransactionCheckPartialMatch
                    If PartialMatchFound = 1 Then
                        ShowProgress
                        If SuspendTrapFinal Then
                            SearchParse (Manipulations(2, SearchLevelCounter(7)))
                            SearchLevelCounter(7) = SearchLevelCounter(7) + 1
                            Exit Sub
                        Else
                            AppendTrapFile
                        End If
                    End If
                End If
                SearchParse (Manipulations(2, SearchLevelCounter(7)))
                SearchLevelCounter(7) = SearchLevelCounter(7) + 1
            Else
                MatchFoundLabel.Caption = "Match Found!!!"
                MatchFoundLabel.Visible = True
                If Timer < SearchStartTime Then
                    SearchStartTime = SearchStartTime - 86400
                    'error condition is test crosses midnight
                End If
                SearchElapsedTime = SearchElapsedTime + Timer - SearchStartTime
                ShowElapsedTime
                ProgressLabel.Caption = Empty
                Exit Sub
            End If
            If SearchProgressCounter Mod SpeedMod = 0 Then
                SearchProgressCounter = 0
                If SearchProgress.Value < SearchProgress.Max Then
                    SearchProgress.Value = SearchProgress.Value + 1
                Else
                    FirstPass = False
                    SearchProgress.Value = 0
                    If Timer < SearchStartTime Then
                        SearchStartTime = SearchStartTime - 86400
                        'error condition is test crosses midnight
                    End If
                    ManipulationsPerSecond = (SpeedMod * 100) / (Timer - SearchStartTime)
                    SpeedMod = Int(ManipulationsPerSecond / 8)
                    SearchElapsedTime = SearchElapsedTime + Timer - SearchStartTime
                    ShowElapsedTime
                    ShowProgress
                    ShowEstimatedTime
                    SearchStartTime = Timer
                End If
                DoEvents
                If SearchContinueReady = 1 Then
                    Exit Sub
                End If
            End If
            SearchProgressCounter = SearchProgressCounter + 1
        Loop 'level 7
            SearchParse (Manipulations(2, SearchLevelCounter(6)))
            SearchLevelCounter(7) = 1
            SearchLevelCounter(6) = SearchLevelCounter(6) + 1
        Loop 'level 6
            SearchParse (Manipulations(2, SearchLevelCounter(5)))
            SearchLevelCounter(6) = 1
            SearchLevelCounter(5) = SearchLevelCounter(5) + 1
        Loop 'level 5
            SearchParse (Manipulations(2, SearchLevelCounter(4)))
            SearchLevelCounter(5) = 1
            SearchLevelCounter(4) = SearchLevelCounter(4) + 1
        Loop 'level 4
            SearchParse (Manipulations(2, SearchLevelCounter(3)))
            SearchLevelCounter(4) = 1
            SearchLevelCounter(3) = SearchLevelCounter(3) + 1
        Loop 'level 3
            SearchParse (Manipulations(2, SearchLevelCounter(2)))
            SearchLevelCounter(3) = 1
            SearchLevelCounter(2) = SearchLevelCounter(2) + 1
        Loop 'level 2
            SearchLevelCounter(2) = 1
            SearchLevelCounter(1) = SearchLevelCounter(1) + 1
    Loop 'level 1
    SearchLevelCounter(1) = 1
    SearchCurrentLevel = SearchCurrentLevel + 1
End If
End Sub


Private Sub SearchLevel8()
'SEARCH EIGHTH LEVEL (8)
If SearchCurrentLevel = 8 And _
  SearchCurrentLevel <= Val(SearchLevelsTextBox.Text) Then
    Do While SearchingMode And SearchLevelCounter(1) <= SearchCounterMax
        ResetCurrentDeck
        SearchParse (Manipulations(1, SearchLevelCounter(1)))
        Do While SearchingMode And SearchLevelCounter(2) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(2)))
        Do While SearchingMode And SearchLevelCounter(3) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(3)))
        Do While SearchingMode And SearchLevelCounter(4) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(4)))
        Do While SearchingMode And SearchLevelCounter(5) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(5)))
        Do While SearchingMode And SearchLevelCounter(6) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(6)))
        Do While SearchingMode And SearchLevelCounter(7) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(7)))
        Do While SearchingMode And SearchLevelCounter(8) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(8)))
            SearchTransactionCheckMatch
            If MatchFound = 0 Then
                If TrapThreshold Then
                    SearchTransactionCheckPartialMatch
                    If PartialMatchFound = 1 Then
                        ShowProgress
                        If SuspendTrapFinal Then
                            SearchParse (Manipulations(2, SearchLevelCounter(8)))
                            SearchLevelCounter(8) = SearchLevelCounter(8) + 1
                            Exit Sub
                        Else
                            AppendTrapFile
                        End If
                    End If
                End If
                SearchParse (Manipulations(2, SearchLevelCounter(8)))
                SearchLevelCounter(8) = SearchLevelCounter(8) + 1
            Else
                MatchFoundLabel.Caption = "Match Found!!!"
                MatchFoundLabel.Visible = True
                If Timer < SearchStartTime Then
                    SearchStartTime = SearchStartTime - 86400
                    'error condition is test crosses midnight
                End If
                SearchElapsedTime = SearchElapsedTime + Timer - SearchStartTime
                ShowElapsedTime
                ProgressLabel.Caption = Empty
                Exit Sub
            End If
            If SearchProgressCounter Mod SpeedMod = 0 Then
                SearchProgressCounter = 0
                If SearchProgress.Value < SearchProgress.Max Then
                    SearchProgress.Value = SearchProgress.Value + 1
                Else
                    FirstPass = False
                    SearchProgress.Value = 0
                    If Timer < SearchStartTime Then
                        SearchStartTime = SearchStartTime - 86400
                        'error condition is test crosses midnight
                    End If
                    ManipulationsPerSecond = (SpeedMod * 100) / (Timer - SearchStartTime)
                    SpeedMod = Int(ManipulationsPerSecond / 8)
                    SearchElapsedTime = SearchElapsedTime + Timer - SearchStartTime
                    ShowElapsedTime
                    ShowProgress
                    ShowEstimatedTime
                    SearchStartTime = Timer
                End If
                DoEvents
                If SearchContinueReady = 1 Then
                    Exit Sub
                End If
            End If
            SearchProgressCounter = SearchProgressCounter + 1
        Loop 'level 8
            SearchParse (Manipulations(2, SearchLevelCounter(7)))
            SearchLevelCounter(8) = 1
            SearchLevelCounter(7) = SearchLevelCounter(7) + 1
        Loop 'level 7
            SearchParse (Manipulations(2, SearchLevelCounter(6)))
            SearchLevelCounter(7) = 1
            SearchLevelCounter(6) = SearchLevelCounter(6) + 1
        Loop 'level 6
            SearchParse (Manipulations(2, SearchLevelCounter(5)))
            SearchLevelCounter(6) = 1
            SearchLevelCounter(5) = SearchLevelCounter(5) + 1
        Loop 'level 5
            SearchParse (Manipulations(2, SearchLevelCounter(4)))
            SearchLevelCounter(5) = 1
            SearchLevelCounter(4) = SearchLevelCounter(4) + 1
        Loop 'level 4
            SearchParse (Manipulations(2, SearchLevelCounter(3)))
            SearchLevelCounter(4) = 1
            SearchLevelCounter(3) = SearchLevelCounter(3) + 1
        Loop 'level 3
            SearchParse (Manipulations(2, SearchLevelCounter(2)))
            SearchLevelCounter(3) = 1
            SearchLevelCounter(2) = SearchLevelCounter(2) + 1
        Loop 'level 2
            SearchLevelCounter(2) = 1
            SearchLevelCounter(1) = SearchLevelCounter(1) + 1
    Loop 'level 1
    SearchLevelCounter(1) = 1
    SearchCurrentLevel = SearchCurrentLevel + 1
End If
End Sub


Private Sub SearchLevel9()
'SEARCH NINTH LEVEL (9)
If SearchCurrentLevel = 9 And _
  SearchCurrentLevel <= Val(SearchLevelsTextBox.Text) Then
    Do While SearchingMode And SearchLevelCounter(1) <= SearchCounterMax
        ResetCurrentDeck
        SearchParse (Manipulations(1, SearchLevelCounter(1)))
        Do While SearchingMode And SearchLevelCounter(2) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(2)))
        Do While SearchingMode And SearchLevelCounter(3) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(3)))
        Do While SearchingMode And SearchLevelCounter(4) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(4)))
        Do While SearchingMode And SearchLevelCounter(5) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(5)))
        Do While SearchingMode And SearchLevelCounter(6) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(6)))
        Do While SearchingMode And SearchLevelCounter(7) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(7)))
        Do While SearchingMode And SearchLevelCounter(8) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(8)))
        Do While SearchingMode And SearchLevelCounter(9) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(9)))
            SearchTransactionCheckMatch
            If MatchFound = 0 Then
                If TrapThreshold Then
                    SearchTransactionCheckPartialMatch
                    If PartialMatchFound = 1 Then
                        ShowProgress
                        If SuspendTrapFinal Then
                            SearchParse (Manipulations(2, SearchLevelCounter(9)))
                            SearchLevelCounter(9) = SearchLevelCounter(9) + 1
                            Exit Sub
                        Else
                            AppendTrapFile
                        End If
                    End If
                End If
                SearchParse (Manipulations(2, SearchLevelCounter(9)))
                SearchLevelCounter(9) = SearchLevelCounter(9) + 1
            Else
                MatchFoundLabel.Caption = "Match Found!!!"
                MatchFoundLabel.Visible = True
                If Timer < SearchStartTime Then
                    SearchStartTime = SearchStartTime - 86400
                    'error condition is test crosses midnight
                End If
                SearchElapsedTime = SearchElapsedTime + Timer - SearchStartTime
                ShowElapsedTime
                ProgressLabel.Caption = Empty
                Exit Sub
            End If
            If SearchProgressCounter Mod SpeedMod = 0 Then
                SearchProgressCounter = 0
                If SearchProgress.Value < SearchProgress.Max Then
                    SearchProgress.Value = SearchProgress.Value + 1
                Else
                    FirstPass = False
                    SearchProgress.Value = 0
                    If Timer < SearchStartTime Then
                        SearchStartTime = SearchStartTime - 86400
                        'error condition is test crosses midnight
                    End If
                    ManipulationsPerSecond = (SpeedMod * 100) / (Timer - SearchStartTime)
                    SpeedMod = Int(ManipulationsPerSecond / 8)
                    SearchElapsedTime = SearchElapsedTime + Timer - SearchStartTime
                    ShowElapsedTime
                    ShowProgress
                    ShowEstimatedTime
                    SearchStartTime = Timer
                End If
                DoEvents
                If SearchContinueReady = 1 Then
                    Exit Sub
                End If
            End If
            SearchProgressCounter = SearchProgressCounter + 1
        Loop 'level 9
            SearchParse (Manipulations(2, SearchLevelCounter(8)))
            SearchLevelCounter(9) = 1
            SearchLevelCounter(8) = SearchLevelCounter(8) + 1
        Loop 'level 8
            SearchParse (Manipulations(2, SearchLevelCounter(7)))
            SearchLevelCounter(8) = 1
            SearchLevelCounter(7) = SearchLevelCounter(7) + 1
        Loop 'level 7
            SearchParse (Manipulations(2, SearchLevelCounter(6)))
            SearchLevelCounter(7) = 1
            SearchLevelCounter(6) = SearchLevelCounter(6) + 1
        Loop 'level 6
            SearchParse (Manipulations(2, SearchLevelCounter(5)))
            SearchLevelCounter(6) = 1
            SearchLevelCounter(5) = SearchLevelCounter(5) + 1
        Loop 'level 5
            SearchParse (Manipulations(2, SearchLevelCounter(4)))
            SearchLevelCounter(5) = 1
            SearchLevelCounter(4) = SearchLevelCounter(4) + 1
        Loop 'level 4
            SearchParse (Manipulations(2, SearchLevelCounter(3)))
            SearchLevelCounter(4) = 1
            SearchLevelCounter(3) = SearchLevelCounter(3) + 1
        Loop 'level 3
            SearchParse (Manipulations(2, SearchLevelCounter(2)))
            SearchLevelCounter(3) = 1
            SearchLevelCounter(2) = SearchLevelCounter(2) + 1
        Loop 'level 2
            SearchLevelCounter(2) = 1
            SearchLevelCounter(1) = SearchLevelCounter(1) + 1
    Loop 'level 1
    SearchLevelCounter(1) = 1
    SearchCurrentLevel = SearchCurrentLevel + 1
End If
End Sub


Private Sub SearchLevel10()
'SEARCH TENTH LEVEL (10)
If SearchCurrentLevel = 10 And _
  SearchCurrentLevel <= Val(SearchLevelsTextBox.Text) Then
    Do While SearchingMode And SearchLevelCounter(1) <= SearchCounterMax
        ResetCurrentDeck
        SearchParse (Manipulations(1, SearchLevelCounter(1)))
        Do While SearchingMode And SearchLevelCounter(2) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(2)))
        Do While SearchingMode And SearchLevelCounter(3) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(3)))
        Do While SearchingMode And SearchLevelCounter(4) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(4)))
        Do While SearchingMode And SearchLevelCounter(5) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(5)))
        Do While SearchingMode And SearchLevelCounter(6) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(6)))
        Do While SearchingMode And SearchLevelCounter(7) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(7)))
        Do While SearchingMode And SearchLevelCounter(8) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(8)))
        Do While SearchingMode And SearchLevelCounter(9) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(9)))
        Do While SearchingMode And SearchLevelCounter(10) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(10)))
            SearchTransactionCheckMatch
            If MatchFound = 0 Then
                If TrapThreshold Then
                    SearchTransactionCheckPartialMatch
                    If PartialMatchFound = 1 Then
                        ShowProgress
                        If SuspendTrapFinal Then
                            SearchParse (Manipulations(2, SearchLevelCounter(10)))
                            SearchLevelCounter(10) = SearchLevelCounter(10) + 1
                            Exit Sub
                        Else
                            AppendTrapFile
                        End If
                    End If
                End If
                SearchParse (Manipulations(2, SearchLevelCounter(10)))
                SearchLevelCounter(10) = SearchLevelCounter(10) + 1
            Else
                MatchFoundLabel.Caption = "Match Found!!!"
                MatchFoundLabel.Visible = True
                If Timer < SearchStartTime Then
                    SearchStartTime = SearchStartTime - 86400
                    'error condition is test crosses midnight
                End If
                SearchElapsedTime = SearchElapsedTime + Timer - SearchStartTime
                ShowElapsedTime
                ProgressLabel.Caption = Empty
                Exit Sub
            End If
            If SearchProgressCounter Mod SpeedMod = 0 Then
                SearchProgressCounter = 0
                If SearchProgress.Value < SearchProgress.Max Then
                    SearchProgress.Value = SearchProgress.Value + 1
                Else
                    FirstPass = False
                    SearchProgress.Value = 0
                    If Timer < SearchStartTime Then
                        SearchStartTime = SearchStartTime - 86400
                        'error condition is test crosses midnight
                    End If
                    ManipulationsPerSecond = (SpeedMod * 100) / (Timer - SearchStartTime)
                    SpeedMod = Int(ManipulationsPerSecond / 8)
                    SearchElapsedTime = SearchElapsedTime + Timer - SearchStartTime
                    ShowElapsedTime
                    ShowProgress
                    ShowEstimatedTime
                    SearchStartTime = Timer
                End If
                DoEvents
                If SearchContinueReady = 1 Then
                    Exit Sub
                End If
            End If
            SearchProgressCounter = SearchProgressCounter + 1
        Loop 'level 10
            SearchParse (Manipulations(2, SearchLevelCounter(9)))
            SearchLevelCounter(10) = 1
            SearchLevelCounter(9) = SearchLevelCounter(9) + 1
        Loop 'level 9
            SearchParse (Manipulations(2, SearchLevelCounter(8)))
            SearchLevelCounter(9) = 1
            SearchLevelCounter(8) = SearchLevelCounter(8) + 1
        Loop 'level 8
            SearchParse (Manipulations(2, SearchLevelCounter(7)))
            SearchLevelCounter(8) = 1
            SearchLevelCounter(7) = SearchLevelCounter(7) + 1
        Loop 'level 7
            SearchParse (Manipulations(2, SearchLevelCounter(6)))
            SearchLevelCounter(7) = 1
            SearchLevelCounter(6) = SearchLevelCounter(6) + 1
        Loop 'level 6
            SearchParse (Manipulations(2, SearchLevelCounter(5)))
            SearchLevelCounter(6) = 1
            SearchLevelCounter(5) = SearchLevelCounter(5) + 1
        Loop 'level 5
            SearchParse (Manipulations(2, SearchLevelCounter(4)))
            SearchLevelCounter(5) = 1
            SearchLevelCounter(4) = SearchLevelCounter(4) + 1
        Loop 'level 4
            SearchParse (Manipulations(2, SearchLevelCounter(3)))
            SearchLevelCounter(4) = 1
            SearchLevelCounter(3) = SearchLevelCounter(3) + 1
        Loop 'level 3
            SearchParse (Manipulations(2, SearchLevelCounter(2)))
            SearchLevelCounter(3) = 1
            SearchLevelCounter(2) = SearchLevelCounter(2) + 1
        Loop 'level 2
            SearchLevelCounter(2) = 1
            SearchLevelCounter(1) = SearchLevelCounter(1) + 1
    Loop 'level 1
    SearchLevelCounter(1) = 1
    SearchCurrentLevel = SearchCurrentLevel + 1
End If
End Sub


Private Sub SearchLevel11()
'SEARCH ELEVENTH LEVEL (11)
If SearchCurrentLevel = 11 And _
  SearchCurrentLevel <= Val(SearchLevelsTextBox.Text) Then
    Do While SearchingMode And SearchLevelCounter(1) <= SearchCounterMax
        ResetCurrentDeck
        SearchParse (Manipulations(1, SearchLevelCounter(1)))
        Do While SearchingMode And SearchLevelCounter(2) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(2)))
        Do While SearchingMode And SearchLevelCounter(3) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(3)))
        Do While SearchingMode And SearchLevelCounter(4) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(4)))
        Do While SearchingMode And SearchLevelCounter(5) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(5)))
        Do While SearchingMode And SearchLevelCounter(6) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(6)))
        Do While SearchingMode And SearchLevelCounter(7) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(7)))
        Do While SearchingMode And SearchLevelCounter(8) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(8)))
        Do While SearchingMode And SearchLevelCounter(9) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(9)))
        Do While SearchingMode And SearchLevelCounter(10) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(10)))
        Do While SearchingMode And SearchLevelCounter(11) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(11)))
            SearchTransactionCheckMatch
            If MatchFound = 0 Then
                If TrapThreshold Then
                    SearchTransactionCheckPartialMatch
                    If PartialMatchFound = 1 Then
                        ShowProgress
                        If SuspendTrapFinal Then
                            SearchParse (Manipulations(2, SearchLevelCounter(11)))
                            SearchLevelCounter(11) = SearchLevelCounter(11) + 1
                            Exit Sub
                        Else
                            AppendTrapFile
                        End If
                    End If
                End If
                SearchParse (Manipulations(2, SearchLevelCounter(11)))
                SearchLevelCounter(11) = SearchLevelCounter(11) + 1
            Else
                MatchFoundLabel.Caption = "Match Found!!!"
                MatchFoundLabel.Visible = True
                If Timer < SearchStartTime Then
                    SearchStartTime = SearchStartTime - 86400
                    'error condition is test crosses midnight
                End If
                SearchElapsedTime = SearchElapsedTime + Timer - SearchStartTime
                ShowElapsedTime
                ProgressLabel.Caption = Empty
                Exit Sub
            End If
            If SearchProgressCounter Mod SpeedMod = 0 Then
                SearchProgressCounter = 0
                If SearchProgress.Value < SearchProgress.Max Then
                    SearchProgress.Value = SearchProgress.Value + 1
                Else
                    FirstPass = False
                    SearchProgress.Value = 0
                    If Timer < SearchStartTime Then
                        SearchStartTime = SearchStartTime - 86400
                        'error condition is test crosses midnight
                    End If
                    ManipulationsPerSecond = (SpeedMod * 100) / (Timer - SearchStartTime)
                    SpeedMod = Int(ManipulationsPerSecond / 8)
                    SearchElapsedTime = SearchElapsedTime + Timer - SearchStartTime
                    ShowElapsedTime
                    ShowProgress
                    ShowEstimatedTime
                    SearchStartTime = Timer
                End If
                DoEvents
                If SearchContinueReady = 1 Then
                    Exit Sub
                End If
            End If
            SearchProgressCounter = SearchProgressCounter + 1
        Loop 'level 11
            SearchParse (Manipulations(2, SearchLevelCounter(10)))
            SearchLevelCounter(11) = 1
            SearchLevelCounter(10) = SearchLevelCounter(10) + 1
        Loop 'level 10
            SearchParse (Manipulations(2, SearchLevelCounter(9)))
            SearchLevelCounter(10) = 1
            SearchLevelCounter(9) = SearchLevelCounter(9) + 1
        Loop 'level 9
            SearchParse (Manipulations(2, SearchLevelCounter(8)))
            SearchLevelCounter(9) = 1
            SearchLevelCounter(8) = SearchLevelCounter(8) + 1
        Loop 'level 8
            SearchParse (Manipulations(2, SearchLevelCounter(7)))
            SearchLevelCounter(8) = 1
            SearchLevelCounter(7) = SearchLevelCounter(7) + 1
        Loop 'level 7
            SearchParse (Manipulations(2, SearchLevelCounter(6)))
            SearchLevelCounter(7) = 1
            SearchLevelCounter(6) = SearchLevelCounter(6) + 1
        Loop 'level 6
            SearchParse (Manipulations(2, SearchLevelCounter(5)))
            SearchLevelCounter(6) = 1
            SearchLevelCounter(5) = SearchLevelCounter(5) + 1
        Loop 'level 5
            SearchParse (Manipulations(2, SearchLevelCounter(4)))
            SearchLevelCounter(5) = 1
            SearchLevelCounter(4) = SearchLevelCounter(4) + 1
        Loop 'level 4
            SearchParse (Manipulations(2, SearchLevelCounter(3)))
            SearchLevelCounter(4) = 1
            SearchLevelCounter(3) = SearchLevelCounter(3) + 1
        Loop 'level 3
            SearchParse (Manipulations(2, SearchLevelCounter(2)))
            SearchLevelCounter(3) = 1
            SearchLevelCounter(2) = SearchLevelCounter(2) + 1
        Loop 'level 2
            SearchLevelCounter(2) = 1
            SearchLevelCounter(1) = SearchLevelCounter(1) + 1
    Loop 'level 1
    SearchLevelCounter(1) = 1
    SearchCurrentLevel = SearchCurrentLevel + 1
End If
End Sub


Private Sub SearchLevel12()
'SEARCH TWELVTH LEVEL (12)
If SearchCurrentLevel = 12 And _
  SearchCurrentLevel <= Val(SearchLevelsTextBox.Text) Then
    Do While SearchingMode And SearchLevelCounter(1) <= SearchCounterMax
        ResetCurrentDeck
        SearchParse (Manipulations(1, SearchLevelCounter(1)))
        Do While SearchingMode And SearchLevelCounter(2) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(2)))
        Do While SearchingMode And SearchLevelCounter(3) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(3)))
        Do While SearchingMode And SearchLevelCounter(4) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(4)))
        Do While SearchingMode And SearchLevelCounter(5) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(5)))
        Do While SearchingMode And SearchLevelCounter(6) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(6)))
        Do While SearchingMode And SearchLevelCounter(7) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(7)))
        Do While SearchingMode And SearchLevelCounter(8) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(8)))
        Do While SearchingMode And SearchLevelCounter(9) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(9)))
        Do While SearchingMode And SearchLevelCounter(10) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(10)))
        Do While SearchingMode And SearchLevelCounter(11) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(11)))
        Do While SearchingMode And SearchLevelCounter(12) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(12)))
            SearchTransactionCheckMatch
            If MatchFound = 0 Then
                If TrapThreshold Then
                    SearchTransactionCheckPartialMatch
                    If PartialMatchFound = 1 Then
                        ShowProgress
                        If SuspendTrapFinal Then
                            SearchParse (Manipulations(2, SearchLevelCounter(12)))
                            SearchLevelCounter(12) = SearchLevelCounter(12) + 1
                            Exit Sub
                        Else
                            AppendTrapFile
                        End If
                    End If
                End If
                SearchParse (Manipulations(2, SearchLevelCounter(12)))
                SearchLevelCounter(12) = SearchLevelCounter(12) + 1
            Else
                MatchFoundLabel.Caption = "Match Found!!!"
                MatchFoundLabel.Visible = True
                If Timer < SearchStartTime Then
                    SearchStartTime = SearchStartTime - 86400
                    'error condition is test crosses midnight
                End If
                SearchElapsedTime = SearchElapsedTime + Timer - SearchStartTime
                ShowElapsedTime
                ProgressLabel.Caption = Empty
                Exit Sub
            End If
            If SearchProgressCounter Mod SpeedMod = 0 Then
                SearchProgressCounter = 0
                If SearchProgress.Value < SearchProgress.Max Then
                    SearchProgress.Value = SearchProgress.Value + 1
                Else
                    FirstPass = False
                    SearchProgress.Value = 0
                    If Timer < SearchStartTime Then
                        SearchStartTime = SearchStartTime - 86400
                        'error condition is test crosses midnight
                    End If
                    ManipulationsPerSecond = (SpeedMod * 100) / (Timer - SearchStartTime)
                    SpeedMod = Int(ManipulationsPerSecond / 8)
                    SearchElapsedTime = SearchElapsedTime + Timer - SearchStartTime
                    ShowElapsedTime
                    ShowProgress
                    ShowEstimatedTime
                    SearchStartTime = Timer
                End If
                DoEvents
                If SearchContinueReady = 1 Then
                    Exit Sub
                End If
            End If
            SearchProgressCounter = SearchProgressCounter + 1
        Loop 'level 12
            SearchParse (Manipulations(2, SearchLevelCounter(11)))
            SearchLevelCounter(12) = 1
            SearchLevelCounter(11) = SearchLevelCounter(11) + 1
        Loop 'level 11
            SearchParse (Manipulations(2, SearchLevelCounter(10)))
            SearchLevelCounter(11) = 1
            SearchLevelCounter(10) = SearchLevelCounter(10) + 1
        Loop 'level 10
            SearchParse (Manipulations(2, SearchLevelCounter(9)))
            SearchLevelCounter(10) = 1
            SearchLevelCounter(9) = SearchLevelCounter(9) + 1
        Loop 'level 9
            SearchParse (Manipulations(2, SearchLevelCounter(8)))
            SearchLevelCounter(9) = 1
            SearchLevelCounter(8) = SearchLevelCounter(8) + 1
        Loop 'level 8
            SearchParse (Manipulations(2, SearchLevelCounter(7)))
            SearchLevelCounter(8) = 1
            SearchLevelCounter(7) = SearchLevelCounter(7) + 1
        Loop 'level 7
            SearchParse (Manipulations(2, SearchLevelCounter(6)))
            SearchLevelCounter(7) = 1
            SearchLevelCounter(6) = SearchLevelCounter(6) + 1
        Loop 'level 6
            SearchParse (Manipulations(2, SearchLevelCounter(5)))
            SearchLevelCounter(6) = 1
            SearchLevelCounter(5) = SearchLevelCounter(5) + 1
        Loop 'level 5
            SearchParse (Manipulations(2, SearchLevelCounter(4)))
            SearchLevelCounter(5) = 1
            SearchLevelCounter(4) = SearchLevelCounter(4) + 1
        Loop 'level 4
            SearchParse (Manipulations(2, SearchLevelCounter(3)))
            SearchLevelCounter(4) = 1
            SearchLevelCounter(3) = SearchLevelCounter(3) + 1
        Loop 'level 3
            SearchParse (Manipulations(2, SearchLevelCounter(2)))
            SearchLevelCounter(3) = 1
            SearchLevelCounter(2) = SearchLevelCounter(2) + 1
        Loop 'level 2
            SearchLevelCounter(2) = 1
            SearchLevelCounter(1) = SearchLevelCounter(1) + 1
    Loop 'level 1
    SearchLevelCounter(1) = 1
    SearchCurrentLevel = SearchCurrentLevel + 1
End If
End Sub


Private Sub SearchLevel13()
'SEARCH THIRTEENTH LEVEL (13)
If SearchCurrentLevel = 13 And _
  SearchCurrentLevel <= Val(SearchLevelsTextBox.Text) Then
    Do While SearchingMode And SearchLevelCounter(1) <= SearchCounterMax
        ResetCurrentDeck
        SearchParse (Manipulations(1, SearchLevelCounter(1)))
        Do While SearchingMode And SearchLevelCounter(2) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(2)))
        Do While SearchingMode And SearchLevelCounter(3) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(3)))
        Do While SearchingMode And SearchLevelCounter(4) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(4)))
        Do While SearchingMode And SearchLevelCounter(5) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(5)))
        Do While SearchingMode And SearchLevelCounter(6) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(6)))
        Do While SearchingMode And SearchLevelCounter(7) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(7)))
        Do While SearchingMode And SearchLevelCounter(8) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(8)))
        Do While SearchingMode And SearchLevelCounter(9) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(9)))
        Do While SearchingMode And SearchLevelCounter(10) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(10)))
        Do While SearchingMode And SearchLevelCounter(11) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(11)))
        Do While SearchingMode And SearchLevelCounter(12) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(12)))
        Do While SearchingMode And SearchLevelCounter(13) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(13)))
            SearchTransactionCheckMatch
            If MatchFound = 0 Then
                If TrapThreshold Then
                    SearchTransactionCheckPartialMatch
                    If PartialMatchFound = 1 Then
                        ShowProgress
                        If SuspendTrapFinal Then
                            SearchParse (Manipulations(2, SearchLevelCounter(13)))
                            SearchLevelCounter(13) = SearchLevelCounter(13) + 1
                            Exit Sub
                        Else
                            AppendTrapFile
                        End If
                    End If
                End If
                SearchParse (Manipulations(2, SearchLevelCounter(13)))
                SearchLevelCounter(13) = SearchLevelCounter(13) + 1
            Else
                MatchFoundLabel.Caption = "Match Found!!!"
                MatchFoundLabel.Visible = True
                If Timer < SearchStartTime Then
                    SearchStartTime = SearchStartTime - 86400
                    'error condition is test crosses midnight
                End If
                SearchElapsedTime = SearchElapsedTime + Timer - SearchStartTime
                ShowElapsedTime
                ProgressLabel.Caption = Empty
                Exit Sub
            End If
            If SearchProgressCounter Mod SpeedMod = 0 Then
                SearchProgressCounter = 0
                If SearchProgress.Value < SearchProgress.Max Then
                    SearchProgress.Value = SearchProgress.Value + 1
                Else
                    FirstPass = False
                    SearchProgress.Value = 0
                    If Timer < SearchStartTime Then
                        SearchStartTime = SearchStartTime - 86400
                        'error condition is test crosses midnight
                    End If
                    ManipulationsPerSecond = (SpeedMod * 100) / (Timer - SearchStartTime)
                    SpeedMod = Int(ManipulationsPerSecond / 8)
                    SearchElapsedTime = SearchElapsedTime + Timer - SearchStartTime
                    ShowElapsedTime
                    ShowProgress
                    ShowEstimatedTime
                    SearchStartTime = Timer
                End If
                DoEvents
                If SearchContinueReady = 1 Then
                    Exit Sub
                End If
            End If
            SearchProgressCounter = SearchProgressCounter + 1
        Loop 'level 13
            SearchParse (Manipulations(2, SearchLevelCounter(12)))
            SearchLevelCounter(13) = 1
            SearchLevelCounter(12) = SearchLevelCounter(12) + 1
        Loop 'level 12
            SearchParse (Manipulations(2, SearchLevelCounter(11)))
            SearchLevelCounter(12) = 1
            SearchLevelCounter(11) = SearchLevelCounter(11) + 1
        Loop 'level 11
            SearchParse (Manipulations(2, SearchLevelCounter(10)))
            SearchLevelCounter(11) = 1
            SearchLevelCounter(10) = SearchLevelCounter(10) + 1
        Loop 'level 10
            SearchParse (Manipulations(2, SearchLevelCounter(9)))
            SearchLevelCounter(10) = 1
            SearchLevelCounter(9) = SearchLevelCounter(9) + 1
        Loop 'level 9
            SearchParse (Manipulations(2, SearchLevelCounter(8)))
            SearchLevelCounter(9) = 1
            SearchLevelCounter(8) = SearchLevelCounter(8) + 1
        Loop 'level 8
            SearchParse (Manipulations(2, SearchLevelCounter(7)))
            SearchLevelCounter(8) = 1
            SearchLevelCounter(7) = SearchLevelCounter(7) + 1
        Loop 'level 7
            SearchParse (Manipulations(2, SearchLevelCounter(6)))
            SearchLevelCounter(7) = 1
            SearchLevelCounter(6) = SearchLevelCounter(6) + 1
        Loop 'level 6
            SearchParse (Manipulations(2, SearchLevelCounter(5)))
            SearchLevelCounter(6) = 1
            SearchLevelCounter(5) = SearchLevelCounter(5) + 1
        Loop 'level 5
            SearchParse (Manipulations(2, SearchLevelCounter(4)))
            SearchLevelCounter(5) = 1
            SearchLevelCounter(4) = SearchLevelCounter(4) + 1
        Loop 'level 4
            SearchParse (Manipulations(2, SearchLevelCounter(3)))
            SearchLevelCounter(4) = 1
            SearchLevelCounter(3) = SearchLevelCounter(3) + 1
        Loop 'level 3
            SearchParse (Manipulations(2, SearchLevelCounter(2)))
            SearchLevelCounter(3) = 1
            SearchLevelCounter(2) = SearchLevelCounter(2) + 1
        Loop 'level 2
            SearchLevelCounter(2) = 1
            SearchLevelCounter(1) = SearchLevelCounter(1) + 1
    Loop 'level 1
    SearchLevelCounter(1) = 1
    SearchCurrentLevel = SearchCurrentLevel + 1
End If
End Sub


Private Sub SearchLevel14()
'SEARCH FOURTEENTH LEVEL (14)
If SearchCurrentLevel = 14 And _
  SearchCurrentLevel <= Val(SearchLevelsTextBox.Text) Then
    Do While SearchingMode And SearchLevelCounter(1) <= SearchCounterMax
        ResetCurrentDeck
        SearchParse (Manipulations(1, SearchLevelCounter(1)))
        Do While SearchingMode And SearchLevelCounter(2) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(2)))
        Do While SearchingMode And SearchLevelCounter(3) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(3)))
        Do While SearchingMode And SearchLevelCounter(4) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(4)))
        Do While SearchingMode And SearchLevelCounter(5) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(5)))
        Do While SearchingMode And SearchLevelCounter(6) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(6)))
        Do While SearchingMode And SearchLevelCounter(7) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(7)))
        Do While SearchingMode And SearchLevelCounter(8) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(8)))
        Do While SearchingMode And SearchLevelCounter(9) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(9)))
        Do While SearchingMode And SearchLevelCounter(10) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(10)))
        Do While SearchingMode And SearchLevelCounter(11) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(11)))
        Do While SearchingMode And SearchLevelCounter(12) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(12)))
        Do While SearchingMode And SearchLevelCounter(13) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(13)))
        Do While SearchingMode And SearchLevelCounter(14) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(14)))
            SearchTransactionCheckMatch
            If MatchFound = 0 Then
                If TrapThreshold Then
                    SearchTransactionCheckPartialMatch
                    If PartialMatchFound = 1 Then
                        ShowProgress
                        If SuspendTrapFinal Then
                            SearchParse (Manipulations(2, SearchLevelCounter(14)))
                            SearchLevelCounter(14) = SearchLevelCounter(14) + 1
                            Exit Sub
                        Else
                            AppendTrapFile
                        End If
                    End If
                End If
                SearchParse (Manipulations(2, SearchLevelCounter(14)))
                SearchLevelCounter(14) = SearchLevelCounter(14) + 1
            Else
                MatchFoundLabel.Caption = "Match Found!!!"
                MatchFoundLabel.Visible = True
                If Timer < SearchStartTime Then
                    SearchStartTime = SearchStartTime - 86400
                    'error condition is test crosses midnight
                End If
                SearchElapsedTime = SearchElapsedTime + Timer - SearchStartTime
                ShowElapsedTime
                ProgressLabel.Caption = Empty
                Exit Sub
            End If
            If SearchProgressCounter Mod SpeedMod = 0 Then
                SearchProgressCounter = 0
                If SearchProgress.Value < SearchProgress.Max Then
                    SearchProgress.Value = SearchProgress.Value + 1
                Else
                    FirstPass = False
                    SearchProgress.Value = 0
                    If Timer < SearchStartTime Then
                        SearchStartTime = SearchStartTime - 86400
                        'error condition is test crosses midnight
                    End If
                    ManipulationsPerSecond = (SpeedMod * 100) / (Timer - SearchStartTime)
                    SpeedMod = Int(ManipulationsPerSecond / 8)
                    SearchElapsedTime = SearchElapsedTime + Timer - SearchStartTime
                    ShowElapsedTime
                    ShowProgress
                    ShowEstimatedTime
                    SearchStartTime = Timer
                End If
                DoEvents
                If SearchContinueReady = 1 Then
                    Exit Sub
                End If
            End If
            SearchProgressCounter = SearchProgressCounter + 1
        Loop 'level 14
            SearchParse (Manipulations(2, SearchLevelCounter(13)))
            SearchLevelCounter(14) = 1
            SearchLevelCounter(13) = SearchLevelCounter(13) + 1
        Loop 'level 13
            SearchParse (Manipulations(2, SearchLevelCounter(12)))
            SearchLevelCounter(13) = 1
            SearchLevelCounter(12) = SearchLevelCounter(12) + 1
        Loop 'level 12
            SearchParse (Manipulations(2, SearchLevelCounter(11)))
            SearchLevelCounter(12) = 1
            SearchLevelCounter(11) = SearchLevelCounter(11) + 1
        Loop 'level 11
            SearchParse (Manipulations(2, SearchLevelCounter(10)))
            SearchLevelCounter(11) = 1
            SearchLevelCounter(10) = SearchLevelCounter(10) + 1
        Loop 'level 10
            SearchParse (Manipulations(2, SearchLevelCounter(9)))
            SearchLevelCounter(10) = 1
            SearchLevelCounter(9) = SearchLevelCounter(9) + 1
        Loop 'level 9
            SearchParse (Manipulations(2, SearchLevelCounter(8)))
            SearchLevelCounter(9) = 1
            SearchLevelCounter(8) = SearchLevelCounter(8) + 1
        Loop 'level 8
            SearchParse (Manipulations(2, SearchLevelCounter(7)))
            SearchLevelCounter(8) = 1
            SearchLevelCounter(7) = SearchLevelCounter(7) + 1
        Loop 'level 7
            SearchParse (Manipulations(2, SearchLevelCounter(6)))
            SearchLevelCounter(7) = 1
            SearchLevelCounter(6) = SearchLevelCounter(6) + 1
        Loop 'level 6
            SearchParse (Manipulations(2, SearchLevelCounter(5)))
            SearchLevelCounter(6) = 1
            SearchLevelCounter(5) = SearchLevelCounter(5) + 1
        Loop 'level 5
            SearchParse (Manipulations(2, SearchLevelCounter(4)))
            SearchLevelCounter(5) = 1
            SearchLevelCounter(4) = SearchLevelCounter(4) + 1
        Loop 'level 4
            SearchParse (Manipulations(2, SearchLevelCounter(3)))
            SearchLevelCounter(4) = 1
            SearchLevelCounter(3) = SearchLevelCounter(3) + 1
        Loop 'level 3
            SearchParse (Manipulations(2, SearchLevelCounter(2)))
            SearchLevelCounter(3) = 1
            SearchLevelCounter(2) = SearchLevelCounter(2) + 1
        Loop 'level 2
            SearchLevelCounter(2) = 1
            SearchLevelCounter(1) = SearchLevelCounter(1) + 1
    Loop 'level 1
    SearchLevelCounter(1) = 1
    SearchCurrentLevel = SearchCurrentLevel + 1
End If
End Sub


Private Sub SearchLevel15()
'SEARCH FIFTEENTH LEVEL (15)
If SearchCurrentLevel = 15 And _
  SearchCurrentLevel <= Val(SearchLevelsTextBox.Text) Then
    Do While SearchingMode And SearchLevelCounter(1) <= SearchCounterMax
        ResetCurrentDeck
        SearchParse (Manipulations(1, SearchLevelCounter(1)))
        Do While SearchingMode And SearchLevelCounter(2) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(2)))
        Do While SearchingMode And SearchLevelCounter(3) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(3)))
        Do While SearchingMode And SearchLevelCounter(4) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(4)))
        Do While SearchingMode And SearchLevelCounter(5) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(5)))
        Do While SearchingMode And SearchLevelCounter(6) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(6)))
        Do While SearchingMode And SearchLevelCounter(7) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(7)))
        Do While SearchingMode And SearchLevelCounter(8) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(8)))
        Do While SearchingMode And SearchLevelCounter(9) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(9)))
        Do While SearchingMode And SearchLevelCounter(10) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(10)))
        Do While SearchingMode And SearchLevelCounter(11) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(11)))
        Do While SearchingMode And SearchLevelCounter(12) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(12)))
        Do While SearchingMode And SearchLevelCounter(13) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(13)))
        Do While SearchingMode And SearchLevelCounter(14) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(14)))
        Do While SearchingMode And SearchLevelCounter(15) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(15)))
            SearchTransactionCheckMatch
            If MatchFound = 0 Then
                If TrapThreshold Then
                    SearchTransactionCheckPartialMatch
                    If PartialMatchFound = 1 Then
                        ShowProgress
                        If SuspendTrapFinal Then
                            SearchParse (Manipulations(2, SearchLevelCounter(15)))
                            SearchLevelCounter(15) = SearchLevelCounter(15) + 1
                            Exit Sub
                        Else
                            AppendTrapFile
                        End If
                    End If
                End If
                SearchParse (Manipulations(2, SearchLevelCounter(15)))
                SearchLevelCounter(15) = SearchLevelCounter(15) + 1
            Else
                MatchFoundLabel.Caption = "Match Found!!!"
                MatchFoundLabel.Visible = True
                If Timer < SearchStartTime Then
                    SearchStartTime = SearchStartTime - 86400
                    'error condition is test crosses midnight
                End If
                SearchElapsedTime = SearchElapsedTime + Timer - SearchStartTime
                ShowElapsedTime
                ProgressLabel.Caption = Empty
                Exit Sub
            End If
            If SearchProgressCounter Mod SpeedMod = 0 Then
                SearchProgressCounter = 0
                If SearchProgress.Value < SearchProgress.Max Then
                    SearchProgress.Value = SearchProgress.Value + 1
                Else
                    FirstPass = False
                    SearchProgress.Value = 0
                    If Timer < SearchStartTime Then
                        SearchStartTime = SearchStartTime - 86400
                        'error condition is test crosses midnight
                    End If
                    ManipulationsPerSecond = (SpeedMod * 100) / (Timer - SearchStartTime)
                    SpeedMod = Int(ManipulationsPerSecond / 8)
                    SearchElapsedTime = SearchElapsedTime + Timer - SearchStartTime
                    ShowElapsedTime
                    ShowProgress
                    ShowEstimatedTime
                    SearchStartTime = Timer
                End If
                DoEvents
                If SearchContinueReady = 1 Then
                    Exit Sub
                End If
            End If
            SearchProgressCounter = SearchProgressCounter + 1
        Loop 'level 15
            SearchParse (Manipulations(2, SearchLevelCounter(14)))
            SearchLevelCounter(15) = 1
            SearchLevelCounter(14) = SearchLevelCounter(14) + 1
        Loop 'level 14
            SearchParse (Manipulations(2, SearchLevelCounter(13)))
            SearchLevelCounter(14) = 1
            SearchLevelCounter(13) = SearchLevelCounter(13) + 1
        Loop 'level 13
            SearchParse (Manipulations(2, SearchLevelCounter(12)))
            SearchLevelCounter(13) = 1
            SearchLevelCounter(12) = SearchLevelCounter(12) + 1
        Loop 'level 12
            SearchParse (Manipulations(2, SearchLevelCounter(11)))
            SearchLevelCounter(12) = 1
            SearchLevelCounter(11) = SearchLevelCounter(11) + 1
        Loop 'level 11
            SearchParse (Manipulations(2, SearchLevelCounter(10)))
            SearchLevelCounter(11) = 1
            SearchLevelCounter(10) = SearchLevelCounter(10) + 1
        Loop 'level 10
            SearchParse (Manipulations(2, SearchLevelCounter(9)))
            SearchLevelCounter(10) = 1
            SearchLevelCounter(9) = SearchLevelCounter(9) + 1
        Loop 'level 9
            SearchParse (Manipulations(2, SearchLevelCounter(8)))
            SearchLevelCounter(9) = 1
            SearchLevelCounter(8) = SearchLevelCounter(8) + 1
        Loop 'level 8
            SearchParse (Manipulations(2, SearchLevelCounter(7)))
            SearchLevelCounter(8) = 1
            SearchLevelCounter(7) = SearchLevelCounter(7) + 1
        Loop 'level 7
            SearchParse (Manipulations(2, SearchLevelCounter(6)))
            SearchLevelCounter(7) = 1
            SearchLevelCounter(6) = SearchLevelCounter(6) + 1
        Loop 'level 6
            SearchParse (Manipulations(2, SearchLevelCounter(5)))
            SearchLevelCounter(6) = 1
            SearchLevelCounter(5) = SearchLevelCounter(5) + 1
        Loop 'level 5
            SearchParse (Manipulations(2, SearchLevelCounter(4)))
            SearchLevelCounter(5) = 1
            SearchLevelCounter(4) = SearchLevelCounter(4) + 1
        Loop 'level 4
            SearchParse (Manipulations(2, SearchLevelCounter(3)))
            SearchLevelCounter(4) = 1
            SearchLevelCounter(3) = SearchLevelCounter(3) + 1
        Loop 'level 3
            SearchParse (Manipulations(2, SearchLevelCounter(2)))
            SearchLevelCounter(3) = 1
            SearchLevelCounter(2) = SearchLevelCounter(2) + 1
        Loop 'level 2
            SearchLevelCounter(2) = 1
            SearchLevelCounter(1) = SearchLevelCounter(1) + 1
    Loop 'level 1
    SearchLevelCounter(1) = 1
    SearchCurrentLevel = SearchCurrentLevel + 1
End If
End Sub

Private Sub SearchLevel16()
'SEARCH SIXTEENTH LEVEL (16)
If SearchCurrentLevel = 16 And _
  SearchCurrentLevel <= Val(SearchLevelsTextBox.Text) Then
    Do While SearchingMode And SearchLevelCounter(1) <= SearchCounterMax
        ResetCurrentDeck
        SearchParse (Manipulations(1, SearchLevelCounter(1)))
        Do While SearchingMode And SearchLevelCounter(2) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(2)))
        Do While SearchingMode And SearchLevelCounter(3) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(3)))
        Do While SearchingMode And SearchLevelCounter(4) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(4)))
        Do While SearchingMode And SearchLevelCounter(5) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(5)))
        Do While SearchingMode And SearchLevelCounter(6) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(6)))
        Do While SearchingMode And SearchLevelCounter(7) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(7)))
        Do While SearchingMode And SearchLevelCounter(8) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(8)))
        Do While SearchingMode And SearchLevelCounter(9) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(9)))
        Do While SearchingMode And SearchLevelCounter(10) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(10)))
        Do While SearchingMode And SearchLevelCounter(11) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(11)))
        Do While SearchingMode And SearchLevelCounter(12) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(12)))
        Do While SearchingMode And SearchLevelCounter(13) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(13)))
        Do While SearchingMode And SearchLevelCounter(14) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(14)))
        Do While SearchingMode And SearchLevelCounter(15) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(15)))
        Do While SearchingMode And SearchLevelCounter(16) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(16)))
            SearchTransactionCheckMatch
            If MatchFound = 0 Then
                If TrapThreshold Then
                    SearchTransactionCheckPartialMatch
                    If PartialMatchFound = 1 Then
                        ShowProgress
                        If SuspendTrapFinal Then
                            SearchParse (Manipulations(2, SearchLevelCounter(16)))
                            SearchLevelCounter(16) = SearchLevelCounter(16) + 1
                            Exit Sub
                        Else
                            AppendTrapFile
                        End If
                    End If
                End If
                SearchParse (Manipulations(2, SearchLevelCounter(16)))
                SearchLevelCounter(16) = SearchLevelCounter(16) + 1
            Else
                MatchFoundLabel.Caption = "Match Found!!!"
                MatchFoundLabel.Visible = True
                If Timer < SearchStartTime Then
                    SearchStartTime = SearchStartTime - 86400
                    'error condition is test crosses midnight
                End If
                SearchElapsedTime = SearchElapsedTime + Timer - SearchStartTime
                ShowElapsedTime
                ProgressLabel.Caption = Empty
                Exit Sub
            End If
            If SearchProgressCounter Mod SpeedMod = 0 Then
                SearchProgressCounter = 0
                If SearchProgress.Value < SearchProgress.Max Then
                    SearchProgress.Value = SearchProgress.Value + 1
                Else
                    FirstPass = False
                    SearchProgress.Value = 0
                    If Timer < SearchStartTime Then
                        SearchStartTime = SearchStartTime - 86400
                        'error condition is test crosses midnight
                    End If
                    ManipulationsPerSecond = (SpeedMod * 100) / (Timer - SearchStartTime)
                    SpeedMod = Int(ManipulationsPerSecond / 8)
                    SearchElapsedTime = SearchElapsedTime + Timer - SearchStartTime
                    ShowElapsedTime
                    ShowProgress
                    ShowEstimatedTime
                    SearchStartTime = Timer
                End If
                DoEvents
                If SearchContinueReady = 1 Then
                    Exit Sub
                End If
            End If
            SearchProgressCounter = SearchProgressCounter + 1
        Loop 'level 16
            SearchParse (Manipulations(2, SearchLevelCounter(15)))
            SearchLevelCounter(16) = 1
            SearchLevelCounter(15) = SearchLevelCounter(15) + 1
        Loop 'level 15
            SearchParse (Manipulations(2, SearchLevelCounter(14)))
            SearchLevelCounter(15) = 1
            SearchLevelCounter(14) = SearchLevelCounter(14) + 1
        Loop 'level 14
            SearchParse (Manipulations(2, SearchLevelCounter(13)))
            SearchLevelCounter(14) = 1
            SearchLevelCounter(13) = SearchLevelCounter(13) + 1
        Loop 'level 13
            SearchParse (Manipulations(2, SearchLevelCounter(12)))
            SearchLevelCounter(13) = 1
            SearchLevelCounter(12) = SearchLevelCounter(12) + 1
        Loop 'level 12
            SearchParse (Manipulations(2, SearchLevelCounter(11)))
            SearchLevelCounter(12) = 1
            SearchLevelCounter(11) = SearchLevelCounter(11) + 1
        Loop 'level 11
            SearchParse (Manipulations(2, SearchLevelCounter(10)))
            SearchLevelCounter(11) = 1
            SearchLevelCounter(10) = SearchLevelCounter(10) + 1
        Loop 'level 10
            SearchParse (Manipulations(2, SearchLevelCounter(9)))
            SearchLevelCounter(10) = 1
            SearchLevelCounter(9) = SearchLevelCounter(9) + 1
        Loop 'level 9
            SearchParse (Manipulations(2, SearchLevelCounter(8)))
            SearchLevelCounter(9) = 1
            SearchLevelCounter(8) = SearchLevelCounter(8) + 1
        Loop 'level 8
            SearchParse (Manipulations(2, SearchLevelCounter(7)))
            SearchLevelCounter(8) = 1
            SearchLevelCounter(7) = SearchLevelCounter(7) + 1
        Loop 'level 7
            SearchParse (Manipulations(2, SearchLevelCounter(6)))
            SearchLevelCounter(7) = 1
            SearchLevelCounter(6) = SearchLevelCounter(6) + 1
        Loop 'level 6
            SearchParse (Manipulations(2, SearchLevelCounter(5)))
            SearchLevelCounter(6) = 1
            SearchLevelCounter(5) = SearchLevelCounter(5) + 1
        Loop 'level 5
            SearchParse (Manipulations(2, SearchLevelCounter(4)))
            SearchLevelCounter(5) = 1
            SearchLevelCounter(4) = SearchLevelCounter(4) + 1
        Loop 'level 4
            SearchParse (Manipulations(2, SearchLevelCounter(3)))
            SearchLevelCounter(4) = 1
            SearchLevelCounter(3) = SearchLevelCounter(3) + 1
        Loop 'level 3
            SearchParse (Manipulations(2, SearchLevelCounter(2)))
            SearchLevelCounter(3) = 1
            SearchLevelCounter(2) = SearchLevelCounter(2) + 1
        Loop 'level 2
            SearchLevelCounter(2) = 1
            SearchLevelCounter(1) = SearchLevelCounter(1) + 1
    Loop 'level 1
    SearchLevelCounter(1) = 1
    SearchCurrentLevel = SearchCurrentLevel + 1
End If
End Sub


Private Sub SearchLevel17()
'SEARCH SEVENTEENTH LEVEL (17)
If SearchCurrentLevel = 17 And _
  SearchCurrentLevel <= Val(SearchLevelsTextBox.Text) Then
    Do While SearchingMode And SearchLevelCounter(1) <= SearchCounterMax
        ResetCurrentDeck
        SearchParse (Manipulations(1, SearchLevelCounter(1)))
        Do While SearchingMode And SearchLevelCounter(2) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(2)))
        Do While SearchingMode And SearchLevelCounter(3) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(3)))
        Do While SearchingMode And SearchLevelCounter(4) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(4)))
        Do While SearchingMode And SearchLevelCounter(5) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(5)))
        Do While SearchingMode And SearchLevelCounter(6) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(6)))
        Do While SearchingMode And SearchLevelCounter(7) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(7)))
        Do While SearchingMode And SearchLevelCounter(8) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(8)))
        Do While SearchingMode And SearchLevelCounter(9) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(9)))
        Do While SearchingMode And SearchLevelCounter(10) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(10)))
        Do While SearchingMode And SearchLevelCounter(11) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(11)))
        Do While SearchingMode And SearchLevelCounter(12) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(12)))
        Do While SearchingMode And SearchLevelCounter(13) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(13)))
        Do While SearchingMode And SearchLevelCounter(14) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(14)))
        Do While SearchingMode And SearchLevelCounter(15) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(15)))
        Do While SearchingMode And SearchLevelCounter(16) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(16)))
        Do While SearchingMode And SearchLevelCounter(17) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(17)))
            SearchTransactionCheckMatch
            If MatchFound = 0 Then
                If TrapThreshold Then
                    SearchTransactionCheckPartialMatch
                    If PartialMatchFound = 1 Then
                        ShowProgress
                        If SuspendTrapFinal Then
                            SearchParse (Manipulations(2, SearchLevelCounter(17)))
                            SearchLevelCounter(17) = SearchLevelCounter(17) + 1
                            Exit Sub
                        Else
                            AppendTrapFile
                        End If
                    End If
                End If
                SearchParse (Manipulations(2, SearchLevelCounter(17)))
                SearchLevelCounter(17) = SearchLevelCounter(17) + 1
            Else
                MatchFoundLabel.Caption = "Match Found!!!"
                MatchFoundLabel.Visible = True
                If Timer < SearchStartTime Then
                    SearchStartTime = SearchStartTime - 86400
                    'error condition is test crosses midnight
                End If
                SearchElapsedTime = SearchElapsedTime + Timer - SearchStartTime
                ShowElapsedTime
                ProgressLabel.Caption = Empty
                Exit Sub
            End If
            If SearchProgressCounter Mod SpeedMod = 0 Then
                SearchProgressCounter = 0
                If SearchProgress.Value < SearchProgress.Max Then
                    SearchProgress.Value = SearchProgress.Value + 1
                Else
                    FirstPass = False
                    SearchProgress.Value = 0
                    If Timer < SearchStartTime Then
                        SearchStartTime = SearchStartTime - 86400
                        'error condition is test crosses midnight
                    End If
                    ManipulationsPerSecond = (SpeedMod * 100) / (Timer - SearchStartTime)
                    SpeedMod = Int(ManipulationsPerSecond / 8)
                    SearchElapsedTime = SearchElapsedTime + Timer - SearchStartTime
                    ShowElapsedTime
                    ShowProgress
                    ShowEstimatedTime
                    SearchStartTime = Timer
                End If
                DoEvents
                If SearchContinueReady = 1 Then
                    Exit Sub
                End If
            End If
            SearchProgressCounter = SearchProgressCounter + 1
        Loop 'level 17
            SearchParse (Manipulations(2, SearchLevelCounter(16)))
            SearchLevelCounter(17) = 1
            SearchLevelCounter(16) = SearchLevelCounter(16) + 1
        Loop 'level 16
            SearchParse (Manipulations(2, SearchLevelCounter(15)))
            SearchLevelCounter(16) = 1
            SearchLevelCounter(15) = SearchLevelCounter(15) + 1
        Loop 'level 15
            SearchParse (Manipulations(2, SearchLevelCounter(14)))
            SearchLevelCounter(15) = 1
            SearchLevelCounter(14) = SearchLevelCounter(14) + 1
        Loop 'level 14
            SearchParse (Manipulations(2, SearchLevelCounter(13)))
            SearchLevelCounter(14) = 1
            SearchLevelCounter(13) = SearchLevelCounter(13) + 1
        Loop 'level 13
            SearchParse (Manipulations(2, SearchLevelCounter(12)))
            SearchLevelCounter(13) = 1
            SearchLevelCounter(12) = SearchLevelCounter(12) + 1
        Loop 'level 12
            SearchParse (Manipulations(2, SearchLevelCounter(11)))
            SearchLevelCounter(12) = 1
            SearchLevelCounter(11) = SearchLevelCounter(11) + 1
        Loop 'level 11
            SearchParse (Manipulations(2, SearchLevelCounter(10)))
            SearchLevelCounter(11) = 1
            SearchLevelCounter(10) = SearchLevelCounter(10) + 1
        Loop 'level 10
            SearchParse (Manipulations(2, SearchLevelCounter(9)))
            SearchLevelCounter(10) = 1
            SearchLevelCounter(9) = SearchLevelCounter(9) + 1
        Loop 'level 9
            SearchParse (Manipulations(2, SearchLevelCounter(8)))
            SearchLevelCounter(9) = 1
            SearchLevelCounter(8) = SearchLevelCounter(8) + 1
        Loop 'level 8
            SearchParse (Manipulations(2, SearchLevelCounter(7)))
            SearchLevelCounter(8) = 1
            SearchLevelCounter(7) = SearchLevelCounter(7) + 1
        Loop 'level 7
            SearchParse (Manipulations(2, SearchLevelCounter(6)))
            SearchLevelCounter(7) = 1
            SearchLevelCounter(6) = SearchLevelCounter(6) + 1
        Loop 'level 6
            SearchParse (Manipulations(2, SearchLevelCounter(5)))
            SearchLevelCounter(6) = 1
            SearchLevelCounter(5) = SearchLevelCounter(5) + 1
        Loop 'level 5
            SearchParse (Manipulations(2, SearchLevelCounter(4)))
            SearchLevelCounter(5) = 1
            SearchLevelCounter(4) = SearchLevelCounter(4) + 1
        Loop 'level 4
            SearchParse (Manipulations(2, SearchLevelCounter(3)))
            SearchLevelCounter(4) = 1
            SearchLevelCounter(3) = SearchLevelCounter(3) + 1
        Loop 'level 3
            SearchParse (Manipulations(2, SearchLevelCounter(2)))
            SearchLevelCounter(3) = 1
            SearchLevelCounter(2) = SearchLevelCounter(2) + 1
        Loop 'level 2
            SearchLevelCounter(2) = 1
            SearchLevelCounter(1) = SearchLevelCounter(1) + 1
    Loop 'level 1
    SearchLevelCounter(1) = 1
    SearchCurrentLevel = SearchCurrentLevel + 1
End If
End Sub


Private Sub SearchLevel18()
'SEARCH EIGHTEENTH LEVEL (18)
If SearchCurrentLevel = 18 And _
  SearchCurrentLevel <= Val(SearchLevelsTextBox.Text) Then
    Do While SearchingMode And SearchLevelCounter(1) <= SearchCounterMax
        ResetCurrentDeck
        SearchParse (Manipulations(1, SearchLevelCounter(1)))
        Do While SearchingMode And SearchLevelCounter(2) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(2)))
        Do While SearchingMode And SearchLevelCounter(3) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(3)))
        Do While SearchingMode And SearchLevelCounter(4) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(4)))
        Do While SearchingMode And SearchLevelCounter(5) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(5)))
        Do While SearchingMode And SearchLevelCounter(6) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(6)))
        Do While SearchingMode And SearchLevelCounter(7) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(7)))
        Do While SearchingMode And SearchLevelCounter(8) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(8)))
        Do While SearchingMode And SearchLevelCounter(9) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(9)))
        Do While SearchingMode And SearchLevelCounter(10) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(10)))
        Do While SearchingMode And SearchLevelCounter(11) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(11)))
        Do While SearchingMode And SearchLevelCounter(12) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(12)))
        Do While SearchingMode And SearchLevelCounter(13) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(13)))
        Do While SearchingMode And SearchLevelCounter(14) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(14)))
        Do While SearchingMode And SearchLevelCounter(15) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(15)))
        Do While SearchingMode And SearchLevelCounter(16) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(16)))
        Do While SearchingMode And SearchLevelCounter(17) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(17)))
        Do While SearchingMode And SearchLevelCounter(18) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(18)))
            SearchTransactionCheckMatch
            If MatchFound = 0 Then
                If TrapThreshold Then
                    SearchTransactionCheckPartialMatch
                    If PartialMatchFound = 1 Then
                        ShowProgress
                        If SuspendTrapFinal Then
                            SearchParse (Manipulations(2, SearchLevelCounter(18)))
                            SearchLevelCounter(18) = SearchLevelCounter(18) + 1
                            Exit Sub
                        Else
                            AppendTrapFile
                        End If
                    End If
                End If
                SearchParse (Manipulations(2, SearchLevelCounter(18)))
                SearchLevelCounter(18) = SearchLevelCounter(18) + 1
            Else
                MatchFoundLabel.Caption = "Match Found!!!"
                MatchFoundLabel.Visible = True
                If Timer < SearchStartTime Then
                    SearchStartTime = SearchStartTime - 86400
                    'error condition is test crosses midnight
                End If
                SearchElapsedTime = SearchElapsedTime + Timer - SearchStartTime
                ShowElapsedTime
                ProgressLabel.Caption = Empty
                Exit Sub
            End If
            If SearchProgressCounter Mod SpeedMod = 0 Then
                SearchProgressCounter = 0
                If SearchProgress.Value < SearchProgress.Max Then
                    SearchProgress.Value = SearchProgress.Value + 1
                Else
                    FirstPass = False
                    SearchProgress.Value = 0
                    If Timer < SearchStartTime Then
                        SearchStartTime = SearchStartTime - 86400
                        'error condition is test crosses midnight
                    End If
                    ManipulationsPerSecond = (SpeedMod * 100) / (Timer - SearchStartTime)
                    SpeedMod = Int(ManipulationsPerSecond / 8)
                    SearchElapsedTime = SearchElapsedTime + Timer - SearchStartTime
                    ShowElapsedTime
                    ShowProgress
                    ShowEstimatedTime
                    SearchStartTime = Timer
                End If
                DoEvents
                If SearchContinueReady = 1 Then
                    Exit Sub
                End If
            End If
            SearchProgressCounter = SearchProgressCounter + 1
        Loop 'level 18
            SearchParse (Manipulations(2, SearchLevelCounter(17)))
            SearchLevelCounter(18) = 1
            SearchLevelCounter(17) = SearchLevelCounter(17) + 1
        Loop 'level 17
            SearchParse (Manipulations(2, SearchLevelCounter(16)))
            SearchLevelCounter(17) = 1
            SearchLevelCounter(16) = SearchLevelCounter(16) + 1
        Loop 'level 16
            SearchParse (Manipulations(2, SearchLevelCounter(15)))
            SearchLevelCounter(16) = 1
            SearchLevelCounter(15) = SearchLevelCounter(15) + 1
        Loop 'level 15
            SearchParse (Manipulations(2, SearchLevelCounter(14)))
            SearchLevelCounter(15) = 1
            SearchLevelCounter(14) = SearchLevelCounter(14) + 1
        Loop 'level 14
            SearchParse (Manipulations(2, SearchLevelCounter(13)))
            SearchLevelCounter(14) = 1
            SearchLevelCounter(13) = SearchLevelCounter(13) + 1
        Loop 'level 13
            SearchParse (Manipulations(2, SearchLevelCounter(12)))
            SearchLevelCounter(13) = 1
            SearchLevelCounter(12) = SearchLevelCounter(12) + 1
        Loop 'level 12
            SearchParse (Manipulations(2, SearchLevelCounter(11)))
            SearchLevelCounter(12) = 1
            SearchLevelCounter(11) = SearchLevelCounter(11) + 1
        Loop 'level 11
            SearchParse (Manipulations(2, SearchLevelCounter(10)))
            SearchLevelCounter(11) = 1
            SearchLevelCounter(10) = SearchLevelCounter(10) + 1
        Loop 'level 10
            SearchParse (Manipulations(2, SearchLevelCounter(9)))
            SearchLevelCounter(10) = 1
            SearchLevelCounter(9) = SearchLevelCounter(9) + 1
        Loop 'level 9
            SearchParse (Manipulations(2, SearchLevelCounter(8)))
            SearchLevelCounter(9) = 1
            SearchLevelCounter(8) = SearchLevelCounter(8) + 1
        Loop 'level 8
            SearchParse (Manipulations(2, SearchLevelCounter(7)))
            SearchLevelCounter(8) = 1
            SearchLevelCounter(7) = SearchLevelCounter(7) + 1
        Loop 'level 7
            SearchParse (Manipulations(2, SearchLevelCounter(6)))
            SearchLevelCounter(7) = 1
            SearchLevelCounter(6) = SearchLevelCounter(6) + 1
        Loop 'level 6
            SearchParse (Manipulations(2, SearchLevelCounter(5)))
            SearchLevelCounter(6) = 1
            SearchLevelCounter(5) = SearchLevelCounter(5) + 1
        Loop 'level 5
            SearchParse (Manipulations(2, SearchLevelCounter(4)))
            SearchLevelCounter(5) = 1
            SearchLevelCounter(4) = SearchLevelCounter(4) + 1
        Loop 'level 4
            SearchParse (Manipulations(2, SearchLevelCounter(3)))
            SearchLevelCounter(4) = 1
            SearchLevelCounter(3) = SearchLevelCounter(3) + 1
        Loop 'level 3
            SearchParse (Manipulations(2, SearchLevelCounter(2)))
            SearchLevelCounter(3) = 1
            SearchLevelCounter(2) = SearchLevelCounter(2) + 1
        Loop 'level 2
            SearchLevelCounter(2) = 1
            SearchLevelCounter(1) = SearchLevelCounter(1) + 1
    Loop 'level 1
    SearchLevelCounter(1) = 1
    SearchCurrentLevel = SearchCurrentLevel + 1
End If
End Sub


Private Sub SearchLevel19()
'SEARCH NINETEENTH LEVEL (19)
If SearchCurrentLevel = 19 And _
  SearchCurrentLevel <= Val(SearchLevelsTextBox.Text) Then
    Do While SearchingMode And SearchLevelCounter(1) <= SearchCounterMax
        ResetCurrentDeck
        SearchParse (Manipulations(1, SearchLevelCounter(1)))
        Do While SearchingMode And SearchLevelCounter(2) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(2)))
        Do While SearchingMode And SearchLevelCounter(3) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(3)))
        Do While SearchingMode And SearchLevelCounter(4) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(4)))
        Do While SearchingMode And SearchLevelCounter(5) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(5)))
        Do While SearchingMode And SearchLevelCounter(6) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(6)))
        Do While SearchingMode And SearchLevelCounter(7) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(7)))
        Do While SearchingMode And SearchLevelCounter(8) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(8)))
        Do While SearchingMode And SearchLevelCounter(9) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(9)))
        Do While SearchingMode And SearchLevelCounter(10) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(10)))
        Do While SearchingMode And SearchLevelCounter(11) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(11)))
        Do While SearchingMode And SearchLevelCounter(12) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(12)))
        Do While SearchingMode And SearchLevelCounter(13) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(13)))
        Do While SearchingMode And SearchLevelCounter(14) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(14)))
        Do While SearchingMode And SearchLevelCounter(15) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(15)))
        Do While SearchingMode And SearchLevelCounter(16) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(16)))
        Do While SearchingMode And SearchLevelCounter(17) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(17)))
        Do While SearchingMode And SearchLevelCounter(18) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(18)))
        Do While SearchingMode And SearchLevelCounter(19) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(19)))
            SearchTransactionCheckMatch
            If MatchFound = 0 Then
                If TrapThreshold Then
                    SearchTransactionCheckPartialMatch
                    If PartialMatchFound = 1 Then
                        ShowProgress
                        If SuspendTrapFinal Then
                            SearchParse (Manipulations(2, SearchLevelCounter(19)))
                            SearchLevelCounter(19) = SearchLevelCounter(19) + 1
                            Exit Sub
                        Else
                            AppendTrapFile
                        End If
                    End If
                End If
                SearchParse (Manipulations(2, SearchLevelCounter(19)))
                SearchLevelCounter(19) = SearchLevelCounter(19) + 1
            Else
                MatchFoundLabel.Caption = "Match Found!!!"
                MatchFoundLabel.Visible = True
                If Timer < SearchStartTime Then
                    SearchStartTime = SearchStartTime - 86400
                    'error condition is test crosses midnight
                End If
                SearchElapsedTime = SearchElapsedTime + Timer - SearchStartTime
                ShowElapsedTime
                ProgressLabel.Caption = Empty
                Exit Sub
            End If
            If SearchProgressCounter Mod SpeedMod = 0 Then
                SearchProgressCounter = 0
                If SearchProgress.Value < SearchProgress.Max Then
                    SearchProgress.Value = SearchProgress.Value + 1
                Else
                    FirstPass = False
                    SearchProgress.Value = 0
                    If Timer < SearchStartTime Then
                        SearchStartTime = SearchStartTime - 86400
                        'error condition is test crosses midnight
                    End If
                    ManipulationsPerSecond = (SpeedMod * 100) / (Timer - SearchStartTime)
                    SpeedMod = Int(ManipulationsPerSecond / 8)
                    SearchElapsedTime = SearchElapsedTime + Timer - SearchStartTime
                    ShowElapsedTime
                    ShowProgress
                    ShowEstimatedTime
                    SearchStartTime = Timer
                End If
                DoEvents
                If SearchContinueReady = 1 Then
                    Exit Sub
                End If
            End If
            SearchProgressCounter = SearchProgressCounter + 1
        Loop 'level 19
            SearchParse (Manipulations(2, SearchLevelCounter(18)))
            SearchLevelCounter(19) = 1
            SearchLevelCounter(18) = SearchLevelCounter(18) + 1
        Loop 'level 18
            SearchParse (Manipulations(2, SearchLevelCounter(17)))
            SearchLevelCounter(18) = 1
            SearchLevelCounter(17) = SearchLevelCounter(17) + 1
        Loop 'level 17
            SearchParse (Manipulations(2, SearchLevelCounter(16)))
            SearchLevelCounter(17) = 1
            SearchLevelCounter(16) = SearchLevelCounter(16) + 1
        Loop 'level 16
            SearchParse (Manipulations(2, SearchLevelCounter(15)))
            SearchLevelCounter(16) = 1
            SearchLevelCounter(15) = SearchLevelCounter(15) + 1
        Loop 'level 15
            SearchParse (Manipulations(2, SearchLevelCounter(14)))
            SearchLevelCounter(15) = 1
            SearchLevelCounter(14) = SearchLevelCounter(14) + 1
        Loop 'level 14
            SearchParse (Manipulations(2, SearchLevelCounter(13)))
            SearchLevelCounter(14) = 1
            SearchLevelCounter(13) = SearchLevelCounter(13) + 1
        Loop 'level 13
            SearchParse (Manipulations(2, SearchLevelCounter(12)))
            SearchLevelCounter(13) = 1
            SearchLevelCounter(12) = SearchLevelCounter(12) + 1
        Loop 'level 12
            SearchParse (Manipulations(2, SearchLevelCounter(11)))
            SearchLevelCounter(12) = 1
            SearchLevelCounter(11) = SearchLevelCounter(11) + 1
        Loop 'level 11
            SearchParse (Manipulations(2, SearchLevelCounter(10)))
            SearchLevelCounter(11) = 1
            SearchLevelCounter(10) = SearchLevelCounter(10) + 1
        Loop 'level 10
            SearchParse (Manipulations(2, SearchLevelCounter(9)))
            SearchLevelCounter(10) = 1
            SearchLevelCounter(9) = SearchLevelCounter(9) + 1
        Loop 'level 9
            SearchParse (Manipulations(2, SearchLevelCounter(8)))
            SearchLevelCounter(9) = 1
            SearchLevelCounter(8) = SearchLevelCounter(8) + 1
        Loop 'level 8
            SearchParse (Manipulations(2, SearchLevelCounter(7)))
            SearchLevelCounter(8) = 1
            SearchLevelCounter(7) = SearchLevelCounter(7) + 1
        Loop 'level 7
            SearchParse (Manipulations(2, SearchLevelCounter(6)))
            SearchLevelCounter(7) = 1
            SearchLevelCounter(6) = SearchLevelCounter(6) + 1
        Loop 'level 6
            SearchParse (Manipulations(2, SearchLevelCounter(5)))
            SearchLevelCounter(6) = 1
            SearchLevelCounter(5) = SearchLevelCounter(5) + 1
        Loop 'level 5
            SearchParse (Manipulations(2, SearchLevelCounter(4)))
            SearchLevelCounter(5) = 1
            SearchLevelCounter(4) = SearchLevelCounter(4) + 1
        Loop 'level 4
            SearchParse (Manipulations(2, SearchLevelCounter(3)))
            SearchLevelCounter(4) = 1
            SearchLevelCounter(3) = SearchLevelCounter(3) + 1
        Loop 'level 3
            SearchParse (Manipulations(2, SearchLevelCounter(2)))
            SearchLevelCounter(3) = 1
            SearchLevelCounter(2) = SearchLevelCounter(2) + 1
        Loop 'level 2
            SearchLevelCounter(2) = 1
            SearchLevelCounter(1) = SearchLevelCounter(1) + 1
    Loop 'level 1
    SearchLevelCounter(1) = 1
    SearchCurrentLevel = SearchCurrentLevel + 1
End If
End Sub


Private Sub SearchLevel20()
'SEARCH TWENTIETH LEVEL (20)
If SearchCurrentLevel = 20 And _
  SearchCurrentLevel <= Val(SearchLevelsTextBox.Text) Then
    Do While SearchingMode And SearchLevelCounter(1) <= SearchCounterMax
        ResetCurrentDeck
        SearchParse (Manipulations(1, SearchLevelCounter(1)))
        Do While SearchingMode And SearchLevelCounter(2) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(2)))
        Do While SearchingMode And SearchLevelCounter(3) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(3)))
        Do While SearchingMode And SearchLevelCounter(4) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(4)))
        Do While SearchingMode And SearchLevelCounter(5) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(5)))
        Do While SearchingMode And SearchLevelCounter(6) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(6)))
        Do While SearchingMode And SearchLevelCounter(7) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(7)))
        Do While SearchingMode And SearchLevelCounter(8) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(8)))
        Do While SearchingMode And SearchLevelCounter(9) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(9)))
        Do While SearchingMode And SearchLevelCounter(10) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(10)))
        Do While SearchingMode And SearchLevelCounter(11) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(11)))
        Do While SearchingMode And SearchLevelCounter(12) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(12)))
        Do While SearchingMode And SearchLevelCounter(13) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(13)))
        Do While SearchingMode And SearchLevelCounter(14) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(14)))
        Do While SearchingMode And SearchLevelCounter(15) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(15)))
        Do While SearchingMode And SearchLevelCounter(16) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(16)))
        Do While SearchingMode And SearchLevelCounter(17) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(17)))
        Do While SearchingMode And SearchLevelCounter(18) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(18)))
        Do While SearchingMode And SearchLevelCounter(19) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(19)))
        Do While SearchingMode And SearchLevelCounter(20) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(20)))
            SearchTransactionCheckMatch
            If MatchFound = 0 Then
                If TrapThreshold Then
                    SearchTransactionCheckPartialMatch
                    If PartialMatchFound = 1 Then
                        ShowProgress
                        If SuspendTrapFinal Then
                            SearchParse (Manipulations(2, SearchLevelCounter(20)))
                            SearchLevelCounter(20) = SearchLevelCounter(20) + 1
                            Exit Sub
                        Else
                            AppendTrapFile
                        End If
                    End If
                End If
                SearchParse (Manipulations(2, SearchLevelCounter(20)))
                SearchLevelCounter(20) = SearchLevelCounter(20) + 1
            Else
                MatchFoundLabel.Caption = "Match Found!!!"
                MatchFoundLabel.Visible = True
                If Timer < SearchStartTime Then
                    SearchStartTime = SearchStartTime - 86400
                    'error condition is test crosses midnight
                End If
                SearchElapsedTime = SearchElapsedTime + Timer - SearchStartTime
                ShowElapsedTime
                ProgressLabel.Caption = Empty
                Exit Sub
            End If
            If SearchProgressCounter Mod SpeedMod = 0 Then
                SearchProgressCounter = 0
                If SearchProgress.Value < SearchProgress.Max Then
                    SearchProgress.Value = SearchProgress.Value + 1
                Else
                    FirstPass = False
                    SearchProgress.Value = 0
                    If Timer < SearchStartTime Then
                        SearchStartTime = SearchStartTime - 86400
                        'error condition is test crosses midnight
                    End If
                    ManipulationsPerSecond = (SpeedMod * 100) / (Timer - SearchStartTime)
                    SpeedMod = Int(ManipulationsPerSecond / 8)
                    SearchElapsedTime = SearchElapsedTime + Timer - SearchStartTime
                    ShowElapsedTime
                    ShowProgress
                    ShowEstimatedTime
                    SearchStartTime = Timer
                End If
                DoEvents
                If SearchContinueReady = 1 Then
                    Exit Sub
                End If
            End If
            SearchProgressCounter = SearchProgressCounter + 1
        Loop 'level 20
            SearchParse (Manipulations(2, SearchLevelCounter(19)))
            SearchLevelCounter(20) = 1
            SearchLevelCounter(19) = SearchLevelCounter(19) + 1
        Loop 'level 19
            SearchParse (Manipulations(2, SearchLevelCounter(18)))
            SearchLevelCounter(19) = 1
            SearchLevelCounter(18) = SearchLevelCounter(18) + 1
        Loop 'level 18
            SearchParse (Manipulations(2, SearchLevelCounter(17)))
            SearchLevelCounter(18) = 1
            SearchLevelCounter(17) = SearchLevelCounter(17) + 1
        Loop 'level 17
            SearchParse (Manipulations(2, SearchLevelCounter(16)))
            SearchLevelCounter(17) = 1
            SearchLevelCounter(16) = SearchLevelCounter(16) + 1
        Loop 'level 16
            SearchParse (Manipulations(2, SearchLevelCounter(15)))
            SearchLevelCounter(16) = 1
            SearchLevelCounter(15) = SearchLevelCounter(15) + 1
        Loop 'level 15
            SearchParse (Manipulations(2, SearchLevelCounter(14)))
            SearchLevelCounter(15) = 1
            SearchLevelCounter(14) = SearchLevelCounter(14) + 1
        Loop 'level 14
            SearchParse (Manipulations(2, SearchLevelCounter(13)))
            SearchLevelCounter(14) = 1
            SearchLevelCounter(13) = SearchLevelCounter(13) + 1
        Loop 'level 13
            SearchParse (Manipulations(2, SearchLevelCounter(12)))
            SearchLevelCounter(13) = 1
            SearchLevelCounter(12) = SearchLevelCounter(12) + 1
        Loop 'level 12
            SearchParse (Manipulations(2, SearchLevelCounter(11)))
            SearchLevelCounter(12) = 1
            SearchLevelCounter(11) = SearchLevelCounter(11) + 1
        Loop 'level 11
            SearchParse (Manipulations(2, SearchLevelCounter(10)))
            SearchLevelCounter(11) = 1
            SearchLevelCounter(10) = SearchLevelCounter(10) + 1
        Loop 'level 10
            SearchParse (Manipulations(2, SearchLevelCounter(9)))
            SearchLevelCounter(10) = 1
            SearchLevelCounter(9) = SearchLevelCounter(9) + 1
        Loop 'level 9
            SearchParse (Manipulations(2, SearchLevelCounter(8)))
            SearchLevelCounter(9) = 1
            SearchLevelCounter(8) = SearchLevelCounter(8) + 1
        Loop 'level 8
            SearchParse (Manipulations(2, SearchLevelCounter(7)))
            SearchLevelCounter(8) = 1
            SearchLevelCounter(7) = SearchLevelCounter(7) + 1
        Loop 'level 7
            SearchParse (Manipulations(2, SearchLevelCounter(6)))
            SearchLevelCounter(7) = 1
            SearchLevelCounter(6) = SearchLevelCounter(6) + 1
        Loop 'level 6
            SearchParse (Manipulations(2, SearchLevelCounter(5)))
            SearchLevelCounter(6) = 1
            SearchLevelCounter(5) = SearchLevelCounter(5) + 1
        Loop 'level 5
            SearchParse (Manipulations(2, SearchLevelCounter(4)))
            SearchLevelCounter(5) = 1
            SearchLevelCounter(4) = SearchLevelCounter(4) + 1
        Loop 'level 4
            SearchParse (Manipulations(2, SearchLevelCounter(3)))
            SearchLevelCounter(4) = 1
            SearchLevelCounter(3) = SearchLevelCounter(3) + 1
        Loop 'level 3
            SearchParse (Manipulations(2, SearchLevelCounter(2)))
            SearchLevelCounter(3) = 1
            SearchLevelCounter(2) = SearchLevelCounter(2) + 1
        Loop 'level 2
            SearchLevelCounter(2) = 1
            SearchLevelCounter(1) = SearchLevelCounter(1) + 1
    Loop 'level 1
    SearchLevelCounter(1) = 1
    SearchCurrentLevel = SearchCurrentLevel + 1
End If
End Sub


Private Sub SearchLevel21()
'SEARCH TWENTYFIRST LEVEL (21)
If SearchCurrentLevel = 21 And _
  SearchCurrentLevel <= Val(SearchLevelsTextBox.Text) Then
    Do While SearchingMode And SearchLevelCounter(1) <= SearchCounterMax
        ResetCurrentDeck
        SearchParse (Manipulations(1, SearchLevelCounter(1)))
        Do While SearchingMode And SearchLevelCounter(2) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(2)))
        Do While SearchingMode And SearchLevelCounter(3) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(3)))
        Do While SearchingMode And SearchLevelCounter(4) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(4)))
        Do While SearchingMode And SearchLevelCounter(5) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(5)))
        Do While SearchingMode And SearchLevelCounter(6) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(6)))
        Do While SearchingMode And SearchLevelCounter(7) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(7)))
        Do While SearchingMode And SearchLevelCounter(8) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(8)))
        Do While SearchingMode And SearchLevelCounter(9) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(9)))
        Do While SearchingMode And SearchLevelCounter(10) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(10)))
        Do While SearchingMode And SearchLevelCounter(11) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(11)))
        Do While SearchingMode And SearchLevelCounter(12) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(12)))
        Do While SearchingMode And SearchLevelCounter(13) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(13)))
        Do While SearchingMode And SearchLevelCounter(14) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(14)))
        Do While SearchingMode And SearchLevelCounter(15) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(15)))
        Do While SearchingMode And SearchLevelCounter(16) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(16)))
        Do While SearchingMode And SearchLevelCounter(17) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(17)))
        Do While SearchingMode And SearchLevelCounter(18) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(18)))
        Do While SearchingMode And SearchLevelCounter(19) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(19)))
        Do While SearchingMode And SearchLevelCounter(20) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(20)))
        Do While SearchingMode And SearchLevelCounter(21) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(21)))
            SearchTransactionCheckMatch
            If MatchFound = 0 Then
                If TrapThreshold Then
                    SearchTransactionCheckPartialMatch
                    If PartialMatchFound = 1 Then
                        ShowProgress
                        If SuspendTrapFinal Then
                            SearchParse (Manipulations(2, SearchLevelCounter(21)))
                            SearchLevelCounter(21) = SearchLevelCounter(21) + 1
                            Exit Sub
                        Else
                            AppendTrapFile
                        End If
                    End If
                End If
                SearchParse (Manipulations(2, SearchLevelCounter(21)))
                SearchLevelCounter(21) = SearchLevelCounter(21) + 1
            Else
                MatchFoundLabel.Caption = "Match Found!!!"
                MatchFoundLabel.Visible = True
                If Timer < SearchStartTime Then
                    SearchStartTime = SearchStartTime - 86400
                    'error condition is test crosses midnight
                End If
                SearchElapsedTime = SearchElapsedTime + Timer - SearchStartTime
                ShowElapsedTime
                ProgressLabel.Caption = Empty
                Exit Sub
            End If
            If SearchProgressCounter Mod SpeedMod = 0 Then
                SearchProgressCounter = 0
                If SearchProgress.Value < SearchProgress.Max Then
                    SearchProgress.Value = SearchProgress.Value + 1
                Else
                    FirstPass = False
                    SearchProgress.Value = 0
                    If Timer < SearchStartTime Then
                        SearchStartTime = SearchStartTime - 86400
                        'error condition is test crosses midnight
                    End If
                    ManipulationsPerSecond = (SpeedMod * 100) / (Timer - SearchStartTime)
                    SpeedMod = Int(ManipulationsPerSecond / 8)
                    SearchElapsedTime = SearchElapsedTime + Timer - SearchStartTime
                    ShowElapsedTime
                    ShowProgress
                    ShowEstimatedTime
                    SearchStartTime = Timer
                End If
                DoEvents
                If SearchContinueReady = 1 Then
                    Exit Sub
                End If
            End If
            SearchProgressCounter = SearchProgressCounter + 1
        Loop 'level 21
            SearchParse (Manipulations(2, SearchLevelCounter(20)))
            SearchLevelCounter(21) = 1
            SearchLevelCounter(20) = SearchLevelCounter(20) + 1
        Loop 'level 20
            SearchParse (Manipulations(2, SearchLevelCounter(19)))
            SearchLevelCounter(20) = 1
            SearchLevelCounter(19) = SearchLevelCounter(19) + 1
        Loop 'level 19
            SearchParse (Manipulations(2, SearchLevelCounter(18)))
            SearchLevelCounter(19) = 1
            SearchLevelCounter(18) = SearchLevelCounter(18) + 1
        Loop 'level 18
            SearchParse (Manipulations(2, SearchLevelCounter(17)))
            SearchLevelCounter(18) = 1
            SearchLevelCounter(17) = SearchLevelCounter(17) + 1
        Loop 'level 17
            SearchParse (Manipulations(2, SearchLevelCounter(16)))
            SearchLevelCounter(17) = 1
            SearchLevelCounter(16) = SearchLevelCounter(16) + 1
        Loop 'level 16
            SearchParse (Manipulations(2, SearchLevelCounter(15)))
            SearchLevelCounter(16) = 1
            SearchLevelCounter(15) = SearchLevelCounter(15) + 1
        Loop 'level 15
            SearchParse (Manipulations(2, SearchLevelCounter(14)))
            SearchLevelCounter(15) = 1
            SearchLevelCounter(14) = SearchLevelCounter(14) + 1
        Loop 'level 14
            SearchParse (Manipulations(2, SearchLevelCounter(13)))
            SearchLevelCounter(14) = 1
            SearchLevelCounter(13) = SearchLevelCounter(13) + 1
        Loop 'level 13
            SearchParse (Manipulations(2, SearchLevelCounter(12)))
            SearchLevelCounter(13) = 1
            SearchLevelCounter(12) = SearchLevelCounter(12) + 1
        Loop 'level 12
            SearchParse (Manipulations(2, SearchLevelCounter(11)))
            SearchLevelCounter(12) = 1
            SearchLevelCounter(11) = SearchLevelCounter(11) + 1
        Loop 'level 11
            SearchParse (Manipulations(2, SearchLevelCounter(10)))
            SearchLevelCounter(11) = 1
            SearchLevelCounter(10) = SearchLevelCounter(10) + 1
        Loop 'level 10
            SearchParse (Manipulations(2, SearchLevelCounter(9)))
            SearchLevelCounter(10) = 1
            SearchLevelCounter(9) = SearchLevelCounter(9) + 1
        Loop 'level 9
            SearchParse (Manipulations(2, SearchLevelCounter(8)))
            SearchLevelCounter(9) = 1
            SearchLevelCounter(8) = SearchLevelCounter(8) + 1
        Loop 'level 8
            SearchParse (Manipulations(2, SearchLevelCounter(7)))
            SearchLevelCounter(8) = 1
            SearchLevelCounter(7) = SearchLevelCounter(7) + 1
        Loop 'level 7
            SearchParse (Manipulations(2, SearchLevelCounter(6)))
            SearchLevelCounter(7) = 1
            SearchLevelCounter(6) = SearchLevelCounter(6) + 1
        Loop 'level 6
            SearchParse (Manipulations(2, SearchLevelCounter(5)))
            SearchLevelCounter(6) = 1
            SearchLevelCounter(5) = SearchLevelCounter(5) + 1
        Loop 'level 5
            SearchParse (Manipulations(2, SearchLevelCounter(4)))
            SearchLevelCounter(5) = 1
            SearchLevelCounter(4) = SearchLevelCounter(4) + 1
        Loop 'level 4
            SearchParse (Manipulations(2, SearchLevelCounter(3)))
            SearchLevelCounter(4) = 1
            SearchLevelCounter(3) = SearchLevelCounter(3) + 1
        Loop 'level 3
            SearchParse (Manipulations(2, SearchLevelCounter(2)))
            SearchLevelCounter(3) = 1
            SearchLevelCounter(2) = SearchLevelCounter(2) + 1
        Loop 'level 2
            SearchLevelCounter(2) = 1
            SearchLevelCounter(1) = SearchLevelCounter(1) + 1
    Loop 'level 1
    SearchLevelCounter(1) = 1
    SearchCurrentLevel = SearchCurrentLevel + 1
End If
End Sub

Private Sub SearchLevel22()
'SEARCH TWENTYSECOND LEVEL (22)
If SearchCurrentLevel = 22 And _
  SearchCurrentLevel <= Val(SearchLevelsTextBox.Text) Then
    Do While SearchingMode And SearchLevelCounter(1) <= SearchCounterMax
        ResetCurrentDeck
        SearchParse (Manipulations(1, SearchLevelCounter(1)))
        Do While SearchingMode And SearchLevelCounter(2) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(2)))
        Do While SearchingMode And SearchLevelCounter(3) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(3)))
        Do While SearchingMode And SearchLevelCounter(4) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(4)))
        Do While SearchingMode And SearchLevelCounter(5) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(5)))
        Do While SearchingMode And SearchLevelCounter(6) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(6)))
        Do While SearchingMode And SearchLevelCounter(7) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(7)))
        Do While SearchingMode And SearchLevelCounter(8) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(8)))
        Do While SearchingMode And SearchLevelCounter(9) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(9)))
        Do While SearchingMode And SearchLevelCounter(10) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(10)))
        Do While SearchingMode And SearchLevelCounter(11) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(11)))
        Do While SearchingMode And SearchLevelCounter(12) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(12)))
        Do While SearchingMode And SearchLevelCounter(13) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(13)))
        Do While SearchingMode And SearchLevelCounter(14) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(14)))
        Do While SearchingMode And SearchLevelCounter(15) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(15)))
        Do While SearchingMode And SearchLevelCounter(16) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(16)))
        Do While SearchingMode And SearchLevelCounter(17) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(17)))
        Do While SearchingMode And SearchLevelCounter(18) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(18)))
        Do While SearchingMode And SearchLevelCounter(19) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(19)))
        Do While SearchingMode And SearchLevelCounter(20) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(20)))
        Do While SearchingMode And SearchLevelCounter(21) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(21)))
        Do While SearchingMode And SearchLevelCounter(22) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(22)))
            SearchTransactionCheckMatch
            If MatchFound = 0 Then
                If TrapThreshold Then
                    SearchTransactionCheckPartialMatch
                    If PartialMatchFound = 1 Then
                        ShowProgress
                        If SuspendTrapFinal Then
                            SearchParse (Manipulations(2, SearchLevelCounter(22)))
                            SearchLevelCounter(22) = SearchLevelCounter(22) + 1
                            Exit Sub
                        Else
                            AppendTrapFile
                        End If
                    End If
                End If
                SearchParse (Manipulations(2, SearchLevelCounter(22)))
                SearchLevelCounter(22) = SearchLevelCounter(22) + 1
            Else
                MatchFoundLabel.Caption = "Match Found!!!"
                MatchFoundLabel.Visible = True
                If Timer < SearchStartTime Then
                    SearchStartTime = SearchStartTime - 86400
                    'error condition is test crosses midnight
                End If
                SearchElapsedTime = SearchElapsedTime + Timer - SearchStartTime
                ShowElapsedTime
                ProgressLabel.Caption = Empty
                Exit Sub
            End If
            If SearchProgressCounter Mod SpeedMod = 0 Then
                SearchProgressCounter = 0
                If SearchProgress.Value < SearchProgress.Max Then
                    SearchProgress.Value = SearchProgress.Value + 1
                Else
                    FirstPass = False
                    SearchProgress.Value = 0
                    If Timer < SearchStartTime Then
                        SearchStartTime = SearchStartTime - 86400
                        'error condition is test crosses midnight
                    End If
                    ManipulationsPerSecond = (SpeedMod * 100) / (Timer - SearchStartTime)
                    SpeedMod = Int(ManipulationsPerSecond / 8)
                    SearchElapsedTime = SearchElapsedTime + Timer - SearchStartTime
                    ShowElapsedTime
                    ShowProgress
                    ShowEstimatedTime
                    SearchStartTime = Timer
                End If
                DoEvents
                If SearchContinueReady = 1 Then
                    Exit Sub
                End If
            End If
            SearchProgressCounter = SearchProgressCounter + 1
        Loop 'level 22
            SearchParse (Manipulations(2, SearchLevelCounter(21)))
            SearchLevelCounter(22) = 1
            SearchLevelCounter(21) = SearchLevelCounter(21) + 1
        Loop 'level 21
            SearchParse (Manipulations(2, SearchLevelCounter(20)))
            SearchLevelCounter(21) = 1
            SearchLevelCounter(20) = SearchLevelCounter(20) + 1
        Loop 'level 20
            SearchParse (Manipulations(2, SearchLevelCounter(19)))
            SearchLevelCounter(20) = 1
            SearchLevelCounter(19) = SearchLevelCounter(19) + 1
        Loop 'level 19
            SearchParse (Manipulations(2, SearchLevelCounter(18)))
            SearchLevelCounter(19) = 1
            SearchLevelCounter(18) = SearchLevelCounter(18) + 1
        Loop 'level 18
            SearchParse (Manipulations(2, SearchLevelCounter(17)))
            SearchLevelCounter(18) = 1
            SearchLevelCounter(17) = SearchLevelCounter(17) + 1
        Loop 'level 17
            SearchParse (Manipulations(2, SearchLevelCounter(16)))
            SearchLevelCounter(17) = 1
            SearchLevelCounter(16) = SearchLevelCounter(16) + 1
        Loop 'level 16
            SearchParse (Manipulations(2, SearchLevelCounter(15)))
            SearchLevelCounter(16) = 1
            SearchLevelCounter(15) = SearchLevelCounter(15) + 1
        Loop 'level 15
            SearchParse (Manipulations(2, SearchLevelCounter(14)))
            SearchLevelCounter(15) = 1
            SearchLevelCounter(14) = SearchLevelCounter(14) + 1
        Loop 'level 14
            SearchParse (Manipulations(2, SearchLevelCounter(13)))
            SearchLevelCounter(14) = 1
            SearchLevelCounter(13) = SearchLevelCounter(13) + 1
        Loop 'level 13
            SearchParse (Manipulations(2, SearchLevelCounter(12)))
            SearchLevelCounter(13) = 1
            SearchLevelCounter(12) = SearchLevelCounter(12) + 1
        Loop 'level 12
            SearchParse (Manipulations(2, SearchLevelCounter(11)))
            SearchLevelCounter(12) = 1
            SearchLevelCounter(11) = SearchLevelCounter(11) + 1
        Loop 'level 11
            SearchParse (Manipulations(2, SearchLevelCounter(10)))
            SearchLevelCounter(11) = 1
            SearchLevelCounter(10) = SearchLevelCounter(10) + 1
        Loop 'level 10
            SearchParse (Manipulations(2, SearchLevelCounter(9)))
            SearchLevelCounter(10) = 1
            SearchLevelCounter(9) = SearchLevelCounter(9) + 1
        Loop 'level 9
            SearchParse (Manipulations(2, SearchLevelCounter(8)))
            SearchLevelCounter(9) = 1
            SearchLevelCounter(8) = SearchLevelCounter(8) + 1
        Loop 'level 8
            SearchParse (Manipulations(2, SearchLevelCounter(7)))
            SearchLevelCounter(8) = 1
            SearchLevelCounter(7) = SearchLevelCounter(7) + 1
        Loop 'level 7
            SearchParse (Manipulations(2, SearchLevelCounter(6)))
            SearchLevelCounter(7) = 1
            SearchLevelCounter(6) = SearchLevelCounter(6) + 1
        Loop 'level 6
            SearchParse (Manipulations(2, SearchLevelCounter(5)))
            SearchLevelCounter(6) = 1
            SearchLevelCounter(5) = SearchLevelCounter(5) + 1
        Loop 'level 5
            SearchParse (Manipulations(2, SearchLevelCounter(4)))
            SearchLevelCounter(5) = 1
            SearchLevelCounter(4) = SearchLevelCounter(4) + 1
        Loop 'level 4
            SearchParse (Manipulations(2, SearchLevelCounter(3)))
            SearchLevelCounter(4) = 1
            SearchLevelCounter(3) = SearchLevelCounter(3) + 1
        Loop 'level 3
            SearchParse (Manipulations(2, SearchLevelCounter(2)))
            SearchLevelCounter(3) = 1
            SearchLevelCounter(2) = SearchLevelCounter(2) + 1
        Loop 'level 2
            SearchLevelCounter(2) = 1
            SearchLevelCounter(1) = SearchLevelCounter(1) + 1
    Loop 'level 1
    SearchLevelCounter(1) = 1
    SearchCurrentLevel = SearchCurrentLevel + 1
End If
End Sub

Private Sub SearchLevel23()
'SEARCH TWENTYTHIRD LEVEL (23)
If SearchCurrentLevel = 23 And _
  SearchCurrentLevel <= Val(SearchLevelsTextBox.Text) Then
    Do While SearchingMode And SearchLevelCounter(1) <= SearchCounterMax
        ResetCurrentDeck
        SearchParse (Manipulations(1, SearchLevelCounter(1)))
        Do While SearchingMode And SearchLevelCounter(2) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(2)))
        Do While SearchingMode And SearchLevelCounter(3) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(3)))
        Do While SearchingMode And SearchLevelCounter(4) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(4)))
        Do While SearchingMode And SearchLevelCounter(5) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(5)))
        Do While SearchingMode And SearchLevelCounter(6) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(6)))
        Do While SearchingMode And SearchLevelCounter(7) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(7)))
        Do While SearchingMode And SearchLevelCounter(8) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(8)))
        Do While SearchingMode And SearchLevelCounter(9) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(9)))
        Do While SearchingMode And SearchLevelCounter(10) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(10)))
        Do While SearchingMode And SearchLevelCounter(11) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(11)))
        Do While SearchingMode And SearchLevelCounter(12) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(12)))
        Do While SearchingMode And SearchLevelCounter(13) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(13)))
        Do While SearchingMode And SearchLevelCounter(14) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(14)))
        Do While SearchingMode And SearchLevelCounter(15) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(15)))
        Do While SearchingMode And SearchLevelCounter(16) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(16)))
        Do While SearchingMode And SearchLevelCounter(17) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(17)))
        Do While SearchingMode And SearchLevelCounter(18) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(18)))
        Do While SearchingMode And SearchLevelCounter(19) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(19)))
        Do While SearchingMode And SearchLevelCounter(20) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(20)))
        Do While SearchingMode And SearchLevelCounter(21) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(21)))
        Do While SearchingMode And SearchLevelCounter(22) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(22)))
        Do While SearchingMode And SearchLevelCounter(23) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(23)))
            SearchTransactionCheckMatch
            If MatchFound = 0 Then
                If TrapThreshold Then
                    SearchTransactionCheckPartialMatch
                    If PartialMatchFound = 1 Then
                        ShowProgress
                        If SuspendTrapFinal Then
                            SearchParse (Manipulations(2, SearchLevelCounter(23)))
                            SearchLevelCounter(23) = SearchLevelCounter(23) + 1
                            Exit Sub
                        Else
                            AppendTrapFile
                        End If
                    End If
                End If
                SearchParse (Manipulations(2, SearchLevelCounter(23)))
                SearchLevelCounter(23) = SearchLevelCounter(23) + 1
            Else
                MatchFoundLabel.Caption = "Match Found!!!"
                MatchFoundLabel.Visible = True
                If Timer < SearchStartTime Then
                    SearchStartTime = SearchStartTime - 86400
                    'error condition is test crosses midnight
                End If
                SearchElapsedTime = SearchElapsedTime + Timer - SearchStartTime
                ShowElapsedTime
                ProgressLabel.Caption = Empty
                Exit Sub
            End If
            If SearchProgressCounter Mod SpeedMod = 0 Then
                SearchProgressCounter = 0
                If SearchProgress.Value < SearchProgress.Max Then
                    SearchProgress.Value = SearchProgress.Value + 1
                Else
                    FirstPass = False
                    SearchProgress.Value = 0
                    If Timer < SearchStartTime Then
                        SearchStartTime = SearchStartTime - 86400
                        'error condition is test crosses midnight
                    End If
                    ManipulationsPerSecond = (SpeedMod * 100) / (Timer - SearchStartTime)
                    SpeedMod = Int(ManipulationsPerSecond / 8)
                    SearchElapsedTime = SearchElapsedTime + Timer - SearchStartTime
                    ShowElapsedTime
                    ShowProgress
                    ShowEstimatedTime
                    SearchStartTime = Timer
                End If
                DoEvents
                If SearchContinueReady = 1 Then
                    Exit Sub
                End If
            End If
            SearchProgressCounter = SearchProgressCounter + 1
        Loop 'level 23
            SearchParse (Manipulations(2, SearchLevelCounter(22)))
            SearchLevelCounter(23) = 1
            SearchLevelCounter(22) = SearchLevelCounter(22) + 1
        Loop 'level 22
            SearchParse (Manipulations(2, SearchLevelCounter(21)))
            SearchLevelCounter(22) = 1
            SearchLevelCounter(21) = SearchLevelCounter(21) + 1
        Loop 'level 21
            SearchParse (Manipulations(2, SearchLevelCounter(20)))
            SearchLevelCounter(21) = 1
            SearchLevelCounter(20) = SearchLevelCounter(20) + 1
        Loop 'level 20
            SearchParse (Manipulations(2, SearchLevelCounter(19)))
            SearchLevelCounter(20) = 1
            SearchLevelCounter(19) = SearchLevelCounter(19) + 1
        Loop 'level 19
            SearchParse (Manipulations(2, SearchLevelCounter(18)))
            SearchLevelCounter(19) = 1
            SearchLevelCounter(18) = SearchLevelCounter(18) + 1
        Loop 'level 18
            SearchParse (Manipulations(2, SearchLevelCounter(17)))
            SearchLevelCounter(18) = 1
            SearchLevelCounter(17) = SearchLevelCounter(17) + 1
        Loop 'level 17
            SearchParse (Manipulations(2, SearchLevelCounter(16)))
            SearchLevelCounter(17) = 1
            SearchLevelCounter(16) = SearchLevelCounter(16) + 1
        Loop 'level 16
            SearchParse (Manipulations(2, SearchLevelCounter(15)))
            SearchLevelCounter(16) = 1
            SearchLevelCounter(15) = SearchLevelCounter(15) + 1
        Loop 'level 15
            SearchParse (Manipulations(2, SearchLevelCounter(14)))
            SearchLevelCounter(15) = 1
            SearchLevelCounter(14) = SearchLevelCounter(14) + 1
        Loop 'level 14
            SearchParse (Manipulations(2, SearchLevelCounter(13)))
            SearchLevelCounter(14) = 1
            SearchLevelCounter(13) = SearchLevelCounter(13) + 1
        Loop 'level 13
            SearchParse (Manipulations(2, SearchLevelCounter(12)))
            SearchLevelCounter(13) = 1
            SearchLevelCounter(12) = SearchLevelCounter(12) + 1
        Loop 'level 12
            SearchParse (Manipulations(2, SearchLevelCounter(11)))
            SearchLevelCounter(12) = 1
            SearchLevelCounter(11) = SearchLevelCounter(11) + 1
        Loop 'level 11
            SearchParse (Manipulations(2, SearchLevelCounter(10)))
            SearchLevelCounter(11) = 1
            SearchLevelCounter(10) = SearchLevelCounter(10) + 1
        Loop 'level 10
            SearchParse (Manipulations(2, SearchLevelCounter(9)))
            SearchLevelCounter(10) = 1
            SearchLevelCounter(9) = SearchLevelCounter(9) + 1
        Loop 'level 9
            SearchParse (Manipulations(2, SearchLevelCounter(8)))
            SearchLevelCounter(9) = 1
            SearchLevelCounter(8) = SearchLevelCounter(8) + 1
        Loop 'level 8
            SearchParse (Manipulations(2, SearchLevelCounter(7)))
            SearchLevelCounter(8) = 1
            SearchLevelCounter(7) = SearchLevelCounter(7) + 1
        Loop 'level 7
            SearchParse (Manipulations(2, SearchLevelCounter(6)))
            SearchLevelCounter(7) = 1
            SearchLevelCounter(6) = SearchLevelCounter(6) + 1
        Loop 'level 6
            SearchParse (Manipulations(2, SearchLevelCounter(5)))
            SearchLevelCounter(6) = 1
            SearchLevelCounter(5) = SearchLevelCounter(5) + 1
        Loop 'level 5
            SearchParse (Manipulations(2, SearchLevelCounter(4)))
            SearchLevelCounter(5) = 1
            SearchLevelCounter(4) = SearchLevelCounter(4) + 1
        Loop 'level 4
            SearchParse (Manipulations(2, SearchLevelCounter(3)))
            SearchLevelCounter(4) = 1
            SearchLevelCounter(3) = SearchLevelCounter(3) + 1
        Loop 'level 3
            SearchParse (Manipulations(2, SearchLevelCounter(2)))
            SearchLevelCounter(3) = 1
            SearchLevelCounter(2) = SearchLevelCounter(2) + 1
        Loop 'level 2
            SearchLevelCounter(2) = 1
            SearchLevelCounter(1) = SearchLevelCounter(1) + 1
    Loop 'level 1
    SearchLevelCounter(1) = 1
    SearchCurrentLevel = SearchCurrentLevel + 1
End If
End Sub

Private Sub SearchLevel24()
'SEARCH TWENTYFOURTH LEVEL (24)
If SearchCurrentLevel = 24 And _
  SearchCurrentLevel <= Val(SearchLevelsTextBox.Text) Then
    Do While SearchingMode And SearchLevelCounter(1) <= SearchCounterMax
        ResetCurrentDeck
        SearchParse (Manipulations(1, SearchLevelCounter(1)))
        Do While SearchingMode And SearchLevelCounter(2) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(2)))
        Do While SearchingMode And SearchLevelCounter(3) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(3)))
        Do While SearchingMode And SearchLevelCounter(4) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(4)))
        Do While SearchingMode And SearchLevelCounter(5) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(5)))
        Do While SearchingMode And SearchLevelCounter(6) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(6)))
        Do While SearchingMode And SearchLevelCounter(7) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(7)))
        Do While SearchingMode And SearchLevelCounter(8) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(8)))
        Do While SearchingMode And SearchLevelCounter(9) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(9)))
        Do While SearchingMode And SearchLevelCounter(10) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(10)))
        Do While SearchingMode And SearchLevelCounter(11) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(11)))
        Do While SearchingMode And SearchLevelCounter(12) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(12)))
        Do While SearchingMode And SearchLevelCounter(13) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(13)))
        Do While SearchingMode And SearchLevelCounter(14) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(14)))
        Do While SearchingMode And SearchLevelCounter(15) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(15)))
        Do While SearchingMode And SearchLevelCounter(16) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(16)))
        Do While SearchingMode And SearchLevelCounter(17) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(17)))
        Do While SearchingMode And SearchLevelCounter(18) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(18)))
        Do While SearchingMode And SearchLevelCounter(19) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(19)))
        Do While SearchingMode And SearchLevelCounter(20) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(20)))
        Do While SearchingMode And SearchLevelCounter(21) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(21)))
        Do While SearchingMode And SearchLevelCounter(22) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(22)))
        Do While SearchingMode And SearchLevelCounter(23) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(23)))
        Do While SearchingMode And SearchLevelCounter(24) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(24)))
            SearchTransactionCheckMatch
            If MatchFound = 0 Then
                If TrapThreshold Then
                    SearchTransactionCheckPartialMatch
                    If PartialMatchFound = 1 Then
                        ShowProgress
                        If SuspendTrapFinal Then
                            SearchParse (Manipulations(2, SearchLevelCounter(24)))
                            SearchLevelCounter(24) = SearchLevelCounter(24) + 1
                            Exit Sub
                        Else
                            AppendTrapFile
                        End If
                    End If
                End If
                SearchParse (Manipulations(2, SearchLevelCounter(24)))
                SearchLevelCounter(24) = SearchLevelCounter(24) + 1
            Else
                MatchFoundLabel.Caption = "Match Found!!!"
                MatchFoundLabel.Visible = True
                If Timer < SearchStartTime Then
                    SearchStartTime = SearchStartTime - 86400
                    'error condition is test crosses midnight
                End If
                SearchElapsedTime = SearchElapsedTime + Timer - SearchStartTime
                ShowElapsedTime
                ProgressLabel.Caption = Empty
                Exit Sub
            End If
            If SearchProgressCounter Mod SpeedMod = 0 Then
                SearchProgressCounter = 0
                If SearchProgress.Value < SearchProgress.Max Then
                    SearchProgress.Value = SearchProgress.Value + 1
                Else
                    FirstPass = False
                    SearchProgress.Value = 0
                    If Timer < SearchStartTime Then
                        SearchStartTime = SearchStartTime - 86400
                        'error condition is test crosses midnight
                    End If
                    ManipulationsPerSecond = (SpeedMod * 100) / (Timer - SearchStartTime)
                    SpeedMod = Int(ManipulationsPerSecond / 8)
                    SearchElapsedTime = SearchElapsedTime + Timer - SearchStartTime
                    ShowElapsedTime
                    ShowProgress
                    ShowEstimatedTime
                    SearchStartTime = Timer
                End If
                DoEvents
                If SearchContinueReady = 1 Then
                    Exit Sub
                End If
            End If
            SearchProgressCounter = SearchProgressCounter + 1
        Loop 'level 24
            SearchParse (Manipulations(2, SearchLevelCounter(23)))
            SearchLevelCounter(24) = 1
            SearchLevelCounter(23) = SearchLevelCounter(23) + 1
        Loop 'level 23
            SearchParse (Manipulations(2, SearchLevelCounter(22)))
            SearchLevelCounter(23) = 1
            SearchLevelCounter(22) = SearchLevelCounter(22) + 1
        Loop 'level 22
            SearchParse (Manipulations(2, SearchLevelCounter(21)))
            SearchLevelCounter(22) = 1
            SearchLevelCounter(21) = SearchLevelCounter(21) + 1
        Loop 'level 21
            SearchParse (Manipulations(2, SearchLevelCounter(20)))
            SearchLevelCounter(21) = 1
            SearchLevelCounter(20) = SearchLevelCounter(20) + 1
        Loop 'level 20
            SearchParse (Manipulations(2, SearchLevelCounter(19)))
            SearchLevelCounter(20) = 1
            SearchLevelCounter(19) = SearchLevelCounter(19) + 1
        Loop 'level 19
            SearchParse (Manipulations(2, SearchLevelCounter(18)))
            SearchLevelCounter(19) = 1
            SearchLevelCounter(18) = SearchLevelCounter(18) + 1
        Loop 'level 18
            SearchParse (Manipulations(2, SearchLevelCounter(17)))
            SearchLevelCounter(18) = 1
            SearchLevelCounter(17) = SearchLevelCounter(17) + 1
        Loop 'level 17
            SearchParse (Manipulations(2, SearchLevelCounter(16)))
            SearchLevelCounter(17) = 1
            SearchLevelCounter(16) = SearchLevelCounter(16) + 1
        Loop 'level 16
            SearchParse (Manipulations(2, SearchLevelCounter(15)))
            SearchLevelCounter(16) = 1
            SearchLevelCounter(15) = SearchLevelCounter(15) + 1
        Loop 'level 15
            SearchParse (Manipulations(2, SearchLevelCounter(14)))
            SearchLevelCounter(15) = 1
            SearchLevelCounter(14) = SearchLevelCounter(14) + 1
        Loop 'level 14
            SearchParse (Manipulations(2, SearchLevelCounter(13)))
            SearchLevelCounter(14) = 1
            SearchLevelCounter(13) = SearchLevelCounter(13) + 1
        Loop 'level 13
            SearchParse (Manipulations(2, SearchLevelCounter(12)))
            SearchLevelCounter(13) = 1
            SearchLevelCounter(12) = SearchLevelCounter(12) + 1
        Loop 'level 12
            SearchParse (Manipulations(2, SearchLevelCounter(11)))
            SearchLevelCounter(12) = 1
            SearchLevelCounter(11) = SearchLevelCounter(11) + 1
        Loop 'level 11
            SearchParse (Manipulations(2, SearchLevelCounter(10)))
            SearchLevelCounter(11) = 1
            SearchLevelCounter(10) = SearchLevelCounter(10) + 1
        Loop 'level 10
            SearchParse (Manipulations(2, SearchLevelCounter(9)))
            SearchLevelCounter(10) = 1
            SearchLevelCounter(9) = SearchLevelCounter(9) + 1
        Loop 'level 9
            SearchParse (Manipulations(2, SearchLevelCounter(8)))
            SearchLevelCounter(9) = 1
            SearchLevelCounter(8) = SearchLevelCounter(8) + 1
        Loop 'level 8
            SearchParse (Manipulations(2, SearchLevelCounter(7)))
            SearchLevelCounter(8) = 1
            SearchLevelCounter(7) = SearchLevelCounter(7) + 1
        Loop 'level 7
            SearchParse (Manipulations(2, SearchLevelCounter(6)))
            SearchLevelCounter(7) = 1
            SearchLevelCounter(6) = SearchLevelCounter(6) + 1
        Loop 'level 6
            SearchParse (Manipulations(2, SearchLevelCounter(5)))
            SearchLevelCounter(6) = 1
            SearchLevelCounter(5) = SearchLevelCounter(5) + 1
        Loop 'level 5
            SearchParse (Manipulations(2, SearchLevelCounter(4)))
            SearchLevelCounter(5) = 1
            SearchLevelCounter(4) = SearchLevelCounter(4) + 1
        Loop 'level 4
            SearchParse (Manipulations(2, SearchLevelCounter(3)))
            SearchLevelCounter(4) = 1
            SearchLevelCounter(3) = SearchLevelCounter(3) + 1
        Loop 'level 3
            SearchParse (Manipulations(2, SearchLevelCounter(2)))
            SearchLevelCounter(3) = 1
            SearchLevelCounter(2) = SearchLevelCounter(2) + 1
        Loop 'level 2
            SearchLevelCounter(2) = 1
            SearchLevelCounter(1) = SearchLevelCounter(1) + 1
    Loop 'level 1
    SearchLevelCounter(1) = 1
    SearchCurrentLevel = SearchCurrentLevel + 1
End If
End Sub

Private Sub SearchLevel25()
'SEARCH TWENTYFIFTH LEVEL (25)
If SearchCurrentLevel = 25 And _
  SearchCurrentLevel <= Val(SearchLevelsTextBox.Text) Then
    Do While SearchingMode And SearchLevelCounter(1) <= SearchCounterMax
        ResetCurrentDeck
        SearchParse (Manipulations(1, SearchLevelCounter(1)))
        Do While SearchingMode And SearchLevelCounter(2) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(2)))
        Do While SearchingMode And SearchLevelCounter(3) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(3)))
        Do While SearchingMode And SearchLevelCounter(4) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(4)))
        Do While SearchingMode And SearchLevelCounter(5) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(5)))
        Do While SearchingMode And SearchLevelCounter(6) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(6)))
        Do While SearchingMode And SearchLevelCounter(7) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(7)))
        Do While SearchingMode And SearchLevelCounter(8) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(8)))
        Do While SearchingMode And SearchLevelCounter(9) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(9)))
        Do While SearchingMode And SearchLevelCounter(10) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(10)))
        Do While SearchingMode And SearchLevelCounter(11) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(11)))
        Do While SearchingMode And SearchLevelCounter(12) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(12)))
        Do While SearchingMode And SearchLevelCounter(13) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(13)))
        Do While SearchingMode And SearchLevelCounter(14) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(14)))
        Do While SearchingMode And SearchLevelCounter(15) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(15)))
        Do While SearchingMode And SearchLevelCounter(16) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(16)))
        Do While SearchingMode And SearchLevelCounter(17) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(17)))
        Do While SearchingMode And SearchLevelCounter(18) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(18)))
        Do While SearchingMode And SearchLevelCounter(19) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(19)))
        Do While SearchingMode And SearchLevelCounter(20) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(20)))
        Do While SearchingMode And SearchLevelCounter(21) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(21)))
        Do While SearchingMode And SearchLevelCounter(22) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(22)))
        Do While SearchingMode And SearchLevelCounter(23) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(23)))
        Do While SearchingMode And SearchLevelCounter(24) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(24)))
        Do While SearchingMode And SearchLevelCounter(25) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(25)))
            SearchTransactionCheckMatch
            If MatchFound = 0 Then
                If TrapThreshold Then
                    SearchTransactionCheckPartialMatch
                    If PartialMatchFound = 1 Then
                        ShowProgress
                        If SuspendTrapFinal Then
                            SearchParse (Manipulations(2, SearchLevelCounter(25)))
                            SearchLevelCounter(25) = SearchLevelCounter(25) + 1
                            Exit Sub
                        Else
                            AppendTrapFile
                        End If
                    End If
                End If
                SearchParse (Manipulations(2, SearchLevelCounter(25)))
                SearchLevelCounter(25) = SearchLevelCounter(25) + 1
            Else
                MatchFoundLabel.Caption = "Match Found!!!"
                MatchFoundLabel.Visible = True
                If Timer < SearchStartTime Then
                    SearchStartTime = SearchStartTime - 86400
                    'error condition is test crosses midnight
                End If
                SearchElapsedTime = SearchElapsedTime + Timer - SearchStartTime
                ShowElapsedTime
                ProgressLabel.Caption = Empty
                Exit Sub
            End If
            If SearchProgressCounter Mod SpeedMod = 0 Then
                SearchProgressCounter = 0
                If SearchProgress.Value < SearchProgress.Max Then
                    SearchProgress.Value = SearchProgress.Value + 1
                Else
                    FirstPass = False
                    SearchProgress.Value = 0
                    If Timer < SearchStartTime Then
                        SearchStartTime = SearchStartTime - 86400
                        'error condition is test crosses midnight
                    End If
                    ManipulationsPerSecond = (SpeedMod * 100) / (Timer - SearchStartTime)
                    SpeedMod = Int(ManipulationsPerSecond / 8)
                    SearchElapsedTime = SearchElapsedTime + Timer - SearchStartTime
                    ShowElapsedTime
                    ShowProgress
                    ShowEstimatedTime
                    SearchStartTime = Timer
                End If
                DoEvents
                If SearchContinueReady = 1 Then
                    Exit Sub
                End If
            End If
            SearchProgressCounter = SearchProgressCounter + 1
        Loop 'level 25
            SearchParse (Manipulations(2, SearchLevelCounter(24)))
            SearchLevelCounter(25) = 1
            SearchLevelCounter(24) = SearchLevelCounter(24) + 1
        Loop 'level 24
            SearchParse (Manipulations(2, SearchLevelCounter(23)))
            SearchLevelCounter(24) = 1
            SearchLevelCounter(23) = SearchLevelCounter(23) + 1
        Loop 'level 23
            SearchParse (Manipulations(2, SearchLevelCounter(22)))
            SearchLevelCounter(23) = 1
            SearchLevelCounter(22) = SearchLevelCounter(22) + 1
        Loop 'level 22
            SearchParse (Manipulations(2, SearchLevelCounter(21)))
            SearchLevelCounter(22) = 1
            SearchLevelCounter(21) = SearchLevelCounter(21) + 1
        Loop 'level 21
            SearchParse (Manipulations(2, SearchLevelCounter(20)))
            SearchLevelCounter(21) = 1
            SearchLevelCounter(20) = SearchLevelCounter(20) + 1
        Loop 'level 20
            SearchParse (Manipulations(2, SearchLevelCounter(19)))
            SearchLevelCounter(20) = 1
            SearchLevelCounter(19) = SearchLevelCounter(19) + 1
        Loop 'level 19
            SearchParse (Manipulations(2, SearchLevelCounter(18)))
            SearchLevelCounter(19) = 1
            SearchLevelCounter(18) = SearchLevelCounter(18) + 1
        Loop 'level 18
            SearchParse (Manipulations(2, SearchLevelCounter(17)))
            SearchLevelCounter(18) = 1
            SearchLevelCounter(17) = SearchLevelCounter(17) + 1
        Loop 'level 17
            SearchParse (Manipulations(2, SearchLevelCounter(16)))
            SearchLevelCounter(17) = 1
            SearchLevelCounter(16) = SearchLevelCounter(16) + 1
        Loop 'level 16
            SearchParse (Manipulations(2, SearchLevelCounter(15)))
            SearchLevelCounter(16) = 1
            SearchLevelCounter(15) = SearchLevelCounter(15) + 1
        Loop 'level 15
            SearchParse (Manipulations(2, SearchLevelCounter(14)))
            SearchLevelCounter(15) = 1
            SearchLevelCounter(14) = SearchLevelCounter(14) + 1
        Loop 'level 14
            SearchParse (Manipulations(2, SearchLevelCounter(13)))
            SearchLevelCounter(14) = 1
            SearchLevelCounter(13) = SearchLevelCounter(13) + 1
        Loop 'level 13
            SearchParse (Manipulations(2, SearchLevelCounter(12)))
            SearchLevelCounter(13) = 1
            SearchLevelCounter(12) = SearchLevelCounter(12) + 1
        Loop 'level 12
            SearchParse (Manipulations(2, SearchLevelCounter(11)))
            SearchLevelCounter(12) = 1
            SearchLevelCounter(11) = SearchLevelCounter(11) + 1
        Loop 'level 11
            SearchParse (Manipulations(2, SearchLevelCounter(10)))
            SearchLevelCounter(11) = 1
            SearchLevelCounter(10) = SearchLevelCounter(10) + 1
        Loop 'level 10
            SearchParse (Manipulations(2, SearchLevelCounter(9)))
            SearchLevelCounter(10) = 1
            SearchLevelCounter(9) = SearchLevelCounter(9) + 1
        Loop 'level 9
            SearchParse (Manipulations(2, SearchLevelCounter(8)))
            SearchLevelCounter(9) = 1
            SearchLevelCounter(8) = SearchLevelCounter(8) + 1
        Loop 'level 8
            SearchParse (Manipulations(2, SearchLevelCounter(7)))
            SearchLevelCounter(8) = 1
            SearchLevelCounter(7) = SearchLevelCounter(7) + 1
        Loop 'level 7
            SearchParse (Manipulations(2, SearchLevelCounter(6)))
            SearchLevelCounter(7) = 1
            SearchLevelCounter(6) = SearchLevelCounter(6) + 1
        Loop 'level 6
            SearchParse (Manipulations(2, SearchLevelCounter(5)))
            SearchLevelCounter(6) = 1
            SearchLevelCounter(5) = SearchLevelCounter(5) + 1
        Loop 'level 5
            SearchParse (Manipulations(2, SearchLevelCounter(4)))
            SearchLevelCounter(5) = 1
            SearchLevelCounter(4) = SearchLevelCounter(4) + 1
        Loop 'level 4
            SearchParse (Manipulations(2, SearchLevelCounter(3)))
            SearchLevelCounter(4) = 1
            SearchLevelCounter(3) = SearchLevelCounter(3) + 1
        Loop 'level 3
            SearchParse (Manipulations(2, SearchLevelCounter(2)))
            SearchLevelCounter(3) = 1
            SearchLevelCounter(2) = SearchLevelCounter(2) + 1
        Loop 'level 2
            SearchLevelCounter(2) = 1
            SearchLevelCounter(1) = SearchLevelCounter(1) + 1
    Loop 'level 1
    SearchLevelCounter(1) = 1
    SearchCurrentLevel = SearchCurrentLevel + 1
End If
End Sub


Private Sub SearchLevel26()
'SEARCH TWENTYSIXTH LEVEL (26)
If SearchCurrentLevel = 26 And _
  SearchCurrentLevel <= Val(SearchLevelsTextBox.Text) Then
    Do While SearchingMode And SearchLevelCounter(1) <= SearchCounterMax
        ResetCurrentDeck
        SearchParse (Manipulations(1, SearchLevelCounter(1)))
        Do While SearchingMode And SearchLevelCounter(2) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(2)))
        Do While SearchingMode And SearchLevelCounter(3) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(3)))
        Do While SearchingMode And SearchLevelCounter(4) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(4)))
        Do While SearchingMode And SearchLevelCounter(5) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(5)))
        Do While SearchingMode And SearchLevelCounter(6) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(6)))
        Do While SearchingMode And SearchLevelCounter(7) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(7)))
        Do While SearchingMode And SearchLevelCounter(8) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(8)))
        Do While SearchingMode And SearchLevelCounter(9) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(9)))
        Do While SearchingMode And SearchLevelCounter(10) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(10)))
        Do While SearchingMode And SearchLevelCounter(11) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(11)))
        Do While SearchingMode And SearchLevelCounter(12) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(12)))
        Do While SearchingMode And SearchLevelCounter(13) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(13)))
        Do While SearchingMode And SearchLevelCounter(14) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(14)))
        Do While SearchingMode And SearchLevelCounter(15) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(15)))
        Do While SearchingMode And SearchLevelCounter(16) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(16)))
        Do While SearchingMode And SearchLevelCounter(17) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(17)))
        Do While SearchingMode And SearchLevelCounter(18) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(18)))
        Do While SearchingMode And SearchLevelCounter(19) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(19)))
        Do While SearchingMode And SearchLevelCounter(20) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(20)))
        Do While SearchingMode And SearchLevelCounter(21) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(21)))
        Do While SearchingMode And SearchLevelCounter(22) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(22)))
        Do While SearchingMode And SearchLevelCounter(23) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(23)))
        Do While SearchingMode And SearchLevelCounter(24) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(24)))
        Do While SearchingMode And SearchLevelCounter(25) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(25)))
        Do While SearchingMode And SearchLevelCounter(26) <= SearchCounterMax
            SearchParse (Manipulations(1, SearchLevelCounter(26)))
            SearchTransactionCheckMatch
            If MatchFound = 0 Then
                If TrapThreshold Then
                    SearchTransactionCheckPartialMatch
                    If PartialMatchFound = 1 Then
                        ShowProgress
                        If SuspendTrapFinal Then
                            SearchParse (Manipulations(2, SearchLevelCounter(26)))
                            SearchLevelCounter(26) = SearchLevelCounter(26) + 1
                            Exit Sub
                        Else
                            AppendTrapFile
                        End If
                    End If
                End If
                SearchParse (Manipulations(2, SearchLevelCounter(26)))
                SearchLevelCounter(26) = SearchLevelCounter(26) + 1
            Else
                MatchFoundLabel.Caption = "Match Found!!!"
                MatchFoundLabel.Visible = True
                If Timer < SearchStartTime Then
                    SearchStartTime = SearchStartTime - 86400
                    'error condition is test crosses midnight
                End If
                SearchElapsedTime = SearchElapsedTime + Timer - SearchStartTime
                ShowElapsedTime
                ProgressLabel.Caption = Empty
                Exit Sub
            End If
            If SearchProgressCounter Mod SpeedMod = 0 Then
                SearchProgressCounter = 0
                If SearchProgress.Value < SearchProgress.Max Then
                    SearchProgress.Value = SearchProgress.Value + 1
                Else
                    FirstPass = False
                    SearchProgress.Value = 0
                    If Timer < SearchStartTime Then
                        SearchStartTime = SearchStartTime - 86400
                        'error condition is test crosses midnight
                    End If
                    ManipulationsPerSecond = (SpeedMod * 100) / (Timer - SearchStartTime)
                    SpeedMod = Int(ManipulationsPerSecond / 8)
                    SearchElapsedTime = SearchElapsedTime + Timer - SearchStartTime
                    ShowElapsedTime
                    ShowProgress
                    ShowEstimatedTime
                    SearchStartTime = Timer
                End If
                DoEvents
                If SearchContinueReady = 1 Then
                    Exit Sub
                End If
            End If
            SearchProgressCounter = SearchProgressCounter + 1
        Loop 'level 26
            SearchParse (Manipulations(2, SearchLevelCounter(25)))
            SearchLevelCounter(26) = 1
            SearchLevelCounter(25) = SearchLevelCounter(25) + 1
        Loop 'level 25
            SearchParse (Manipulations(2, SearchLevelCounter(24)))
            SearchLevelCounter(25) = 1
            SearchLevelCounter(24) = SearchLevelCounter(24) + 1
        Loop 'level 24
            SearchParse (Manipulations(2, SearchLevelCounter(23)))
            SearchLevelCounter(24) = 1
            SearchLevelCounter(23) = SearchLevelCounter(23) + 1
        Loop 'level 23
            SearchParse (Manipulations(2, SearchLevelCounter(22)))
            SearchLevelCounter(23) = 1
            SearchLevelCounter(22) = SearchLevelCounter(22) + 1
        Loop 'level 22
            SearchParse (Manipulations(2, SearchLevelCounter(21)))
            SearchLevelCounter(22) = 1
            SearchLevelCounter(21) = SearchLevelCounter(21) + 1
        Loop 'level 21
            SearchParse (Manipulations(2, SearchLevelCounter(20)))
            SearchLevelCounter(21) = 1
            SearchLevelCounter(20) = SearchLevelCounter(20) + 1
        Loop 'level 20
            SearchParse (Manipulations(2, SearchLevelCounter(19)))
            SearchLevelCounter(20) = 1
            SearchLevelCounter(19) = SearchLevelCounter(19) + 1
        Loop 'level 19
            SearchParse (Manipulations(2, SearchLevelCounter(18)))
            SearchLevelCounter(19) = 1
            SearchLevelCounter(18) = SearchLevelCounter(18) + 1
        Loop 'level 18
            SearchParse (Manipulations(2, SearchLevelCounter(17)))
            SearchLevelCounter(18) = 1
            SearchLevelCounter(17) = SearchLevelCounter(17) + 1
        Loop 'level 17
            SearchParse (Manipulations(2, SearchLevelCounter(16)))
            SearchLevelCounter(17) = 1
            SearchLevelCounter(16) = SearchLevelCounter(16) + 1
        Loop 'level 16
            SearchParse (Manipulations(2, SearchLevelCounter(15)))
            SearchLevelCounter(16) = 1
            SearchLevelCounter(15) = SearchLevelCounter(15) + 1
        Loop 'level 15
            SearchParse (Manipulations(2, SearchLevelCounter(14)))
            SearchLevelCounter(15) = 1
            SearchLevelCounter(14) = SearchLevelCounter(14) + 1
        Loop 'level 14
            SearchParse (Manipulations(2, SearchLevelCounter(13)))
            SearchLevelCounter(14) = 1
            SearchLevelCounter(13) = SearchLevelCounter(13) + 1
        Loop 'level 13
            SearchParse (Manipulations(2, SearchLevelCounter(12)))
            SearchLevelCounter(13) = 1
            SearchLevelCounter(12) = SearchLevelCounter(12) + 1
        Loop 'level 12
            SearchParse (Manipulations(2, SearchLevelCounter(11)))
            SearchLevelCounter(12) = 1
            SearchLevelCounter(11) = SearchLevelCounter(11) + 1
        Loop 'level 11
            SearchParse (Manipulations(2, SearchLevelCounter(10)))
            SearchLevelCounter(11) = 1
            SearchLevelCounter(10) = SearchLevelCounter(10) + 1
        Loop 'level 10
            SearchParse (Manipulations(2, SearchLevelCounter(9)))
            SearchLevelCounter(10) = 1
            SearchLevelCounter(9) = SearchLevelCounter(9) + 1
        Loop 'level 9
            SearchParse (Manipulations(2, SearchLevelCounter(8)))
            SearchLevelCounter(9) = 1
            SearchLevelCounter(8) = SearchLevelCounter(8) + 1
        Loop 'level 8
            SearchParse (Manipulations(2, SearchLevelCounter(7)))
            SearchLevelCounter(8) = 1
            SearchLevelCounter(7) = SearchLevelCounter(7) + 1
        Loop 'level 7
            SearchParse (Manipulations(2, SearchLevelCounter(6)))
            SearchLevelCounter(7) = 1
            SearchLevelCounter(6) = SearchLevelCounter(6) + 1
        Loop 'level 6
            SearchParse (Manipulations(2, SearchLevelCounter(5)))
            SearchLevelCounter(6) = 1
            SearchLevelCounter(5) = SearchLevelCounter(5) + 1
        Loop 'level 5
            SearchParse (Manipulations(2, SearchLevelCounter(4)))
            SearchLevelCounter(5) = 1
            SearchLevelCounter(4) = SearchLevelCounter(4) + 1
        Loop 'level 4
            SearchParse (Manipulations(2, SearchLevelCounter(3)))
            SearchLevelCounter(4) = 1
            SearchLevelCounter(3) = SearchLevelCounter(3) + 1
        Loop 'level 3
            SearchParse (Manipulations(2, SearchLevelCounter(2)))
            SearchLevelCounter(3) = 1
            SearchLevelCounter(2) = SearchLevelCounter(2) + 1
        Loop 'level 2
            SearchLevelCounter(2) = 1
            SearchLevelCounter(1) = SearchLevelCounter(1) + 1
    Loop 'level 1
    SearchLevelCounter(1) = 1
    SearchCurrentLevel = SearchCurrentLevel + 1
End If
End Sub

Private Sub SearchParse(searchparameter)
stringpointer = InStr(searchparameter, "(")
If stringpointer = 0 Then
    myCall = searchparameter
Else
    myCall = Left(searchparameter, stringpointer - 1)
    commapointer = InStr(searchparameter, ",")
    If commapointer <> 0 Then
        searchparam1 = Mid(searchparameter, stringpointer + 1, _
            commapointer - stringpointer - 1)
        searchparam2 = Mid(searchparameter, commapointer + 2, _
            Len(searchparameter) - commapointer - 2)
        If IsNumeric(searchparam1) And _
            (Val(searchparam1) < 1 Or _
            Val(searchparam1) > 52) Then
                MsgBox ("First search parameter is out of range." & Chr(13) & _
                "It must be between 1 and 52.")
                SearchParseError = True
                Exit Sub
        End If
        If IsNumeric(searchparam2) And _
            (Val(searchparam2) < 1 Or _
            Val(searchparam2) > 52) Then
                MsgBox ("First search parameter is out of range." & Chr(13) & _
                "It must be between 1 and 52.")
                SearchParseError = True
                Exit Sub
        End If
    Else
        searchparam = Mid(searchparameter, stringpointer + 1, _
                        Len(searchparameter) - stringpointer - 1)
        If IsNumeric(searchparam) And _
            (Val(searchparam) < 1 Or _
            Val(searchparam) > 52) Then
                MsgBox ("Seaerch parameter is out of range." & Chr(13) & _
                "It must be between 1 and 52.")
                SearchParseError = True
                Exit Sub
        ElseIf Not IsNumeric(searchparam) Then
            SearchParseError = True
            'assume an error condition, and then the next loop sets the error
            'contition back to false if the text-based event parameter is valid
            'For i% = 1 To SessionNumParameters
            '    If SessionAllowableParameters(i%) = sessioneventparam Then
            '        SessionParseError = False
            '    End If
            'Next i%
            If SessionParseError = True Then
                MsgBox ("Event parameter is invalid.")
                Exit Sub
            End If
        End If
    End If
End If
Select Case myCall
    Case "CutDeckPrecise"
        Call CutDeckPrecise(searchparam1, "X")
    Case "ShiftTopBlock"
        Call ShiftTopBlock(searchparam1, searchparam2)
    Case "ShiftTopBlockInverse"
        Call ShiftTopBlockInverse(searchparam1, searchparam2)
    Case "MoveCard"
        Call MoveCard(searchparam1, searchparam2)
'    Case "CutSpecialRandom"
'        Call CutSpecialRandom(sessioneventparam)
'    Case "CutDeckRandom"
'        Call CutDeckRandom
'    Case "ForceCard"
'        Call ForceCard(sessioneventparam)
'    Case "ReturnCard"
'        Call ReturnCard(sessioneventparam)
'    Case "SelectCardsCutSelectNext1"
'        Call SelectCardsCutSelectNext1
'    Case "SelectCardsCutSelectNext2"
'        Call SelectCardsCutSelectNext2
'    Case "SelectCardsCutSelectNext3"
'        Call SelectCardsCutSelectNext3
'    Case "SelectCardsCutSelectFace1"
'        Call SelectCardsCutSelectFace1
'    Case "SelectCardsCutSelectFace2"
'        Call SelectCardsCutSelectFace2
'    Case "SelectCardsCutSelectFace3"
'        Call SelectCardsCutSelectFace3
'    Case "SelectCardsCutSelectNextRepeat"
'        Call SelectCardsCutSelectNextRepeat
'    Case "SelectCardsCutSelectNextRepeat2"
'        Call SelectCardsCutSelectNextRepeat2
'    Case "FreeChoiceSpreadSelect"
'        Call FreeChoiceSpreadSelect(sessioneventparam)
    Case "InFaro"
        Call InFaro
    Case "InFaroSpecialTop"
        Call InFaroSpecialTop(searchparam1, searchparam2)
    Case "InFaroSpecialBottom"
        Call InFaroSpecialBottom(searchparam1, searchparam2)
    Case "OutFaro"
        Call OutFaro
    Case "OutFaroSpecialTop"
        Call OutFaroSpecialTop(searchparam1, searchparam2)
    Case "OutFaroSpecialBottom"
        Call OutFaroSpecialBottom(searchparam1, searchparam2)
'    Case "OHShuffle"
'        Call OHShuffle
'    Case "OHShuffleTop"
'        Call OHShuffleTop(sessioneventparam)
'    Case "OHShuffleBottom"
'        Call OHShuffleBottom(sessioneventparam)
'    Case "PokerDeal"
'        Call PokerDeal(sessioneventparam)
'    Case "AssemblePokerDeal"
'        Call AssemblePokerDeal(sessioneventparam)
'    Case "SetStack"
'        Call SetStack(sessioneventparam)
'    Case "ResetCurrentDeck"
'        Call ResetCurrentDeck
    Case "InverseInFaro"
        Call InverseInFaro
    Case "InverseInFaroSpecialTop"
        Call InverseInFaroSpecialTop(searchparam1, searchparam2)
    Case "InverseInFaroSpecialBottom"
        Call InverseInFaroSpecialBottom(searchparam1, searchparam2)
    Case "InverseOutFaro"
        Call InverseOutFaro
    Case "InverseOutFaroSpecialTop"
        Call InverseOutFaroSpecialTop(searchparam1, searchparam2)
    Case "InverseOutFaroSpecialBottom"
        Call InverseOutFaroSpecialBottom(searchparam1, searchparam2)
'    Case "RiffleShuffle"
'        Call RiffleShuffle
'    Case "RiffleShuffleTop"
'        Call RiffleShuffleTop(sessioneventparam)
'    Case "RiffleShuffleBottom"
'        Call RiffleShuffleBottom(sessioneventparam)
    Case "RunSingleCards"
        Call RunSingleCards(searchparam)
    Case "RunSingleCardsInverse"
        Call RunSingleCardsInverse(searchparam)
    Case Else
        MsgBox ("frmSearch: This is an unknown Search entry." & _
        Chr(13) & myCall & "error")
        SearchParseError = True
End Select

End Sub


Public Sub ShowProgress()
ProgressLabel.Caption = SearchCurrentLevel & " ["
For i% = 1 To SearchCurrentLevel
    If SearchLevelCounter(i%) <= SearchCounterMax Then
        ProgressLabel.Caption = ProgressLabel.Caption & SearchLevelCounter(i%) & ","
    Else
        ProgressLabel.Caption = ProgressLabel.Caption & SearchCounterMax & ","
    End If
Next i%
ProgressLabel.Caption = Left(ProgressLabel.Caption, Len(ProgressLabel.Caption) - 1)
ProgressLabel.Caption = ProgressLabel.Caption & "]"
MatchListBox.Clear
If SearchingMode Then
    For i% = 0 To SearchCurrentLevel - 1
        If SearchLevelCounter(i% + 1) <= SearchCounterMax Then
            MatchListBox.List(i%) = Manipulations(1, SearchLevelCounter(i% + 1))
        Else
            MatchListBox.List(i%) = Manipulations(1, SearchCounterMax)
        End If
    Next i%
    If PartialMatchFound = 1 Then
        MatchFoundLabel.Caption = "Threshold: " & PartialMatchCounter
    Else
        MatchFoundLabel.Caption = "-- recent test --"
    End If
    MatchFoundLabel.Visible = True
End If
End Sub

Public Sub ShowElapsedTime()
Dim sMillenia As Double
Dim sYears As Double
Dim sDays As Double
Dim sHours As Double
Dim sMinutes As Double
Dim sSeconds As Double
Dim sFront As Double
Dim sElapsed As Double
sElapsed = SearchElapsedTime
sSeconds = (((sElapsed / 60) - Int(sElapsed / 60)) * 60)
sFront = (sElapsed / 60)
'this is now in minutes
sMinutes = (((sFront / 60) - Int(sFront / 60)) * 60)
sFront = (sFront / 60)
'this is now in hours
sHours = (((sFront / 24) - Int(sFront / 24)) * 24)
sFront = (sFront / 24)
'this is now in days
sDays = (((sFront / 365) - Int(sFront / 365)) * 365)
sFront = (sFront / 365)
'this is now in years
If sFront > 10000 Then
    sYears = (((sFront / 1000) - Int(sFront / 1000)) * 1000)
    sFront = (sFront / 1000)
    'this is now in millenia
    sMillenia = sFront
Else
    sYears = sFront
    sMillenia = Empty
End If
'this next section displays meaningful (not excessive detail) results
TimerResult.Caption = ""
If Int(sMillenia) > 0 Then
    TimerResult.Caption = Format$(sMillenia, "Scientific") & " Millenia "
    Exit Sub
End If
If Int(sYears) = 1 Then
    TimerResult.Caption = Int(sYears) & " year "
ElseIf Int(sYears) > 1 Then
    TimerResult.Caption = Int(sYears) & " years "
End If
If Int(sDays) = 1 Then
    TimerResult.Caption = TimerResult.Caption & Int(sDays) & " day "
ElseIf Int(sDays) > 1 Then
    TimerResult.Caption = TimerResult.Caption & Int(sDays) & " days "
End If
If Int(sHours) = 1 Then
    TimerResult.Caption = TimerResult.Caption & Int(sHours) & " hr "
ElseIf Int(sHours) > 1 Then
    TimerResult.Caption = TimerResult.Caption & Int(sHours) & " hrs "
End If
If Int(sMinutes) = 1 Then
    TimerResult.Caption = TimerResult.Caption & Int(sMinutes) & " min "
ElseIf Int(sMinutes) > 1 Then
    TimerResult.Caption = TimerResult.Caption & Int(sMinutes) & " mins "
End If
If Int(sSeconds) = 0 Then
    TimerResult.Caption = TimerResult.Caption & Int(sSeconds) & " secs"
ElseIf Int(sSeconds) = 1 Then
    TimerResult.Caption = TimerResult.Caption & Int(sSeconds) & " sec"
ElseIf Int(sSeconds) > 1 Then
    TimerResult.Caption = TimerResult.Caption & Int(sSeconds) & " secs"
ElseIf Int(sSeconds) < 1 Then
    TimerResult.Caption = TimerResult.Caption & Int(sSeconds) & " secs"
End If
If FirstPass Then
    TimerResult.Caption = "Calibrating to computer speed on first pass..."
End If
End Sub

Public Sub ShowEstimatedTime()
ManipulationsLabel.Caption = SearchCounterMax
Dim sMillenia As Double
Dim sYears As Double
Dim sDays As Double
Dim sHours As Double
Dim sMinutes As Double
Dim sSeconds As Double
Dim sFront As Double
SearchTotalManipulations = 0
For i% = 1 To Val(SearchLevelsTextBox.Text)
    SearchTotalManipulations = SearchTotalManipulations + _
        SearchCounterMax ^ i%
Next i%
SearchTotalPossibleTime = SearchTotalManipulations / ManipulationsPerSecond
sSeconds = (((SearchTotalPossibleTime / 60) - Int(SearchTotalPossibleTime / 60)) * 60)
sFront = (SearchTotalPossibleTime / 60)
'this is now in minutes
sMinutes = (((sFront / 60) - Int(sFront / 60)) * 60)
sFront = (sFront / 60)
'this is now in hours
sHours = (((sFront / 24) - Int(sFront / 24)) * 24)
sFront = (sFront / 24)
'this is now in days
sDays = (((sFront / 365) - Int(sFront / 365)) * 365)
sFront = (sFront / 365)
'this is now in years
If sFront > 10000 Then
    sYears = (((sFront / 1000) - Int(sFront / 1000)) * 1000)
    sFront = (sFront / 1000)
    'this is now in millenia
    sMillenia = sFront
Else
    sYears = sFront
    sMillenia = Empty
End If
'this next section displays meaningful (not excessive detail) results
SearchTimeLabel.Caption = ""
If Int(sMillenia) > 0 Then
    SearchTimeLabel.Caption = Format$(sMillenia, "Scientific") & " Millenia "
    Exit Sub
End If
If Int(sYears) = 1 Then
    SearchTimeLabel.Caption = Int(sYears) & " year "
    Exit Sub
ElseIf Int(sYears) > 1 Then
    SearchTimeLabel.Caption = Int(sYears) & " years "
    Exit Sub
End If
If Int(sDays) > 10 Then
    SearchTimeLabel.Caption = Int(sDays) & " days "
    Exit Sub
End If
If Int(sDays) = 1 Then
    SearchTimeLabel.Caption = Int(sDays) & " day "
ElseIf Int(sDays) > 1 Then
    SearchTimeLabel.Caption = Int(sDays) & " days "
End If
If Int(sHours) = 1 Then
    SearchTimeLabel.Caption = SearchTimeLabel.Caption & Int(sHours) & " hour "
ElseIf Int(sHours) > 1 Then
    SearchTimeLabel.Caption = SearchTimeLabel.Caption & Int(sHours) & " hours "
End If
If Int(sMinutes) = 1 Then
    SearchTimeLabel.Caption = SearchTimeLabel.Caption & Int(sMinutes) & " min "
ElseIf Int(sMinutes) > 1 Then
    SearchTimeLabel.Caption = SearchTimeLabel.Caption & Int(sMinutes) & " mins "
End If
If Int(sSeconds) = 1 Then
    SearchTimeLabel.Caption = SearchTimeLabel.Caption & Int(sSeconds) & " sec "
ElseIf Int(sSeconds) > 1 Then
    SearchTimeLabel.Caption = SearchTimeLabel.Caption & Int(sSeconds) & " secs "
    Exit Sub
End If
If sSeconds < 1 Then
    SearchTimeLabel.Caption = " Less then 1 second"
    Exit Sub
End If
End Sub



Private Sub SearchUncheckAllButton_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
SearchSpecialCancel = False
SearchDisableContinueCheck
If SearchContinueReady = 1 Then
    Exit Sub
End If
End Sub

Private Sub SearchUncheckAllButton_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
SearchCutDeckPrecise.Value = 0
SearchMoveCard.Value = 0
SearchRunSingleCards.Value = 0
SearchRunSingleCardsInv.Value = 0
SearchShiftTopBlock.Value = 0
SearchShiftTopBlockInv.Value = 0
SearchOutFaro.Value = 0
SearchOutFaroInverse.Value = 0
SearchOutFaroSpecialTop.Value = 0
SearchOutFaroSpecialTopInv.Value = 0
SearchOutFaroSpecialBottom.Value = 0
SearchOutFaroSpecialBottomInv.Value = 0
SearchInFaro.Value = 0
SearchInFaroInverse.Value = 0
SearchInFaroSpecialTop.Value = 0
SearchInFaroSpecialTopInv.Value = 0
SearchInFaroSpecialBottom.Value = 0
SearchInFaroSpecialBottomInv.Value = 0
SearchCutDeckPreciseAll.Value = 0
SearchMoveCardAll.Value = 0
SearchRunSingleCardsAll.Value = 0
SearchRunSingleCardsInvAll.Value = 0
SearchShiftTopBlockAll.Value = 0
SearchShiftTopBlockInvAll.Value = 0
SearchOutFaroAll.Value = 0
SearchOutFaroInverseAll.Value = 0
SearchOutFaroSpecialTopAll.Value = 0
SearchOutFaroSpecialTopInvAll.Value = 0
SearchOutFaroSpecialBottomAll.Value = 0
SearchOutFaroSpecialBottomInvAll.Value = 0
SearchInFaroAll.Value = 0
SearchInFaroInverseAll.Value = 0
SearchInFaroSpecialTopAll.Value = 0
SearchInFaroSpecialTopInvAll.Value = 0
SearchInFaroSpecialBottomAll.Value = 0
SearchInFaroSpecialBottomInvAll.Value = 0
'Debug.Print SearchCounterMax
End Sub

Private Sub StartDeckButton_Click()
SearchDisableContinueCheck
If SearchContinueReady = 1 Then
    Exit Sub
End If
    Dim sFile As String
    With SearchCommonDialog
        .DialogTitle = "Open Deck"
        .CancelError = False
        'ToDo: set the flags and attributes of the common dialog control
        .Filter = "StackView Deck Files (*.svf)|*.svf"
        .ShowOpen
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
        sFile = .FileName
        tFile = .FileTitle
    End With
    If Not ((Right(tFile, 3) = "svf") Or (Right(tFile, 3) = "SVF")) Then
        MsgBox ("Invalid file type.  Deck files have a .svf extension.")
        Exit Sub
    End If
    On Error GoTo DeckOpenError
    Dim fso As New FileSystemObject, deckfile As File, ts As TextStream
    Set deckfile = fso.GetFile(sFile)
    Set ts = deckfile.OpenAsTextStream(ForReading)
    For i% = 1 To 52
        StartDeckInitial(1, i%) = Val(ts.ReadLine)
        StartDeckInitial(2, i%) = ts.ReadLine
    Next i%
    ts.Close
    StartDeckLabel.Caption = tFile
    StartDeckLabel.Visible = True
    StartDeckName = tFile
    SearchStartDeckSet = 1
    Exit Sub
DeckOpenError:
MsgBox ("Error opening Deck file.")
End Sub

Private Sub TargetDeckButton_Click()
SearchDisableContinueCheck
If SearchContinueReady = 1 Then
    Exit Sub
End If
    Dim sFile As String
    With SearchCommonDialog
        .DialogTitle = "Open Deck"
        .CancelError = False
        'ToDo: set the flags and attributes of the common dialog control
        .Filter = "StackView Deck Files (*.svf)|*.svf"
        .ShowOpen
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
        sFile = .FileName
        tFile = .FileTitle
    End With
    If Not ((Right(tFile, 3) = "svf") Or (Right(tFile, 3) = "SVF")) Then
        MsgBox ("Invalid file type.  Deck files have a .svf extension.")
        Exit Sub
    End If
    On Error GoTo DeckOpenError
    Dim fso As New FileSystemObject, deckfile As File, ts As TextStream
    Set deckfile = fso.GetFile(sFile)
    Set ts = deckfile.OpenAsTextStream(ForReading)
    For i% = 1 To 52
        TargetDeck(1, i%) = Val(ts.ReadLine)
        TargetDeck(2, i%) = ts.ReadLine
    Next i%
    ts.Close
    TargetDeckLabel.Caption = tFile
    TargetDeckLabel.Visible = True
    TargetDeckName = tFile
    SearchTargetDeckSet = 1
    Exit Sub
DeckOpenError:
MsgBox ("Error opening Deck file.")
End Sub

Private Sub CutDeckPrecise(cutdepthparameter, pReverse)
    pReverse = "X"
    CutDepth = Val(cutdepthparameter)
    If CutDepth = 0 Then
        Exit Sub
    End If
    i = 1
    For j% = CutDepth + 1 To 52
        For z% = 1 To 2
            ChangedDeck(z%, i) = StartDeck(z%, j%)
        Next z%
        i = i + 1
    Next j%
    For k% = 1 To CutDepth
        For z% = 1 To 2
            ChangedDeck(z%, i) = StartDeck(z%, k%)
        Next z%
        i = i + 1
    Next k%
    For m% = 1 To 52
        For z% = 1 To 2
            StartDeck(z%, m%) = ChangedDeck(z%, m%)
        Next z%
    Next m%
End Sub

Private Sub ResetCurrentDeck()
For m% = 1 To 52
    For p% = 1 To 2
        StartDeck(p%, m%) = StartDeckInitial(p%, m%)
    Next p%
Next m%
'For k% = 1 To 52
'    For m% = 1 To 52
'        If StartDeck(1, m%) = k% Then
'            For p% = 1 To 2
'                ChangedDeck(p%, k%) = StartDeck(p%, m%)
'            Next p%
'        End If
'    Next m%
'Next k%
'For m% = 1 To 52
'    For z% = 1 To 2
'        StartDeck(z%, m%) = ChangedDeck(z%, m%)
'    Next z%
'Next m%
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
            For z% = 1 To 2
                ChangedDeck(z%, 2 * i% - 1) = StartDeck(z%, i%)
                ChangedDeck(z%, 2 * i%) = StartDeck(z%, i% + ProtectedBlock)
            Next z%
        Next i%
        MeshedBlock = 2 * ProtectedBlock
        For k% = 1 To 52 - MeshedBlock
            For z% = 1 To 2
                ChangedDeck(z%, k% + MeshedBlock) = StartDeck(z%, k% + MeshedBlock)
            Next z%
        Next k%
        For m% = 1 To 52
            For z% = 1 To 2
                StartDeck(z%, m%) = ChangedDeck(z%, m%)
            Next z%
        Next m%
    Else
        InteriorPosition = InteriorCard + ProtectedBlock
        If ProtectedBlock + InteriorPosition <= 52 Then
            For i% = 1 To InteriorPosition - ProtectedBlock - 1
            'added the "- 1" from InFaro solution
                For z% = 1 To 2
                    ChangedDeck(z%, i%) = StartDeck(z%, ProtectedBlock + i%)
                Next z%
            Next i%
            MeshedBlockPosition = InteriorPosition - ProtectedBlock - 1
            'added the "- 1" from InFaro solution
            For i% = 1 To ProtectedBlock
                For z% = 1 To 2
                    ChangedDeck(z%, MeshedBlockPosition + (2 * i% - 1)) = _
                        StartDeck(z%, i%)
                    ChangedDeck(z%, MeshedBlockPosition + (2 * i%)) = _
                        StartDeck(z%, i% + InteriorPosition - 1)
                        'added the "- 1" from InFaro solution
                Next z%
            Next i%
            For i% = ProtectedBlock * 2 + InteriorPosition - ProtectedBlock To 52
            'took out a "+ 1" right before the "To" from InFaro solution
                For z% = 1 To 2
                    ChangedDeck(z%, i%) = StartDeck(z%, i%)
                Next z%
            Next i%
            For m% = 1 To 52
                For z% = 1 To 2
                    StartDeck(z%, m%) = ChangedDeck(z%, m%)
                Next z%
            Next m%
        Else
            For i% = 1 To InteriorPosition - ProtectedBlock - 1
            'added the "- 1" from InFaro solution
                For z% = 1 To 2
                    ChangedDeck(z%, i%) = StartDeck(z%, ProtectedBlock + i%)
                Next z%
            Next i%
            MeshedBlockPosition = InteriorPosition - ProtectedBlock - 1
            'added the "- 1" from InFaro solution
            MeshedBlock = 2 * (52 - InteriorPosition + 1)
            'added the "+ 1" from InFaro solution
            For i% = 1 To 52 - InteriorPosition + 1
            'added the "+ 1" from InFaro solution
                For z% = 1 To 2
                    ChangedDeck(z%, MeshedBlockPosition + (2 * i% - 1)) = _
                        StartDeck(z%, i%)
                    ChangedDeck(z%, MeshedBlockPosition + (2 * i%)) = _
                        StartDeck(z%, i% + InteriorPosition - 1)
                        'added the "- 1" from InFaro solution
                Next z%
            Next i%
            For i% = 1 To ProtectedBlock - (52 - InteriorPosition + 1)
            'added the "+ 1" from InFaro solution
                For z% = 1 To 2
                    ChangedDeck(z%, MeshedBlockPosition + MeshedBlock + i%) = _
                        StartDeck(z%, 52 - InteriorPosition + 1 + i%)
                        'added the "+ 1" from InFaro solution
                Next z%
            Next i%
            For m% = 1 To 52
                For z% = 1 To 2
                    StartDeck(z%, m%) = ChangedDeck(z%, m%)
                Next z%
            Next m%
        End If
    End If
ElseIf ProtectedBlock >= 26 Then
    If InteriorCard = Empty Then
        For i% = 1 To 52 - ProtectedBlock
            For z% = 1 To 2
                ChangedDeck(z%, 2 * i% - 1) = StartDeck(z%, i%)
                ChangedDeck(z%, 2 * i%) = StartDeck(z%, i% + ProtectedBlock)
            Next z%
        Next i%
        MeshedBlock = 2 * (52 - ProtectedBlock)
        For k% = 1 To 52 - MeshedBlock
            For z% = 1 To 2
                ChangedDeck(z%, k% + MeshedBlock) = StartDeck(z%, k% + 52 - ProtectedBlock)
            Next z%
        Next k%
        For m% = 1 To 52
            For z% = 1 To 2
                StartDeck(z%, m%) = ChangedDeck(z%, m%)
            Next z%
        Next m%
    Else
        InteriorPosition = InteriorCard + ProtectedBlock
        If ProtectedBlock + InteriorPosition <= 52 Then
            For i% = 1 To InteriorPosition - ProtectedBlock - 1
            'added the "- 1" from InFaro solution
                For z% = 1 To 2
                    ChangedDeck(z%, i%) = StartDeck(z%, ProtectedBlock + i%)
                Next z%
            Next i%
            MeshedBlockPosition = InteriorPosition - ProtectedBlock - 1
            'added the "- 1" from InFaro solution
            For i% = 1 To ProtectedBlock
                For z% = 1 To 2
                    ChangedDeck(z%, MeshedBlockPosition + (2 * i% - 1)) = _
                        StartDeck(z%, i%)
                    ChangedDeck(z%, MeshedBlockPosition + (2 * i%)) = _
                        StartDeck(z%, i% + InteriorPosition - 1)
                        'added the "- 1" from InFaro solution
                Next z%
            Next i%
            For i% = ProtectedBlock * 2 + InteriorPosition - ProtectedBlock To 52
            'took out a "+ 1" right before the "To" from InFaro solution
                For z% = 1 To 2
                    ChangedDeck(z%, i%) = StartDeck(z%, i%)
                Next z%
            Next i%
            For m% = 1 To 52
                For z% = 1 To 2
                    StartDeck(z%, m%) = ChangedDeck(z%, m%)
                Next z%
            Next m%
        Else
            For i% = 1 To InteriorPosition - ProtectedBlock - 1
            'added the "- 1" from InFaro solution
                For z% = 1 To 2
                    ChangedDeck(z%, i%) = StartDeck(z%, ProtectedBlock + i%)
                Next z%
            Next i%
            MeshedBlockPosition = InteriorPosition - ProtectedBlock - 1
            'added the "- 1" from InFaro solution
            MeshedBlock = 2 * (52 - InteriorPosition + 1)
            'added the "+ 1" from InFaro solution
            For i% = 1 To 52 - InteriorPosition + 1
            'added the "+ 1" from InFaro solution
                For z% = 1 To 2
                    ChangedDeck(z%, MeshedBlockPosition + (2 * i% - 1)) = _
                        StartDeck(z%, i%)
                    ChangedDeck(z%, MeshedBlockPosition + (2 * i%)) = _
                        StartDeck(z%, i% + InteriorPosition - 1)
                        'added the "- 1" from InFaro solution
                Next z%
            Next i%
            For i% = 1 To ProtectedBlock - (52 - InteriorPosition + 1)
            'added the "+ 1" from InFaro solution
                For z% = 1 To 2
                    ChangedDeck(z%, MeshedBlockPosition + MeshedBlock + i%) = _
                        StartDeck(z%, 52 - InteriorPosition + 1 + i%)
                        'added the "+ 1" from InFaro solution
                Next z%
            Next i%
            For m% = 1 To 52
                For z% = 1 To 2
                    StartDeck(z%, m%) = ChangedDeck(z%, m%)
                Next z%
            Next m%
        End If
    End If
End If
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
        For i% = 1 To 52 - MeshedBlock
            For z% = 1 To 2
                ChangedDeck(z%, i% + MeshedBlock) = StartDeck(z%, i% + MeshedBlock)
            Next z%
        Next i%
        For k% = 1 To ProtectedBlock
            For z% = 1 To 2
                ChangedDeck(z%, k%) = StartDeck(z%, 2 * k% - 1)
                ChangedDeck(z%, k% + ProtectedBlock) = StartDeck(z%, 2 * k%)
            Next z%
        Next k%
        For m% = 1 To 52
            For z% = 1 To 2
                StartDeck(z%, m%) = ChangedDeck(z%, m%)
            Next z%
        Next m%
    Else
        'complete
        InteriorPosition = InteriorCard + ProtectedBlock
        If ProtectedBlock + InteriorPosition <= 52 Then
            For j% = 1 To ProtectedBlock
                For z% = 1 To 2
                    ChangedDeck(z%, j%) = StartDeck(z%, InteriorCard + 2 * (j% - 1))
                Next z%
            Next j%
            For k% = 1 To InteriorCard - 1
                For z% = 1 To 2
                    ChangedDeck(z%, ProtectedBlock + k%) = StartDeck(z%, k%)
                Next z%
            Next k%
            For b% = 1 To ProtectedBlock
                For z% = 1 To 2
                    ChangedDeck(z%, InteriorPosition + b% - 1) = StartDeck(z%, InteriorCard + (2 * b%) - 1)
                Next z%
            Next b%
            For p% = InteriorPosition + ProtectedBlock To 52
                For z% = 1 To 2
                    ChangedDeck(z%, p%) = StartDeck(z%, p%)
                Next z%
            Next p%
            For m% = 1 To 52
                For z% = 1 To 2
                    StartDeck(z%, m%) = ChangedDeck(z%, m%)
                Next z%
            Next m%
        Else
            'complete
            For j% = 1 To 52 - InteriorPosition + 1
                For z% = 1 To 2
                    ChangedDeck(z%, j%) = StartDeck(z%, InteriorCard + 2 * (j% - 1))
                Next z%
            Next j%
            For k% = 1 To ProtectedBlock - (52 - InteriorPosition + 1)
                For z% = 1 To 2
                    ChangedDeck(z%, 52 - InteriorPosition + 1 + k%) = _
                        StartDeck(z%, 2 * 52 - 2 * ProtectedBlock - InteriorCard + 1 + k%)
                Next z%
            Next k%
            For b% = 1 To InteriorCard - 1
                For z% = 1 To 2
                    ChangedDeck(z%, ProtectedBlock + b%) = _
                        StartDeck(z%, b%)
                Next z%
            Next b%
            For p% = 1 To 52 - InteriorPosition + 1
                For z% = 1 To 2
                    ChangedDeck(z%, InteriorPosition - 1 + p%) = _
                        StartDeck(z%, InteriorCard + 2 * p% - 1)
                Next z%
            Next p%
            For m% = 1 To 52
                For z% = 1 To 2
                    StartDeck(z%, m%) = ChangedDeck(z%, m%)
                Next z%
            Next m%
        End If
    End If
ElseIf ProtectedBlock >= 26 Then
    If InteriorCard = 0 Then
        'complete
        MeshedBlock = 2 * (52 - ProtectedBlock)
        For i% = 1 To 52 - ProtectedBlock
            For z% = 1 To 2
                ChangedDeck(z%, i%) = StartDeck(z%, 2 * i% - 1)
            Next z%
        Next i%
        For j% = 1 To 2 * ProtectedBlock - 52
            For z% = 1 To 2
                ChangedDeck(z%, j% + 52 - ProtectedBlock) = StartDeck(z%, j% + MeshedBlock)
            Next z%
        Next j%
        For k% = 1 To 52 - ProtectedBlock
            For z% = 1 To 2
                ChangedDeck(z%, k% + ProtectedBlock) = StartDeck(z%, 2 * k%)
            Next z%
        Next k%
        For m% = 1 To 52
            For z% = 1 To 2
                StartDeck(z%, m%) = ChangedDeck(z%, m%)
            Next z%
        Next m%
    Else
        'complete
        InteriorPosition = InteriorCard + ProtectedBlock
            For j% = 1 To 52 - InteriorPosition + 1
                For z% = 1 To 2
                    ChangedDeck(z%, j%) = StartDeck(z%, InteriorCard + 2 * (j% - 1))
                Next z%
            Next j%
            For k% = 1 To ProtectedBlock - (52 - InteriorPosition + 1)
                For z% = 1 To 2
                    ChangedDeck(z%, 52 - InteriorPosition + 1 + k%) = _
                        StartDeck(z%, 2 * 52 - 2 * ProtectedBlock - InteriorCard + 1 + k%)
                Next z%
            Next k%
            For b% = 1 To InteriorCard - 1
                For z% = 1 To 2
                    ChangedDeck(z%, ProtectedBlock + b%) = _
                        StartDeck(z%, b%)
                Next z%
            Next b%
            For p% = 1 To 52 - InteriorPosition + 1
                For z% = 1 To 2
                    ChangedDeck(z%, InteriorPosition - 1 + p%) = _
                        StartDeck(z%, InteriorCard + 2 * p% - 1)
                Next z%
            Next p%
            For m% = 1 To 52
                For z% = 1 To 2
                    StartDeck(z%, m%) = ChangedDeck(z%, m%)
                Next z%
            Next m%
    End If
End If
End Sub


Private Sub OutFaro()
For i% = 1 To 26
    For z% = 1 To 2
        ChangedDeck(z%, 2 * i% - 1) = StartDeck(z%, i%)
        ChangedDeck(z%, 2 * i%) = StartDeck(z%, i% + 26)
    Next z%
Next i%
For m% = 1 To 52
    For z% = 1 To 2
        StartDeck(z%, m%) = ChangedDeck(z%, m%)
    Next z%
Next m%
End Sub

Private Sub InverseOutFaro()
For i% = 1 To 26
    For z% = 1 To 2
        ChangedDeck(z%, i%) = StartDeck(z%, 2 * i% - 1)
        ChangedDeck(z%, i% + 26) = StartDeck(z%, 2 * i%)
    Next z%
Next i%
For m% = 1 To 52
    For z% = 1 To 2
        StartDeck(z%, m%) = ChangedDeck(z%, m%)
    Next z%
Next m%
End Sub

Private Sub ShiftTopBlock(sblock, sdepth)
ShiftBlock = Val(sblock)
ShiftDepth = Val(sdepth)
    i = 1
    For j% = ShiftBlock + 1 To ShiftBlock + ShiftDepth
        For z% = 1 To 2
            ChangedDeck(z%, i) = StartDeck(z%, j%)
        Next z%
        i = i + 1
    Next j%
    For k% = 1 To ShiftBlock
        For z% = 1 To 2
            ChangedDeck(z%, i) = StartDeck(z%, k%)
        Next z%
        i = i + 1
    Next k%
    If ShiftBlock + ShiftDepth < 52 Then
        For p% = ShiftBlock + ShiftDepth + 1 To 52
            For z% = 1 To 2
                ChangedDeck(z%, i) = StartDeck(z%, p%)
            Next z%
            i = i + 1
        Next p%
    End If
    For m% = 1 To 52
        For z% = 1 To 2
            StartDeck(z%, m%) = ChangedDeck(z%, m%)
        Next z%
    Next m%
End Sub

Private Sub ShiftTopBlockInverse(sblock, sdepth)
'for the Inverse, just Inverse the first two declarations
'to the opposite assignments from the regular subroutine
ShiftBlock = Val(sdepth)
ShiftDepth = Val(sblock)
    i = 1
    For j% = ShiftBlock + 1 To ShiftBlock + ShiftDepth
        For z% = 1 To 2
            ChangedDeck(z%, i) = StartDeck(z%, j%)
        Next z%
        i = i + 1
    Next j%
    For k% = 1 To ShiftBlock
        For z% = 1 To 2
            ChangedDeck(z%, i) = StartDeck(z%, k%)
        Next z%
        i = i + 1
    Next k%
    If ShiftBlock + ShiftDepth < 52 Then
        For p% = ShiftBlock + ShiftDepth + 1 To 52
            For z% = 1 To 2
                ChangedDeck(z%, i) = StartDeck(z%, p%)
            Next z%
            i = i + 1
        Next p%
    End If
    For m% = 1 To 52
        For z% = 1 To 2
            StartDeck(z%, m%) = ChangedDeck(z%, m%)
        Next z%
    Next m%
End Sub

Private Sub MoveCard(fromcard, tocard)
fromCardParam = Val(fromcard)
toCardParam = Val(tocard)
For z% = 1 To 2
    ChangedDeck(z%, toCardParam) = StartDeck(z%, fromCardParam)
Next z%
If toCardParam < fromCardParam Then
    For j% = 1 To toCardParam - 1
        For z% = 1 To 2
            ChangedDeck(z%, j%) = StartDeck(z%, j%)
        Next z%
    Next j%
    For k% = 1 To fromCardParam - toCardParam
        For z% = 1 To 2
            ChangedDeck(z%, toCardParam + k%) = _
            StartDeck(z%, toCardParam - 1 + k%)
        Next z%
    Next k%
    For n% = 1 To 52 - fromCardParam
        For z% = 1 To 2
            ChangedDeck(z%, fromCardParam + n%) = _
            StartDeck(z%, fromCardParam + n%)
        Next z%
    Next n%
    For m% = 1 To 52
        For z% = 1 To 2
            StartDeck(z%, m%) = ChangedDeck(z%, m%)
        Next z%
    Next m%
ElseIf toCardParam > fromCardParam Then
    For k% = 1 To fromCardParam - 1
        For z% = 1 To 2
            ChangedDeck(z%, k%) = _
            StartDeck(z%, k%)
        Next z%
    Next k%
    For j% = 1 To toCardParam - fromCardParam
        For z% = 1 To 2
            ChangedDeck(z%, fromCardParam - 1 + j%) = _
            StartDeck(z%, fromCardParam + j%)
        Next z%
    Next j%
    For n% = 1 To 52 - toCardParam
        For z% = 1 To 2
            ChangedDeck(z%, toCardParam + n%) = _
            StartDeck(z%, toCardParam + n%)
        Next z%
    Next n%
    For m% = 1 To 52
        For z% = 1 To 2
            StartDeck(z%, m%) = ChangedDeck(z%, m%)
        Next z%
    Next m%
End If
End Sub


Private Sub InFaro()
For i% = 1 To 26
    For z% = 1 To 2
        ChangedDeck(z%, 2 * i% - 1) = StartDeck(z%, i% + 26)
        ChangedDeck(z%, 2 * i%) = StartDeck(z%, i%)
    Next z%
Next i%
For m% = 1 To 52
    For z% = 1 To 2
        StartDeck(z%, m%) = ChangedDeck(z%, m%)
    Next z%
Next m%
End Sub

Private Sub InverseInFaro()
For i% = 1 To 26
    For z% = 1 To 2
        ChangedDeck(z%, i% + 26) = StartDeck(z%, 2 * i% - 1)
        ChangedDeck(z%, i%) = StartDeck(z%, 2 * i%)
    Next z%
Next i%
For m% = 1 To 52
    For z% = 1 To 2
        StartDeck(z%, m%) = ChangedDeck(z%, m%)
    Next z%
Next m%
End Sub

Private Sub InverseDeck()
For i% = 1 To 52
    For z% = 1 To 2
        ChangedDeck(z%, i%) = StartDeck(z%, DeckCount + 1 - i%)
    Next z%
Next i%
For m% = 1 To 52
    For z% = 1 To 2
        StartDeck(z%, m%) = ChangedDeck(z%, m%)
    Next z%
Next m%
End Sub

Private Sub RunSingleCards(runcards)
i = 1
tmp = Val(runcards)
For j% = 52 - tmp + 1 To 52
    For z% = 1 To 2
        ChangedDeck(z%, j%) = StartDeck(z%, tmp - i + 1)
    Next z%
    i = i + 1
Next j%
For k% = 1 To 52 - tmp
    For z% = 1 To 2
        ChangedDeck(z%, k%) = StartDeck(z%, i)
    Next z%
    i = i + 1
Next k%
For m% = 1 To 52
    For z% = 1 To 2
        StartDeck(z%, m%) = ChangedDeck(z%, m%)
    Next z%
Next m%
End Sub

Private Sub RunSingleCardsInverse(runcards)
i = 1
tmp = Val(runcards)
For j% = 52 - tmp + 1 To 52
    For z% = 1 To 2
        ChangedDeck(z%, tmp - i + 1) = StartDeck(z%, j%)
    Next z%
    i = i + 1
Next j%
For k% = 1 To 52 - tmp
    For z% = 1 To 2
        ChangedDeck(z%, i) = StartDeck(z%, k%)
    Next z%
    i = i + 1
Next k%
For m% = 1 To 52
    For z% = 1 To 2
        StartDeck(z%, m%) = ChangedDeck(z%, m%)
    Next z%
Next m%
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
            For z% = 1 To 2
                ChangedDeck(z%, 2 * i%) = StartDeck(z%, i%)
                ChangedDeck(z%, 2 * i% - 1) = StartDeck(z%, i% + ProtectedBlock)
            Next z%
        Next i%
        MeshedBlock = 2 * ProtectedBlock
        For k% = 1 To 52 - MeshedBlock
            For z% = 1 To 2
                ChangedDeck(z%, k% + MeshedBlock) = StartDeck(z%, k% + MeshedBlock)
            Next z%
        Next k%
        For m% = 1 To 52
            For z% = 1 To 2
                StartDeck(z%, m%) = ChangedDeck(z%, m%)
            Next z%
        Next m%
    Else
        InteriorPosition = InteriorCard + ProtectedBlock
        If ProtectedBlock + InteriorPosition <= 52 Then
            For i% = 1 To InteriorPosition - ProtectedBlock
            'added the "- 1" from InFaro solution
                For z% = 1 To 2
                    ChangedDeck(z%, i%) = StartDeck(z%, ProtectedBlock + i%)
                Next z%
            Next i%
            MeshedBlockPosition = InteriorPosition - ProtectedBlock
            'added the "- 1" from InFaro solution
            For i% = 1 To ProtectedBlock
                For z% = 1 To 2
                    ChangedDeck(z%, MeshedBlockPosition + (2 * i% - 1)) = _
                        StartDeck(z%, i%)
                    ChangedDeck(z%, MeshedBlockPosition + (2 * i%)) = _
                        StartDeck(z%, i% + InteriorPosition)
                        'added the "- 1" from InFaro solution
                Next z%
            Next i%
            For i% = ProtectedBlock * 2 + InteriorPosition - ProtectedBlock + 1 To 52
            'took out a "+ 1" right before the "To" from InFaro solution
                For z% = 1 To 2
                    ChangedDeck(z%, i%) = StartDeck(z%, i%)
                Next z%
            Next i%
            For m% = 1 To 52
                For z% = 1 To 2
                    StartDeck(z%, m%) = ChangedDeck(z%, m%)
                Next z%
            Next m%
        Else
            For i% = 1 To InteriorPosition - ProtectedBlock
            'added the "- 1" from InFaro solution
                For z% = 1 To 2
                    ChangedDeck(z%, i%) = StartDeck(z%, ProtectedBlock + i%)
                Next z%
            Next i%
            MeshedBlockPosition = InteriorPosition - ProtectedBlock
            'added the "- 1" from InFaro solution
            MeshedBlock = 2 * (52 - InteriorPosition)
            'added the "+ 1" from InFaro solution (inside parens after Interior Position)
            For i% = 1 To 52 - InteriorPosition
            'added the "+ 1" from InFaro solution
                For z% = 1 To 2
                    ChangedDeck(z%, MeshedBlockPosition + (2 * i% - 1)) = _
                        StartDeck(z%, i%)
                    ChangedDeck(z%, MeshedBlockPosition + (2 * i%)) = _
                        StartDeck(z%, i% + InteriorPosition)
                        'added the "- 1" from InFaro solution (after Interior Position)
                Next z%
            Next i%
            For i% = 1 To ProtectedBlock - (52 - InteriorPosition)
            'added the "+ 1" from InFaro solution (inside parens after Interior Position)
                For z% = 1 To 2
                    ChangedDeck(z%, MeshedBlockPosition + MeshedBlock + i%) = _
                        StartDeck(z%, 52 - InteriorPosition + i%)
                        'added the "+ 1" from InFaro solution after Interior Position
                Next z%
            Next i%
            For m% = 1 To 52
                For z% = 1 To 2
                    StartDeck(z%, m%) = ChangedDeck(z%, m%)
                Next z%
            Next m%
        End If
    End If
ElseIf ProtectedBlock >= 26 Then
    If InteriorCard = 0 Then
        For i% = 1 To 52 - ProtectedBlock
            For z% = 1 To 2
                ChangedDeck(z%, 2 * i%) = StartDeck(z%, i%)
                ChangedDeck(z%, 2 * i% - 1) = StartDeck(z%, i% + ProtectedBlock)
            Next z%
        Next i%
        MeshedBlock = 2 * (52 - ProtectedBlock)
        For k% = 1 To 52 - MeshedBlock
            For z% = 1 To 2
                ChangedDeck(z%, k% + MeshedBlock) = StartDeck(z%, k% + 52 - ProtectedBlock)
            Next z%
        Next k%
        For m% = 1 To 52
            For z% = 1 To 2
                StartDeck(z%, m%) = ChangedDeck(z%, m%)
            Next z%
        Next m%
    Else
        InteriorPosition = InteriorCard + ProtectedBlock
        If ProtectedBlock + InteriorPosition <= 52 Then
            For i% = 1 To InteriorPosition - ProtectedBlock
            'added the "- 1" from InFaro solution
                For z% = 1 To 2
                    ChangedDeck(z%, i%) = StartDeck(z%, ProtectedBlock + i%)
                Next z%
            Next i%
            MeshedBlockPosition = InteriorPosition - ProtectedBlock
            'added the "- 1" from InFaro solution
            For i% = 1 To ProtectedBlock
                For z% = 1 To 2
                    ChangedDeck(z%, MeshedBlockPosition + (2 * i% - 1)) = _
                        StartDeck(z%, i%)
                    ChangedDeck(z%, MeshedBlockPosition + (2 * i%)) = _
                        StartDeck(z%, i% + InteriorPosition)
                        'added the "- 1" from InFaro solution (inside parens after Interior Position)
                Next z%
            Next i%
            For i% = ProtectedBlock * 2 + InteriorPosition - ProtectedBlock + 1 To 52
            'took out a "+ 1" right before the "To" from InFaro solution
                For z% = 1 To 2
                    ChangedDeck(z%, i%) = StartDeck(z%, i%)
                Next z%
            Next i%
            For m% = 1 To 52
                For z% = 1 To 2
                    StartDeck(z%, m%) = ChangedDeck(z%, m%)
                Next z%
            Next m%
        Else
            For i% = 1 To InteriorPosition - ProtectedBlock
            'added the "- 1" from InFaro solution
                For z% = 1 To 2
                    ChangedDeck(z%, i%) = StartDeck(z%, ProtectedBlock + i%)
                Next z%
            Next i%
            MeshedBlockPosition = InteriorPosition - ProtectedBlock
            'added the "- 1" from InFaro solution
            MeshedBlock = 2 * (52 - InteriorPosition)
            'added the "+ 1" from InFaro solution inside parens after InteriorPosition)
            For i% = 1 To 52 - InteriorPosition
            'added the "+ 1" from InFaro solution
                For z% = 1 To 2
                    ChangedDeck(z%, MeshedBlockPosition + (2 * i% - 1)) = _
                        StartDeck(z%, i%)
                    ChangedDeck(z%, MeshedBlockPosition + (2 * i%)) = _
                        StartDeck(z%, i% + InteriorPosition)
                        'added the "- 1" from InFaro solution inside parens after Interior Position
                Next z%
            Next i%
            For i% = 1 To ProtectedBlock - (52 - InteriorPosition)
            'added the "+ 1" from InFaro solution inside parens after InteriorPosition)
                For z% = 1 To 2
                    ChangedDeck(z%, MeshedBlockPosition + MeshedBlock + i%) = _
                        StartDeck(z%, 52 - InteriorPosition + i%)
                        'added the "+ 1" from InFaro solution after InteriorPosition
                Next z%
            Next i%
            For m% = 1 To 52
                For z% = 1 To 2
                    StartDeck(z%, m%) = ChangedDeck(z%, m%)
                Next z%
            Next m%
        End If
    End If
End If
InverseDeck
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
        For i% = 1 To 52 - MeshedBlock
            For z% = 1 To 2
                ChangedDeck(z%, i% + MeshedBlock) = StartDeck(z%, i% + MeshedBlock)
            Next z%
        Next i%
        For k% = 1 To ProtectedBlock
            For z% = 1 To 2
                ChangedDeck(z%, k%) = StartDeck(z%, 2 * k%)
                'for infaro, tool out the -1 after k% in Deck only
                ChangedDeck(z%, k% + ProtectedBlock) = StartDeck(z%, 2 * k% - 1)
                'for infaro added the -1 after k%
            Next z%
        Next k%
        For m% = 1 To 52
            For z% = 1 To 2
                StartDeck(z%, m%) = ChangedDeck(z%, m%)
            Next z%
        Next m%
    Else
        'complete
        InteriorPosition = InteriorCard + ProtectedBlock
        If ProtectedBlock + InteriorPosition <= 52 Then
            For j% = 1 To ProtectedBlock
                For z% = 1 To 2
                    ChangedDeck(z%, j%) = StartDeck(z%, InteriorCard + 2 * j% - 1)
                    'for infaro, removed the -1 after (j%) in Deck only
                Next z%
            Next j%
            For k% = 1 To InteriorCard
                'for infaro removed the -1 after InteriorCard
                For z% = 1 To 2
                    ChangedDeck(z%, ProtectedBlock + k%) = StartDeck(z%, k%)
                Next z%
            Next k%
            For b% = 1 To ProtectedBlock
                For z% = 1 To 2
                    ChangedDeck(z%, InteriorPosition + b%) = StartDeck(z%, InteriorCard + (2 * b%))
                    'for infaro took out the -1 after b%) in both terms
                Next z%
            Next b%
            For p% = InteriorPosition + ProtectedBlock To 52
                For z% = 1 To 2
                    ChangedDeck(z%, p%) = StartDeck(z%, p%)
                Next z%
            Next p%
            For m% = 1 To 52
                For z% = 1 To 2
                    StartDeck(z%, m%) = ChangedDeck(z%, m%)
                Next z%
            Next m%
        Else
            'complete
            For j% = 1 To 52 - InteriorPosition
            'for infaro removed the + 1 after InteriorPosition
                For z% = 1 To 2
                    ChangedDeck(z%, j%) = StartDeck(z%, InteriorCard + 2 * j% - 1)
                    'for infaro changed from 2* (j% - 1)
                Next z%
            Next j%
            For k% = 1 To ProtectedBlock - (52 - InteriorPosition)
            'for infaro removed the + 1 after InteriorPosition
                For z% = 1 To 2
                    ChangedDeck(z%, 52 - InteriorPosition + k%) = _
                        StartDeck(z%, 2 * 52 - 2 * ProtectedBlock - InteriorCard + k%)
                        'for infaro took out + 1 after InteriorCard in both terms
                Next z%
            Next k%
            For b% = 1 To InteriorCard
            'for infaro removed the -1 after InteriorCard
                For z% = 1 To 2
                    ChangedDeck(z%, ProtectedBlock + b%) = _
                        StartDeck(z%, b%)
                Next z%
            Next b%
            For p% = 1 To 52 - InteriorPosition
                'for infaro removed the + 1 after InteriorPosition
                For z% = 1 To 2
                    ChangedDeck(z%, InteriorPosition + p%) = _
                        StartDeck(z%, InteriorCard + 2 * p%)
                        'for infaro removed the -1 after InteriorPosition
                        'and changed from 2 * p% - 1
                Next z%
            Next p%
            For m% = 1 To 52
                For z% = 1 To 2
                    StartDeck(z%, m%) = ChangedDeck(z%, m%)
                Next z%
            Next m%
        End If
    End If
ElseIf ProtectedBlock >= 26 Then
    If InteriorCard = 0 Then
        'complete
        MeshedBlock = 2 * (52 - ProtectedBlock)
        For i% = 1 To 52 - ProtectedBlock
            For z% = 1 To 2
                ChangedDeck(z%, i%) = StartDeck(z%, 2 * i%)
                'for infaro removed the -1 after i% in Deck only
            Next z%
        Next i%
        For j% = 1 To 2 * ProtectedBlock - 52
            For z% = 1 To 2
                ChangedDeck(z%, j% + 52 - ProtectedBlock) = StartDeck(z%, j% + MeshedBlock)
            Next z%
        Next j%
        For k% = 1 To 52 - ProtectedBlock
            For z% = 1 To 2
                ChangedDeck(z%, k% + ProtectedBlock) = StartDeck(z%, 2 * k% - 1)
                'for infaro added the - 1 after k% in Deck only
            Next z%
        Next k%
        For m% = 1 To 52
            For z% = 1 To 2
                StartDeck(z%, m%) = ChangedDeck(z%, m%)
            Next z%
        Next m%
    Else
        'complete
        InteriorPosition = InteriorCard + ProtectedBlock
            For j% = 1 To 52 - InteriorPosition
            'for infaro removed the + 1 after InteriorPostion
                For z% = 1 To 2
                    ChangedDeck(z%, j%) = StartDeck(z%, InteriorCard + 2 * j% - 1)
                    'for infaro changed from 2 * (j% - 1)
                Next z%
            Next j%
            For k% = 1 To ProtectedBlock - (52 - InteriorPosition)
                    'for infaro removed the + 1 after InteriorPostion
                For z% = 1 To 2
                    ChangedDeck(z%, 52 - InteriorPosition + k%) = _
                        StartDeck(z%, 2 * 52 - 2 * ProtectedBlock - InteriorCard + k%)
                        'for infaro removed the + 1 after InteriorPostion and Interior card
                Next z%
            Next k%
            For b% = 1 To InteriorCard
            'for infaro removed the - 1 after InteriorCard
                For z% = 1 To 2
                    ChangedDeck(z%, ProtectedBlock + b%) = _
                        StartDeck(z%, b%)
                Next z%
            Next b%
            For p% = 1 To 52 - InteriorPosition
            'for infaro removed the + 1 after InteriorPostion
                For z% = 1 To 2
                    ChangedDeck(z%, InteriorPosition + p%) = _
                        StartDeck(z%, InteriorCard + 2 * p%)
                        'for infaro removed the -1 after InteriorPosition
                        'and changed from 2 * p% - 1
                Next z%
            Next p%
            For m% = 1 To 52
                For z% = 1 To 2
                    StartDeck(z%, m%) = ChangedDeck(z%, m%)
                Next z%
            Next m%
    End If
End If
Call CutDeckPrecise(ProtectedBlock, "X")
InverseDeck
End Sub


Private Sub InFaroSpecialTop(istnumber, ifinumber)
ProtectedBlock = Val(istnumber)
InteriorCard = Val(ifinumber)
If ProtectedBlock = 26 And InteriorCard = Empty Then
    InFaro
ElseIf ProtectedBlock < 26 Then
    If InteriorCard = Empty Then
        For i% = 1 To ProtectedBlock
            For z% = 1 To 2
                ChangedDeck(z%, 2 * i% - 1) = StartDeck(z%, i% + ProtectedBlock)
                ChangedDeck(z%, 2 * i%) = StartDeck(z%, i%)
            Next z%
        Next i%
        MeshedBlock = 2 * ProtectedBlock
        For k% = 1 To 52 - MeshedBlock
            For z% = 1 To 2
                ChangedDeck(z%, k% + MeshedBlock) = StartDeck(z%, k% + MeshedBlock)
            Next z%
        Next k%
        For m% = 1 To 52
            For z% = 1 To 2
                StartDeck(z%, m%) = ChangedDeck(z%, m%)
            Next z%
        Next m%
    Else
        InteriorPosition = InteriorCard + ProtectedBlock
        If ProtectedBlock + InteriorPosition <= 52 Then
            For i% = 1 To InteriorPosition - ProtectedBlock
            'added the "- 1" from InFaro solution
                For z% = 1 To 2
                    ChangedDeck(z%, i%) = StartDeck(z%, ProtectedBlock + i%)
                Next z%
            Next i%
            MeshedBlockPosition = InteriorPosition - ProtectedBlock
            'added the "- 1" from InFaro solution
            For i% = 1 To ProtectedBlock
                For z% = 1 To 2
                    ChangedDeck(z%, MeshedBlockPosition + (2 * i% - 1)) = _
                        StartDeck(z%, i%)
                    ChangedDeck(z%, MeshedBlockPosition + (2 * i%)) = _
                        StartDeck(z%, i% + InteriorPosition)
                        'added the "- 1" from InFaro solution
                Next z%
            Next i%
            For i% = ProtectedBlock * 2 + InteriorPosition - ProtectedBlock + 1 To 52
            'took out a "+ 1" right before the "To" from InFaro solution
                For z% = 1 To 2
                    ChangedDeck(z%, i%) = StartDeck(z%, i%)
                Next z%
            Next i%
            For m% = 1 To 52
                For z% = 1 To 2
                    StartDeck(z%, m%) = ChangedDeck(z%, m%)
                Next z%
            Next m%
        Else
            For i% = 1 To InteriorPosition - ProtectedBlock
            'added the "- 1" from InFaro solution
                For z% = 1 To 2
                    ChangedDeck(z%, i%) = StartDeck(z%, ProtectedBlock + i%)
                Next z%
            Next i%
            MeshedBlockPosition = InteriorPosition - ProtectedBlock
            'added the "- 1" from InFaro solution
            MeshedBlock = 2 * (52 - InteriorPosition)
            'added the "+ 1" from InFaro solution (inside parens after Interior Position)
            For i% = 1 To 52 - InteriorPosition
            'added the "+ 1" from InFaro solution
                For z% = 1 To 2
                    ChangedDeck(z%, MeshedBlockPosition + (2 * i% - 1)) = _
                        StartDeck(z%, i%)
                    ChangedDeck(z%, MeshedBlockPosition + (2 * i%)) = _
                        StartDeck(z%, i% + InteriorPosition)
                        'added the "- 1" from InFaro solution (after Interior Position)
                Next z%
            Next i%
            For i% = 1 To ProtectedBlock - (52 - InteriorPosition)
            'added the "+ 1" from InFaro solution (inside parens after Interior Position)
                For z% = 1 To 2
                    ChangedDeck(z%, MeshedBlockPosition + MeshedBlock + i%) = _
                        StartDeck(z%, 52 - InteriorPosition + i%)
                        'added the "+ 1" from InFaro solution after Interior Position
                Next z%
            Next i%
            For m% = 1 To 52
                For z% = 1 To 2
                    StartDeck(z%, m%) = ChangedDeck(z%, m%)
                Next z%
            Next m%
        End If
    End If
ElseIf ProtectedBlock >= 26 Then
    If InteriorCard = Empty Then
        For i% = 1 To 52 - ProtectedBlock
            For z% = 1 To 2
                ChangedDeck(z%, 2 * i% - 1) = StartDeck(z%, i% + ProtectedBlock)
                ChangedDeck(z%, 2 * i%) = StartDeck(z%, i%)
            Next z%
        Next i%
        MeshedBlock = 2 * (52 - ProtectedBlock)
        For k% = 1 To 52 - MeshedBlock
            For z% = 1 To 2
                ChangedDeck(z%, k% + MeshedBlock) = StartDeck(z%, k% + 52 - ProtectedBlock)
            Next z%
        Next k%
        For m% = 1 To 52
            For z% = 1 To 2
                StartDeck(z%, m%) = ChangedDeck(z%, m%)
            Next z%
        Next m%
    Else
        InteriorPosition = InteriorCard + ProtectedBlock
        If ProtectedBlock + InteriorPosition <= 52 Then
            For i% = 1 To InteriorPosition - ProtectedBlock
            'added the "- 1" from InFaro solution
                For z% = 1 To 2
                    ChangedDeck(z%, i%) = StartDeck(z%, ProtectedBlock + i%)
                Next z%
            Next i%
            MeshedBlockPosition = InteriorPosition - ProtectedBlock
            'added the "- 1" from InFaro solution
            For i% = 1 To ProtectedBlock
                For z% = 1 To 2
                    ChangedDeck(z%, MeshedBlockPosition + (2 * i% - 1)) = _
                        StartDeck(z%, i%)
                    ChangedDeck(z%, MeshedBlockPosition + (2 * i%)) = _
                        StartDeck(z%, i% + InteriorPosition)
                        'added the "- 1" from InFaro solution (inside parens after Interior Position)
                Next z%
            Next i%
            For i% = ProtectedBlock * 2 + InteriorPosition - ProtectedBlock + 1 To 52
            'took out a "+ 1" right before the "To" from InFaro solution
                For z% = 1 To 2
                    ChangedDeck(z%, i%) = StartDeck(z%, i%)
                Next z%
            Next i%
            For m% = 1 To 52
                For z% = 1 To 2
                    StartDeck(z%, m%) = ChangedDeck(z%, m%)
                Next z%
            Next m%
        Else
            For i% = 1 To InteriorPosition - ProtectedBlock
            'added the "- 1" from InFaro solution
                For z% = 1 To 2
                    ChangedDeck(z%, i%) = StartDeck(z%, ProtectedBlock + i%)
                Next z%
            Next i%
            MeshedBlockPosition = InteriorPosition - ProtectedBlock
            'added the "- 1" from InFaro solution
            MeshedBlock = 2 * (52 - InteriorPosition)
            'added the "+ 1" from InFaro solution inside parens after InteriorPosition)
            For i% = 1 To 52 - InteriorPosition
            'added the "+ 1" from InFaro solution
                For z% = 1 To 2
                    ChangedDeck(z%, MeshedBlockPosition + (2 * i% - 1)) = _
                        StartDeck(z%, i%)
                    ChangedDeck(z%, MeshedBlockPosition + (2 * i%)) = _
                        StartDeck(z%, i% + InteriorPosition)
                        'added the "- 1" from InFaro solution inside parens after Interior Position
                Next z%
            Next i%
            For i% = 1 To ProtectedBlock - (52 - InteriorPosition)
            'added the "+ 1" from InFaro solution inside parens after InteriorPosition)
                For z% = 1 To 2
                    ChangedDeck(z%, MeshedBlockPosition + MeshedBlock + i%) = _
                        StartDeck(z%, 52 - InteriorPosition + i%)
                        'added the "+ 1" from InFaro solution after InteriorPosition
                Next z%
            Next i%
            For m% = 1 To 52
                For z% = 1 To 2
                    StartDeck(z%, m%) = ChangedDeck(z%, m%)
                Next z%
            Next m%
        End If
    End If
End If
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
        For i% = 1 To 52 - MeshedBlock
            For z% = 1 To 2
                ChangedDeck(z%, i% + MeshedBlock) = StartDeck(z%, i% + MeshedBlock)
            Next z%
        Next i%
        For k% = 1 To ProtectedBlock
            For z% = 1 To 2
                ChangedDeck(z%, k%) = StartDeck(z%, 2 * k%)
                ChangedDeck(z%, k% + ProtectedBlock) = StartDeck(z%, 2 * k% - 1)
            Next z%
        Next k%
        For m% = 1 To 52
            For z% = 1 To 2
                StartDeck(z%, m%) = ChangedDeck(z%, m%)
            Next z%
        Next m%
    Else
        'complete
        InteriorPosition = InteriorCard + ProtectedBlock
        If ProtectedBlock + InteriorPosition <= 52 Then
            For j% = 1 To ProtectedBlock
                For z% = 1 To 2
                    ChangedDeck(z%, j%) = StartDeck(z%, InteriorCard + 2 * j% - 1)
                Next z%
            Next j%
            For k% = 1 To InteriorCard
                For z% = 1 To 2
                    ChangedDeck(z%, ProtectedBlock + k%) = StartDeck(z%, k%)
                Next z%
            Next k%
            For b% = 1 To ProtectedBlock
                For z% = 1 To 2
                    ChangedDeck(z%, InteriorPosition + b%) = StartDeck(z%, InteriorCard + (2 * b%))
                Next z%
            Next b%
            For p% = InteriorPosition + ProtectedBlock + 1 To 52
                For z% = 1 To 2
                    ChangedDeck(z%, p%) = StartDeck(z%, p%)
                Next z%
            Next p%
            For m% = 1 To 52
                For z% = 1 To 2
                    StartDeck(z%, m%) = ChangedDeck(z%, m%)
                Next z%
            Next m%
        Else
            'complete
            For j% = 1 To 52 - InteriorPosition
                For z% = 1 To 2
                    ChangedDeck(z%, j%) = StartDeck(z%, InteriorCard + 2 * j% - 1)
                Next z%
            Next j%
            For k% = 1 To ProtectedBlock - (52 - InteriorPosition)
                For z% = 1 To 2
                    ChangedDeck(z%, 52 - InteriorPosition + k%) = _
                        StartDeck(z%, 2 * 52 - 2 * ProtectedBlock - InteriorCard + k%)
                Next z%
            Next k%
            For b% = 1 To InteriorCard
                For z% = 1 To 2
                    ChangedDeck(z%, ProtectedBlock + b%) = _
                        StartDeck(z%, b%)
                Next z%
            Next b%
            For p% = 1 To 52 - InteriorPosition
                For z% = 1 To 2
                    ChangedDeck(z%, InteriorPosition + p%) = _
                        StartDeck(z%, InteriorCard + 2 * p%)
                Next z%
            Next p%
            For m% = 1 To 52
                For z% = 1 To 2
                    StartDeck(z%, m%) = ChangedDeck(z%, m%)
                Next z%
            Next m%
        End If
    End If
ElseIf ProtectedBlock >= 26 Then
    If InteriorCard = 0 Then
        'complete
        MeshedBlock = 2 * (52 - ProtectedBlock)
        For i% = 1 To 52 - ProtectedBlock
            For z% = 1 To 2
                ChangedDeck(z%, i%) = StartDeck(z%, 2 * i%)
            Next z%
        Next i%
        For j% = 1 To 2 * ProtectedBlock - 52
            For z% = 1 To 2
                ChangedDeck(z%, j% + 52 - ProtectedBlock) = StartDeck(z%, j% + MeshedBlock)
            Next z%
        Next j%
        For k% = 1 To 52 - ProtectedBlock
            For z% = 1 To 2
                ChangedDeck(z%, k% + ProtectedBlock) = StartDeck(z%, 2 * k% - 1)
            Next z%
        Next k%
        For m% = 1 To 52
            For z% = 1 To 2
                StartDeck(z%, m%) = ChangedDeck(z%, m%)
            Next z%
        Next m%
    Else
        'complete
        InteriorPosition = InteriorCard + ProtectedBlock
            For j% = 1 To 52 - InteriorPosition
                For z% = 1 To 2
                    ChangedDeck(z%, j%) = StartDeck(z%, InteriorCard + 2 * j% - 1)
                Next z%
            Next j%
            For k% = 1 To ProtectedBlock - (52 - InteriorPosition)
                For z% = 1 To 2
                    ChangedDeck(z%, 52 - InteriorPosition + k%) = _
                        StartDeck(z%, 2 * 52 - 2 * ProtectedBlock - InteriorCard + k%)
                Next z%
            Next k%
            For b% = 1 To InteriorCard
                For z% = 1 To 2
                    ChangedDeck(z%, ProtectedBlock + b%) = _
                        StartDeck(z%, b%)
                Next z%
            Next b%
            For p% = 1 To 52 - InteriorPosition
                For z% = 1 To 2
                    ChangedDeck(z%, InteriorPosition + p%) = _
                        StartDeck(z%, InteriorCard + 2 * p%)
                Next z%
            Next p%
            For m% = 1 To 52
                For z% = 1 To 2
                    StartDeck(z%, m%) = ChangedDeck(z%, m%)
                Next z%
            Next m%
    End If
End If
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
            For z% = 1 To 2
                ChangedDeck(z%, 2 * i% - 1) = StartDeck(z%, i%)
                ChangedDeck(z%, 2 * i%) = StartDeck(z%, i% + ProtectedBlock)
            Next z%
        Next i%
        MeshedBlock = 2 * ProtectedBlock
        For k% = 1 To 52 - MeshedBlock
            For z% = 1 To 2
                ChangedDeck(z%, k% + MeshedBlock) = StartDeck(z%, k% + MeshedBlock)
            Next z%
        Next k%
        For m% = 1 To 52
            For z% = 1 To 2
                StartDeck(z%, m%) = ChangedDeck(z%, m%)
            Next z%
        Next m%
    Else
        InteriorPosition = InteriorCard + ProtectedBlock
        If ProtectedBlock + InteriorPosition <= 52 Then
            For i% = 1 To InteriorPosition - ProtectedBlock - 1
            'added the "- 1" from InFaro solution
                For z% = 1 To 2
                    ChangedDeck(z%, i%) = StartDeck(z%, ProtectedBlock + i%)
                Next z%
            Next i%
            MeshedBlockPosition = InteriorPosition - ProtectedBlock - 1
            'added the "- 1" from InFaro solution
            For i% = 1 To ProtectedBlock
                For z% = 1 To 2
                    ChangedDeck(z%, MeshedBlockPosition + (2 * i% - 1)) = _
                        StartDeck(z%, i%)
                    ChangedDeck(z%, MeshedBlockPosition + (2 * i%)) = _
                        StartDeck(z%, i% + InteriorPosition - 1)
                        'added the "- 1" from InFaro solution
                Next z%
            Next i%
            For i% = ProtectedBlock * 2 + InteriorPosition - ProtectedBlock To 52
            'took out a "+ 1" right before the "To" from InFaro solution
                For z% = 1 To 2
                    ChangedDeck(z%, i%) = StartDeck(z%, i%)
                Next z%
            Next i%
            For m% = 1 To 52
                For z% = 1 To 2
                    StartDeck(z%, m%) = ChangedDeck(z%, m%)
                Next z%
            Next m%
        Else
            For i% = 1 To InteriorPosition - ProtectedBlock - 1
            'added the "- 1" from InFaro solution
                For z% = 1 To 2
                    ChangedDeck(z%, i%) = StartDeck(z%, ProtectedBlock + i%)
                Next z%
            Next i%
            MeshedBlockPosition = InteriorPosition - ProtectedBlock - 1
            'added the "- 1" from InFaro solution
            MeshedBlock = 2 * (52 - InteriorPosition + 1)
            'added the "+ 1" from InFaro solution
            For i% = 1 To 52 - InteriorPosition + 1
            'added the "+ 1" from InFaro solution
                For z% = 1 To 2
                    ChangedDeck(z%, MeshedBlockPosition + (2 * i% - 1)) = _
                        StartDeck(z%, i%)
                    ChangedDeck(z%, MeshedBlockPosition + (2 * i%)) = _
                        StartDeck(z%, i% + InteriorPosition - 1)
                        'added the "- 1" from InFaro solution
                Next z%
            Next i%
            For i% = 1 To ProtectedBlock - (52 - InteriorPosition + 1)
            'added the "+ 1" from InFaro solution
                For z% = 1 To 2
                    ChangedDeck(z%, MeshedBlockPosition + MeshedBlock + i%) = _
                        StartDeck(z%, 52 - InteriorPosition + 1 + i%)
                        'added the "+ 1" from InFaro solution
                Next z%
            Next i%
            For m% = 1 To 52
                For z% = 1 To 2
                    StartDeck(z%, m%) = ChangedDeck(z%, m%)
                Next z%
            Next m%
        End If
    End If
ElseIf ProtectedBlock >= 26 Then
    If InteriorCard = Empty Then
        For i% = 1 To 52 - ProtectedBlock
            For z% = 1 To 2
                ChangedDeck(z%, 2 * i% - 1) = StartDeck(z%, i%)
                ChangedDeck(z%, 2 * i%) = StartDeck(z%, i% + ProtectedBlock)
            Next z%
        Next i%
        MeshedBlock = 2 * (52 - ProtectedBlock)
        For k% = 1 To 52 - MeshedBlock
            For z% = 1 To 2
                ChangedDeck(z%, k% + MeshedBlock) = StartDeck(z%, k% + 52 - ProtectedBlock)
            Next z%
        Next k%
        For m% = 1 To 52
            For z% = 1 To 2
                StartDeck(z%, m%) = ChangedDeck(z%, m%)
            Next z%
        Next m%
    Else
        InteriorPosition = InteriorCard + ProtectedBlock
        If ProtectedBlock + InteriorPosition <= 52 Then
            For i% = 1 To InteriorPosition - ProtectedBlock - 1
            'added the "- 1" from InFaro solution
                For z% = 1 To 2
                    ChangedDeck(z%, i%) = StartDeck(z%, ProtectedBlock + i%)
                Next z%
            Next i%
            MeshedBlockPosition = InteriorPosition - ProtectedBlock - 1
            'added the "- 1" from InFaro solution
            For i% = 1 To ProtectedBlock
                For z% = 1 To 2
                    ChangedDeck(z%, MeshedBlockPosition + (2 * i% - 1)) = _
                        StartDeck(z%, i%)
                    ChangedDeck(z%, MeshedBlockPosition + (2 * i%)) = _
                        StartDeck(z%, i% + InteriorPosition - 1)
                        'added the "- 1" from InFaro solution
                Next z%
            Next i%
            For i% = ProtectedBlock * 2 + InteriorPosition - ProtectedBlock To 52
            'took out a "+ 1" right before the "To" from InFaro solution
                For z% = 1 To 2
                    ChangedDeck(z%, i%) = StartDeck(z%, i%)
                Next z%
            Next i%
            For m% = 1 To 52
                For z% = 1 To 2
                    StartDeck(z%, m%) = ChangedDeck(z%, m%)
                Next z%
            Next m%
        Else
            For i% = 1 To InteriorPosition - ProtectedBlock - 1
            'added the "- 1" from InFaro solution
                For z% = 1 To 2
                    ChangedDeck(z%, i%) = StartDeck(z%, ProtectedBlock + i%)
                Next z%
            Next i%
            MeshedBlockPosition = InteriorPosition - ProtectedBlock - 1
            'added the "- 1" from InFaro solution
            MeshedBlock = 2 * (52 - InteriorPosition + 1)
            'added the "+ 1" from InFaro solution
            For i% = 1 To 52 - InteriorPosition + 1
            'added the "+ 1" from InFaro solution
                For z% = 1 To 2
                    ChangedDeck(z%, MeshedBlockPosition + (2 * i% - 1)) = _
                        StartDeck(z%, i%)
                    ChangedDeck(z%, MeshedBlockPosition + (2 * i%)) = _
                        StartDeck(z%, i% + InteriorPosition - 1)
                        'added the "- 1" from InFaro solution
                Next z%
            Next i%
            For i% = 1 To ProtectedBlock - (52 - InteriorPosition + 1)
            'added the "+ 1" from InFaro solution
                For z% = 1 To 2
                    ChangedDeck(z%, MeshedBlockPosition + MeshedBlock + i%) = _
                        StartDeck(z%, 52 - InteriorPosition + 1 + i%)
                        'added the "+ 1" from InFaro solution
                Next z%
            Next i%
            For m% = 1 To 52
                For z% = 1 To 2
                    StartDeck(z%, m%) = ChangedDeck(z%, m%)
                Next z%
            Next m%
        End If
    End If
End If
InverseDeck
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
        For i% = 1 To 52 - MeshedBlock
            For z% = 1 To 2
                ChangedDeck(z%, i% + MeshedBlock) = StartDeck(z%, i% + MeshedBlock)
            Next z%
        Next i%
        For k% = 1 To ProtectedBlock
            For z% = 1 To 2
                ChangedDeck(z%, k%) = StartDeck(z%, 2 * k% - 1)
                'for infaro, tool out the -1 after k% in Deck only
                ChangedDeck(z%, k% + ProtectedBlock) = StartDeck(z%, 2 * k%)
                'for infaro added the -1 after k%
            Next z%
        Next k%
        For m% = 1 To 52
            For z% = 1 To 2
                StartDeck(z%, m%) = ChangedDeck(z%, m%)
            Next z%
        Next m%
    Else
        'complete
        InteriorPosition = InteriorCard + ProtectedBlock
        If ProtectedBlock + InteriorPosition <= 52 Then
            For j% = 1 To ProtectedBlock
                For z% = 1 To 2
                    ChangedDeck(z%, j%) = StartDeck(z%, InteriorCard + 2 * (j% - 1))
                    'for infaro, removed the -1 after (j%) in Deck only
                Next z%
            Next j%
            For k% = 1 To InteriorCard - 1
                'for infaro removed the -1 after InteriorCard
                For z% = 1 To 2
                    ChangedDeck(z%, ProtectedBlock + k%) = StartDeck(z%, k%)
                Next z%
            Next k%
            For b% = 1 To ProtectedBlock
                For z% = 1 To 2
                    ChangedDeck(z%, InteriorPosition + b% - 1) = StartDeck(z%, InteriorCard + (2 * b%) - 1)
                    'for infaro took out the -1 after b%) in both terms
                Next z%
            Next b%
            For p% = InteriorPosition + ProtectedBlock To 52
                For z% = 1 To 2
                    ChangedDeck(z%, p%) = StartDeck(z%, p%)
                Next z%
            Next p%
            For m% = 1 To 52
                For z% = 1 To 2
                    StartDeck(z%, m%) = ChangedDeck(z%, m%)
                Next z%
            Next m%
        Else
            'complete
            For j% = 1 To 52 - InteriorPosition + 1
            'for infaro removed the + 1 after InteriorPosition
                For z% = 1 To 2
                    ChangedDeck(z%, j%) = StartDeck(z%, InteriorCard + 2 * (j% - 1))
                    'for infaro changed from 2* (j% - 1)
                Next z%
            Next j%
            For k% = 1 To ProtectedBlock - (52 - InteriorPosition + 1)
            'for infaro removed the + 1 after InteriorPosition
                For z% = 1 To 2
                    ChangedDeck(z%, 52 - InteriorPosition + 1 + k%) = _
                        StartDeck(z%, 2 * 52 - 2 * ProtectedBlock - InteriorCard + 1 + k%)
                        'for infaro took out + 1 after InteriorCard in both terms
                Next z%
            Next k%
            For b% = 1 To InteriorCard - 1
            'for infaro removed the -1 after InteriorCard
                For z% = 1 To 2
                    ChangedDeck(z%, ProtectedBlock + b%) = _
                        StartDeck(z%, b%)
                Next z%
            Next b%
            For p% = 1 To 52 - InteriorPosition + 1
                'for infaro removed the + 1 after InteriorPosition
                For z% = 1 To 2
                    ChangedDeck(z%, InteriorPosition - 1 + p%) = _
                        StartDeck(z%, InteriorCard + 2 * p% - 1)
                        'for infaro removed the -1 after InteriorPosition
                        'and changed from 2 * p% - 1
                Next z%
            Next p%
            For m% = 1 To 52
                For z% = 1 To 2
                    StartDeck(z%, m%) = ChangedDeck(z%, m%)
                Next z%
            Next m%
        End If
    End If
ElseIf ProtectedBlock >= 26 Then
    If InteriorCard = 0 Then
        'complete
        MeshedBlock = 2 * (52 - ProtectedBlock)
        For i% = 1 To 52 - ProtectedBlock
            For z% = 1 To 2
                ChangedDeck(z%, i%) = StartDeck(z%, 2 * i% - 1)
                'for infaro removed the -1 after i% in Deck only
            Next z%
        Next i%
        For j% = 1 To 2 * ProtectedBlock - 52
            For z% = 1 To 2
                ChangedDeck(z%, j% + 52 - ProtectedBlock) = StartDeck(z%, j% + MeshedBlock)
            Next z%
        Next j%
        For k% = 1 To 52 - ProtectedBlock
            For z% = 1 To 2
                ChangedDeck(z%, k% + ProtectedBlock) = StartDeck(z%, 2 * k%)
                'for infaro added the - 1 after k% in Deck only
            Next z%
        Next k%
        For m% = 1 To 52
            For z% = 1 To 2
                StartDeck(z%, m%) = ChangedDeck(z%, m%)
            Next z%
        Next m%
    Else
        'complete
        InteriorPosition = InteriorCard + ProtectedBlock
            For j% = 1 To 52 - InteriorPosition + 1
            'for infaro removed the + 1 after InteriorPostion
                For z% = 1 To 2
                    ChangedDeck(z%, j%) = StartDeck(z%, InteriorCard + 2 * (j% - 1))
                    'for infaro changed from 2 * (j% - 1)
                Next z%
            Next j%
            For k% = 1 To ProtectedBlock - (52 - InteriorPosition + 1)
                    'for infaro removed the + 1 after InteriorPostion
                For z% = 1 To 2
                    ChangedDeck(z%, 52 - InteriorPosition + 1 + k%) = _
                        StartDeck(z%, 2 * 52 - 2 * ProtectedBlock - InteriorCard + 1 + k%)
                        'for infaro removed the + 1 after InteriorPostion and Interior card
                Next z%
            Next k%
            For b% = 1 To InteriorCard - 1
            'for infaro removed the - 1 after InteriorCard
                For z% = 1 To 2
                    ChangedDeck(z%, ProtectedBlock + b%) = _
                        StartDeck(z%, b%)
                Next z%
            Next b%
            For p% = 1 To 52 - InteriorPosition + 1
            'for infaro removed the + 1 after InteriorPostion
                For z% = 1 To 2
                    ChangedDeck(z%, InteriorPosition - 1 + p%) = _
                        StartDeck(z%, InteriorCard + 2 * p% - 1)
                        'for infaro removed the -1 after InteriorPosition
                        'and changed from 2 * p% - 1
                Next z%
            Next p%
            For m% = 1 To 52
                For z% = 1 To 2
                    StartDeck(z%, m%) = ChangedDeck(z%, m%)
                Next z%
            Next m%
    End If
End If
Call CutDeckPrecise(ProtectedBlock, "X")
InverseDeck
End Sub

Function Min(firstParam, secondParam)
If firstParam <= secondParam Then
    Min = firstParam
Else
    Min = secondParam
End If
End Function

Private Sub TransferToSession_Click()
If MatchListBox.ListCount = 0 Then
    MsgBox ("There is nothing to transfer.")
    Exit Sub
End If
If SessionRecordMode = True Then
    MsgBox ("The Session module is still in Recording mode." & Chr(13) & _
        "You must turn off the recording mode before you can transfer.")
    Exit Sub
End If
If SessionSaved = 0 Then
    Dim Msg, Style, Title, Help, Ctxt, Response, MyString
    Msg = "There are still unsaved Session Events on the Sessions screen" & _
        Chr(13) & _
        "This will CLEAR ALL Session Events before transfering the results." & Chr(13) & _
        Chr(13) & "Do you want to continue?"   ' Define message.
    Style = vbYesNo + vbQuestion + vbDefaultButton2   ' Define buttons.
    Title = "Clear ALL Session Events"   ' Define title.
    Response = MsgBox(Msg, Style, Title)
    If Response = vbYes Then   ' User chose Yes.
        frmStackView.SessionListBox.Clear
        sessioncount = MatchListBox.ListCount
        For i% = 0 To sessioncount - 1
            frmStackView.SessionListBox.AddItem MatchListBox.List(i%)
        Next i%
        frmStackView.SessionStatusUpdate (0)
        SearchSessionTransferred = 1
    End If
Else
    frmStackView.SessionListBox.Clear
    sessioncount = MatchListBox.ListCount
    For i% = 0 To sessioncount - 1
        frmStackView.SessionListBox.AddItem MatchListBox.List(i%)
    Next i%
    frmStackView.SessionStatusUpdate (0)
    SearchSessionTransferred = 1
End If
End Sub

Private Sub SearchDisableContinueCheck()
If SearchContinueReady = 1 Then
    Dim Msg, Style, Title, Help, Ctxt, Response, MyString
    Msg = "Changing Search parameters will require a new start." & _
        Chr(13) & _
        "You will not be able to continue the current search." & Chr(13) & _
        Chr(13) & "Do you want to proceed with the change?"   ' Define message.
    Style = vbYesNo + vbQuestion + vbDefaultButton2   ' Define buttons.
    Title = "Change Search Parameters?"   ' Define title.
    Response = MsgBox(Msg, Style, Title)
    If Response = vbYes Then   ' User chose Yes.
        SearchContinueReady = 0
        ContinueSearchToggle(0).Visible = False
        ContinueSearchToggle(1).Visible = False
        ProgressLabel.Caption = Empty
        TimerResult.Caption = Empty
    End If
End If
End Sub
Private Sub RestartSearchCheck()
If SearchContinueReady = 1 And SearchSaved = 0 Then
    Dim Msg, Style, Title, Help, Ctxt, Response, MyString
    Msg = "This will restart the entire search." & _
        Chr(13) & _
        "You have not saved the current search parameters." & Chr(13) & _
        Chr(13) & "Do you want to proceed with the restart?"   ' Define message.
    Style = vbYesNo + vbQuestion + vbDefaultButton2   ' Define buttons.
    Title = "Restart Search?"   ' Define title.
    Response = MsgBox(Msg, Style, Title)
    If Response = vbYes Then   ' User chose Yes.
        SearchContinueReady = 0
        ContinueSearchToggle(0).Visible = False
        ContinueSearchToggle(1).Visible = False
        ProgressLabel.Caption = Empty
        TimerResult.Caption = Empty
    End If
End If
End Sub

Public Sub OpenSearchCheck()
If SearchContinueReady = 1 And SearchSaved = 0 Then
    Dim Msg, Style, Title, Help, Ctxt, Response, MyString
    Msg = "This will load a new search file." & _
        Chr(13) & _
        "You have not saved the current search parameters." & Chr(13) & _
        Chr(13) & "Do you want to proceed with the Open command?"   ' Define message.
    Style = vbYesNo + vbQuestion + vbDefaultButton2   ' Define buttons.
    Title = "Restart Search?"   ' Define title.
    Response = MsgBox(Msg, Style, Title)
    If Response = vbYes Then   ' User chose Yes.
        SearchContinueReady = 0
        ContinueSearchToggle(0).Visible = False
        ContinueSearchToggle(1).Visible = False
    End If
End If
End Sub
