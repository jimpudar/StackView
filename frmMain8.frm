VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "StackView"
   ClientHeight    =   3165
   ClientLeft      =   165
   ClientTop       =   480
   ClientWidth     =   4680
   Icon            =   "frmMain8.frx":0000
   LinkTopic       =   "MDIForm1"
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin MSComctlLib.Toolbar tbToolBar 
      Align           =   1  'Align Top
      Height          =   1080
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   1905
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   17
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Object.ToolTipText     =   "New"
            ImageKey        =   "New"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Object.ToolTipText     =   "Open"
            ImageKey        =   "Open"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Object.ToolTipText     =   "Save"
            ImageKey        =   "Save"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Print"
            Object.ToolTipText     =   "Print"
            ImageKey        =   "Print"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cut"
            Object.ToolTipText     =   "Cut"
            ImageKey        =   "Cut"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copy"
            Object.ToolTipText     =   "Copy"
            ImageKey        =   "Copy"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Paste"
            Object.ToolTipText     =   "Paste"
            ImageKey        =   "Paste"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Bold"
            Object.ToolTipText     =   "Bold"
            ImageKey        =   "Bold"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Italic"
            Object.ToolTipText     =   "Italic"
            ImageKey        =   "Italic"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Underline"
            Object.ToolTipText     =   "Underline"
            ImageKey        =   "Underline"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Align Left"
            Object.ToolTipText     =   "Align Left"
            ImageKey        =   "Align Left"
            Style           =   2
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Center"
            Object.ToolTipText     =   "Center"
            ImageKey        =   "Center"
            Style           =   2
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Align Right"
            Object.ToolTipText     =   "Align Right"
            ImageKey        =   "Align Right"
            Style           =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   2895
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   2593
            Text            =   "Status"
            TextSave        =   "Status"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "9/23/2006"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "4:15 PM"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   1740
      Top             =   1335
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   1740
      Top             =   1335
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain8.frx":1CCA
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain8.frx":1DDC
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain8.frx":1EEE
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain8.frx":2000
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain8.frx":2112
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain8.frx":2224
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain8.frx":2336
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain8.frx":2448
            Key             =   "Bold"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain8.frx":255A
            Key             =   "Italic"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain8.frx":266C
            Key             =   "Underline"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain8.frx":277E
            Key             =   "Align Left"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain8.frx":2890
            Key             =   "Center"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain8.frx":29A2
            Key             =   "Align Right"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
         Enabled         =   0   'False
         Shortcut        =   ^N
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open Deck..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "&Save Deck As..."
      End
      Begin VB.Menu mnuFileSaveAsDefault 
         Caption         =   "Save Deck As &Default"
      End
      Begin VB.Menu mnuFileBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileOpenSession 
         Caption         =   "Open Sessio&n..."
      End
      Begin VB.Menu mnuFileSaveSessionAs 
         Caption         =   "Save Session As..."
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileOpenSearch 
         Caption         =   "Open Searc&h..."
      End
      Begin VB.Menu mnuFileSaveSearchAs 
         Caption         =   "Save Search As..."
      End
      Begin VB.Menu mnuFileBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpenMnemonic 
         Caption         =   "Open Mnemonic..."
      End
      Begin VB.Menu mnuSaveMnemonic 
         Caption         =   "Save Mnemonic As..."
      End
      Begin VB.Menu mnuSaveMnemonicAsDefault 
         Caption         =   "Save Mnemonic As Default"
      End
      Begin VB.Menu mnuFileBar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "&Close"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileSaveAll 
         Caption         =   "Save A&ll"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileProperties 
         Caption         =   "Propert&ies"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileBar3 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFilePageSetup 
         Caption         =   "Page Set&up..."
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFilePrintPreview 
         Caption         =   "Print Pre&view"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print..."
      End
      Begin VB.Menu mnuFileBar4 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileSend 
         Caption         =   "Sen&d..."
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileBar5 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileBar6 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnuEditUndo 
         Caption         =   "&Undo"
      End
      Begin VB.Menu mnuEditBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cu&t"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEditPasteSpecial 
         Caption         =   "Paste &Special..."
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuDeck 
         Caption         =   "&Deck"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuControl 
         Caption         =   "&Controls"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewToolbar 
         Caption         =   "&Toolbar"
         Checked         =   -1  'True
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuViewBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewDeckFile 
         Caption         =   "Deck File"
      End
      Begin VB.Menu mnuViewSessionFile 
         Caption         =   "Session File"
      End
      Begin VB.Menu mnuViewMnemonicFile 
         Caption         =   "Mnemonic File"
      End
      Begin VB.Menu mnuViewSearchFile 
         Caption         =   "Search File"
      End
      Begin VB.Menu mnuViewTrapFile 
         Caption         =   "Threshhold Trap File"
      End
      Begin VB.Menu mnuViewRefresh 
         Caption         =   "&Refresh"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuViewOptions 
         Caption         =   "&Options..."
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuViewWebBrowser 
         Caption         =   "&Web Browser"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuCustomDeck 
         Caption         =   "&Custom Deck"
      End
      Begin VB.Menu mnuMnemonics 
         Caption         =   "Custom &Mnemonics"
      End
      Begin VB.Menu mnuShuffleMeter 
         Caption         =   "Joyal ShuffleMeter"
      End
      Begin VB.Menu mnuPiles 
         Caption         =   "&Piles Control"
      End
      Begin VB.Menu mnuRecord 
         Caption         =   "&Record"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuBackDesign 
         Caption         =   "Set &Back Design"
      End
      Begin VB.Menu mnuTest 
         Caption         =   "StackView &Test"
      End
      Begin VB.Menu mnuAdvancedTest 
         Caption         =   "StackView &Advanced Test"
      End
      Begin VB.Menu mnuSearch 
         Caption         =   "StackView &Search"
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnuWindowNewWindow 
         Caption         =   "&New Window"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuWindowCascade 
         Caption         =   "&Cascade"
      End
      Begin VB.Menu mnuWindowTileHorizontal 
         Caption         =   "Tile &Horizontal"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuWindowTileVertical 
         Caption         =   "Tile &Vertical"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuWindowArrangeIcons 
         Caption         =   "&Arrange Icons"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuOpenUserGuide 
         Caption         =   "&Open User Guide"
      End
      Begin VB.Menu mnuHelpSearchForHelpOn 
         Caption         =   "&Search For Help On..."
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About "
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
'Const EM_UNDO = &HC7
Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hwnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)
   
Private Declare Function WinHelp Lib "user32" Alias "WinHelpA" _
    (ByVal hwnd As Long, ByVal lpHelpFile As String, _
    ByVal wCommand As Long, ByVal dwData As Long) As Long

Private Const HELP_CONTENTS = 3
Private Const HELP_FINDER = 11
Private Const SW_SHOW = 1

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
    'If for any reason the ShellExecute() doesn't execute the application
    'properly, the function returns a value less than or equal to 32. Otherwise,
    'it returns a value that points to the launched application.


Private Sub MDIForm_Load()
    BackDesignCurrent = "BicycleRiderRed"
    'set this as the default back design and color
    Call frmBackDesignDialog.LoadBackDesign(BackDesignCurrent)
    frmSplash.Show
    frmSplash.Refresh
    frmSplash.ZOrder
    frmDeck.Visible = False
    frmDeck.Top = 105
    frmDeck.Left = 105
    frmStackView.Visible = False
    frmStackView.Top = 2000
    frmStackView.Left = 2000
    frmTest.Visible = False
    frmTest.Top = 1500
    frmTest.Left = 1500
    mnuTest.Checked = False
    frmCustomDeck.Visible = False
    frmCustomDeck.Top = 1600
    frmCustomDeck.Left = 1400
    mnuCustomDeck.Checked = False
    frmShuffleMeter.Visible = False
    frmShuffleMeter.Top = 1600
    frmShuffleMeter.Left = 1400
    mnuShuffleMeter.Checked = False
    frmTestAdvanced.Visible = False
    frmTestAdvanced.Top = 1500
    frmTestAdvanced.Left = 1500
    mnuAdvancedTest.Checked = False
    frmPiles.Visible = False
    frmPiles.Top = 1500
    frmPiles.Left = 1500
    mnuPiles.Checked = False
    frmMnemonic.Visible = False
    frmMnemonic.Top = 1500
    frmMnemonic.Left = 1500
    mnuMnemonics.Checked = False
    SessionParseError = False
    SessionSaved = 1
    SearchSaved = 1
    GilbreathActive = False
    frmDeck.Show
    frmStackView.Show
End Sub

Private Sub LoadSessionAllowableParameters()
'if new text based parameters are added, update the declaration line in the
'MDIForm_Load procedure
ReDim SessionAllowableParameters(SessionNumParameters)
SessionAllowableParameters(1) = Chr(34) & "Default" & Chr(34)
SessionAllowableParameters(2) = Chr(34) & "Aronson" & Chr(34)
SessionAllowableParameters(3) = Chr(34) & "Eight Kings" & Chr(34)
SessionAllowableParameters(4) = Chr(34) & "Joyal (CHaSeD)" & Chr(34)
SessionAllowableParameters(5) = Chr(34) & "Joyal (SHoCkeD)" & Chr(34)
SessionAllowableParameters(6) = Chr(34) & "New Deck (Bicycle)" & Chr(34)
SessionAllowableParameters(7) = Chr(34) & "New Deck (Fournier)" & Chr(34)
SessionAllowableParameters(8) = Chr(34) & "Nikola" & Chr(34)
SessionAllowableParameters(9) = Chr(34) & "Osterlind" & Chr(34)
SessionAllowableParameters(10) = Chr(34) & "Si Stebbins (3)" & Chr(34)
SessionAllowableParameters(11) = Chr(34) & "Si Stebbins (4)" & Chr(34)
SessionAllowableParameters(12) = Chr(34) & "Stanyon" & Chr(34)
SessionAllowableParameters(13) = Chr(34) & "Tamariz" & Chr(34)
SessionAllowableParameters(14) = "Quarter"
SessionAllowableParameters(15) = "Third"
SessionAllowableParameters(16) = "Half"
SessionAllowableParameters(17) = "Two Thirds"
SessionAllowableParameters(18) = "Three Quarters"
SessionAllowableParameters(19) = "Shallow"
SessionAllowableParameters(20) = "Deep"
SessionAllowableParameters(21) = Chr(34) & "Anywhere" & Chr(34)
SessionAllowableParameters(22) = Chr(34) & "Any Card" & Chr(34)
SessionAllowableParameters(23) = Chr(34) & "Original Position" & Chr(34)
SessionAllowableParameters(24) = "Top Third"
SessionAllowableParameters(25) = "Middle Third"
SessionAllowableParameters(26) = "Bottom Third"
SessionAllowableParameters(27) = "Forwards"
SessionAllowableParameters(28) = "Backwards"
SessionAllowableParameters(29) = "Unwind"
End Sub



Private Sub MDIForm_Unload(Cancel As Integer)
    If Me.Name = "frmDeck" Then
        mnuDeck.Checked = False
    ElseIf Me.Name = "frmStackView" Then
        mnuControl.Checked = False
    End If
    If Me.WindowState <> vbMinimized Then
        SaveSetting App.Title, "Settings", "MainLeft", Me.Left
        SaveSetting App.Title, "Settings", "MainTop", Me.Top
        SaveSetting App.Title, "Settings", "MainWidth", Me.Width
        SaveSetting App.Title, "Settings", "MainHeight", Me.Height
        SaveSetting App.Title, "Settings", "DeckPreference", _
            frmStackView.SetStackCombo.ListIndex
        SaveSetting App.Title, "Settings", "HighlightSelectionsPreference", _
            frmStackView.HighlightSelectionsCheck
        SaveSetting App.Title, "Settings", "ShowIndexPreference", _
            frmStackView.ShowIndexValues
    End If
End Sub

Private Sub mnuAdvancedTest_Click()
    mnuAdvancedTest.Checked = Not mnuAdvancedTest.Checked
    frmTestAdvanced.Visible = mnuAdvancedTest.Checked
End Sub

Private Sub mnuBackDesign_Click()
frmBackDesignDialog.Show vbModal
End Sub

Private Sub mnuFileOpenSearch_Click()
    frmSearch.OpenSearchCheck
    If SearchContinueReady = 1 Then
        Exit Sub
    End If
    frmSearch.Visible = True
    SearchFileLoading = True
    'this is used in the checkbox_click procedures to prevent the processing
    'from occuring
    SearchCounterMax = 0
    'this sets the countermax back to zero so that the manipulations checkboxes
    'can recalculate what it should be (in case the 'open search' was initiated
    'after another search was stopped.
    Dim sFile As String
    Dim tFile As String
    Dim junker As String
    On Error GoTo CancelOpenSearch:
    With dlgCommonDialog
        .DialogTitle = "Open Search"
        .CancelError = True
        .Filter = "StackView Search Files (*.svh)|*.svh"
        .ShowOpen
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
        sFile = .FileName
        tFile = .FileTitle
    End With
    If Not ((Right(tFile, 3) = "svh") Or (Right(tFile, 3) = "SVH")) Then
        MsgBox ("Invalid file type.  Search files have a .svh extension.")
        Exit Sub
    End If
    On Error GoTo SearchOpenError
    Dim fso As New FileSystemObject, searchfile As File, ts As TextStream
    Set searchfile = fso.GetFile(sFile)
    Set ts = searchfile.OpenAsTextStream(ForReading)
    
    'this section reads in the decks used in the search
    junker = ts.ReadLine        'StartDeckInitial
    For i% = 1 To 52
        StartDeckInitial(1, i%) = Val(ts.ReadLine)
        StartDeckInitial(2, i%) = ts.ReadLine
    Next i%
    junker = ts.ReadLine        'StartDeck
    For i% = 1 To 52
        StartDeck(1, i%) = Val(ts.ReadLine)
        StartDeck(2, i%) = ts.ReadLine
    Next i%
    junker = ts.ReadLine        'TargetDeck
    For i% = 1 To 52
        TargetDeck(1, i%) = Val(ts.ReadLine)
        TargetDeck(2, i%) = ts.ReadLine
    Next i%
    
    'this section reads the file names and search levels
    junker = ts.ReadLine        'StartDeckName
    StartDeckName = ts.ReadLine
    junker = ts.ReadLine        'TargetDeckName
    TargetDeckName = ts.ReadLine
    junker = ts.ReadLine        'SearchLevels
    frmSearch.SearchLevelsTextBox.Text = ts.ReadLine
    
    'this section reads in the manipulations check boxes and parameter values
    junker = ts.ReadLine        'SearchCutDeckPrecise
    SearchCDP = ts.ReadLine
    SearchCDPMin = ts.ReadLine
    SearchCDPMax = ts.ReadLine
    frmSearch.SearchCutDeckPrecise.Value = ts.ReadLine
    frmSearch.SearchCutDeckPreciseAll.Value = ts.ReadLine
    junker = ts.ReadLine        'SearchRunSingleCards
    SearchRSC = ts.ReadLine
    SearchRSCMin = ts.ReadLine
    SearchRSCMax = ts.ReadLine
    frmSearch.SearchRunSingleCards.Value = ts.ReadLine
    frmSearch.SearchRunSingleCardsAll.Value = ts.ReadLine
    junker = ts.ReadLine        'SearchRunSingleCardsInv
    SearchRSCR = ts.ReadLine
    SearchRSCRMin = ts.ReadLine
    SearchRSCRMax = ts.ReadLine
    frmSearch.SearchRunSingleCardsInv.Value = ts.ReadLine
    frmSearch.SearchRunSingleCardsInvAll.Value = ts.ReadLine
    junker = ts.ReadLine        'SearchMoveCard
    SearchMC = ts.ReadLine
    SearchMC1Min = ts.ReadLine
    SearchMC1Max = ts.ReadLine
    SearchMC2Min = ts.ReadLine
    SearchMC2Max = ts.ReadLine
    frmSearch.SearchMoveCard.Value = ts.ReadLine
    frmSearch.SearchMoveCardAll.Value = ts.ReadLine
    junker = ts.ReadLine        'SearchShiftTopBlock
    SearchSTB = ts.ReadLine
    SearchSTB1Min = ts.ReadLine
    SearchSTB1Max = ts.ReadLine
    SearchSTB2Min = ts.ReadLine
    SearchSTB2Max = ts.ReadLine
    frmSearch.SearchShiftTopBlock.Value = ts.ReadLine
    frmSearch.SearchShiftTopBlockAll.Value = ts.ReadLine
    junker = ts.ReadLine        'SearchShiftTopBlockInv
    SearchSTBR = ts.ReadLine
    SearchSTBR1Min = ts.ReadLine
    SearchSTBR1Max = ts.ReadLine
    SearchSTBR2Min = ts.ReadLine
    SearchSTBR2Max = ts.ReadLine
    frmSearch.SearchShiftTopBlockInv.Value = ts.ReadLine
    frmSearch.SearchShiftTopBlockInvAll.Value = ts.ReadLine
    junker = ts.ReadLine        'SearchCutOutFaro
    frmSearch.SearchOutFaroAll.Value = ts.ReadLine
    junker = ts.ReadLine        'SearchCutOutFaroInverse
    frmSearch.SearchOutFaroInverseAll.Value = ts.ReadLine
    junker = ts.ReadLine        'SearchOutFaroSpecialTop
    SearchOFST = ts.ReadLine
    SearchOFST1Min = ts.ReadLine
    SearchOFST1Max = ts.ReadLine
    SearchOFST2Min = ts.ReadLine
    SearchOFST2Max = ts.ReadLine
    frmSearch.SearchOutFaroSpecialTop.Value = ts.ReadLine
    frmSearch.SearchOutFaroSpecialTopAll.Value = ts.ReadLine
    junker = ts.ReadLine        'SearchOutFaroSpecialTopInv
    SearchOFSTR = ts.ReadLine
    SearchOFSTR1Min = ts.ReadLine
    SearchOFSTR1Max = ts.ReadLine
    SearchOFSTR2Min = ts.ReadLine
    SearchOFSTR2Max = ts.ReadLine
    frmSearch.SearchOutFaroSpecialTopInv.Value = ts.ReadLine
    frmSearch.SearchOutFaroSpecialTopInvAll.Value = ts.ReadLine
    junker = ts.ReadLine        'SearchOutFaroSpecialBottom
    SearchOFSB = ts.ReadLine
    SearchOFSB1Min = ts.ReadLine
    SearchOFSB1Max = ts.ReadLine
    SearchOFSB2Min = ts.ReadLine
    SearchOFSB2Max = ts.ReadLine
    frmSearch.SearchOutFaroSpecialBottom.Value = ts.ReadLine
    frmSearch.SearchOutFaroSpecialBottomAll.Value = ts.ReadLine
    junker = ts.ReadLine        'SearchOutFaroSpecialBottomInv
    SearchOFSBR = ts.ReadLine
    SearchOFSBR1Min = ts.ReadLine
    SearchOFSBR1Max = ts.ReadLine
    SearchOFSBR2Min = ts.ReadLine
    SearchOFSBR2Max = ts.ReadLine
    frmSearch.SearchOutFaroSpecialBottomInv.Value = ts.ReadLine
    frmSearch.SearchOutFaroSpecialBottomInvAll.Value = ts.ReadLine
    junker = ts.ReadLine        'SearchCutInFaro
    frmSearch.SearchInFaroAll.Value = ts.ReadLine
    junker = ts.ReadLine        'SearchCutInFaroInverse
    frmSearch.SearchInFaroInverseAll.Value = ts.ReadLine
    junker = ts.ReadLine        'SearchInFaroSpecialTop
    SearchIFST = ts.ReadLine
    SearchIFST1Min = ts.ReadLine
    SearchIFST1Max = ts.ReadLine
    SearchIFST2Min = ts.ReadLine
    SearchIFST2Max = ts.ReadLine
    frmSearch.SearchInFaroSpecialTop.Value = ts.ReadLine
    frmSearch.SearchInFaroSpecialTopAll.Value = ts.ReadLine
    junker = ts.ReadLine        'SearchInFaroSpecialTopInv
    SearchIFSTR = ts.ReadLine
    SearchIFSTR1Min = ts.ReadLine
    SearchIFSTR1Max = ts.ReadLine
    SearchIFSTR2Min = ts.ReadLine
    SearchIFSTR2Max = ts.ReadLine
    frmSearch.SearchInFaroSpecialTopInv.Value = ts.ReadLine
    frmSearch.SearchInFaroSpecialTopInvAll.Value = ts.ReadLine
    junker = ts.ReadLine        'SearchInFaroSpecialBottom
    SearchIFSB = ts.ReadLine
    SearchIFSB1Min = ts.ReadLine
    SearchIFSB1Max = ts.ReadLine
    SearchIFSB2Min = ts.ReadLine
    SearchIFSB2Max = ts.ReadLine
    frmSearch.SearchInFaroSpecialBottom.Value = ts.ReadLine
    frmSearch.SearchInFaroSpecialBottomAll.Value = ts.ReadLine
    junker = ts.ReadLine        'SearchInFaroSpecialBottomInv
    SearchIFSBR = ts.ReadLine
    SearchIFSBR1Min = ts.ReadLine
    SearchIFSBR1Max = ts.ReadLine
    SearchIFSBR2Min = ts.ReadLine
    SearchIFSBR2Max = ts.ReadLine
    frmSearch.SearchInFaroSpecialBottomInv.Value = ts.ReadLine
    frmSearch.SearchInFaroSpecialBottomInvAll.Value = ts.ReadLine
    
    'this section reads the level counters
    junker = ts.ReadLine        'SearchLevelCounter(x)
    For i% = 1 To 26
        SearchLevelCounter(i%) = ts.ReadLine
    Next i%
    
    junker = ts.ReadLine        ' "MatchFound"
    MatchFound = ts.ReadLine
    junker = ts.ReadLine        ' "NoMatchFound"
    NoMatchFound = ts.ReadLine
    junker = ts.ReadLine        ' "SearchStartDeckSet"
    SearchStartDeckSet = ts.ReadLine
    junker = ts.ReadLine        ' "SearchTargetDeckSet"
    SearchTargetDeckSet = ts.ReadLine
    junker = ts.ReadLine        ' "SearchCurrentLevel"
    SearchCurrentLevel = ts.ReadLine
    junker = ts.ReadLine        ' "SearchCurrentLevelRestart"
    SearchCurrentLevelRestart = ts.ReadLine
    junker = ts.ReadLine        ' "SearchProcessing"
    SearchProcessing = ts.ReadLine
    junker = ts.ReadLine        ' "SearchProgressCounter"
    SearchProgressCounter = ts.ReadLine
    junker = ts.ReadLine        ' "SearchElapsedTime"
    SearchElapsedTime = ts.ReadLine
    junker = ts.ReadLine        ' "SearchSpecialName"
    SearchSpecialName = ts.ReadLine
    junker = ts.ReadLine        ' "SearchSessionTransferred"
    SearchSessionTransferred = ts.ReadLine
    junker = ts.ReadLine        ' "SearchingMode"
    SearchingMode = ts.ReadLine
    junker = ts.ReadLine        ' "SearchCounter"
    SearchCounter = ts.ReadLine
    junker = ts.ReadLine        ' "SearchContinueReady"
    SearchContinueReady = ts.ReadLine
    junker = ts.ReadLine        ' "SearchCounterMax"
    SearchCounterMax = ts.ReadLine
    junker = ts.ReadLine        ' "SearchParseError"
    SearchParseError = ts.ReadLine
    junker = ts.ReadLine        ' "SearchMatchStartCard"
    SearchMatchStartCard = ts.ReadLine
    junker = ts.ReadLine        ' "SearchMatchEndCard"
    SearchMatchEndCard = ts.ReadLine
    junker = ts.ReadLine        ' "SearchMatchWholeOption"
    frmSearch.SearchMatchWholeOption.Value = ts.ReadLine
    junker = ts.ReadLine        ' "SearchMatchPartialOption"
    frmSearch.SearchMatchPartialOption.Value = ts.ReadLine
    
    junker = ts.ReadLine        ' "ThresholdMatchCards"
    ThresholdMatchCards = ts.ReadLine
    junker = ts.ReadLine        ' "TrapThreshold"
    TrapThreshold = ts.ReadLine
    junker = ts.ReadLine        ' "WholeDeckMatchSet"
    WholeDeckMatchSet = ts.ReadLine
    junker = ts.ReadLine        ' "PartialMatchFound"
    PartialMatchFound = ts.ReadLine
    junker = ts.ReadLine        ' "ContinueSearchToggle(0)"
    frmSearch.ContinueSearchToggle(0).Visible = ts.ReadLine
    junker = ts.ReadLine        ' "ContinueSearchToggle(1)"
    frmSearch.ContinueSearchToggle(1).Visible = ts.ReadLine
    junker = ts.ReadLine        ' "TimerResult"
    frmSearch.TimerResult.Caption = ts.ReadLine
    junker = ts.ReadLine        ' "ProgressLabel"
    frmSearch.ProgressLabel.Caption = ts.ReadLine
    junker = ts.ReadLine        ' "SearchTimeLabel"
    frmSearch.SearchTimeLabel.Caption = ts.ReadLine
    junker = ts.ReadLine        ' "ManipulationsLabel"
    frmSearch.ManipulationsLabel.Caption = ts.ReadLine
    junker = ts.ReadLine        ' "TrapFileWhole"
    TrapFileWhole = ts.ReadLine
    junker = ts.ReadLine        ' "TrapFileWholeFinal"
    TrapFileWholeFinal = ts.ReadLine
    junker = ts.ReadLine        ' "TrapFileFinal"
    TrapFileFinal = ts.ReadLine
    junker = ts.ReadLine        ' "TrapPathWhole"
    TrapPathWhole = ts.ReadLine
    junker = ts.ReadLine        ' "TrapPathWholeFinal"
    TrapPathWholeFinal = ts.ReadLine
    junker = ts.ReadLine        ' "TrapPathFinal"
    TrapPathFinal = ts.ReadLine
    junker = ts.ReadLine        ' "TrapFilePartial"
    TrapFilePartial = ts.ReadLine
    junker = ts.ReadLine        ' "TrapFilePartialFinal"
    TrapFilePartialFinal = ts.ReadLine
    junker = ts.ReadLine        ' "TrapPathPartial"
    TrapPathPartial = ts.ReadLine
    junker = ts.ReadLine        ' "TrapPathPartialFinal"
    TrapPathPartialFinal = ts.ReadLine
    junker = ts.ReadLine        ' "SuspendTrapWhole"
    SuspendTrapWhole = ts.ReadLine
    junker = ts.ReadLine        ' "SuspendTrapPartial"
    SuspendTrapPartial = ts.ReadLine
    junker = ts.ReadLine        ' "SuspendTrapWholeFinal"
    SuspendTrapWholeFinal = ts.ReadLine
    junker = ts.ReadLine        ' "SuspendTrapPartialFinal"
    SuspendTrapPartialFinal = ts.ReadLine
    junker = ts.ReadLine        ' "SuspendTrapFinal"
    SuspendTrapFinal = ts.ReadLine
    junker = ts.ReadLine        ' "PartialDeckMatchStart"
    frmPartialDeckMatch.SearchInputOneMin.Text = ts.ReadLine
    junker = ts.ReadLine        ' "PartialDeckMatchEnd"
    frmPartialDeckMatch.SearchInputOneMax.Text = ts.ReadLine
    junker = ts.ReadLine        ' "PartialDeckMatchThresholdCards"
    frmPartialDeckMatch.ThresholdCards.Text = ts.ReadLine
    junker = ts.ReadLine        ' "PartialDeckMatchThresholdCheck"
    frmPartialDeckMatch.ThresholdCheck.Value = ts.ReadLine
    junker = ts.ReadLine        ' "WholeDeckMatchThresholdCards"
    frmWholeDeckMatch.ThresholdCards.Text = ts.ReadLine
    junker = ts.ReadLine        ' "WholeDeckMatchThresholdCheck"
    frmWholeDeckMatch.ThresholdCheck.Value = ts.ReadLine
    'new section to identify found solution from Save Search
    If MatchFound = 1 And Not ts.AtEndOfStream Then
        junker = ts.ReadLine        ' "SolutionResults"
        frmSearch.MatchListBox.Clear
            For i% = 0 To SearchCurrentLevel - 1
                frmSearch.MatchListBox.List(i%) = ts.ReadLine
            Next i%
    End If
    
    ts.Close
    SearchFileLoading = False
'    frmSearch.ContinueSearchToggle(0).Visible = True
'    frmSearch.ContinueSearchToggle(1).Visible = False
    frmSearch.StartDeckLabel.Caption = StartDeckName
    frmSearch.StartDeckLabel.Visible = True
    frmSearch.TargetDeckLabel.Caption = TargetDeckName
    frmSearch.TargetDeckLabel.Visible = True
    frmSearch.LoadManipulations
    'new section
    If MatchFound = 1 Then
        frmSearch.MatchFoundLabel.Caption = "Match Found!!!"
        frmSearch.MatchFoundLabel.Visible = True
    End If
    If NoMatchFound = 1 Then
        frmSearch.MatchFoundLabel.Caption = "No Match Found"
        frmSearch.MatchFoundLabel.Visible = True
    End If
'    frmSearch.ShowEstimatedTime
'    frmSearch.ShowElapsedTime
'    frmSearch.ShowProgress
    'frmStackView.SessionStatusUpdate (tFile)
    Exit Sub
SearchOpenError:
SearchFileLoading = False
MsgBox ("Error reading file.  File is not recognized." & Chr(13) & _
        "The Search file must have been saved " & Chr(13) & _
        "with Version 5.0 of StackView.")
Exit Sub
CancelOpenSearch:
End Sub

Private Sub mnuFileOpenSession_Click()
    OpenSessionCheck
    If SessionSaved = 0 Then
        Exit Sub
    End If
    Dim sFile As String
    Dim tFile As String
    On Error GoTo CancelOpenSession
    With dlgCommonDialog
        .DialogTitle = "Open Session"
        .CancelError = True
        'ToDo: set the flags and attributes of the common dialog control
        .Filter = "StackView Session Files (*.svs)|*.svs"
        .ShowOpen
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
        sFile = .FileName
        tFile = .FileTitle
    End With
    If Not ((Right(tFile, 3) = "svs") Or (Right(tFile, 3) = "SVS")) Then
        MsgBox ("Invalid file type.  Session files have a .svs extension.")
        Exit Sub
    End If
    'ActiveForm.rtfText.LoadFile sFile
    'ActiveForm.Caption = sFile
    On Error GoTo SessionOpenError
    frmStackView.SessionListBox.Clear
    Dim fso As New FileSystemObject, sessionfile As File, ts As TextStream
    Set sessionfile = fso.GetFile(sFile)
    Set ts = sessionfile.OpenAsTextStream(ForReading)
    i = 0 'set file index counter to top line of SessionListBox
    Do While ts.AtEndOfStream <> True
        frmStackView.SessionListBox.List(i) = ts.ReadLine
        i = i + 1
    Loop
    ts.Close
    frmStackView.SessionStatusUpdate (tFile)
    frmStackView.StackViewDialog.Tab = 2
    Exit Sub
SessionOpenError:
MsgBox ("Error reading file.  File may be corrupt.")
Exit Sub
CancelOpenSession:
End Sub

Private Sub mnuFileSaveSearchAs_Click()
'    If SearchContinueReady = 0 Then
'        MsgBox ("There is no Search progress to save.")
'        Exit Sub
'    End If
    Dim sFile As String
    Dim tFile As String
    On Error GoTo CancelSaveSearch
    With dlgCommonDialog
        .DialogTitle = "Save Search As"
        .CancelError = True
        .Filter = "StackView Search Files (*.svh)|*.svh"
        .ShowSave
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
        sFile = .FileName
        tFile = .FileTitle
    End With
    On Error GoTo SearchSaveError
    Dim fso, searchfile
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set searchfile = fso.CreateTextFile(sFile, True)
    
    'this section writes the file contents
    
    'this section writes the decks used in search
    searchfile.WriteLine "StartDeckInitial"
    For i% = 1 To 52
        searchfile.WriteLine (StartDeckInitial(1, i%))
        searchfile.WriteLine (StartDeckInitial(2, i%))
    Next i%
    searchfile.WriteLine "StartDeck"
    For i% = 1 To 52
        searchfile.WriteLine (StartDeck(1, i%))
        searchfile.WriteLine (StartDeck(2, i%))
    Next i%
    searchfile.WriteLine "TargetDeck"
    For i% = 1 To 52
        searchfile.WriteLine (TargetDeck(1, i%))
        searchfile.WriteLine (TargetDeck(2, i%))
    Next i%
    
    'this section writes the file names and search levels
    searchfile.WriteLine "StartDeckName"
    searchfile.WriteLine StartDeckName
    searchfile.WriteLine "TargetDeckName"
    searchfile.WriteLine TargetDeckName
    searchfile.WriteLine "SearchLevels"
    searchfile.WriteLine frmSearch.SearchLevelsTextBox.Text
    
    
    'this section sets the manipulations checkboxes and parameter values
    searchfile.WriteLine "CutDeckPrecise"
    searchfile.WriteLine SearchCDP
    searchfile.WriteLine SearchCDPMin
    searchfile.WriteLine SearchCDPMax
    searchfile.WriteLine frmSearch.SearchCutDeckPrecise.Value
    searchfile.WriteLine frmSearch.SearchCutDeckPreciseAll.Value
    searchfile.WriteLine "RunSingleCards"
    searchfile.WriteLine SearchRSC
    searchfile.WriteLine SearchRSCMin
    searchfile.WriteLine SearchRSCMax
    searchfile.WriteLine frmSearch.SearchRunSingleCards.Value
    searchfile.WriteLine frmSearch.SearchRunSingleCardsAll.Value
    searchfile.WriteLine "RunSingleCardsInv"
    searchfile.WriteLine SearchRSCR
    searchfile.WriteLine SearchRSCRMin
    searchfile.WriteLine SearchRSCRMax
    searchfile.WriteLine frmSearch.SearchRunSingleCardsInv.Value
    searchfile.WriteLine frmSearch.SearchRunSingleCardsInvAll.Value
    searchfile.WriteLine "SearchMoveCard"
    searchfile.WriteLine SearchMC
    searchfile.WriteLine SearchMC1Min
    searchfile.WriteLine SearchMC1Max
    searchfile.WriteLine SearchMC2Min
    searchfile.WriteLine SearchMC2Max
    searchfile.WriteLine frmSearch.SearchMoveCard.Value
    searchfile.WriteLine frmSearch.SearchMoveCardAll.Value
    searchfile.WriteLine "SearchShiftTopBlock"
    searchfile.WriteLine SearchSTB
    searchfile.WriteLine SearchSTB1Min
    searchfile.WriteLine SearchSTB1Max
    searchfile.WriteLine SearchSTB2Min
    searchfile.WriteLine SearchSTB2Max
    searchfile.WriteLine frmSearch.SearchShiftTopBlock.Value
    searchfile.WriteLine frmSearch.SearchShiftTopBlockAll.Value
    searchfile.WriteLine "SearchShiftTopBlockInv"
    searchfile.WriteLine SearchSTBR
    searchfile.WriteLine SearchSTBR1Min
    searchfile.WriteLine SearchSTBR1Max
    searchfile.WriteLine SearchSTBR2Min
    searchfile.WriteLine SearchSTBR2Max
    searchfile.WriteLine frmSearch.SearchShiftTopBlockInv.Value
    searchfile.WriteLine frmSearch.SearchShiftTopBlockInvAll.Value
    searchfile.WriteLine "SearchOutFaro"
    searchfile.WriteLine frmSearch.SearchOutFaroAll.Value
    searchfile.WriteLine "SearchOutFaroInverse"
    searchfile.WriteLine frmSearch.SearchOutFaroInverseAll.Value
    searchfile.WriteLine "SearchOutFaroSpecialTop"
    searchfile.WriteLine SearchOFST
    searchfile.WriteLine SearchOFST1Min
    searchfile.WriteLine SearchOFST1Max
    searchfile.WriteLine SearchOFST2Min
    searchfile.WriteLine SearchOFST2Max
    searchfile.WriteLine frmSearch.SearchOutFaroSpecialTop.Value
    searchfile.WriteLine frmSearch.SearchOutFaroSpecialTopAll.Value
    searchfile.WriteLine "SearchOutFaroSpecialTopInv"
    searchfile.WriteLine SearchOFSTR
    searchfile.WriteLine SearchOFSTR1Min
    searchfile.WriteLine SearchOFSTR1Max
    searchfile.WriteLine SearchOFSTR2Min
    searchfile.WriteLine SearchOFSTR2Max
    searchfile.WriteLine frmSearch.SearchOutFaroSpecialTopInv.Value
    searchfile.WriteLine frmSearch.SearchOutFaroSpecialTopInvAll.Value
    searchfile.WriteLine "SearchOutFaroSpecialBottom"
    searchfile.WriteLine SearchOFSB
    searchfile.WriteLine SearchOFSB1Min
    searchfile.WriteLine SearchOFSB1Max
    searchfile.WriteLine SearchOFSB2Min
    searchfile.WriteLine SearchOFSB2Max
    searchfile.WriteLine frmSearch.SearchOutFaroSpecialBottom.Value
    searchfile.WriteLine frmSearch.SearchOutFaroSpecialBottomAll.Value
    searchfile.WriteLine "SearchOutFaroSpecialBottomInv"
    searchfile.WriteLine SearchOFSBR
    searchfile.WriteLine SearchOFSBR1Min
    searchfile.WriteLine SearchOFSBR1Max
    searchfile.WriteLine SearchOFSBR2Min
    searchfile.WriteLine SearchOFSBR2Max
    searchfile.WriteLine frmSearch.SearchOutFaroSpecialBottomInv.Value
    searchfile.WriteLine frmSearch.SearchOutFaroSpecialBottomInvAll.Value
    searchfile.WriteLine "SearchInFaro"
    searchfile.WriteLine frmSearch.SearchInFaroAll.Value
    searchfile.WriteLine "SearchInFaroInverse"
    searchfile.WriteLine frmSearch.SearchInFaroInverseAll.Value
    searchfile.WriteLine "SearchInFaroSpecialTop"
    searchfile.WriteLine SearchIFST
    searchfile.WriteLine SearchIFST1Min
    searchfile.WriteLine SearchIFST1Max
    searchfile.WriteLine SearchIFST2Min
    searchfile.WriteLine SearchIFST2Max
    searchfile.WriteLine frmSearch.SearchInFaroSpecialTop.Value
    searchfile.WriteLine frmSearch.SearchInFaroSpecialTopAll.Value
    searchfile.WriteLine "SearchInFaroSpecialTopInv"
    searchfile.WriteLine SearchIFSTR
    searchfile.WriteLine SearchIFSTR1Min
    searchfile.WriteLine SearchIFSTR1Max
    searchfile.WriteLine SearchIFSTR2Min
    searchfile.WriteLine SearchIFSTR2Max
    searchfile.WriteLine frmSearch.SearchInFaroSpecialTopInv.Value
    searchfile.WriteLine frmSearch.SearchInFaroSpecialTopInvAll.Value
    searchfile.WriteLine "SearchInFaroSpecialBottom"
    searchfile.WriteLine SearchIFSB
    searchfile.WriteLine SearchIFSB1Min
    searchfile.WriteLine SearchIFSB1Max
    searchfile.WriteLine SearchIFSB2Min
    searchfile.WriteLine SearchIFSB2Max
    searchfile.WriteLine frmSearch.SearchInFaroSpecialBottom.Value
    searchfile.WriteLine frmSearch.SearchInFaroSpecialBottomAll.Value
    searchfile.WriteLine "SearchInFaroSpecialBottomInv"
    searchfile.WriteLine SearchIFSBR
    searchfile.WriteLine SearchIFSBR1Min
    searchfile.WriteLine SearchIFSBR1Max
    searchfile.WriteLine SearchIFSBR2Min
    searchfile.WriteLine SearchIFSBR2Max
    searchfile.WriteLine frmSearch.SearchInFaroSpecialBottomInv.Value
    searchfile.WriteLine frmSearch.SearchInFaroSpecialBottomInvAll.Value
    
    'this section writes the level counters
    searchfile.WriteLine "SearchLevelCounter(x)"
    For i% = 1 To 26
        searchfile.WriteLine SearchLevelCounter(i%)
    Next i%
    
    'this section writes miscellaneous search variables
    searchfile.WriteLine "MatchFound"
    searchfile.WriteLine MatchFound
    searchfile.WriteLine "NoMatchFound"
    searchfile.WriteLine NoMatchFound
    searchfile.WriteLine "SearchStartDeckSet"
    searchfile.WriteLine SearchStartDeckSet
    searchfile.WriteLine "SearchTargetDeckSet"
    searchfile.WriteLine SearchTargetDeckSet
    searchfile.WriteLine "SearchCurrentLevel"
    searchfile.WriteLine SearchCurrentLevel
    searchfile.WriteLine "SearchCurrentLevelRestart"
    searchfile.WriteLine SearchCurrentLevelRestart
    searchfile.WriteLine "SearchProcessing"
    searchfile.WriteLine SearchProcessing
    searchfile.WriteLine "SearchProgressCounter"
    searchfile.WriteLine SearchProgressCounter
    searchfile.WriteLine "SearchElapsedTime"
    searchfile.WriteLine SearchElapsedTime
    searchfile.WriteLine "SearchSpecialName"
    searchfile.WriteLine SearchSpecialName
    searchfile.WriteLine "SearchSessionTransferred"
    searchfile.WriteLine SearchSessionTransferred
    searchfile.WriteLine "SearchingMode"
    searchfile.WriteLine SearchingMode
    searchfile.WriteLine "SearchCounter"
    searchfile.WriteLine SearchCounter
    searchfile.WriteLine "SearchContinueReady"
    searchfile.WriteLine SearchContinueReady
    searchfile.WriteLine "SearchCounterMax"
    searchfile.WriteLine SearchCounterMax
    searchfile.WriteLine "SearchParseError"
    searchfile.WriteLine SearchParseError
    searchfile.WriteLine "SearchMatchStartCard"
    searchfile.WriteLine SearchMatchStartCard
    searchfile.WriteLine "SearchMatchEndCard"
    searchfile.WriteLine SearchMatchEndCard
    searchfile.WriteLine "WholeDeckMatch"
    searchfile.WriteLine frmSearch.SearchMatchWholeOption.Value
    searchfile.WriteLine "PartialDeckMatch"
    searchfile.WriteLine frmSearch.SearchMatchPartialOption.Value
    
    searchfile.WriteLine "ThresholdMatchCards"
    searchfile.WriteLine ThresholdMatchCards
    searchfile.WriteLine "TrapThreshold"
    searchfile.WriteLine TrapThreshold
    searchfile.WriteLine "WholeDeckMatchSet"
    searchfile.WriteLine WholeDeckMatchSet
    searchfile.WriteLine "PartialMatchFound"
    searchfile.WriteLine PartialMatchFound
    searchfile.WriteLine "ContinueSearchToggle(0)"
    searchfile.WriteLine frmSearch.ContinueSearchToggle(0).Visible
    searchfile.WriteLine "ContinueSearchToggle(1)"
    searchfile.WriteLine frmSearch.ContinueSearchToggle(1).Visible
    searchfile.WriteLine "TimerResult"
    searchfile.WriteLine frmSearch.TimerResult.Caption
    searchfile.WriteLine "ProgressLabel"
    searchfile.WriteLine frmSearch.ProgressLabel.Caption
    searchfile.WriteLine "SearchTimeLabel"
    searchfile.WriteLine frmSearch.SearchTimeLabel.Caption
    searchfile.WriteLine "ManipulationsLabel"
    searchfile.WriteLine frmSearch.ManipulationsLabel.Caption
    searchfile.WriteLine "TrapFileWhole"
    searchfile.WriteLine TrapFileWhole
    searchfile.WriteLine "TrapFileWholeFinal"
    searchfile.WriteLine TrapFileWholeFinal
    searchfile.WriteLine "TrapFileFinal"
    searchfile.WriteLine TrapFileFinal
    searchfile.WriteLine "TrapPathWhole"
    searchfile.WriteLine TrapPathWhole
    searchfile.WriteLine "TrapPathWholeFinal"
    searchfile.WriteLine TrapPathWholeFinal
    searchfile.WriteLine "TrapPathFinal"
    searchfile.WriteLine TrapPathFinal
    searchfile.WriteLine "TrapFilePartial"
    searchfile.WriteLine TrapFilePartial
    searchfile.WriteLine "TrapFilePartialFinal"
    searchfile.WriteLine TrapFilePartialFinal
    searchfile.WriteLine "TrapPathPartial"
    searchfile.WriteLine TrapPathPartial
    searchfile.WriteLine "TrapPathPartialFinal"
    searchfile.WriteLine TrapPathPartialFinal
    searchfile.WriteLine "SuspendTrapWhole"
    searchfile.WriteLine SuspendTrapWhole
    searchfile.WriteLine "SuspendTrapPartial"
    searchfile.WriteLine SuspendTrapPartial
    searchfile.WriteLine "SuspendTrapWholeFinal"
    searchfile.WriteLine SuspendTrapWholeFinal
    searchfile.WriteLine "SuspendTrapPartialFinal"
    searchfile.WriteLine SuspendTrapPartialFinal
    searchfile.WriteLine "SuspendTrapFinal"
    searchfile.WriteLine SuspendTrapFinal
    searchfile.WriteLine "PartialDeckMatchStart"
    searchfile.WriteLine frmPartialDeckMatch.SearchInputOneMin.Text
    searchfile.WriteLine "PartialDeckMatchEnd"
    searchfile.WriteLine frmPartialDeckMatch.SearchInputOneMax.Text
    searchfile.WriteLine "PartialDeckMatchThresholdCards"
    searchfile.WriteLine frmPartialDeckMatch.ThresholdCards.Text
    searchfile.WriteLine "PartialDeckMatchThresholdCheck"
    searchfile.WriteLine frmPartialDeckMatch.ThresholdCheck.Value
    searchfile.WriteLine "WholeDeckMatchThresholdCards"
    searchfile.WriteLine frmWholeDeckMatch.ThresholdCards.Text
    searchfile.WriteLine "WholeDeckMatchThresholdCheck"
    searchfile.WriteLine frmWholeDeckMatch.ThresholdCheck.Value
    'new section to identify found solution for Open Search
    If MatchFound = 1 And _
        frmSearch.MatchListBox.ListCount = SearchCurrentLevel Then
        searchfile.WriteLine "SolutionResults"
        For i% = 0 To SearchCurrentLevel - 1
            searchfile.WriteLine frmSearch.MatchListBox.List(i%)
        Next i%
    End If
        
    searchfile.Close
    'frmStackView.SessionStatusUpdate (tFile)
    'should have a similar status text for search
    SearchSaved = 1
    Exit Sub
SearchSaveError:
SearchSaved = 0
MsgBox ("Error saving file.  File may be corrupt.")
Exit Sub
CancelSaveSearch:
End Sub

Private Sub mnuFileSaveSessionAs_Click()
    If frmStackView.SessionListBox.ListCount = 0 Then
        MsgBox ("There is no Session to save.")
        Exit Sub
    End If
    Dim sFile As String
    Dim tFile As String
    On Error GoTo CancelSaveSession:
    With dlgCommonDialog
        .DialogTitle = "Save Session As"
        .CancelError = True
        .Filter = "StackView Session Files (*.svs)|*.svs"
        .ShowSave
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
        sFile = .FileName
        tFile = .FileTitle
    End With
    'ActiveForm.Caption = sFile
    'ActiveForm.rtfText.SaveFile sFile
    On Error GoTo SessionSaveError
    Dim fso, sessionfile
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set sessionfile = fso.CreateTextFile(sFile, True)
    For i% = 0 To frmStackView.SessionListBox.ListCount - 1
        sessionfile.WriteLine frmStackView.SessionListBox.List(i%)
    Next i%
    sessionfile.Close
    frmStackView.SessionStatusUpdate (tFile)
    Exit Sub
SessionSaveError:
MsgBox ("Error saving file.  File may be corrupt.")
Exit Sub
CancelSaveSession:
End Sub


Private Sub mnuMnemonics_Click()
    mnuMnemonics.Checked = Not mnuMnemonics.Checked
    frmMnemonic.Visible = mnuMnemonics.Checked
End Sub

Private Sub mnuOpenMnemonic_Click()
    Dim sFile As String
    On Error GoTo CancelMnemonicOpen
    With dlgCommonDialog
        .DialogTitle = "Open Mnemonics"
        .CancelError = True
        'ToDo: set the flags and attributes of the common dialog control
        .Filter = "StackView Mnemonic Files (*.svm)|*.svm"
        .ShowOpen
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
        sFile = .FileName
        tFile = .FileTitle
    End With
    'ActiveForm.rtfText.LoadFile sFile
    'ActiveForm.Caption = sFile
    If Not ((Right(tFile, 3) = "svm") Or (Right(tFile, 3) = "SVM")) Then
        MsgBox ("Invalid file type.  Mnemonic files have a .svm extension.")
        Exit Sub
    End If
    On Error GoTo MnemonicOpenError
    Dim fso As New FileSystemObject, mnemonicfile As File, ts As TextStream
    Set mnemonicfile = fso.GetFile(sFile)
    Set ts = mnemonicfile.OpenAsTextStream(ForReading)
    For i% = 1 To 52
        MnemonicCards(i%) = ts.ReadLine
    Next i%
    For i% = 1 To 52
        MnemonicPositions(i%) = ts.ReadLine
    Next i%
    ts.Close
    Call frmMnemonic.LoadMnemonicTable
    'show the form
    mnuMnemonics.Checked = True
    frmMnemonic.Visible = True
    MnemonicSaved = 1
    Exit Sub
MnemonicOpenError:
MsgBox ("Error opening Deck file.")
Exit Sub
CancelMnemonicOpen:
End Sub

Private Sub mnuOpenUserGuide_Click()
ShellCode = ShellExecute(Me.hwnd, "open", App.Path & "\StackView User Guide.pdf", "", "", 1)
If ShellCode <= 32 Then
    MsgBox ("There was an error in opening the User Guide." & Chr(13) & _
        "Either the 'StackView User Guide.pdf' file is missing" & Chr(13) & _
        "from the StackView application directory, or Adobe Acrobat Reader" & Chr(13) & _
        "is not installed on your computer.  Visit www.adobe.com" & Chr(13) & Chr(13) & _
        "You may download the Users Guide from www.stackview.com")
End If
End Sub

Private Sub mnuPiles_Click()
    mnuPiles.Checked = Not mnuPiles.Checked
    frmPiles.Visible = mnuPiles.Checked
End Sub

Private Sub mnuSaveMnemonic_Click()
    Dim sFile As String
    On Error GoTo CancelSaveMnemonic
    With dlgCommonDialog
        .DialogTitle = "Save Mnemonic As"
        .CancelError = True
        .Filter = "StackView Mnemonic Files (*.svm)|*.svm"
        .ShowSave
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
        sFile = .FileName
    End With
    'ActiveForm.Caption = sFile
    'ActiveForm.rtfText.SaveFile sFile
    On Error GoTo MnemonicSaveError
    Dim fso, mnemonicfile
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set mnemonicfile = fso.CreateTextFile(sFile, True)
    For i% = 1 To 52
        mnemonicfile.WriteLine (MnemonicCards(i%))
    Next i%
    For i% = 1 To 52
        mnemonicfile.WriteLine (MnemonicPositions(i%))
    Next i%
    mnemonicfile.Close
    MnemonicSaved = 1
    Exit Sub
MnemonicSaveError:
MsgBox ("Error saving Mnemonic file.")
Exit Sub
CancelSaveMnemonic:
End Sub

Private Sub mnuSaveMnemonicAsDefault_Click()
Dim Msg, Style, Title, Help, Ctxt, Response, MyString
Msg = "This will save the current mnemonics as the Default." & Chr(13) & _
    "Do you want to continue?"   ' Define message.
Style = vbYesNo + vbQuestion + vbDefaultButton2   ' Define buttons.
Title = "Save As Default"   ' Define title.
'Help = "DEMO.HLP"   ' Define Help file.
'Ctxt = 1000   ' Define topic
      ' context.
      ' Display message.
Response = MsgBox(Msg, Style, Title) ', Help, Ctxt)
If Response = vbYes Then   ' User chose Yes.
    On Error GoTo DefaultMnemonicSaveError
    Dim fso, mnemonicfile
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set mnemonicfile = fso.CreateTextFile(App.Path & "\stackview.svm", True)
    For i% = 1 To 52
        mnemonicfile.WriteLine (MnemonicCards(i%))
    Next i%
    For i% = 1 To 52
        mnemonicfile.WriteLine (MnemonicPositions(i%))
    Next i%
    mnemonicfile.Close
End If
MnemonicSaved = 1
Exit Sub
DefaultMnemonicSaveError:
MsgBox ("Error saving Default Mnemonic file.")
End Sub

'Public Sub mnuRecord_Click()
'    If SessionRecordMode Then
'        frmStackView.SessionRecordLabel.Caption = "Start Recording"
'        'change the button label
'    Else
'        frmStackView.SessionRecordLabel.Caption = "Stop Recording"
'        'change the button label
'    End If
'    mnuRecord.Checked = Not mnuRecord.Checked
'    'toggle the check mark
'    SessionRecordMode = Not SessionRecordMode
'    'toggle the logical value for the button
'End Sub

Private Sub mnuSearch_Click()
    mnuSearch.Checked = Not mnuSearch.Checked
    frmSearch.Visible = mnuSearch.Checked
End Sub
Private Sub mnuCustomDeck_Click()
    mnuCustomDeck.Checked = Not mnuCustomDeck.Checked
    frmCustomDeck.Visible = mnuCustomDeck.Checked
End Sub

Private Sub mnuShuffleMeter_Click()
    mnuShuffleMeter.Checked = Not mnuShuffleMeter.Checked
    frmShuffleMeter.Visible = mnuShuffleMeter.Checked
End Sub

Private Sub mnuTest_Click()
    mnuTest.Checked = Not mnuTest.Checked
    frmTest.Visible = mnuTest.Checked
End Sub

Private Sub mnuViewDeckFile_Click()
    Dim sFile As String
    Dim tFile As String
    Wrap$ = Chr$(13) + Chr$(10)  'create wrap character
    On Error GoTo CancelViewDeck
    With dlgCommonDialog
        .DialogTitle = "Open Deck File"
        .CancelError = True
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
    'On Error GoTo DeckOpenError
    On Error GoTo TooBig:    'set error handler
    Dim fso As New FileSystemObject, deckfile As File, ts As TextStream
    Set deckfile = fso.GetFile(sFile)
    Set ts = deckfile.OpenAsTextStream(ForReading)
    Do While ts.AtEndOfStream <> True       'then read lines from file
        LineOfText$ = ts.ReadLine
        AllText$ = AllText$ & LineOfText$ & Wrap$
    Loop
    ts.Close
    Dim frmD As frmViewFile
    Set frmD = New frmViewFile
    frmD.Caption = tFile
    frmD.FileTextBox = AllText$
    frmD.Show
    'frmViewFile.Caption = tFile
    'frmViewFile.FileTextBox.Text = AllText$  'display file
Exit Sub
TooBig:             'error handler displays message
MsgBox ("The specified file is too large.")
Exit Sub
CancelViewDeck:
End Sub

Private Sub mnuViewMnemonicFile_Click()
    Dim sFile As String
    Dim tFile As String
    Wrap$ = Chr$(13) + Chr$(10)  'create wrap character
    On Error GoTo CancelViewMnemonic
    With dlgCommonDialog
        .DialogTitle = "Open Mnemonic File"
        .CancelError = True
        'ToDo: set the flags and attributes of the common dialog control
        .Filter = "StackView Mnemonic Files (*.svm)|*.svm"
        .ShowOpen
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
        sFile = .FileName
        tFile = .FileTitle
    End With
    If Not ((Right(tFile, 3) = "svm") Or (Right(tFile, 3) = "SVM")) Then
        MsgBox ("Invalid file type.  Mnemonic files have a .svm extension.")
        Exit Sub
    End If
    'On Error GoTo DeckOpenError
    On Error GoTo MnemonicTooBig:    'set error handler
    Dim fso As New FileSystemObject, mnemonicfile As File, ts As TextStream
    Set mnemonicfile = fso.GetFile(sFile)
    Set ts = mnemonicfile.OpenAsTextStream(ForReading)
    Do While ts.AtEndOfStream <> True       'then read lines from file
        LineOfText$ = ts.ReadLine
        AllText$ = AllText$ & LineOfText$ & Wrap$
    Loop
    ts.Close
    Dim frmD As frmViewFile
    Set frmD = New frmViewFile
    frmD.Caption = tFile
    frmD.FileTextBox = AllText$
    frmD.Show
    'frmViewFile.Caption = tFile
    'frmViewFile.FileTextBox.Text = AllText$  'display file
Exit Sub
MnemonicTooBig:             'error handler displays message
MsgBox ("The specified file is too large.")
Exit Sub
CancelViewMnemonic:
End Sub

Private Sub mnuViewSearchFile_Click()
    Dim sFile As String
    Dim tFile As String
    Wrap$ = Chr$(13) + Chr$(10)  'create wrap character
    On Error GoTo CancelViewSearch
    With dlgCommonDialog
        .DialogTitle = "Open Search File"
        .CancelError = True
        'ToDo: set the flags and attributes of the common dialog control
        .Filter = "StackView Search Files (*.svh)|*.svh"
        .ShowOpen
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
        sFile = .FileName
        tFile = .FileTitle
    End With
    If Not ((Right(tFile, 3) = "svh") Or (Right(tFile, 3) = "SVH")) Then
        MsgBox ("Invalid file type.  Search files have a .svh extension.")
        Exit Sub
    End If
    'On Error GoTo DeckOpenError
    On Error GoTo TooBig:    'set error handler
    Dim fso As New FileSystemObject, deckfile As File, ts As TextStream
    Set deckfile = fso.GetFile(sFile)
    Set ts = deckfile.OpenAsTextStream(ForReading)
    Do While ts.AtEndOfStream <> True       'then read lines from file
        LineOfText$ = ts.ReadLine
        AllText$ = AllText$ & LineOfText$ & Wrap$
    Loop
    ts.Close
    Dim frmD As frmViewFile
    Set frmD = New frmViewFile
    frmD.Caption = tFile
    frmD.FileTextBox = AllText$
    frmD.Show
    'frmViewFile.Caption = tFile
    'frmViewFile.FileTextBox.Text = AllText$  'display file
Exit Sub
TooBig:             'error handler displays message
MsgBox ("The specified file is too large.")
Exit Sub
CancelViewSearch:
End Sub

Private Sub mnuViewSessionFile_Click()
    Dim sFile As String
    Dim tFile As String
    Wrap$ = Chr$(13) + Chr$(10)  'create wrap character
    On Error GoTo CancelViewSession
    With dlgCommonDialog
        .DialogTitle = "Open Session File"
        .CancelError = True
        'ToDo: set the flags and attributes of the common dialog control
        .Filter = "StackView Session Files (*.svs)|*.svs"
        .ShowOpen
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
        sFile = .FileName
        tFile = .FileTitle
    End With
    If Not ((Right(tFile, 3) = "svs") Or (Right(tFile, 3) = "SVS")) Then
        MsgBox ("Invalid file type.  Session files have a .svs extension.")
        Exit Sub
    End If
    'On Error GoTo DeckOpenError
    On Error GoTo TooBig:    'set error handler
    Dim fso As New FileSystemObject, deckfile As File, ts As TextStream
    Set deckfile = fso.GetFile(sFile)
    Set ts = deckfile.OpenAsTextStream(ForReading)
    Do While ts.AtEndOfStream <> True       'then read lines from file
        LineOfText$ = ts.ReadLine
        AllText$ = AllText$ & LineOfText$ & Wrap$
    Loop
    ts.Close
    Dim frmD As frmViewFile
    Set frmD = New frmViewFile
    frmD.Caption = tFile
    frmD.FileTextBox = AllText$
    frmD.Show
    'frmViewFile.Caption = tFile
    'frmViewFile.FileTextBox.Text = AllText$  'display file
Exit Sub
TooBig:             'error handler displays message
MsgBox ("The specified file is too large.")
Exit Sub
CancelViewSession:
End Sub

Private Sub mnuViewTrapFile_Click()
    Dim sFile As String
    Dim tFile As String
    Wrap$ = Chr$(13) + Chr$(10)  'create wrap character
    On Error GoTo CancelViewTrap
    With dlgCommonDialog
        .DialogTitle = "Open Threshold Trap File"
        .CancelError = True
        'ToDo: set the flags and attributes of the common dialog control
        .Filter = "StackView Threshold Trap Files (*.svt)|*.svt"
        .ShowOpen
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
        sFile = .FileName
        tFile = .FileTitle
    End With
    If Not ((Right(tFile, 3) = "svt") Or (Right(tFile, 3) = "SVT")) Then
        MsgBox ("Invalid file type.  Threshold Trap files have a .svt extension.")
        Exit Sub
    End If
    'On Error GoTo DeckOpenError
    On Error GoTo TooBig:    'set error handler
    Dim fso As New FileSystemObject, deckfile As File, ts As TextStream
    Set deckfile = fso.GetFile(sFile)
    Set ts = deckfile.OpenAsTextStream(ForReading)
    Do While ts.AtEndOfStream <> True       'then read lines from file
        LineOfText$ = ts.ReadLine
        AllText$ = AllText$ & LineOfText$ & Wrap$
    Loop
    ts.Close
    Dim frmD As frmViewFile
    Set frmD = New frmViewFile
    frmD.Caption = tFile
    frmD.FileTextBox = AllText$
    frmD.Show
    'frmViewFile.Caption = tFile
    'frmViewFile.FileTextBox.Text = AllText$  'display file
Exit Sub
TooBig:             'error handler displays message
MsgBox ("The specified file is too large.")
Exit Sub
CancelViewTrap:
End Sub

'Private Sub tbToolBar_ButtonClick(ByVal Button As MSComCtlLib.Button)
'    On Error Resume Next
'    Select Case Button.Key
'        Case "New"
'            LoadNewDoc
'        Case "Open"
'            mnuFileOpen_Click
'        Case "Save"
'            mnuFileSave_Click
'        Case "Print"
'            mnuFilePrint_Click
'        Case "Cut"
'            mnuEditCut_Click
'        Case "Copy"
'            mnuEditCopy_Click
'        Case "Paste"
'            mnuEditPaste_Click
'        Case "Bold"
'            ActiveForm.rtfText.SelBold = Not ActiveForm.rtfText.SelBold
'            Button.Value = IIf(ActiveForm.rtfText.SelBold, tbrPressed, tbrUnpressed)
'        Case "Italic"
'            ActiveForm.rtfText.SelItalic = Not ActiveForm.rtfText.SelItalic
'            Button.Value = IIf(ActiveForm.rtfText.SelItalic, tbrPressed, tbrUnpressed)
'        Case "Underline"
'            ActiveForm.rtfText.SelUnderline = Not ActiveForm.rtfText.SelUnderline
'            Button.Value = IIf(ActiveForm.rtfText.SelUnderline, tbrPressed, tbrUnpressed)
'        Case "Align Left"
'            ActiveForm.rtfText.SelAlignment = rtfLeft
'        Case "Center"
'            ActiveForm.rtfText.SelAlignment = rtfCenter
'        Case "Align Right"
'            ActiveForm.rtfText.SelAlignment = rtfRight
'    End Select
'End Sub

Private Sub mnuHelpAbout_Click()
    frmAbout.Show vbModal, Me
End Sub

'Private Sub mnuHelpSearchForHelpOn_Click()
'    Dim nRet As Integer'


    'if there is no helpfile for this project display a message to the user
    'you can set the HelpFile for your application in the
    'Project Properties dialog
'    If Len(App.HelpFile) = 0 Then
'        MsgBox "Unable to display Help Contents. There is no Help associated with this project.", vbInformation, Me.Caption
'    Else
'        On Error Resume Next
'        nRet = OSWinHelp(Me.hwnd, App.HelpFile, 261, 0)
'        If Err Then
'            MsgBox Err.Description
'        End If
'    End If

'End Sub

Private Sub mnuHelpContents_Click()
    Dim nRet As Integer


    'if there is no helpfile for this project display a message to the user
    'you can set the HelpFile for your application in the
    'Project Properties dialog
    If Len(App.HelpFile) = 0 Then
        MsgBox "Unable to display Help Contents. There is no Help associated with this project.", vbInformation, Me.Caption
    Else
        On Error Resume Next
      Dim lResult As Long
      Dim sHelpFile As String
      Dim lCommand As Long, lOption As Long
      sHelpFile = App.HelpFile
      lCommand = HELP_FINDER
      lOption = 0
      lResult = WinHelp(Me.hwnd, sHelpFile, lCommand, lOption)
        
        'nRet = OSWinHelp(Me.hwnd, App.HelpFile, 3, 0)
        If Err Then
            MsgBox Err.Description
        End If
    End If

End Sub


'Private Sub mnuWindowArrangeIcons_Click()
'    Me.Arrange vbArrangeIcons
'End Sub

'Private Sub mnuWindowTileVertical_Click()
'    Me.Arrange vbTileVertical
'End Sub
'
'Private Sub mnuWindowTileHorizontal_Click()
'    Me.Arrange vbTileHorizontal
'End Sub
'
Private Sub mnuWindowCascade_Click()
    Me.Arrange vbCascade
End Sub

'Private Sub mnuWindowNewWindow_Click()
'    LoadNewDoc
'End Sub

'Private Sub mnuToolsOptions_Click()
'    frmOptions.Show
'End Sub

'Private Sub mnuViewWebBrowser_Click()
'    'ToDo: Add 'mnuViewWebBrowser_Click' code.
'    MsgBox "Add 'mnuViewWebBrowser_Click' code."
'End Sub

'Private Sub mnuViewOptions_Click()
'    frmOptions.Show vbModal, Me
'End Sub

'Private Sub mnuViewRefresh_Click()
'    'ToDo: Add 'mnuViewRefresh_Click' code.
'    MsgBox "Add 'mnuViewRefresh_Click' code."
'End Sub

'Private Sub mnuViewStatusBar_Click()
'    mnuViewStatusBar.Checked = Not mnuViewStatusBar.Checked
'    sbStatusBar.Visible = mnuViewStatusBar.Checked
'End Sub

'Private Sub mnuViewToolbar_Click()
'    mnuViewToolbar.Checked = Not mnuViewToolbar.Checked
'    tbToolBar.Visible = mnuViewToolbar.Checked
'End Sub

Private Sub mnuDeck_Click()
    mnuDeck.Checked = Not mnuDeck.Checked
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
    'Call frmDeck.DisplayCards
    frmDeck.Visible = mnuDeck.Checked
End Sub

Private Sub mnuControl_Click()
    mnuControl.Checked = Not mnuControl.Checked
    frmStackView.Visible = mnuControl.Checked
End Sub

'Private Sub mnuEditPasteSpecial_Click()
'    'ToDo: Add 'mnuEditPasteSpecial_Click' code.
'    MsgBox "Add 'mnuEditPasteSpecial_Click' code."
'End Sub

'Private Sub mnuEditPaste_Click()'
'    On Error Resume Next
'    ActiveForm.rtfText.SelRTF = Clipboard.GetText
'
'End Sub

'Private Sub mnuEditCopy_Click()
'    On Error Resume Next
'    Clipboard.SetText ActiveForm.rtfText.SelRTF
'
'End Sub

'Private Sub mnuEditCut_Click()
'    On Error Resume Next
'    Clipboard.SetText ActiveForm.rtfText.SelRTF
'    ActiveForm.rtfText.SelText = vbNullString
'
'End Sub

'Private Sub mnuEditUndo_Click()
'    'ToDo: Add 'mnuEditUndo_Click' code.
'    MsgBox "Add 'mnuEditUndo_Click' code."
'End Sub


Private Sub mnuFileExit_Click()
    'unload the form
    Unload Me

End Sub

Private Sub mnuFileSend_Click()
    'ToDo: Add 'mnuFileSend_Click' code.
    MsgBox "Add 'mnuFileSend_Click' code."
End Sub

Private Sub mnuFilePrint_Click()
    On Error Resume Next
    If ActiveForm Is Nothing Then Exit Sub
    With dlgCommonDialog
        .DialogTitle = "Print"
        .CancelError = True
        .Flags = cdlPDReturnDC + cdlPDNoPageNums + _
            cdlPDAllPages + cdlPDNoSelection + cdlPDDisablePrintToFile
        'If ActiveForm.rtfText.SelLength = 0 Then
        '    .Flags = .Flags + cdlPDAllPages
        'Else
        '    .Flags = .Flags + cdlPDSelection
        'End If
        .ShowPrinter
        If Err <> MSComDlg.cdlCancel Then
            'ActiveForm.rtfText.SelPrint .hDC
            ActiveForm.PrintForm
        End If
    End With

End Sub

'Private Sub mnuFilePrintPreview_Click()
'    'ToDo: Add 'mnuFilePrintPreview_Click' code.
'    MsgBox "Add 'mnuFilePrintPreview_Click' code."
'End Sub

'Private Sub mnuFilePageSetup_Click()
'    On Error Resume Next
'    With dlgCommonDialog
'        .DialogTitle = "Page Setup"
'        .CancelError = True
'        .ShowPrinter
'    End With
'
'End Sub

'Private Sub mnuFileProperties_Click()
'    'ToDo: Add 'mnuFileProperties_Click' code.
'    MsgBox "Add 'mnuFileProperties_Click' code."
'End Sub

'Private Sub mnuFileSaveAll_Click()
'    'ToDo: Add 'mnuFileSaveAll_Click' code.
'    MsgBox "Add 'mnuFileSaveAll_Click' code."
'End Sub

Private Sub mnuFileSaveAs_Click()
    Dim sFile As String
    On Error GoTo CancelSaveDeck
    With dlgCommonDialog
        .DialogTitle = "Save Deck As"
        .CancelError = True
        .Filter = "StackView Deck Files (*.svf)|*.svf"
        .ShowSave
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
        sFile = .FileName
    End With
    'ActiveForm.Caption = sFile
    'ActiveForm.rtfText.SaveFile sFile
    On Error GoTo DeckSaveError
    Dim fso, deckfile
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set deckfile = fso.CreateTextFile(sFile, True)
    For i% = 1 To 52
        deckfile.WriteLine (Deck(1, i%))
        deckfile.WriteLine (Deck(2, i%))
    Next i%
    deckfile.WriteLine ("BackDesign=" & BackDesignCurrent)
    deckfile.Close
    Exit Sub
DeckSaveError:
MsgBox ("Error saving Deck file.")
Exit Sub
CancelSaveDeck:
End Sub

Private Sub mnuFileSave_Click()
    Dim sFile As String
    If Left$(ActiveForm.Caption, 8) = "Document" Then
        With dlgCommonDialog
            .DialogTitle = "Save"
            .CancelError = False
            'ToDo: set the flags and attributes of the common dialog control
            .Filter = "All Files (*.*)|*.*"
            .ShowSave
            If Len(.FileName) = 0 Then
                Exit Sub
            End If
            sFile = .FileName
        End With
        ActiveForm.rtfText.SaveFile sFile
    Else
        sFile = ActiveForm.Caption
        ActiveForm.rtfText.SaveFile sFile
    End If

End Sub
Private Sub mnuFileSaveAsDefault_Click()
Dim Msg, Style, Title, Help, Ctxt, Response, MyString
Msg = "This will save the current deck as the Default." & Chr(13) & _
    "Do you want to continue?"   ' Define message.
Style = vbYesNo + vbQuestion + vbDefaultButton2   ' Define buttons.
Title = "Save As Default"   ' Define title.
'Help = "DEMO.HLP"   ' Define Help file.
'Ctxt = 1000   ' Define topic
      ' context.
      ' Display message.
Response = MsgBox(Msg, Style, Title) ', Help, Ctxt)
If Response = vbYes Then   ' User chose Yes.
    On Error GoTo DefaultSaveError
    Dim fso, deckfile
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set deckfile = fso.CreateTextFile(App.Path & "\stackview.svf", True)
    For i% = 1 To 52
        deckfile.WriteLine (Deck(1, i%))
        deckfile.WriteLine (Deck(2, i%))
    Next i%
    deckfile.WriteLine ("BackDesign=" & BackDesignCurrent)
    deckfile.Close
End If
Exit Sub
DefaultSaveError:
MsgBox ("Error saving Default Deck file.")
End Sub

Private Sub mnuFileClose_Click()
    'ToDo: Add 'mnuFileClose_Click' code.
    MsgBox "Add 'mnuFileClose_Click' code."
End Sub

Public Sub SessionInsertMacro()
    Dim sFile As String
    Dim tFile As String
    On Error GoTo MacroOpenCancel
    'get the file name
    With dlgCommonDialog
        .DialogTitle = "Insert Macro"
        .CancelError = True
        'ToDo: set the flags and attributes of the common dialog control
        .Filter = "StackView Session Files (*.svs)|*.svs"
        .ShowOpen
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
        sFile = .FileName
        tFile = .FileTitle
    End With
    'check for the right extension
    If Not ((Right(tFile, 3) = "svs") Or (Right(tFile, 3) = "SVS")) Then
        InsertMacroError = True
        MsgBox ("Invalid file type.  Session files have a .svs extension.")
        Exit Sub
    End If
    'check that the file is in the App directory
    If Left(sFile, Len(sFile) - Len(tFile) - 1) <> App.Path Then
        InsertMacroError = True
        MsgBox ("The Macro Session file must be in the " & Chr(13) & _
                "same directory as Stackview, which is:" & Chr(13) & Chr(13) & _
                App.Path)
        Exit Sub
    End If
    On Error GoTo MacroOpenError
    'insert the Macro statement in the list
    SessionCommand = "Macro(" & tFile & ")"
    frmStackView.SessionListBox.AddItem SessionCommand
    Call frmStackView.SessionStatusUpdate(0)
    Exit Sub
MacroOpenError:
InsertMacroError = True
MsgBox ("Error reading file.  File may be corrupt." & Chr(13) & _
        "Also, the Macro Session file must be in the " & Chr(13) & _
        "same directory as Stackview, which is:" & Chr(13) & Chr(13) & _
        App.Path)
MacroOpenCancel:
End Sub


Private Sub mnuFileOpen_Click()
    Dim sFile As String
    On Error GoTo CancelOpen
    With dlgCommonDialog
        .DialogTitle = "Open Deck"
        .CancelError = True
        'ToDo: set the flags and attributes of the common dialog control
        .Filter = "StackView Deck Files (*.svf)|*.svf"
        .ShowOpen
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
        sFile = .FileName
        tFile = .FileTitle
    End With
    'ActiveForm.rtfText.LoadFile sFile
    'ActiveForm.Caption = sFile
    If Not ((Right(tFile, 3) = "svf") Or (Right(tFile, 3) = "SVF")) Then
        MsgBox ("Invalid file type.  Deck files have a .svf extension.")
        Exit Sub
    End If
    On Error GoTo DeckOpenError
    Dim fso As New FileSystemObject, deckfile As File, ts As TextStream
    Set deckfile = fso.GetFile(sFile)
    Set ts = deckfile.OpenAsTextStream(ForReading)
    For i% = 1 To 52
        Deck(1, i%) = Val(ts.ReadLine)
        Deck(2, i%) = ts.ReadLine
        Deck(6, i%) = False
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
    Call frmStackView.ClearSelections_Click
    Call frmStackView.ShowCards
    frmDeck.Show
    Exit Sub
DeckOpenError:
MsgBox ("Error opening Deck file.")
Exit Sub
CancelOpen:
End Sub

'Private Sub mnuFileNew_Click()
'    LoadNewDoc
'End Sub

'Public Sub UnCheckDeck()
'    mnuDeck.Checked = False
'End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim Msg   ' Declare variable.
    ' Set the message text.
    Msg = ""
    If SessionSaved = 0 And SearchSaved = 0 And MnemonicSaved = 0 Then
        Msg = Msg & "You have unsaved Session, Search, and Mnemonic activities still present." & Chr(13)
    ElseIf SessionSaved = 0 And SearchSaved = 0 And MnemonicSaved = 1 Then
        Msg = Msg & "You have unsaved Session and Search activities still present." & Chr(13)
    ElseIf SessionSaved = 0 And SearchSaved = 1 And MnemonicSaved = 0 Then
        Msg = Msg & "You have unsaved Session and Mnemonic activities still present." & Chr(13)
    ElseIf SessionSaved = 0 And SearchSaved = 1 And MnemonicSaved = 1 Then
        Msg = Msg & "You have unsaved Session activity still present." & Chr(13)
    ElseIf SessionSaved = 1 And SearchSaved = 0 And MnemonicSaved = 0 Then
        Msg = Msg & "You have unsaved Search and Mnemonic activities still present." & Chr(13)
    ElseIf SessionSaved = 1 And SearchSaved = 0 And MnemonicSaved = 1 Then
        Msg = Msg & "You have unsaved Search activity still present." & Chr(13)
    ElseIf SessionSaved = 1 And SearchSaved = 1 And MnemonicSaved = 0 Then
        Msg = Msg & "You have unsaved Mnemonic activity still present." & Chr(13)
    End If
    Msg = Msg & Chr(13) & "Do you really want to exit Stackview?" & Chr(13)
    ' If user clicks the No button, stop QueryUnload.
    If MsgBox(Msg, vbQuestion + vbYesNo, Me.Caption) = vbNo Then
        Cancel = True
    Else
        SessionSaved = 1
        SearchSaved = 1
        MnemonicSaved = 1
        'End
    End If
    Dim frm As Form
    For Each frm In Forms
      If frm.Name <> Me.Name Then ' Unload this form LAST
        Unload frm
        Set frm = Nothing
      End If
    Next
    Unload Me
End Sub

Private Sub OpenSessionCheck()
If SessionSaved = 0 Then
    Dim Msg, Style, Title, Help, Ctxt, Response, MyString
    Msg = "This will load a new session file." & _
        Chr(13) & _
        "You have not saved the current session events." & Chr(13) & _
        Chr(13) & "Do you want to proceed with the Open command?"   ' Define message.
    Style = vbYesNo + vbQuestion + vbDefaultButton2   ' Define buttons.
    Title = "Restart Search?"   ' Define title.
    Response = MsgBox(Msg, Style, Title)
    If Response = vbYes Then   ' User chose Yes.
        SessionSaved = 1
    End If
End If
End Sub
