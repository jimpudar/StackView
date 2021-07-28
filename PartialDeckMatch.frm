VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmPartialDeckMatch 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Partial Deck Match"
   ClientHeight    =   3420
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   Icon            =   "PartialDeckMatch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.Frame TrapFrame 
      Height          =   1050
      Left            =   315
      TabIndex        =   13
      Top             =   1770
      Width           =   3870
      Begin VB.OptionButton SuspendTrapOption 
         Caption         =   "Suspend Search on Trap"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   120
         TabIndex        =   6
         Top             =   135
         Value           =   -1  'True
         Width           =   2580
      End
      Begin VB.OptionButton SaveTrapOption 
         Caption         =   "Save Traps to File"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   120
         TabIndex        =   7
         Top             =   390
         Width           =   2580
      End
      Begin VB.Label TrapFileLabel 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   270
         Left            =   405
         TabIndex        =   14
         Top             =   675
         Width           =   3345
      End
   End
   Begin VB.TextBox ThresholdCards 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   3195
      TabIndex        =   5
      Top             =   1440
      Width           =   510
   End
   Begin VB.CheckBox ThresholdCheck 
      Caption         =   "Trap threshold matches of"
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
      Left            =   405
      TabIndex        =   4
      Top             =   1440
      Width           =   2745
   End
   Begin VB.TextBox SearchInputOneMax 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   3225
      TabIndex        =   3
      Top             =   420
      Width           =   825
   End
   Begin VB.TextBox SearchInputOneMin 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   1020
      TabIndex        =   2
      Top             =   420
      Width           =   825
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4695
      TabIndex        =   1
      Top             =   2925
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   3255
      TabIndex        =   0
      Top             =   2925
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   5385
      Top             =   1590
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label5 
      Caption         =   "Click the checkbox below if you would like to have Stackview trap the search at a specified threshhold match amount."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   255
      TabIndex        =   12
      Top             =   855
      Width           =   5550
   End
   Begin VB.Label Label4 
      Caption         =   "cards"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3780
      TabIndex        =   11
      Top             =   1470
      Width           =   570
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "End"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2535
      TabIndex        =   10
      Top             =   435
      Width           =   570
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   360
      TabIndex        =   9
      Top             =   435
      Width           =   570
   End
   Begin VB.Label Label1 
      Caption         =   "Enter the Partial Deck Match range for the search:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   225
      TabIndex        =   8
      Top             =   90
      Width           =   5550
   End
End
Attribute VB_Name = "frmPartialDeckMatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CancelButton_Click()
SearchSpecialCancel = True
If WholeDeckMatchSet = True Then
    frmSearch.SearchMatchWholeOption.Value = True
    SearchMatchStartCard = 1
    SearchMatchEndCard = 52
    ThresholdCards.Text = Empty
    ThresholdCheck.Value = 0
End If
If SuspendTrapPartialFinal Then
    TrapFilePartial = Empty
    TrapFileLabel.Caption = Empty
End If
Unload Me
End Sub

Private Sub Form_Load()
If WholeDeckMatchSet = True Then
    SearchInputOneMin.Text = Empty
    SearchInputOneMax.Text = Empty
    ThresholdCards.Text = Empty
    ThresholdCheck.Value = 0
    TrapFileLabel.Caption = Empty
    TrapFilePartial = Empty
    SuspendTrapOption.Value = True
    SaveTrapOption.Value = False
    TrapFrame.Enabled = False
    SuspendTrapOption.Enabled = False
    SaveTrapOption.Enabled = False
Else
    SearchInputOneMin.Text = SearchMatchStartCard
    SearchInputOneMax.Text = SearchMatchEndCard
    ThresholdCards.Text = ThresholdMatchCards
    If TrapThreshold Then
        ThresholdCheck.Value = 1
        TrapFilePartial = TrapFilePartialFinal
        TrapFileLabel.Caption = TrapFilePartial
        SuspendTrapPartial = SuspendTrapPartialFinal
        SuspendTrapOption.Value = SuspendTrapPartial
        SaveTrapOption.Value = Not SuspendTrapPartial
        TrapFrame.Enabled = True
        SuspendTrapOption.Enabled = True
        SaveTrapOption.Enabled = True
    Else
        ThresholdCheck.Value = 0
        TrapFileLabel.Caption = Empty
        TrapFilePartial = Empty
        SuspendTrapOption.Value = True
        SaveTrapOption.Value = False
        TrapFrame.Enabled = False
        SuspendTrapOption.Enabled = False
        SaveTrapOption.Enabled = False
    End If
End If
End Sub

Private Sub OKButton_Click()
If SearchInputOneMin.Text = Empty Then
    MsgBox "Please enter a valid number (1 to 52)" & Chr(13) _
        & "in the 'Start' Input Box"
    SearchInputOneMin.SetFocus
    Exit Sub
End If
If SearchInputOneMin.Text <> Empty And _
    (Not IsNumeric(SearchInputOneMin.Text) Or _
    Val(SearchInputOneMin.Text) < 1 Or _
    Val(SearchInputOneMin.Text) > 52) Then
    SearchInputOneMin.Text = Empty
    MsgBox "Please enter a valid number (1 to 52)" & Chr(13) _
        & "in the 'Start' Input Box"
    SearchInputOneMin.SetFocus
    Exit Sub
End If
If SearchInputOneMax.Text = Empty Then
    MsgBox "Please enter a valid number (1 to 52)" & Chr(13) _
        & "in the 'End' Input Box"
    SearchInputOneMax.SetFocus
    Exit Sub
End If
If SearchInputOneMax.Text <> Empty And _
    (Not IsNumeric(SearchInputOneMax.Text) Or _
    Val(SearchInputOneMax.Text) < 1 Or _
    Val(SearchInputOneMax.Text) > 52) Then
    SearchInputOneMax.Text = Empty
    MsgBox "Please enter a valid number (1 to 52)" & Chr(13) _
        & "in the 'End' Input Box"
    SearchInputOneMax.SetFocus
    Exit Sub
End If
If Val(SearchInputOneMax.Text) < Val(SearchInputOneMin.Text) Then
    MsgBox "The Start value must be less than or equal to the End value."
    SearchInputOneMin.Text = Empty
    SearchInputOneMax.Text = Empty
    SearchInputOneMin.SetFocus
    Exit Sub
End If
If ThresholdCheck.Value = 1 And ThresholdCards.Text = Empty Then
    MsgBox "Please enter a valid number less than " _
        & (Val(SearchInputOneMax.Text) - Val(SearchInputOneMin.Text) + 1) & Chr(13) _
        & "in the 'cards' Input Box"
    ThresholdCards.SetFocus
    Exit Sub
End If
If ThresholdCheck.Value = 1 And ThresholdCards.Text <> Empty And _
    (Not IsNumeric(ThresholdCards.Text) Or _
    Val(ThresholdCards.Text) < 1 Or _
    Val(ThresholdCards.Text) > _
        (Val(SearchInputOneMax.Text) - Val(SearchInputOneMin.Text))) Then
    ThresholdCards.Text = Empty
    MsgBox "Please enter a valid number less than " _
        & (Val(SearchInputOneMax.Text) - Val(SearchInputOneMin.Text) + 1) & Chr(13) _
        & "in the 'cards' Input Box"
    ThresholdCards.SetFocus
    Exit Sub
End If
SearchMatchStartCard = Val(SearchInputOneMin.Text)
SearchMatchEndCard = Val(SearchInputOneMax.Text)
SearchSpecialCancel = False
ThresholdMatchCards = Val(ThresholdCards.Text)
If ThresholdCheck.Value = 1 Then
    TrapThreshold = True
    If SuspendTrapOption Then
        SuspendTrapPartialFinal = True
        SuspendTrapFinal = True
        TrapFilePartialFinal = Empty
        TrapFileLabel.Caption = Empty
    Else
        SuspendTrapPartialFinal = False
        SuspendTrapFinal = False
        TrapFilePartialFinal = TrapFilePartial
        TrapFileFinal = TrapFilePartial
        TrapPathPartialFinal = TrapPathPartial
        TrapPathFinal = TrapPathPartial
    End If
Else
    TrapThreshold = False
End If
WholeDeckMatchSet = False
Unload Me
End Sub


Private Sub SaveTrapOption_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If SaveTrapOption.Value = True Then
    SuspendTrapPartial = False
    SaveTrapFilePartial
End If
End Sub

Private Sub SuspendTrapOption_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If SuspendTrapOption.Value Then
    SuspendTrapPartial = True
    TrapFilePartial = Empty
    TrapFileLabel.Caption = Empty
End If
End Sub

Private Sub SaveTrapFilePartial()
    Dim sFile As String
    Dim tFile As String
    dlgCommonDialog.CancelError = True
    On Error GoTo TrapCancelError
    dlgCommonDialog.DialogTitle = "Create Trap File As"
    dlgCommonDialog.Filter = "StackView Trap Files (*.svt)|*.svt"
    dlgCommonDialog.ShowSave
    If Len(dlgCommonDialog.FileName) = 0 Then
        Exit Sub
    End If
    sFile = dlgCommonDialog.FileName
    tFile = dlgCommonDialog.FileTitle
    'ActiveForm.Caption = sFile
    'ActiveForm.rtfText.SaveFile sFile
    On Error GoTo TrapSaveError
    Dim fso, trapfile
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(sFile) Then
        Dim Msg   ' Declare variable.
        ' Set the message text.
        Msg = "The file " & tFile & " already exists." & Chr(13) & _
            "Do you want to overwrite the file?"
        ' If user clicks the No button, stop QueryUnload.
        If MsgBox(Msg, vbQuestion + vbYesNo, Me.Caption) = vbNo Then
            Cancel = True
            Exit Sub
        Else
            Set trapfile = fso.CreateTextFile(sFile, True)
            trapfile.Close
            TrapFileLabel.Caption = tFile
            TrapFilePartial = tFile
            TrapPathPartial = sFile
            TrapFileLabel.Visible = True
            Exit Sub
        End If
    Else
        Set trapfile = fso.CreateTextFile(sFile, True)
        trapfile.Close
        TrapFileLabel.Caption = tFile
        TrapFilePartial = tFile
        TrapPathPartial = sFile
        TrapFileLabel.Visible = True
        Exit Sub
    End If
TrapCancelError:
If TrapFileLabel.Caption = Empty Then
    SuspendTrapOption = True
End If
Exit Sub
TrapSaveError:
MsgBox ("Error creating Trap file.")
End Sub


Private Sub ThresholdCheck_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If ThresholdCheck.Value = 0 Then
    ThresholdCards.Text = Empty
End If
If ThresholdCheck.Value = 1 Then
    TrapFrame.Enabled = True
    SaveTrapOption.Enabled = True
    SuspendTrapOption.Enabled = True
    TrapFileLabel.Visible = True
    SuspendTrapOption.Value = True
    SuspendTrapPartial = True
Else
    TrapFrame.Enabled = False
    SaveTrapOption.Enabled = False
    SuspendTrapOption.Enabled = False
    TrapFileLabel.Visible = False
End If
End Sub
