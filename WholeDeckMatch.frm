VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmWholeDeckMatch 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Whole Deck Match"
   ClientHeight    =   3195
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   Icon            =   "WholeDeckMatch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   4965
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame TrapFrame 
      Height          =   1050
      Left            =   330
      TabIndex        =   9
      Top             =   1515
      Width           =   3870
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
         TabIndex        =   5
         Top             =   390
         Width           =   2580
      End
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
         TabIndex        =   4
         Top             =   135
         Value           =   -1  'True
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
         TabIndex        =   10
         Top             =   675
         Width           =   3345
      End
   End
   Begin VB.TextBox ThresholdCards 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   3120
      TabIndex        =   3
      Top             =   1260
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
      Left            =   330
      TabIndex        =   2
      Top             =   1260
      Width           =   2745
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4695
      TabIndex        =   1
      Top             =   2625
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   3255
      TabIndex        =   0
      Top             =   2625
      Width           =   1215
   End
   Begin VB.Label Label2 
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
      Left            =   210
      TabIndex        =   8
      Top             =   570
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
      Left            =   3705
      TabIndex        =   7
      Top             =   1290
      Width           =   570
   End
   Begin VB.Label Label1 
      Caption         =   "The Match range will be the whole deck (all 52 cards)."
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
      Left            =   195
      TabIndex        =   6
      Top             =   165
      Width           =   5550
   End
End
Attribute VB_Name = "frmWholeDeckMatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CancelButton_Click()
SearchSpecialCancel = True
If WholeDeckMatchSet = False Then
    frmSearch.SearchMatchPartialOption.Value = True
    ThresholdCards.Text = Empty
    ThresholdCheck.Value = 0
End If
If SuspendTrapWholeFinal Then
    TrapFileWhole = Empty
    TrapFileLabel.Caption = Empty
End If
Unload Me
End Sub

Private Sub Form_Load()
If WholeDeckMatchSet = False Then
    ThresholdCards.Text = Empty
    ThresholdCheck.Value = 0
    TrapFileLabel.Caption = Empty
    TrapFileWhole = Empty
    SuspendTrapOption.Value = True
    SaveTrapOption.Value = False
    TrapFrame.Enabled = False
    SuspendTrapOption.Enabled = False
    SaveTrapOption.Enabled = False
Else
    ThresholdCards.Text = ThresholdMatchCards
    If TrapThreshold Then
        ThresholdCheck.Value = 1
        TrapFileWhole = TrapFileFinal
        TrapFileLabel.Caption = TrapFileWhole
        SuspendTrapWhole = SuspendTrapFinal
        SuspendTrapOption.Value = SuspendTrapFinal
        SaveTrapOption.Value = Not SuspendTrapFinal
        TrapFrame.Enabled = True
        SuspendTrapOption.Enabled = True
        SaveTrapOption.Enabled = True
    Else
        ThresholdCheck.Value = 0
        TrapFileLabel.Caption = Empty
        TrapFileWhole = Empty
        SuspendTrapOption.Value = True
        SaveTrapOption.Value = False
        TrapFrame.Enabled = False
        SuspendTrapOption.Enabled = False
        SaveTrapOption.Enabled = False
    End If
End If
End Sub

Private Sub OKButton_Click()
If ThresholdCheck.Value = 1 And ThresholdCards.Text = Empty Then
    MsgBox "Please enter a valid number less than 52" _
        & Chr(13) & "in the 'cards' Input Box"
    ThresholdCards.SetFocus
    Exit Sub
End If
If ThresholdCheck.Value = 1 And ThresholdCards.Text <> Empty And _
    (Not IsNumeric(ThresholdCards.Text) Or _
    Val(ThresholdCards.Text) < 1 Or _
    Val(ThresholdCards.Text) > 51) Then
    ThresholdCards.Text = Empty
    MsgBox "Please enter a valid number less than 52" _
        & Chr(13) & "in the 'cards' Input Box"
    ThresholdCards.SetFocus
    Exit Sub
End If
SearchMatchStartCard = 1
SearchMatchEndCard = 52
SearchSpecialCancel = False
ThresholdMatchCards = Val(ThresholdCards.Text)
If ThresholdCheck.Value = 1 Then
    TrapThreshold = True
    If SuspendTrapOption Then
        SuspendTrapWholeFinal = True
        SuspendTrapFinal = True
        TrapFileWholeFinal = Empty
        TrapFileLabel.Caption = Empty
    Else
        SuspendTrapWholeFinal = False
        SuspendTrapFinal = False
        TrapFileWholeFinal = TrapFileWhole
        TrapFileFinal = TrapFileWhole
        TrapPathWholeFinal = TrapPathWhole
        TrapPathFinal = TrapPathWhole
    End If
Else
    TrapThreshold = False
End If
WholeDeckMatchSet = True
Unload Me
End Sub




Private Sub SaveTrapOption_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If SaveTrapOption.Value = True Then
    SuspendTrapWhole = False
    SaveTrapFileWhole
End If
End Sub

Private Sub SuspendTrapOption_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If SuspendTrapOption.Value Then
    SuspendTrapWhole = True
    TrapFileWhole = Empty
    TrapFileLabel.Caption = Empty
End If
End Sub

Private Sub SaveTrapFileWhole()
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
            TrapFileWhole = tFile
            TrapPathWhole = sFile
            TrapFileLabel.Visible = True
            Exit Sub
        End If
    Else
        Set trapfile = fso.CreateTextFile(sFile, True)
        trapfile.Close
        TrapFileLabel.Caption = tFile
        TrapFileWhole = tFile
        TrapPathWhole = sFile
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
    SuspendTrapWhole = True
Else
    TrapFrame.Enabled = False
    SaveTrapOption.Enabled = False
    SuspendTrapOption.Enabled = False
    TrapFileLabel.Visible = False
End If
End Sub
