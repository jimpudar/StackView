VERSION 5.00
Begin VB.Form frmSearchSpecialOne 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Search Special Range"
   ClientHeight    =   2160
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7665
   Icon            =   "frmSearchSpecialOne.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   7665
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox SearchInputOneMax 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   2835
      TabIndex        =   3
      Top             =   840
      Width           =   825
   End
   Begin VB.TextBox SearchInputOneMin 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   840
      TabIndex        =   2
      Top             =   840
      Width           =   825
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4380
      TabIndex        =   1
      Top             =   1620
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   3015
      TabIndex        =   0
      Top             =   1620
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Max"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2265
      TabIndex        =   6
      Top             =   870
      Width           =   450
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Min"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   300
      TabIndex        =   5
      Top             =   870
      Width           =   450
   End
   Begin VB.Label SearchLabelOne 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   240
      TabIndex        =   4
      Top             =   390
      Width           =   7185
   End
End
Attribute VB_Name = "frmSearchSpecialOne"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CancelButton_Click()
Select Case SearchSpecialName
    Case "CutDeckPrecise"
        SearchSpecialCancel = True
        frmSearch.SearchCutDeckPrecise.Value = 0
    Case "RunSingleCards"
        SearchSpecialCancel = True
        frmSearch.SearchRunSingleCards.Value = 0
    Case "RunSingleCardsInv"
        SearchSpecialCancel = True
        frmSearch.SearchRunSingleCardsInv.Value = 0
    Case "MatchPartial"
        SearchSpecialCancel = True
        frmSearch.SearchMatchWholeOption.Value = True
End Select
Unload Me
End Sub

Private Sub Form_Load()
Select Case SearchSpecialName
    Case "CutDeckPrecise"
        SearchLabelOne.Caption = "Enter Search range parameters for Cut Deck Precise:"
        SearchInputOneMin.Text = SearchCDPMin
        SearchInputOneMax.Text = SearchCDPMax
    Case "RunSingleCards"
        SearchLabelOne.Caption = "Enter Search range parameters for Run Single Cards:"
        SearchInputOneMin.Text = SearchRSCMin
        SearchInputOneMax.Text = SearchRSCMax
    Case "RunSingleCardsInv"
        SearchLabelOne.Caption = "Enter Search range parameters for Run Single Cards Inverse:"
        SearchInputOneMin.Text = SearchRSCRMin
        SearchInputOneMax.Text = SearchRSCRMax
    Case "MatchPartial"
        SearchLabelOne.Caption = "Enter desired range of card positions for Partial deck Match:"
        SearchInputOneMin.Text = SearchMatchStartCard
        SearchInputOneMax.Text = SearchMatchEndCard
End Select
End Sub

Private Sub OKButton_Click()
If SearchInputOneMin.Text <> Empty And _
    (Not IsNumeric(SearchInputOneMin.Text) Or _
    Val(SearchInputOneMin.Text) < 1 Or _
    Val(SearchInputOneMin.Text) > 52) Then
    SearchInputOneMin.Text = Empty
    MsgBox "Please enter a valid number (1 to 52)" & Chr(13) _
        & "in the Input Box"
    SearchInputOneMin.SetFocus
    Exit Sub
End If
If SearchInputOneMax.Text <> Empty And _
    (Not IsNumeric(SearchInputOneMax.Text) Or _
    Val(SearchInputOneMax.Text) < 1 Or _
    Val(SearchInputOneMax.Text) > 52) Then
    SearchInputOneMax.Text = Empty
    MsgBox "Please enter a valid number (1 to 52)" & Chr(13) _
        & "in the Input Box"
    SearchInputOneMax.SetFocus
    Exit Sub
End If
If Val(SearchInputOneMax.Text) < Val(SearchInputOneMin.Text) Then
    MsgBox "Min must be less than or equal to Max."
    SearchInputOneMin.Text = Empty
    SearchInputOneMax.Text = Empty
    SearchInputOneMin.SetFocus
    Exit Sub
End If
Select Case SearchSpecialName
    Case "CutDeckPrecise"
        SearchCDPMin = Val(SearchInputOneMin.Text)
        SearchCDPMax = Val(SearchInputOneMax.Text)
        SearchSpecialCancel = False
    Case "RunSingleCards"
        SearchRSCMin = Val(SearchInputOneMin.Text)
        SearchRSCMax = Val(SearchInputOneMax.Text)
        SearchSpecialCancel = False
    Case "RunSingleCardsInv"
        SearchRSCRMin = Val(SearchInputOneMin.Text)
        SearchRSCRMax = Val(SearchInputOneMax.Text)
        SearchSpecialCancel = False
    Case "MatchPartial"
        SearchMatchStartCard = Val(SearchInputOneMin.Text)
        SearchMatchEndCard = Val(SearchInputOneMax.Text)
        SearchSpecialCancel = False
End Select
Unload Me
End Sub
