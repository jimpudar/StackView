VERSION 5.00
Begin VB.Form frmSearchSpecialTwo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Search Special Range"
   ClientHeight    =   3195
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7665
   Icon            =   "frmSearchSpecialTwo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   7665
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox SearchInputTwoMax 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   3135
      TabIndex        =   5
      Top             =   1980
      Width           =   825
   End
   Begin VB.TextBox SearchInputOneMax 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   3135
      TabIndex        =   3
      Top             =   840
      Width           =   825
   End
   Begin VB.TextBox SearchInputTwoMin 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   1170
      TabIndex        =   4
      Top             =   1980
      Width           =   825
   End
   Begin VB.TextBox SearchInputOneMin 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   1170
      TabIndex        =   2
      Top             =   840
      Width           =   825
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4395
      TabIndex        =   1
      Top             =   2610
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   3030
      TabIndex        =   0
      Top             =   2610
      Width           =   1215
   End
   Begin VB.Label Label4 
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
      Left            =   2565
      TabIndex        =   11
      Top             =   1995
      Width           =   450
   End
   Begin VB.Label Label3 
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
      Left            =   2565
      TabIndex        =   10
      Top             =   870
      Width           =   450
   End
   Begin VB.Label Label2 
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
      Left            =   600
      TabIndex        =   9
      Top             =   1995
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
      Left            =   600
      TabIndex        =   8
      Top             =   870
      Width           =   450
   End
   Begin VB.Label SearchLabelTwo 
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
      Left            =   255
      TabIndex        =   7
      Top             =   1530
      Width           =   7185
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
      TabIndex        =   6
      Top             =   390
      Width           =   7185
   End
End
Attribute VB_Name = "frmSearchSpecialTwo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CancelButton_Click()
Select Case SearchSpecialName
    Case "MoveCard"
        SearchSpecialCancel = True
        frmSearch.SearchMoveCard.Value = 0
    Case "ShiftTopBlock"
        SearchSpecialCancel = True
        frmSearch.SearchShiftTopBlock.Value = 0
    Case "ShiftTopBlockInv"
        SearchSpecialCancel = True
        frmSearch.SearchShiftTopBlockInv.Value = 0
    Case "OutFaroSpecialTop"
        SearchSpecialCancel = True
        frmSearch.SearchOutFaroSpecialTop.Value = 0
    Case "OutFaroSpecialTopInv"
        SearchSpecialCancel = True
        frmSearch.SearchOutFaroSpecialTopInv.Value = 0
    Case "OutFaroSpecialBottom"
        SearchSpecialCancel = True
        frmSearch.SearchOutFaroSpecialBottom.Value = 0
    Case "OutFaroSpecialBottomInv"
        SearchSpecialCancel = True
        frmSearch.SearchOutFaroSpecialBottomInv.Value = 0
    Case "InFaroSpecialTop"
        SearchSpecialCancel = True
        frmSearch.SearchInFaroSpecialTop.Value = 0
    Case "InFaroSpecialTopInv"
        SearchSpecialCancel = True
        frmSearch.SearchInFaroSpecialTopInv.Value = 0
    Case "InFaroSpecialBottom"
        SearchSpecialCancel = True
        frmSearch.SearchInFaroSpecialBottom.Value = 0
    Case "InFaroSpecialBottomInv"
        SearchSpecialCancel = True
        frmSearch.SearchInFaroSpecialBottomInv.Value = 0
End Select
Unload Me
End Sub

Private Sub Form_Load()
Select Case SearchSpecialName
    Case "MoveCard"
        SearchLabelOne.Caption = "Enter range of card positions from top for Move Card FROM:"
        SearchInputOneMin.Text = SearchMC1Min
        SearchInputOneMax.Text = SearchMC1Max
        SearchLabelTwo.Caption = "Enter range of card positions from top for Move Card TO:"
        SearchInputTwoMin.Text = SearchMC2Min
        SearchInputTwoMax.Text = SearchMC2Max
    Case "ShiftTopBlock"
        SearchLabelOne.Caption = "Enter range for size of BLOCK for Shift Top Block:"
        SearchInputOneMin.Text = SearchSTB1Min
        SearchInputOneMax.Text = SearchSTB1Max
        SearchLabelTwo.Caption = "Enter range of DEPTH from top for Shift Top Block:"
        SearchInputTwoMin.Text = SearchSTB2Min
        SearchInputTwoMax.Text = SearchSTB2Max
    Case "ShiftTopBlockInv"
        SearchLabelOne.Caption = "Enter range for size of BLOCK for Shift Top Block Inv:"
        SearchInputOneMin.Text = SearchSTBR1Min
        SearchInputOneMax.Text = SearchSTBR1Max
        SearchLabelTwo.Caption = "Enter range of DEPTH from top for Shift Top Block Inv:"
        SearchInputTwoMin.Text = SearchSTBR2Min
        SearchInputTwoMax.Text = SearchSTBR2Max
    Case "OutFaroSpecialTop"
        SearchLabelOne.Caption = "Enter range of cards FROM TOP for Out Faro Special Top:"
        SearchInputOneMin.Text = SearchOFST1Min
        SearchInputOneMax.Text = SearchOFST1Max
        SearchLabelTwo.Caption = "Maximum INTERIOR POSITION from top for Out Faro Special Top:"
        SearchInputTwoMin.Text = SearchOFST2Min
        SearchInputTwoMax.Text = SearchOFST2Max
    Case "OutFaroSpecialTopInv"
        SearchLabelOne.Caption = "Enter range of cards FROM TOP for Out Faro Special Top Inv:"
        SearchInputOneMin.Text = SearchOFSTR1Min
        SearchInputOneMax.Text = SearchOFSTR1Max
        SearchLabelTwo.Caption = "Range for INTERIOR POSITION from top for Out Faro Special Top Inv:"
        SearchInputTwoMin.Text = SearchOFSTR2Min
        SearchInputTwoMax.Text = SearchOFSTR2Max
    Case "OutFaroSpecialBottom"
        SearchLabelOne.Caption = "Enter range of cards FROM TOP for Out Faro Special Bottom?"
        SearchInputOneMin.Text = SearchOFSB1Min
        SearchInputOneMax.Text = SearchOFSB1Max
        SearchLabelTwo.Caption = "Range for INTERIOR POSITION from bottom for Out Faro Special Bottom:"
        SearchInputTwoMin.Text = SearchOFSB2Min
        SearchInputTwoMax.Text = SearchOFSB2Max
    Case "OutFaroSpecialBottomInv"
        SearchLabelOne.Caption = "Enter range of cards FROM TOP for Out Faro Special Bottom Inv:"
        SearchInputOneMin.Text = SearchOFSBR1Min
        SearchInputOneMax.Text = SearchOFSBR1Max
        SearchLabelTwo.Caption = "Range for INTERIOR POSITION from bottom for Out Faro Special Bottom Inv:"
        SearchInputTwoMin.Text = SearchOFSBR2Min
        SearchInputTwoMax.Text = SearchOFSBR2Max
    Case "InFaroSpecialTop"
        SearchLabelOne.Caption = "Enter range of cards FROM TOP for In Faro Special Top:"
        SearchInputOneMin.Text = SearchIFST1Min
        SearchInputOneMax.Text = SearchIFST1Max
        SearchLabelTwo.Caption = "Range for INTERIOR POSITION from top for In Faro Special Top:"
        SearchInputTwoMin.Text = SearchIFST2Min
        SearchInputTwoMax.Text = SearchIFST2Max
    Case "InFaroSpecialTopInv"
        SearchLabelOne.Caption = "Enter range of cards FROM TOP for In Faro Special Top Inv:"
        SearchInputOneMin.Text = SearchIFSTR1Min
        SearchInputOneMax.Text = SearchIFSTR1Max
        SearchLabelTwo.Caption = "Range for INTERIOR POSITION from top for In Faro Special Top Inv:"
        SearchInputTwoMin.Text = SearchIFSTR2Min
        SearchInputTwoMax.Text = SearchIFSTR2Max
    Case "InFaroSpecialBottom"
        SearchLabelOne.Caption = "Enter range of cards FROM TOP for In Faro Special Bottom:"
        SearchInputOneMin.Text = SearchIFSB1Min
        SearchInputOneMax.Text = SearchIFSB1Max
        SearchLabelTwo.Caption = "Range for INTERIOR POSITION from bottom for In Faro Special Bottom:"
        SearchInputTwoMin.Text = SearchIFSB2Min
        SearchInputTwoMax.Text = SearchIFSB2Max
    Case "InFaroSpecialBottomInv"
        SearchLabelOne.Caption = "Enter range of cards FROM TOP for In Faro Special Bottom Inv:"
        SearchInputOneMin.Text = SearchIFSBR1Min
        SearchInputOneMax.Text = SearchIFSBR1Max
        SearchLabelTwo.Caption = "Range for INTERIOR POSITION from bottom for In Faro Special Bottom Inv:"
        SearchInputTwoMin.Text = SearchIFSBR2Min
        SearchInputTwoMax.Text = SearchIFSBR2Max
End Select
End Sub

Private Sub OKButton_Click()
If SearchInputOneMin.Text <> Empty And _
    (Not IsNumeric(SearchInputOneMin.Text) Or _
    Val(SearchInputOneMin.Text) < 1 Or _
    Val(SearchInputOneMin.Text) > 52) Then
    SearchInputOneMin.Text = Empty
    MsgBox "Please enter a valid number (1 to 52)" & Chr(13) _
        & "in the first Min Input Box"
    SearchInputOneMin.SetFocus
    Exit Sub
End If
If SearchInputOneMax.Text <> Empty And _
    (Not IsNumeric(SearchInputOneMax.Text) Or _
    Val(SearchInputOneMax.Text) < 1 Or _
    Val(SearchInputOneMax.Text) > 52) Then
    SearchInputOneMax.Text = Empty
    MsgBox "Please enter a valid number (1 to 52)" & Chr(13) _
        & "in the first Max Input Box"
    SearchInputOneMax.SetFocus
    Exit Sub
End If
If SearchInputTwoMin.Text <> Empty And _
    (Not IsNumeric(SearchInputTwoMin.Text) Or _
    Val(SearchInputTwoMin.Text) < 1 Or _
    Val(SearchInputTwoMin.Text) > 52) Then
    SearchInputTwoMin.Text = Empty
    MsgBox "Please enter a valid number (1 to 52)" & Chr(13) _
        & "in the second Min Input Box"
    SearchInputTwoMin.SetFocus
    Exit Sub
End If
If SearchInputTwoMax.Text <> Empty And _
    (Not IsNumeric(SearchInputTwoMax.Text) Or _
    Val(SearchInputTwoMax.Text) < 1 Or _
    Val(SearchInputTwoMax.Text) > 52) Then
    SearchInputTwoMax.Text = Empty
    MsgBox "Please enter a valid number (1 to 52)" & Chr(13) _
        & "in the second Max Input Box"
    SearchInputTwoMax.SetFocus
    Exit Sub
End If
If Val(SearchInputOneMax.Text) < Val(SearchInputOneMin.Text) Then
    MsgBox "Min must be less than or equal to Max."
    SearchInputOneMin.Text = Empty
    SearchInputOneMax.Text = Empty
    SearchInputOneMin.SetFocus
    Exit Sub
End If
If Val(SearchInputTwoMax.Text) < Val(SearchInputTwoMin.Text) Then
    MsgBox "Min must be less than or equal to Max."
    SearchInputTwoMin.Text = Empty
    SearchInputTwoMax.Text = Empty
    SearchInputTwoMin.SetFocus
    Exit Sub
End If
If SearchSpecialName = "MatchPartial" And _
    Val(SearchInputOneMin.Text) >= Val(SearchInputOneMax.Text) Then
    SearchInputOneMin.Text = Empty
    SearchInputOneMax.Text = Empty
    MsgBox "Min must be less than or equal to Max."
    SearchInputOneMin.SetFocus
    Exit Sub
End If
Select Case SearchSpecialName
    Case "MoveCard"
        SearchMC1Min = Val(SearchInputOneMin.Text)
        SearchMC1Max = Val(SearchInputOneMax.Text)
        SearchMC2Min = Val(SearchInputTwoMin.Text)
        SearchMC2Max = Val(SearchInputTwoMax.Text)
        SearchSpecialCancel = False
    Case "ShiftTopBlock"
        SearchSTB1Min = Val(SearchInputOneMin.Text)
        SearchSTB1Max = Val(SearchInputOneMax.Text)
        SearchSTB2Min = Val(SearchInputTwoMin.Text)
        SearchSTB2Max = Val(SearchInputTwoMax.Text)
        SearchSpecialCancel = False
    Case "ShiftTopBlockInv"
        SearchSTBR1Min = Val(SearchInputOneMin.Text)
        SearchSTBR1Max = Val(SearchInputOneMax.Text)
        SearchSTBR2Min = Val(SearchInputTwoMin.Text)
        SearchSTBR2Max = Val(SearchInputTwoMax.Text)
        SearchSpecialCancel = False
    Case "OutFaroSpecialTop"
        SearchOFST1Min = Val(SearchInputOneMin.Text)
        SearchOFST1Max = Val(SearchInputOneMax.Text)
        SearchOFST2Min = Val(SearchInputTwoMin.Text)
        SearchOFST2Max = Val(SearchInputTwoMax.Text)
        SearchSpecialCancel = False
    Case "OutFaroSpecialTopInv"
        SearchOFSTR1Min = Val(SearchInputOneMin.Text)
        SearchOFSTR1Max = Val(SearchInputOneMax.Text)
        SearchOFSTR2Min = Val(SearchInputTwoMin.Text)
        SearchOFSTR2Max = Val(SearchInputTwoMax.Text)
        SearchSpecialCancel = False
    Case "OutFaroSpecialBottom"
        SearchOFSB1Min = Val(SearchInputOneMin.Text)
        SearchOFSB1Max = Val(SearchInputOneMax.Text)
        SearchOFSB2Min = Val(SearchInputTwoMin.Text)
        SearchOFSB2Max = Val(SearchInputTwoMax.Text)
        SearchSpecialCancel = False
    Case "OutFaroSpecialBottomInv"
        SearchOFSBR1Min = Val(SearchInputOneMin.Text)
        SearchOFSBR1Max = Val(SearchInputOneMax.Text)
        SearchOFSBR2Min = Val(SearchInputTwoMin.Text)
        SearchOFSBR2Max = Val(SearchInputTwoMax.Text)
        SearchSpecialCancel = False
    Case "InFaroSpecialTop"
        SearchIFST1Min = Val(SearchInputOneMin.Text)
        SearchIFST1Max = Val(SearchInputOneMax.Text)
        SearchIFST2Min = Val(SearchInputTwoMin.Text)
        SearchIFST2Max = Val(SearchInputTwoMax.Text)
        SearchSpecialCancel = False
    Case "InFaroSpecialTopInv"
        SearchIFSTR1Min = Val(SearchInputOneMin.Text)
        SearchIFSTR1Max = Val(SearchInputOneMax.Text)
        SearchIFSTR2Min = Val(SearchInputTwoMin.Text)
        SearchIFSTR2Max = Val(SearchInputTwoMax.Text)
        SearchSpecialCancel = False
    Case "InFaroSpecialBottom"
        SearchIFSB1Min = Val(SearchInputOneMin.Text)
        SearchIFSB1Max = Val(SearchInputOneMax.Text)
        SearchIFSB2Min = Val(SearchInputTwoMin.Text)
        SearchIFSB2Max = Val(SearchInputTwoMax.Text)
        SearchSpecialCancel = False
    Case "InFaroSpecialBottomInv"
        SearchIFSBR1Min = Val(SearchInputOneMin.Text)
        SearchIFSBR1Max = Val(SearchInputOneMax.Text)
        SearchIFSBR2Min = Val(SearchInputTwoMin.Text)
        SearchIFSBR2Max = Val(SearchInputTwoMax.Text)
        SearchSpecialCancel = False
End Select
Unload Me
End Sub


