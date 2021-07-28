VERSION 5.00
Begin VB.Form frmBackDesignDialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Back Design"
   ClientHeight    =   4470
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5925
   Icon            =   "frmBackDesignDialog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   5925
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton BackDesignBlueOption 
      Caption         =   "Blue"
      Height          =   315
      Left            =   4140
      TabIndex        =   5
      Top             =   600
      Width           =   840
   End
   Begin VB.OptionButton BackDesignRedOption 
      Caption         =   "Red"
      Height          =   315
      Left            =   4140
      TabIndex        =   4
      Top             =   285
      Value           =   -1  'True
      Width           =   840
   End
   Begin VB.ComboBox BackDesignCombo 
      Height          =   315
      ItemData        =   "frmBackDesignDialog.frx":1CCA
      Left            =   180
      List            =   "frmBackDesignDialog.frx":1CF5
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   585
      Width           =   2580
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4560
      TabIndex        =   1
      Top             =   3870
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   3255
      TabIndex        =   0
      Top             =   3870
      Width           =   1215
   End
   Begin VB.Image TallyHoFanRed 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   3870
      Picture         =   "frmBackDesignDialog.frx":1DBA
      Tag             =   "BackDesign"
      Top             =   1155
      Width           =   1350
   End
   Begin VB.Image AladdinBlue 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   3870
      Picture         =   "frmBackDesignDialog.frx":6578
      Tag             =   "BackDesign"
      Top             =   1155
      Width           =   1350
   End
   Begin VB.Image AladdinBlueSelected 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   3870
      Picture         =   "frmBackDesignDialog.frx":B056
      Tag             =   "BackDesign"
      Top             =   1155
      Width           =   1350
   End
   Begin VB.Image AladdinRed 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   3870
      Picture         =   "frmBackDesignDialog.frx":E869
      Tag             =   "BackDesign"
      Top             =   1155
      Width           =   1350
   End
   Begin VB.Image AladdinRedSelected 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   3870
      Picture         =   "frmBackDesignDialog.frx":13227
      Tag             =   "BackDesign"
      Top             =   1155
      Width           =   1350
   End
   Begin VB.Image AviatorBlue 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   3870
      Picture         =   "frmBackDesignDialog.frx":168FF
      Tag             =   "BackDesign"
      Top             =   1155
      Width           =   1350
   End
   Begin VB.Image AviatorBlueSelected 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   3870
      Picture         =   "frmBackDesignDialog.frx":1AAB9
      Tag             =   "BackDesign"
      Top             =   1155
      Width           =   1350
   End
   Begin VB.Image AviatorRed 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   3870
      Picture         =   "frmBackDesignDialog.frx":1DB87
      Tag             =   "BackDesign"
      Top             =   1155
      Width           =   1350
   End
   Begin VB.Image AviatorRedSelected 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   3870
      Picture         =   "frmBackDesignDialog.frx":21D6A
      Tag             =   "BackDesign"
      Top             =   1155
      Width           =   1350
   End
   Begin VB.Image BicycleAutoBlue 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   3870
      Picture         =   "frmBackDesignDialog.frx":24E65
      Tag             =   "BackDesign"
      Top             =   1155
      Width           =   1350
   End
   Begin VB.Image BicycleAutoBlueSelected 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   3870
      Picture         =   "frmBackDesignDialog.frx":293DF
      Tag             =   "BackDesign"
      Top             =   1155
      Width           =   1350
   End
   Begin VB.Image BicycleAutoRed 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   3870
      Picture         =   "frmBackDesignDialog.frx":2C748
      Tag             =   "BackDesign"
      Top             =   1155
      Width           =   1350
   End
   Begin VB.Image BicycleAutoRedSelected 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   3870
      Picture         =   "frmBackDesignDialog.frx":30A4E
      Tag             =   "BackDesign"
      Top             =   1155
      Width           =   1350
   End
   Begin VB.Image BicycleExpertBlue 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   3870
      Picture         =   "frmBackDesignDialog.frx":33BD1
      Tag             =   "BackDesign"
      Top             =   1155
      Width           =   1350
   End
   Begin VB.Image BicycleExpertBlueSelected 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   3870
      Picture         =   "frmBackDesignDialog.frx":388E4
      Tag             =   "BackDesign"
      Top             =   1155
      Width           =   1350
   End
   Begin VB.Image BicycleExpertRed 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   3870
      Picture         =   "frmBackDesignDialog.frx":3C24A
      Tag             =   "BackDesign"
      Top             =   1155
      Width           =   1350
   End
   Begin VB.Image BicycleExpertRedSelected 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   3870
      Picture         =   "frmBackDesignDialog.frx":40720
      Tag             =   "BackDesign"
      Top             =   1155
      Width           =   1350
   End
   Begin VB.Image BicycleFanBlue 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   3870
      Picture         =   "frmBackDesignDialog.frx":43A2A
      Tag             =   "BackDesign"
      Top             =   1155
      Width           =   1350
   End
   Begin VB.Image BicycleFanBlueSelected 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   3870
      Picture         =   "frmBackDesignDialog.frx":47F4F
      Tag             =   "BackDesign"
      Top             =   1155
      Width           =   1350
   End
   Begin VB.Image BicycleFanRed 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   3870
      Picture         =   "frmBackDesignDialog.frx":4B2B7
      Tag             =   "BackDesign"
      Top             =   1155
      Width           =   1350
   End
   Begin VB.Image BicycleFanRedSelected 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   3870
      Picture         =   "frmBackDesignDialog.frx":4F8CE
      Tag             =   "BackDesign"
      Top             =   1155
      Width           =   1350
   End
   Begin VB.Image BicycleLeagueBlue 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   3870
      Picture         =   "frmBackDesignDialog.frx":52D13
      Tag             =   "BackDesign"
      Top             =   1155
      Width           =   1350
   End
   Begin VB.Image BicycleLeagueBlueSelected 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   3870
      Picture         =   "frmBackDesignDialog.frx":5793C
      Tag             =   "BackDesign"
      Top             =   1155
      Width           =   1350
   End
   Begin VB.Image BicycleLeagueRed 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   3870
      Picture         =   "frmBackDesignDialog.frx":5B1E4
      Tag             =   "BackDesign"
      Top             =   1155
      Width           =   1350
   End
   Begin VB.Image BicycleLeagueRedSelected 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   3870
      Picture         =   "frmBackDesignDialog.frx":5F118
      Tag             =   "BackDesign"
      Top             =   1155
      Width           =   1350
   End
   Begin VB.Image BicycleRacerBlue 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   3870
      Picture         =   "frmBackDesignDialog.frx":62016
      Tag             =   "BackDesign"
      Top             =   1155
      Width           =   1350
   End
   Begin VB.Image BicycleRacerBlueSelected 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   3870
      Picture         =   "frmBackDesignDialog.frx":66746
      Tag             =   "BackDesign"
      Top             =   1155
      Width           =   1350
   End
   Begin VB.Image BicycleRacerRed 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   3870
      Picture         =   "frmBackDesignDialog.frx":69BE6
      Tag             =   "BackDesign"
      Top             =   1155
      Width           =   1350
   End
   Begin VB.Image BicycleRacerRedSelected 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   3870
      Picture         =   "frmBackDesignDialog.frx":6E426
      Tag             =   "BackDesign"
      Top             =   1155
      Width           =   1350
   End
   Begin VB.Image BicycleRiderBlue 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   3870
      Picture         =   "frmBackDesignDialog.frx":719DF
      Tag             =   "BackDesign"
      Top             =   1155
      Width           =   1350
   End
   Begin VB.Image BicycleRiderBlueSelected 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   3870
      Picture         =   "frmBackDesignDialog.frx":75FB5
      Tag             =   "BackDesign"
      Top             =   1155
      Width           =   1350
   End
   Begin VB.Image BicycleRiderRed 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   3870
      Picture         =   "frmBackDesignDialog.frx":79399
      Tag             =   "BackDesign"
      Top             =   1155
      Width           =   1350
   End
   Begin VB.Image BicycleRiderRedSelected 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   3870
      Picture         =   "frmBackDesignDialog.frx":7DC8C
      Tag             =   "BackDesign"
      Top             =   1155
      Width           =   1350
   End
   Begin VB.Image BicycleSolitaireBlue 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   3870
      Picture         =   "frmBackDesignDialog.frx":812A7
      Tag             =   "BackDesign"
      Top             =   1155
      Width           =   1350
   End
   Begin VB.Image BicycleSolitaireBlueSelected 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   3870
      Picture         =   "frmBackDesignDialog.frx":85D2A
      Tag             =   "BackDesign"
      Top             =   1155
      Width           =   1350
   End
   Begin VB.Image BicycleSolitaireRed 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   3870
      Picture         =   "frmBackDesignDialog.frx":8943D
      Tag             =   "BackDesign"
      Top             =   1155
      Width           =   1350
   End
   Begin VB.Image BicycleSolitaireRedSelected 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   3870
      Picture         =   "frmBackDesignDialog.frx":8D979
      Tag             =   "BackDesign"
      Top             =   1155
      Width           =   1350
   End
   Begin VB.Image BulldogSqueezerBlue 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   3870
      Picture         =   "frmBackDesignDialog.frx":90CBC
      Tag             =   "BackDesign"
      Top             =   1155
      Width           =   1350
   End
   Begin VB.Image BulldogSqueezerBlueSelected 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   3870
      Picture         =   "frmBackDesignDialog.frx":94C48
      Tag             =   "BackDesign"
      Top             =   1155
      Width           =   1350
   End
   Begin VB.Image BulldogSqueezerRedSelected 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   3870
      Picture         =   "frmBackDesignDialog.frx":97BFA
      Tag             =   "BackDesign"
      Top             =   1155
      Width           =   1350
   End
   Begin VB.Image StreamlineBlue 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   3870
      Picture         =   "frmBackDesignDialog.frx":9AABC
      Tag             =   "BackDesign"
      Top             =   1155
      Width           =   1350
   End
   Begin VB.Image StreamlineBlueSelected 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   3870
      Picture         =   "frmBackDesignDialog.frx":9F1D1
      Tag             =   "BackDesign"
      Top             =   1155
      Width           =   1350
   End
   Begin VB.Image StreamlineRed 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   3870
      Picture         =   "frmBackDesignDialog.frx":A26AC
      Tag             =   "BackDesign"
      Top             =   1155
      Width           =   1350
   End
   Begin VB.Image StreamlineRedSelected 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   3870
      Picture         =   "frmBackDesignDialog.frx":A6FE3
      Tag             =   "BackDesign"
      Top             =   1155
      Width           =   1350
   End
   Begin VB.Image TallyHoCircleBlue 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   3870
      Picture         =   "frmBackDesignDialog.frx":AA65C
      Tag             =   "BackDesign"
      Top             =   1155
      Width           =   1350
   End
   Begin VB.Image TallyHoCircleBlueSelected 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   3870
      Picture         =   "frmBackDesignDialog.frx":AF613
      Tag             =   "BackDesign"
      Top             =   1155
      Width           =   1350
   End
   Begin VB.Image TallyHoCircleRed 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   3870
      Picture         =   "frmBackDesignDialog.frx":B31B2
      Tag             =   "BackDesign"
      Top             =   1155
      Width           =   1350
   End
   Begin VB.Image TallyHoCircleRedSelected 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   3870
      Picture         =   "frmBackDesignDialog.frx":B803E
      Tag             =   "BackDesign"
      Top             =   1155
      Width           =   1350
   End
   Begin VB.Image TallyHoFanBlue 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   3870
      Picture         =   "frmBackDesignDialog.frx":BBAA6
      Tag             =   "BackDesign"
      Top             =   1155
      Width           =   1350
   End
   Begin VB.Image TallyHoFanBlueSelected 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   3870
      Picture         =   "frmBackDesignDialog.frx":C0144
      Tag             =   "BackDesign"
      Top             =   1155
      Width           =   1350
   End
   Begin VB.Image TallyHoFanRedSelected 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   3870
      Picture         =   "frmBackDesignDialog.frx":C35D9
      Tag             =   "BackDesign"
      Top             =   1155
      Width           =   1350
   End
   Begin VB.Label Label1 
      Caption         =   "Select Back Design"
      Height          =   330
      Left            =   180
      TabIndex        =   3
      Top             =   270
      Width           =   2550
   End
   Begin VB.Image BulldogSqueezerRed 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   3870
      Picture         =   "frmBackDesignDialog.frx":C6B5B
      Tag             =   "BackDesign"
      Top             =   1155
      Width           =   1350
   End
End
Attribute VB_Name = "frmBackDesignDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Since the original coding, the US Playing Card Company identified different
'trademark names for some of the back designs than I had initially understood.
'The code below still uses the incorrect names in the non-user-visible places,
'but the executable code refers to the back designs correctly.
'Incorrect                      Correct
' Original                      Name per
'   Name                         USPC
'-------------------------------------------
'Aladdin                        Feather
'Expert                         Old Fan
'Fan                            New Fan
'Solitaire                      High Wheel


Private Sub BackDesignBlueOption_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case BackDesignCombo.Text
    Case "Feather"
        If BackDesignBlueOption.Value = True Then
            BackDesignSelected = "AladdinBlue"
            AladdinBlue.ZOrder
        ElseIf BackDesignRedOption.Value = True Then
            BackDesignSelected = "AladdinRed"
            AladdinRed.ZOrder
        End If
    Case "Aviator"
        If BackDesignBlueOption.Value = True Then
            BackDesignSelected = "AviatorBlue"
            AviatorBlue.ZOrder
        ElseIf BackDesignRedOption.Value = True Then
            BackDesignSelected = "AviatorRed"
            AviatorRed.ZOrder
        End If
    Case "Bicycle Auto"
        If BackDesignBlueOption.Value = True Then
            BackDesignSelected = "BicycleAutoBlue"
            BicycleAutoBlue.ZOrder
        ElseIf BackDesignRedOption.Value = True Then
            BackDesignSelected = "BicycleAutoRed"
            BicycleAutoRed.ZOrder
        End If
    Case "Bicycle Old Fan"
        If BackDesignBlueOption.Value = True Then
            BackDesignSelected = "BicycleExpertBlue"
            BicycleExpertBlue.ZOrder
        ElseIf BackDesignRedOption.Value = True Then
            BackDesignSelected = "BicycleExpertRed"
            BicycleExpertRed.ZOrder
        End If
    Case "Bicycle New Fan"
        If BackDesignBlueOption.Value = True Then
            BackDesignSelected = "BicycleFanBlue"
            BicycleFanBlue.ZOrder
        ElseIf BackDesignRedOption.Value = True Then
            BackDesignSelected = "BicycleFanRed"
            BicycleFanRed.ZOrder
        End If
    Case "Bicycle League"
        If BackDesignBlueOption.Value = True Then
            BackDesignSelected = "BicycleLeagueBlue"
            BicycleLeagueBlue.ZOrder
        ElseIf BackDesignRedOption.Value = True Then
            BackDesignSelected = "BicycleLeagueRed"
            BicycleLeagueRed.ZOrder
        End If
    Case "Bicycle Racer"
        If BackDesignBlueOption.Value = True Then
            BackDesignSelected = "BicycleRacerBlue"
            BicycleRacerBlue.ZOrder
        ElseIf BackDesignRedOption.Value = True Then
            BackDesignSelected = "BicycleRacerRed"
            BicycleRacerRed.ZOrder
        End If
    Case "Bicycle Rider"
        If BackDesignBlueOption.Value = True Then
            BackDesignSelected = "BicycleRiderBlue"
            BicycleRiderBlue.ZOrder
        ElseIf BackDesignRedOption.Value = True Then
            BackDesignSelected = "BicycleRiderRed"
            BicycleRiderRed.ZOrder
        End If
    Case "Bicycle High Wheel"
        If BackDesignBlueOption.Value = True Then
            BackDesignSelected = "BicycleSolitaireBlue"
            BicycleSolitaireBlue.ZOrder
        ElseIf BackDesignRedOption.Value = True Then
            BackDesignSelected = "BicycleSolitaireRed"
            BicycleSolitaireRed.ZOrder
        End If
    Case "Bulldog Squeezer"
        If BackDesignBlueOption.Value = True Then
            BackDesignSelected = "BulldogSqueezerBlue"
            BulldogSqueezerBlue.ZOrder
        ElseIf BackDesignRedOption.Value = True Then
            BackDesignSelected = "BulldogSqueezerRed"
            BulldogSqueezerRed.ZOrder
        End If
    Case "Streamline"
        If BackDesignBlueOption.Value = True Then
            BackDesignSelected = "StreamlineBlue"
            StreamlineBlue.ZOrder
        ElseIf BackDesignRedOption.Value = True Then
            BackDesignSelected = "StreamlineRed"
            StreamlineRed.ZOrder
        End If
    Case "Tally Ho Circle"
        If BackDesignBlueOption.Value = True Then
            BackDesignSelected = "TallyHoCircleBlue"
            TallyHoCircleBlue.ZOrder
        ElseIf BackDesignRedOption.Value = True Then
            BackDesignSelected = "TallyHoCircleRed"
            TallyHoCircleRed.ZOrder
        End If
    Case "Tally Ho Fan"
        If BackDesignBlueOption.Value = True Then
            BackDesignSelected = "TallyHoFanBlue"
            TallyHoFanBlue.ZOrder
        ElseIf BackDesignRedOption.Value = True Then
            BackDesignSelected = "TallyHoFanRed"
            TallyHoFanRed.ZOrder
        End If
End Select
End Sub

Private Sub BackDesignCombo_Change()
Call ShowBackDesign
End Sub


Private Sub BackDesignCombo_Click()
Call ShowBackDesign
End Sub

Private Sub BackDesignRedOption_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case BackDesignCombo.Text
    Case "Feather"
        If BackDesignBlueOption.Value = True Then
            BackDesignSelected = "AladdinBlue"
            AladdinBlue.ZOrder
        ElseIf BackDesignRedOption.Value = True Then
            BackDesignSelected = "AladdinRed"
            AladdinRed.ZOrder
        End If
    Case "Aviator"
        If BackDesignBlueOption.Value = True Then
            BackDesignSelected = "AviatorBlue"
            AviatorBlue.ZOrder
        ElseIf BackDesignRedOption.Value = True Then
            BackDesignSelected = "AviatorRed"
            AviatorRed.ZOrder
        End If
    Case "Bicycle Auto"
        If BackDesignBlueOption.Value = True Then
            BackDesignSelected = "BicycleAutoBlue"
            BicycleAutoBlue.ZOrder
        ElseIf BackDesignRedOption.Value = True Then
            BackDesignSelected = "BicycleAutoRed"
            BicycleAutoRed.ZOrder
        End If
    Case "Bicycle Old Fan"
        If BackDesignBlueOption.Value = True Then
            BackDesignSelected = "BicycleExpertBlue"
            BicycleExpertBlue.ZOrder
        ElseIf BackDesignRedOption.Value = True Then
            BackDesignSelected = "BicycleExpertRed"
            BicycleExpertRed.ZOrder
        End If
    Case "Bicycle New Fan"
        If BackDesignBlueOption.Value = True Then
            BackDesignSelected = "BicycleFanBlue"
            BicycleFanBlue.ZOrder
        ElseIf BackDesignRedOption.Value = True Then
            BackDesignSelected = "BicycleFanRed"
            BicycleFanRed.ZOrder
        End If
    Case "Bicycle League"
        If BackDesignBlueOption.Value = True Then
            BackDesignSelected = "BicycleLeagueBlue"
            BicycleLeagueBlue.ZOrder
        ElseIf BackDesignRedOption.Value = True Then
            BackDesignSelected = "BicycleLeagueRed"
            BicycleLeagueRed.ZOrder
        End If
    Case "Bicycle Racer"
        If BackDesignBlueOption.Value = True Then
            BackDesignSelected = "BicycleRacerBlue"
            BicycleRacerBlue.ZOrder
        ElseIf BackDesignRedOption.Value = True Then
            BackDesignSelected = "BicycleRacerRed"
            BicycleRacerRed.ZOrder
        End If
    Case "Bicycle Rider"
        If BackDesignBlueOption.Value = True Then
            BackDesignSelected = "BicycleRiderBlue"
            BicycleRiderBlue.ZOrder
        ElseIf BackDesignRedOption.Value = True Then
            BackDesignSelected = "BicycleRiderRed"
            BicycleRiderRed.ZOrder
        End If
    Case "Bicycle High Wheel"
        If BackDesignBlueOption.Value = True Then
            BackDesignSelected = "BicycleSolitaireBlue"
            BicycleSolitaireBlue.ZOrder
        ElseIf BackDesignRedOption.Value = True Then
            BackDesignSelected = "BicycleSolitaireRed"
            BicycleSolitaireRed.ZOrder
        End If
    Case "Bulldog Squeezer"
        If BackDesignBlueOption.Value = True Then
            BackDesignSelected = "BulldogSqueezerBlue"
            BulldogSqueezerBlue.ZOrder
        ElseIf BackDesignRedOption.Value = True Then
            BackDesignSelected = "BulldogSqueezerRed"
            BulldogSqueezerRed.ZOrder
        End If
    Case "Streamline"
        If BackDesignBlueOption.Value = True Then
            BackDesignSelected = "StreamlineBlue"
            StreamlineBlue.ZOrder
        ElseIf BackDesignRedOption.Value = True Then
            BackDesignSelected = "StreamlineRed"
            StreamlineRed.ZOrder
        End If
    Case "Tally Ho Circle"
        If BackDesignBlueOption.Value = True Then
            BackDesignSelected = "TallyHoCircleBlue"
            TallyHoCircleBlue.ZOrder
        ElseIf BackDesignRedOption.Value = True Then
            BackDesignSelected = "TallyHoCircleRed"
            TallyHoCircleRed.ZOrder
        End If
    Case "Tally Ho Fan"
        If BackDesignBlueOption.Value = True Then
            BackDesignSelected = "TallyHoFanBlue"
            TallyHoFanBlue.ZOrder
        ElseIf BackDesignRedOption.Value = True Then
            BackDesignSelected = "TallyHoFanRed"
            TallyHoFanRed.ZOrder
        End If
End Select
End Sub

Private Sub CancelButton_Click()
Unload Me
End Sub

Private Sub Form_Load()
If BackDesignCurrent = "AladdinBlue" Then
    BackDesignBlueOption.Value = True
    BackDesignCombo.Text = "Feather"
ElseIf BackDesignCurrent = "AladdinRed" Then
    BackDesignRedOption.Value = True
    BackDesignCombo.Text = "Feather"
ElseIf BackDesignCurrent = "AviatorBlue" Then
    BackDesignBlueOption.Value = True
    BackDesignCombo.Text = "Aviator"
ElseIf BackDesignCurrent = "AviatorRed" Then
    BackDesignRedOption.Value = True
    BackDesignCombo.Text = "Aviator"
ElseIf BackDesignCurrent = "BicycleAutoBlue" Then
    BackDesignBlueOption.Value = True
    BackDesignCombo.Text = "Bicycle Auto"
ElseIf BackDesignCurrent = "BicycleAutoRed" Then
    BackDesignRedOption.Value = True
    BackDesignCombo.Text = "Bicycle Auto"
ElseIf BackDesignCurrent = "BicycleExpertBlue" Then
    BackDesignBlueOption.Value = True
    BackDesignCombo.Text = "Bicycle Old Fan"
ElseIf BackDesignCurrent = "BicycleExpertRed" Then
    BackDesignRedOption.Value = True
    BackDesignCombo.Text = "Bicycle Old Fan"
ElseIf BackDesignCurrent = "BicycleFanBlue" Then
    BackDesignBlueOption.Value = True
    BackDesignCombo.Text = "Bicycle New Fan"
ElseIf BackDesignCurrent = "BicycleFanRed" Then
    BackDesignRedOption.Value = True
    BackDesignCombo.Text = "Bicycle New Fan"
ElseIf BackDesignCurrent = "BicycleLeagueBlue" Then
    BackDesignBlueOption.Value = True
    BackDesignCombo.Text = "Bicycle League"
ElseIf BackDesignCurrent = "BicycleLeagueRed" Then
    BackDesignRedOption.Value = True
    BackDesignCombo.Text = "Bicycle League"
ElseIf BackDesignCurrent = "BicycleRacerBlue" Then
    BackDesignBlueOption.Value = True
    BackDesignCombo.Text = "Bicycle Racer"
ElseIf BackDesignCurrent = "BicycleRacerRed" Then
    BackDesignRedOption.Value = True
    BackDesignCombo.Text = "Bicycle Racer"
ElseIf BackDesignCurrent = "BicycleRiderBlue" Then
    BackDesignBlueOption.Value = True
    BackDesignCombo.Text = "Bicycle Rider"
ElseIf BackDesignCurrent = "BicycleRiderRed" Then
    BackDesignRedOption.Value = True
    BackDesignCombo.Text = "Bicycle Rider"
ElseIf BackDesignCurrent = "BicycleSolitaireBlue" Then
    BackDesignBlueOption.Value = True
    BackDesignCombo.Text = "Bicycle High Wheel"
ElseIf BackDesignCurrent = "BicycleSolitaireRed" Then
    BackDesignRedOption.Value = True
    BackDesignCombo.Text = "Bicycle High Wheel"
ElseIf BackDesignCurrent = "BulldogSqueezerBlue" Then
    BackDesignBlueOption.Value = True
    BackDesignCombo.Text = "Bulldog Squeezer"
ElseIf BackDesignCurrent = "BulldogSqueezerRed" Then
    BackDesignRedOption.Value = True
    BackDesignCombo.Text = "Bulldog Squeezer"
ElseIf BackDesignCurrent = "StreamlineBlue" Then
    BackDesignBlueOption.Value = True
    BackDesignCombo.Text = "Streamline"
ElseIf BackDesignCurrent = "StreamlineRed" Then
    BackDesignRedOption.Value = True
    BackDesignCombo.Text = "Streamline"
ElseIf BackDesignCurrent = "TallyHoCircleBlue" Then
    BackDesignBlueOption.Value = True
    BackDesignCombo.Text = "Tally Ho Circle"
ElseIf BackDesignCurrent = "TallyHoCircleRed" Then
    BackDesignRedOption.Value = True
    BackDesignCombo.Text = "Tally Ho Circle"
ElseIf BackDesignCurrent = "TallyHoFanBlue" Then
    BackDesignBlueOption.Value = True
    BackDesignCombo.Text = "Tally Ho Fan"
ElseIf BackDesignCurrent = "TallyHoFanRed" Then
    BackDesignRedOption.Value = True
    BackDesignCombo.Text = "Tally Ho Fan"
End If
If Right(BackDesignCurrent, 3) = "Red" Then
    BackDesignRedOption.Value = True
Else
    BackDesignBlueOption.Value = True
End If
Select Case BackDesignCombo.Text
    Case "Feather"
        If BackDesignBlueOption.Value = True Then
            BackDesignSelected = "AladdinBlue"
            AladdinBlue.ZOrder
        ElseIf BackDesignRedOption.Value = True Then
            BackDesignSelected = "AladdinRed"
            AladdinRed.ZOrder
        End If
    Case "Aviator"
        If BackDesignBlueOption.Value = True Then
            BackDesignSelected = "AviatorBlue"
            AviatorBlue.ZOrder
        ElseIf BackDesignRedOption.Value = True Then
            BackDesignSelected = "AviatorRed"
            AviatorRed.ZOrder
        End If
    Case "Bicycle Auto"
        If BackDesignBlueOption.Value = True Then
            BackDesignSelected = "BicycleAutoBlue"
            BicycleAutoBlue.ZOrder
        ElseIf BackDesignRedOption.Value = True Then
            BackDesignSelected = "BicycleAutoRed"
            BicycleAutoRed.ZOrder
        End If
    Case "Bicycle Old Fan"
        If BackDesignBlueOption.Value = True Then
            BackDesignSelected = "BicycleExpertBlue"
            BicycleExpertBlue.ZOrder
        ElseIf BackDesignRedOption.Value = True Then
            BackDesignSelected = "BicycleExpertRed"
            BicycleExpertRed.ZOrder
        End If
    Case "Bicycle New Fan"
        If BackDesignBlueOption.Value = True Then
            BackDesignSelected = "BicycleFanBlue"
            BicycleFanBlue.ZOrder
        ElseIf BackDesignRedOption.Value = True Then
            BackDesignSelected = "BicycleFanRed"
            BicycleFanRed.ZOrder
        End If
    Case "Bicycle League"
        If BackDesignBlueOption.Value = True Then
            BackDesignSelected = "BicycleLeagueBlue"
            BicycleLeagueBlue.ZOrder
        ElseIf BackDesignRedOption.Value = True Then
            BackDesignSelected = "BicycleLeagueRed"
            BicycleLeagueRed.ZOrder
        End If
    Case "Bicycle Racer"
        If BackDesignBlueOption.Value = True Then
            BackDesignSelected = "BicycleRacerBlue"
            BicycleRacerBlue.ZOrder
        ElseIf BackDesignRedOption.Value = True Then
            BackDesignSelected = "BicycleRacerRed"
            BicycleRacerRed.ZOrder
        End If
    Case "Bicycle Rider"
        If BackDesignBlueOption.Value = True Then
            BackDesignSelected = "BicycleRiderBlue"
            BicycleRiderBlue.ZOrder
        ElseIf BackDesignRedOption.Value = True Then
            BackDesignSelected = "BicycleRiderRed"
            BicycleRiderRed.ZOrder
        End If
    Case "Bicycle High Wheel"
        If BackDesignBlueOption.Value = True Then
            BackDesignSelected = "BicycleSolitaireBlue"
            BicycleSolitaireBlue.ZOrder
        ElseIf BackDesignRedOption.Value = True Then
            BackDesignSelected = "BicycleSolitaireRed"
            BicycleSolitaireRed.ZOrder
        End If
    Case "Bulldog Squeezer"
        If BackDesignBlueOption.Value = True Then
            BackDesignSelected = "BulldogSqueezerBlue"
            BulldogSqueezerBlue.ZOrder
        ElseIf BackDesignRedOption.Value = True Then
            BackDesignSelected = "BulldogSqueezerRed"
            BulldogSqueezerRed.ZOrder
        End If
    Case "Streamline"
        If BackDesignBlueOption.Value = True Then
            BackDesignSelected = "StreamlineBlue"
            StreamlineBlue.ZOrder
        ElseIf BackDesignRedOption.Value = True Then
            BackDesignSelected = "StreamlineRed"
            StreamlineRed.ZOrder
        End If
    Case "Tally Ho Circle"
        If BackDesignBlueOption.Value = True Then
            BackDesignSelected = "TallyHoCircleBlue"
            TallyHoCircleBlue.ZOrder
        ElseIf BackDesignRedOption.Value = True Then
            BackDesignSelected = "TallyHoCircleRed"
            TallyHoCircleRed.ZOrder
        End If
    Case "Tally Ho Fan"
        If BackDesignBlueOption.Value = True Then
            BackDesignSelected = "TallyHoFanBlue"
            TallyHoFanBlue.ZOrder
        ElseIf BackDesignRedOption.Value = True Then
            BackDesignSelected = "TallyHoFanRed"
            TallyHoFanRed.ZOrder
        End If
End Select
End Sub

Public Sub LoadBackDesign(pDesign)
Select Case pDesign
    Case "AladdinBlue"
        For i% = 0 To 51
            frmDeck.Back(i%).Picture = AladdinBlue.Picture
            frmDeck.BackSelected(i%).Picture = AladdinBlueSelected.Picture
        Next i%
    Case "AladdinRed"
        For i% = 0 To 51
            frmDeck.Back(i%).Picture = AladdinRed.Picture
            frmDeck.BackSelected(i%).Picture = AladdinRedSelected.Picture
        Next i%
    Case "AviatorBlue"
        For i% = 0 To 51
            frmDeck.Back(i%).Picture = AviatorBlue.Picture
            frmDeck.BackSelected(i%).Picture = AviatorBlueSelected.Picture
        Next i%
    Case "AviatorRed"
        For i% = 0 To 51
            frmDeck.Back(i%).Picture = AviatorRed.Picture
            frmDeck.BackSelected(i%).Picture = AviatorRedSelected.Picture
        Next i%
    Case "BicycleAutoBlue"
        For i% = 0 To 51
            frmDeck.Back(i%).Picture = BicycleAutoBlue.Picture
            frmDeck.BackSelected(i%).Picture = BicycleAutoBlueSelected.Picture
        Next i%
    Case "BicycleAutoRed"
        For i% = 0 To 51
            frmDeck.Back(i%).Picture = BicycleAutoRed.Picture
            frmDeck.BackSelected(i%).Picture = BicycleAutoRedSelected.Picture
        Next i%
    Case "BicycleExpertBlue"
        For i% = 0 To 51
            frmDeck.Back(i%).Picture = BicycleExpertBlue.Picture
            frmDeck.BackSelected(i%).Picture = BicycleExpertBlueSelected.Picture
        Next i%
    Case "BicycleExpertRed"
        For i% = 0 To 51
            frmDeck.Back(i%).Picture = BicycleExpertRed.Picture
            frmDeck.BackSelected(i%).Picture = BicycleExpertRedSelected.Picture
        Next i%
    Case "BicycleFanBlue"
        For i% = 0 To 51
            frmDeck.Back(i%).Picture = BicycleFanBlue.Picture
            frmDeck.BackSelected(i%).Picture = BicycleFanBlueSelected.Picture
        Next i%
    Case "BicycleFanRed"
        For i% = 0 To 51
            frmDeck.Back(i%).Picture = BicycleFanRed.Picture
            frmDeck.BackSelected(i%).Picture = BicycleFanRedSelected.Picture
        Next i%
    Case "BicycleLeagueBlue"
        For i% = 0 To 51
            frmDeck.Back(i%).Picture = BicycleLeagueBlue.Picture
            frmDeck.BackSelected(i%).Picture = BicycleLeagueBlueSelected.Picture
        Next i%
    Case "BicycleLeagueRed"
        For i% = 0 To 51
            frmDeck.Back(i%).Picture = BicycleLeagueRed.Picture
            frmDeck.BackSelected(i%).Picture = BicycleLeagueRedSelected.Picture
        Next i%
    Case "BicycleRacerBlue"
        For i% = 0 To 51
            frmDeck.Back(i%).Picture = BicycleRacerBlue.Picture
            frmDeck.BackSelected(i%).Picture = BicycleRacerBlueSelected.Picture
        Next i%
    Case "BicycleRacerRed"
        For i% = 0 To 51
            frmDeck.Back(i%).Picture = BicycleRacerRed.Picture
            frmDeck.BackSelected(i%).Picture = BicycleRacerRedSelected.Picture
        Next i%
    Case "BicycleRiderBlue"
        For i% = 0 To 51
            frmDeck.Back(i%).Picture = BicycleRiderBlue.Picture
            frmDeck.BackSelected(i%).Picture = BicycleRiderBlueSelected.Picture
        Next i%
    Case "BicycleRiderRed"
        For i% = 0 To 51
            frmDeck.Back(i%).Picture = BicycleRiderRed.Picture
            frmDeck.BackSelected(i%).Picture = BicycleRiderRedSelected.Picture
        Next i%
    Case "BicycleSolitaireBlue"
        For i% = 0 To 51
            frmDeck.Back(i%).Picture = BicycleSolitaireBlue.Picture
            frmDeck.BackSelected(i%).Picture = BicycleSolitaireBlueSelected.Picture
        Next i%
    Case "BicycleSolitaireRed"
        For i% = 0 To 51
            frmDeck.Back(i%).Picture = BicycleSolitaireRed.Picture
            frmDeck.BackSelected(i%).Picture = BicycleSolitaireRedSelected.Picture
        Next i%
    Case "BulldogSqueezerBlue"
        For i% = 0 To 51
            frmDeck.Back(i%).Picture = BulldogSqueezerBlue.Picture
            frmDeck.BackSelected(i%).Picture = BulldogSqueezerBlueSelected.Picture
        Next i%
    Case "BulldogSqueezerRed"
        For i% = 0 To 51
            frmDeck.Back(i%).Picture = BulldogSqueezerRed.Picture
            frmDeck.BackSelected(i%).Picture = BulldogSqueezerRedSelected.Picture
        Next i%
    Case "StreamlineBlue"
        For i% = 0 To 51
            frmDeck.Back(i%).Picture = StreamlineBlue.Picture
            frmDeck.BackSelected(i%).Picture = StreamlineBlueSelected.Picture
        Next i%
    Case "StreamlineRed"
        For i% = 0 To 51
            frmDeck.Back(i%).Picture = StreamlineRed.Picture
            frmDeck.BackSelected(i%).Picture = StreamlineRedSelected.Picture
        Next i%
    Case "TallyHoCircleBlue"
        For i% = 0 To 51
            frmDeck.Back(i%).Picture = TallyHoCircleBlue.Picture
            frmDeck.BackSelected(i%).Picture = TallyHoCircleBlueSelected.Picture
        Next i%
    Case "TallyHoCircleRed"
        For i% = 0 To 51
            frmDeck.Back(i%).Picture = TallyHoCircleRed.Picture
            frmDeck.BackSelected(i%).Picture = TallyHoCircleRedSelected.Picture
        Next i%
    Case "TallyHoFanBlue"
        For i% = 0 To 51
            frmDeck.Back(i%).Picture = TallyHoFanBlue.Picture
            frmDeck.BackSelected(i%).Picture = TallyHoFanBlueSelected.Picture
        Next i%
    Case "TallyHoFanRed"
        For i% = 0 To 51
            frmDeck.Back(i%).Picture = TallyHoFanRed.Picture
            frmDeck.BackSelected(i%).Picture = TallyHoFanRedSelected.Picture
        Next i%
End Select
End Sub

Private Sub OKButton_Click()
'the Deck window must be present
If Not frmMain.mnuDeck.Checked Then
    MsgBox ("You must have the Deck window open" & Chr(13) & _
        "to complete the Set Back Design." & Chr(13) & Chr(13) & _
        "First press Cancel, then " & Chr(13) & _
        "select the Deck option from the View menu.")
    Exit Sub
End If
Call LoadBackDesign(BackDesignSelected)
frmStackView.ShowCards
BackDesignCurrent = BackDesignSelected
Unload Me
End Sub

Private Sub ShowBackDesign()
Select Case BackDesignCombo.ListIndex
    Case 0
        If BackDesignBlueOption.Value = True Then
            BackDesignSelected = "AviatorBlue"
            AviatorBlue.ZOrder
        ElseIf BackDesignRedOption.Value = True Then
            BackDesignSelected = "AviatorRed"
            AviatorRed.ZOrder
        End If
    Case 1
        If BackDesignBlueOption.Value = True Then
            BackDesignSelected = "BicycleAutoBlue"
            BicycleAutoBlue.ZOrder
        ElseIf BackDesignRedOption.Value = True Then
            BackDesignSelected = "BicycleAutoRed"
            BicycleAutoRed.ZOrder
        End If
    Case 2
        If BackDesignBlueOption.Value = True Then
            BackDesignSelected = "BicycleSolitaireBlue"
            BicycleSolitaireBlue.ZOrder
        ElseIf BackDesignRedOption.Value = True Then
            BackDesignSelected = "BicycleSolitaireRed"
            BicycleSolitaireRed.ZOrder
        End If
    Case 3
        If BackDesignBlueOption.Value = True Then
            BackDesignSelected = "BicycleLeagueBlue"
            BicycleLeagueBlue.ZOrder
        ElseIf BackDesignRedOption.Value = True Then
            BackDesignSelected = "BicycleLeagueRed"
            BicycleLeagueRed.ZOrder
        End If
    Case 4
        If BackDesignBlueOption.Value = True Then
            BackDesignSelected = "BicycleFanBlue"
            BicycleFanBlue.ZOrder
        ElseIf BackDesignRedOption.Value = True Then
            BackDesignSelected = "BicycleFanRed"
            BicycleFanRed.ZOrder
        End If
    Case 5
        If BackDesignBlueOption.Value = True Then
            BackDesignSelected = "BicycleExpertBlue"
            BicycleExpertBlue.ZOrder
        ElseIf BackDesignRedOption.Value = True Then
            BackDesignSelected = "BicycleExpertRed"
            BicycleExpertRed.ZOrder
        End If
    Case 6
        If BackDesignBlueOption.Value = True Then
            BackDesignSelected = "BicycleRacerBlue"
            BicycleRacerBlue.ZOrder
        ElseIf BackDesignRedOption.Value = True Then
            BackDesignSelected = "BicycleRacerRed"
            BicycleRacerRed.ZOrder
        End If
    Case 7
        If BackDesignBlueOption.Value = True Then
            BackDesignSelected = "BicycleRiderBlue"
            BicycleRiderBlue.ZOrder
        ElseIf BackDesignRedOption.Value = True Then
            BackDesignSelected = "BicycleRiderRed"
            BicycleRiderRed.ZOrder
        End If
    Case 8
        If BackDesignBlueOption.Value = True Then
            BackDesignSelected = "BulldogSqueezerBlue"
            BulldogSqueezerBlue.ZOrder
        ElseIf BackDesignRedOption.Value = True Then
            BackDesignSelected = "BulldogSqueezerRed"
            BulldogSqueezerRed.ZOrder
        End If
    Case 9
        If BackDesignBlueOption.Value = True Then
            BackDesignSelected = "AladdinBlue"
            AladdinBlue.ZOrder
        ElseIf BackDesignRedOption.Value = True Then
            BackDesignSelected = "AladdinRed"
            AladdinRed.ZOrder
        End If
    Case 10
        If BackDesignBlueOption.Value = True Then
            BackDesignSelected = "StreamlineBlue"
            StreamlineBlue.ZOrder
        ElseIf BackDesignRedOption.Value = True Then
            BackDesignSelected = "StreamlineRed"
            StreamlineRed.ZOrder
        End If
    Case 11
        If BackDesignBlueOption.Value = True Then
            BackDesignSelected = "TallyHoCircleBlue"
            TallyHoCircleBlue.ZOrder
        ElseIf BackDesignRedOption.Value = True Then
            BackDesignSelected = "TallyHoCircleRed"
            TallyHoCircleRed.ZOrder
        End If
    Case 12
        If BackDesignBlueOption.Value = True Then
            BackDesignSelected = "TallyHoFanBlue"
            TallyHoFanBlue.ZOrder
        ElseIf BackDesignRedOption.Value = True Then
            BackDesignSelected = "TallyHoFanRed"
            TallyHoFanRed.ZOrder
        End If
End Select
End Sub
