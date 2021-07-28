VERSION 5.00
Begin VB.Form frmCustomDeck 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Custom Deck"
   ClientHeight    =   6285
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7830
   Icon            =   "frmCustomDeck.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6285
   ScaleWidth      =   7830
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton ReorderRetainedStack 
      Caption         =   "Reorder Retained Stack"
      Enabled         =   0   'False
      Height          =   315
      Left            =   135
      TabIndex        =   59
      Top             =   1260
      Width           =   3180
   End
   Begin VB.CommandButton ResetCustomStack 
      Caption         =   "Reset: New"
      Height          =   315
      Left            =   135
      TabIndex        =   5
      Top             =   900
      Width           =   1560
   End
   Begin VB.CommandButton TransferCustomStack 
      Caption         =   "Transfer: New"
      Height          =   315
      Left            =   135
      TabIndex        =   4
      Top             =   450
      Width           =   1560
   End
   Begin VB.CommandButton StanyonSpecialDeck 
      Caption         =   "Create Stanyon Variation Deck from first Five Cards"
      Height          =   315
      Left            =   3390
      TabIndex        =   3
      Top             =   450
      Width           =   4245
   End
   Begin VB.CommandButton ImportCurrentDeck 
      Caption         =   "Import Current Deck from StackView"
      Height          =   315
      Left            =   135
      TabIndex        =   2
      Top             =   90
      Width           =   3180
   End
   Begin VB.CommandButton TransferCustomStackRetain 
      Caption         =   "Transfer: Retain"
      Height          =   315
      Left            =   1755
      TabIndex        =   1
      Top             =   450
      Width           =   1545
   End
   Begin VB.CommandButton ResetCustomStackRetain 
      Caption         =   "Reset: Retain"
      Height          =   315
      Left            =   1755
      TabIndex        =   0
      Top             =   900
      Width           =   1560
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   135
      X2              =   3285
      Y1              =   825
      Y2              =   825
   End
   Begin VB.Label StackVal1 
      Caption         =   "52"
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
      Index           =   52
      Left            =   7455
      TabIndex        =   58
      Tag             =   "StackVal"
      Top             =   2760
      Width           =   135
   End
   Begin VB.Label StackVal1 
      Caption         =   "51"
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
      Index           =   51
      Left            =   7155
      TabIndex        =   57
      Tag             =   "StackVal"
      Top             =   2760
      Width           =   135
   End
   Begin VB.Label StackVal1 
      Caption         =   "50"
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
      Index           =   50
      Left            =   6855
      TabIndex        =   56
      Tag             =   "StackVal"
      Top             =   2745
      Width           =   135
   End
   Begin VB.Label StackVal1 
      Caption         =   "49"
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
      Index           =   49
      Left            =   6570
      TabIndex        =   55
      Tag             =   "StackVal"
      Top             =   2790
      Width           =   135
   End
   Begin VB.Label StackVal1 
      Caption         =   "48"
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
      Index           =   48
      Left            =   6315
      TabIndex        =   54
      Tag             =   "StackVal"
      Top             =   2775
      Width           =   135
   End
   Begin VB.Label StackVal1 
      Caption         =   "47"
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
      Index           =   47
      Left            =   6015
      TabIndex        =   53
      Tag             =   "StackVal"
      Top             =   2760
      Width           =   135
   End
   Begin VB.Label StackVal1 
      Caption         =   "46"
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
      Index           =   46
      Left            =   5715
      TabIndex        =   52
      Tag             =   "StackVal"
      Top             =   2745
      Width           =   135
   End
   Begin VB.Label StackVal1 
      Caption         =   "45"
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
      Index           =   45
      Left            =   5430
      TabIndex        =   51
      Tag             =   "StackVal"
      Top             =   2745
      Width           =   135
   End
   Begin VB.Label StackVal1 
      Caption         =   "44"
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
      Index           =   44
      Left            =   5145
      TabIndex        =   50
      Tag             =   "StackVal"
      Top             =   2775
      Width           =   135
   End
   Begin VB.Label StackVal1 
      Caption         =   "43"
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
      Index           =   43
      Left            =   4860
      TabIndex        =   49
      Tag             =   "StackVal"
      Top             =   2760
      Width           =   135
   End
   Begin VB.Label StackVal1 
      Caption         =   "42"
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
      Index           =   42
      Left            =   4575
      TabIndex        =   48
      Tag             =   "StackVal"
      Top             =   2775
      Width           =   135
   End
   Begin VB.Label StackVal1 
      Caption         =   "41"
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
      Index           =   41
      Left            =   4290
      TabIndex        =   47
      Tag             =   "StackVal"
      Top             =   2760
      Width           =   135
   End
   Begin VB.Label StackVal1 
      Caption         =   "40"
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
      Index           =   40
      Left            =   4020
      TabIndex        =   46
      Tag             =   "StackVal"
      Top             =   2760
      Width           =   135
   End
   Begin VB.Label StackVal1 
      Caption         =   "39"
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
      Index           =   39
      Left            =   3585
      TabIndex        =   45
      Tag             =   "StackVal"
      Top             =   2745
      Width           =   135
   End
   Begin VB.Label StackVal1 
      Caption         =   "38"
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
      Index           =   38
      Left            =   3330
      TabIndex        =   44
      Tag             =   "StackVal"
      Top             =   2760
      Width           =   135
   End
   Begin VB.Label StackVal1 
      Caption         =   "37"
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
      Index           =   37
      Left            =   3015
      TabIndex        =   43
      Tag             =   "StackVal"
      Top             =   2760
      Width           =   135
   End
   Begin VB.Label StackVal1 
      Caption         =   "36"
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
      Index           =   36
      Left            =   2730
      TabIndex        =   42
      Tag             =   "StackVal"
      Top             =   2775
      Width           =   135
   End
   Begin VB.Label StackVal1 
      Caption         =   "35"
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
      Index           =   35
      Left            =   2460
      TabIndex        =   41
      Tag             =   "StackVal"
      Top             =   2790
      Width           =   135
   End
   Begin VB.Label StackVal1 
      Caption         =   "34"
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
      Index           =   34
      Left            =   2190
      TabIndex        =   40
      Tag             =   "StackVal"
      Top             =   2760
      Width           =   135
   End
   Begin VB.Label StackVal1 
      Caption         =   "33"
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
      Index           =   33
      Left            =   1920
      TabIndex        =   39
      Tag             =   "StackVal"
      Top             =   2775
      Width           =   135
   End
   Begin VB.Label StackVal1 
      Caption         =   "32"
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
      Index           =   32
      Left            =   1605
      TabIndex        =   38
      Tag             =   "StackVal"
      Top             =   2790
      Width           =   135
   End
   Begin VB.Label StackVal1 
      Caption         =   "31"
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
      Index           =   31
      Left            =   1305
      TabIndex        =   37
      Tag             =   "StackVal"
      Top             =   2760
      Width           =   135
   End
   Begin VB.Label StackVal1 
      Caption         =   "30"
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
      Index           =   30
      Left            =   1020
      TabIndex        =   36
      Tag             =   "StackVal"
      Top             =   2745
      Width           =   135
   End
   Begin VB.Label StackVal1 
      Caption         =   "29"
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
      Index           =   29
      Left            =   750
      TabIndex        =   35
      Tag             =   "StackVal"
      Top             =   2775
      Width           =   135
   End
   Begin VB.Label StackVal1 
      Caption         =   "28"
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
      Index           =   28
      Left            =   450
      TabIndex        =   34
      Tag             =   "StackVal"
      Top             =   2775
      Width           =   135
   End
   Begin VB.Label StackVal1 
      Caption         =   "27"
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
      Index           =   27
      Left            =   195
      TabIndex        =   33
      Tag             =   "StackVal"
      Top             =   2790
      Width           =   135
   End
   Begin VB.Label StackVal1 
      Caption         =   "26"
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
      Index           =   26
      Left            =   7410
      TabIndex        =   32
      Tag             =   "StackVal"
      Top             =   1680
      Width           =   135
   End
   Begin VB.Label StackVal1 
      Caption         =   "25"
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
      Index           =   25
      Left            =   7140
      TabIndex        =   31
      Tag             =   "StackVal"
      Top             =   1680
      Width           =   135
   End
   Begin VB.Label StackVal1 
      Caption         =   "24"
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
      Index           =   24
      Left            =   6840
      TabIndex        =   30
      Tag             =   "StackVal"
      Top             =   1665
      Width           =   135
   End
   Begin VB.Label StackVal1 
      Caption         =   "23"
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
      Index           =   23
      Left            =   6600
      TabIndex        =   29
      Tag             =   "StackVal"
      Top             =   1650
      Width           =   135
   End
   Begin VB.Label StackVal1 
      Caption         =   "22"
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
      Index           =   22
      Left            =   6300
      TabIndex        =   28
      Tag             =   "StackVal"
      Top             =   1665
      Width           =   135
   End
   Begin VB.Label StackVal1 
      Caption         =   "21"
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
      Index           =   21
      Left            =   6015
      TabIndex        =   27
      Tag             =   "StackVal"
      Top             =   1680
      Width           =   135
   End
   Begin VB.Label StackVal1 
      Caption         =   "20"
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
      Index           =   20
      Left            =   5730
      TabIndex        =   26
      Tag             =   "StackVal"
      Top             =   1665
      Width           =   135
   End
   Begin VB.Label StackVal1 
      Caption         =   "19"
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
      Index           =   19
      Left            =   5460
      TabIndex        =   25
      Tag             =   "StackVal"
      Top             =   1680
      Width           =   135
   End
   Begin VB.Label StackVal1 
      Caption         =   "18"
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
      Index           =   18
      Left            =   5160
      TabIndex        =   24
      Tag             =   "StackVal"
      Top             =   1680
      Width           =   135
   End
   Begin VB.Label StackVal1 
      Caption         =   "17"
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
      Index           =   17
      Left            =   4890
      TabIndex        =   23
      Tag             =   "StackVal"
      Top             =   1680
      Width           =   135
   End
   Begin VB.Label StackVal1 
      Caption         =   "16"
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
      Index           =   16
      Left            =   4590
      TabIndex        =   22
      Tag             =   "StackVal"
      Top             =   1680
      Width           =   135
   End
   Begin VB.Label StackVal1 
      Caption         =   "15"
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
      Index           =   15
      Left            =   4305
      TabIndex        =   21
      Tag             =   "StackVal"
      Top             =   1665
      Width           =   135
   End
   Begin VB.Label StackVal1 
      Caption         =   "14"
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
      Index           =   14
      Left            =   4035
      TabIndex        =   20
      Tag             =   "StackVal"
      Top             =   1650
      Width           =   135
   End
   Begin VB.Label StackVal1 
      Caption         =   "13"
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
      Index           =   13
      Left            =   3600
      TabIndex        =   19
      Tag             =   "StackVal"
      Top             =   1650
      Width           =   135
   End
   Begin VB.Label StackVal1 
      Caption         =   "12"
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
      Index           =   12
      Left            =   3315
      TabIndex        =   18
      Tag             =   "StackVal"
      Top             =   1635
      Width           =   135
   End
   Begin VB.Label StackVal1 
      Caption         =   "11"
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
      Index           =   11
      Left            =   3045
      TabIndex        =   17
      Tag             =   "StackVal"
      Top             =   1665
      Width           =   135
   End
   Begin VB.Label StackVal1 
      Caption         =   "10"
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
      Index           =   10
      Left            =   2760
      TabIndex        =   16
      Tag             =   "StackVal"
      Top             =   1680
      Width           =   135
   End
   Begin VB.Label StackVal1 
      Caption         =   "9"
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
      Index           =   9
      Left            =   2490
      TabIndex        =   15
      Tag             =   "StackVal"
      Top             =   1665
      Width           =   135
   End
   Begin VB.Label StackVal1 
      Caption         =   "8"
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
      Index           =   8
      Left            =   2175
      TabIndex        =   14
      Tag             =   "StackVal"
      Top             =   1665
      Width           =   135
   End
   Begin VB.Label StackVal1 
      Caption         =   "7"
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
      Index           =   7
      Left            =   1890
      TabIndex        =   13
      Tag             =   "StackVal"
      Top             =   1665
      Width           =   135
   End
   Begin VB.Label StackVal1 
      Caption         =   "6"
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
      Index           =   6
      Left            =   1620
      TabIndex        =   12
      Tag             =   "StackVal"
      Top             =   1650
      Width           =   135
   End
   Begin VB.Label StackVal1 
      Caption         =   "5"
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
      Index           =   5
      Left            =   1320
      TabIndex        =   11
      Tag             =   "StackVal"
      Top             =   1665
      Width           =   135
   End
   Begin VB.Label StackVal1 
      Caption         =   "4"
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
      Index           =   4
      Left            =   1035
      TabIndex        =   10
      Tag             =   "StackVal"
      Top             =   1680
      Width           =   135
   End
   Begin VB.Label StackVal1 
      Caption         =   "3"
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
      Index           =   3
      Left            =   750
      TabIndex        =   9
      Tag             =   "StackVal"
      Top             =   1665
      Width           =   135
   End
   Begin VB.Label StackVal1 
      Caption         =   "2"
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
      Index           =   2
      Left            =   465
      TabIndex        =   8
      Tag             =   "StackVal"
      Top             =   1680
      Width           =   135
   End
   Begin VB.Label StackVal1 
      Caption         =   "1"
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
      Index           =   1
      Left            =   195
      TabIndex        =   7
      Tag             =   "StackVal"
      Top             =   1650
      Width           =   135
   End
   Begin VB.Image IndexAH 
      Height          =   750
      Left            =   150
      Picture         =   "frmCustomDeck.frx":1CCA
      Stretch         =   -1  'True
      Tag             =   "Index"
      Top             =   4305
      Width           =   255
   End
   Begin VB.Image Index2H 
      Height          =   750
      Left            =   435
      Picture         =   "frmCustomDeck.frx":29EF
      Stretch         =   -1  'True
      Tag             =   "Index"
      Top             =   4305
      Width           =   255
   End
   Begin VB.Image Index3H 
      Height          =   750
      Left            =   720
      Picture         =   "frmCustomDeck.frx":3788
      Stretch         =   -1  'True
      Tag             =   "Index"
      Top             =   4305
      Width           =   255
   End
   Begin VB.Image Index4H 
      Height          =   750
      Left            =   1005
      Picture         =   "frmCustomDeck.frx":456E
      Stretch         =   -1  'True
      Tag             =   "Index"
      Top             =   4305
      Width           =   255
   End
   Begin VB.Image Index5H 
      Height          =   750
      Left            =   1290
      Picture         =   "frmCustomDeck.frx":51C9
      Stretch         =   -1  'True
      Tag             =   "Index"
      Top             =   4305
      Width           =   255
   End
   Begin VB.Image Index6H 
      Height          =   750
      Left            =   1575
      Picture         =   "frmCustomDeck.frx":5FB6
      Stretch         =   -1  'True
      Tag             =   "Index"
      Top             =   4305
      Width           =   255
   End
   Begin VB.Image Index7H 
      Height          =   750
      Left            =   1860
      Picture         =   "frmCustomDeck.frx":6D82
      Stretch         =   -1  'True
      Tag             =   "Index"
      Top             =   4305
      Width           =   255
   End
   Begin VB.Image Index8H 
      Height          =   750
      Left            =   2145
      Picture         =   "frmCustomDeck.frx":79C4
      Stretch         =   -1  'True
      Tag             =   "Index"
      Top             =   4305
      Width           =   255
   End
   Begin VB.Image Index9H 
      Height          =   750
      Left            =   2430
      Picture         =   "frmCustomDeck.frx":87E3
      Stretch         =   -1  'True
      Tag             =   "Index"
      Top             =   4305
      Width           =   255
   End
   Begin VB.Image Index10H 
      Height          =   750
      Left            =   2715
      Picture         =   "frmCustomDeck.frx":95DA
      Stretch         =   -1  'True
      Tag             =   "Index"
      Top             =   4305
      Width           =   255
   End
   Begin VB.Image IndexJH 
      Height          =   750
      Left            =   3000
      Picture         =   "frmCustomDeck.frx":A58B
      Stretch         =   -1  'True
      Tag             =   "Index"
      Top             =   4305
      Width           =   255
   End
   Begin VB.Image IndexQH 
      Height          =   750
      Left            =   3285
      Picture         =   "frmCustomDeck.frx":B074
      Stretch         =   -1  'True
      Tag             =   "Index"
      Top             =   4305
      Width           =   255
   End
   Begin VB.Image IndexKH 
      Height          =   750
      Left            =   3570
      Picture         =   "frmCustomDeck.frx":C172
      Stretch         =   -1  'True
      Tag             =   "Index"
      Top             =   4305
      Width           =   255
   End
   Begin VB.Image IndexAD 
      Height          =   750
      Left            =   3990
      Picture         =   "frmCustomDeck.frx":CE90
      Stretch         =   -1  'True
      Tag             =   "Index"
      Top             =   4305
      Width           =   255
   End
   Begin VB.Image Index2D 
      Height          =   750
      Left            =   4275
      Picture         =   "frmCustomDeck.frx":DAFF
      Stretch         =   -1  'True
      Tag             =   "Index"
      Top             =   4305
      Width           =   255
   End
   Begin VB.Image Index3D 
      Height          =   750
      Left            =   4560
      Picture         =   "frmCustomDeck.frx":E7D9
      Stretch         =   -1  'True
      Tag             =   "Index"
      Top             =   4305
      Width           =   255
   End
   Begin VB.Image Index4D 
      Height          =   750
      Left            =   4845
      Picture         =   "frmCustomDeck.frx":F4EB
      Stretch         =   -1  'True
      Tag             =   "Index"
      Top             =   4305
      Width           =   255
   End
   Begin VB.Image Index5D 
      Height          =   750
      Left            =   5130
      Picture         =   "frmCustomDeck.frx":10087
      Stretch         =   -1  'True
      Tag             =   "Index"
      Top             =   4305
      Width           =   255
   End
   Begin VB.Image Index6D 
      Height          =   750
      Left            =   5415
      Picture         =   "frmCustomDeck.frx":10DBF
      Stretch         =   -1  'True
      Tag             =   "Index"
      Top             =   4305
      Width           =   255
   End
   Begin VB.Image Index7D 
      Height          =   750
      Left            =   5700
      Picture         =   "frmCustomDeck.frx":11AD5
      Stretch         =   -1  'True
      Tag             =   "Index"
      Top             =   4305
      Width           =   255
   End
   Begin VB.Image Index8D 
      Height          =   750
      Left            =   5970
      Picture         =   "frmCustomDeck.frx":1265D
      Stretch         =   -1  'True
      Tag             =   "Index"
      Top             =   4305
      Width           =   255
   End
   Begin VB.Image Index9D 
      Height          =   750
      Left            =   6270
      Picture         =   "frmCustomDeck.frx":133C7
      Stretch         =   -1  'True
      Tag             =   "Index"
      Top             =   4305
      Width           =   255
   End
   Begin VB.Image Index10D 
      Height          =   750
      Left            =   6555
      Picture         =   "frmCustomDeck.frx":14104
      Stretch         =   -1  'True
      Tag             =   "Index"
      Top             =   4305
      Width           =   255
   End
   Begin VB.Image IndexJD 
      Height          =   750
      Left            =   6840
      Picture         =   "frmCustomDeck.frx":14FFF
      Stretch         =   -1  'True
      Tag             =   "Index"
      Top             =   4305
      Width           =   255
   End
   Begin VB.Image IndexQD 
      Height          =   750
      Left            =   7125
      Picture         =   "frmCustomDeck.frx":15A31
      Stretch         =   -1  'True
      Tag             =   "Index"
      Top             =   4305
      Width           =   255
   End
   Begin VB.Image IndexKD 
      Height          =   750
      Left            =   7410
      Picture         =   "frmCustomDeck.frx":16A7A
      Stretch         =   -1  'True
      Tag             =   "Index"
      Top             =   4305
      Width           =   255
   End
   Begin VB.Image IndexAS 
      Height          =   750
      Left            =   150
      Picture         =   "frmCustomDeck.frx":176E2
      Stretch         =   -1  'True
      Tag             =   "Index"
      Top             =   5400
      Width           =   255
   End
   Begin VB.Image Index2S 
      Height          =   750
      Left            =   435
      Picture         =   "frmCustomDeck.frx":180A1
      Stretch         =   -1  'True
      Tag             =   "Index"
      Top             =   5400
      Width           =   255
   End
   Begin VB.Image Index3S 
      Height          =   750
      Left            =   720
      Picture         =   "frmCustomDeck.frx":18AE6
      Stretch         =   -1  'True
      Tag             =   "Index"
      Top             =   5400
      Width           =   255
   End
   Begin VB.Image Index4S 
      Height          =   750
      Left            =   1005
      Picture         =   "frmCustomDeck.frx":1951C
      Stretch         =   -1  'True
      Tag             =   "Index"
      Top             =   5400
      Width           =   255
   End
   Begin VB.Image Index5S 
      Height          =   750
      Left            =   1290
      Picture         =   "frmCustomDeck.frx":19DF1
      Stretch         =   -1  'True
      Tag             =   "Index"
      Top             =   5400
      Width           =   255
   End
   Begin VB.Image Index6S 
      Height          =   750
      Left            =   1575
      Picture         =   "frmCustomDeck.frx":1A82D
      Stretch         =   -1  'True
      Tag             =   "Index"
      Top             =   5400
      Width           =   255
   End
   Begin VB.Image Index7S 
      Height          =   750
      Left            =   1860
      Picture         =   "frmCustomDeck.frx":1B271
      Stretch         =   -1  'True
      Tag             =   "Index"
      Top             =   5400
      Width           =   255
   End
   Begin VB.Image Index8S 
      Height          =   750
      Left            =   2145
      Picture         =   "frmCustomDeck.frx":1BB95
      Stretch         =   -1  'True
      Tag             =   "Index"
      Top             =   5400
      Width           =   255
   End
   Begin VB.Image Index9S 
      Height          =   750
      Left            =   2430
      Picture         =   "frmCustomDeck.frx":1C602
      Stretch         =   -1  'True
      Tag             =   "Index"
      Top             =   5400
      Width           =   255
   End
   Begin VB.Image Index10S 
      Height          =   750
      Left            =   2715
      Picture         =   "frmCustomDeck.frx":1D076
      Stretch         =   -1  'True
      Tag             =   "Index"
      Top             =   5400
      Width           =   255
   End
   Begin VB.Image IndexJS 
      Height          =   750
      Left            =   3000
      Picture         =   "frmCustomDeck.frx":1DBCA
      Stretch         =   -1  'True
      Tag             =   "Index"
      Top             =   5400
      Width           =   255
   End
   Begin VB.Image IndexQS 
      Height          =   750
      Left            =   3285
      Picture         =   "frmCustomDeck.frx":1E3E1
      Stretch         =   -1  'True
      Tag             =   "Index"
      Top             =   5400
      Width           =   255
   End
   Begin VB.Image IndexKS 
      Height          =   750
      Left            =   3570
      Picture         =   "frmCustomDeck.frx":1F04D
      Stretch         =   -1  'True
      Tag             =   "Index"
      Top             =   5400
      Width           =   255
   End
   Begin VB.Image IndexAC 
      Height          =   750
      Left            =   3975
      Picture         =   "frmCustomDeck.frx":1F9E0
      Stretch         =   -1  'True
      Tag             =   "Index"
      Top             =   5400
      Width           =   255
   End
   Begin VB.Image Index2C 
      Height          =   750
      Left            =   4275
      Picture         =   "frmCustomDeck.frx":20490
      Stretch         =   -1  'True
      Tag             =   "Index"
      Top             =   5400
      Width           =   255
   End
   Begin VB.Image Index3C 
      Height          =   750
      Left            =   4560
      Picture         =   "frmCustomDeck.frx":20FD3
      Stretch         =   -1  'True
      Tag             =   "Index"
      Top             =   5400
      Width           =   255
   End
   Begin VB.Image Index4C 
      Height          =   750
      Left            =   4845
      Picture         =   "frmCustomDeck.frx":21AF7
      Stretch         =   -1  'True
      Tag             =   "Index"
      Top             =   5400
      Width           =   255
   End
   Begin VB.Image Index5C 
      Height          =   750
      Left            =   5130
      Picture         =   "frmCustomDeck.frx":224BD
      Stretch         =   -1  'True
      Tag             =   "Index"
      Top             =   5400
      Width           =   255
   End
   Begin VB.Image Index6C 
      Height          =   750
      Left            =   5415
      Picture         =   "frmCustomDeck.frx":22FEF
      Stretch         =   -1  'True
      Tag             =   "Index"
      Top             =   5400
      Width           =   255
   End
   Begin VB.Image Index7C 
      Height          =   750
      Left            =   5700
      Picture         =   "frmCustomDeck.frx":23B2D
      Stretch         =   -1  'True
      Tag             =   "Index"
      Top             =   5400
      Width           =   255
   End
   Begin VB.Image Index8C 
      Height          =   750
      Left            =   5985
      Picture         =   "frmCustomDeck.frx":24542
      Stretch         =   -1  'True
      Tag             =   "Index"
      Top             =   5400
      Width           =   255
   End
   Begin VB.Image Index9C 
      Height          =   750
      Left            =   6270
      Picture         =   "frmCustomDeck.frx":2509E
      Stretch         =   -1  'True
      Tag             =   "Index"
      Top             =   5400
      Width           =   255
   End
   Begin VB.Image Index10C 
      Height          =   750
      Left            =   6555
      Picture         =   "frmCustomDeck.frx":25C05
      Stretch         =   -1  'True
      Tag             =   "Index"
      Top             =   5400
      Width           =   255
   End
   Begin VB.Image IndexJC 
      Height          =   750
      Left            =   6840
      Picture         =   "frmCustomDeck.frx":26853
      Stretch         =   -1  'True
      Tag             =   "Index"
      Top             =   5400
      Width           =   255
   End
   Begin VB.Image IndexQC 
      Height          =   750
      Left            =   7125
      Picture         =   "frmCustomDeck.frx":27158
      Stretch         =   -1  'True
      Tag             =   "Index"
      Top             =   5400
      Width           =   255
   End
   Begin VB.Image IndexKC 
      Height          =   750
      Left            =   7410
      Picture         =   "frmCustomDeck.frx":27EB3
      Stretch         =   -1  'True
      Tag             =   "Index"
      Top             =   5400
      Width           =   255
   End
   Begin VB.Label StanyonParameters 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3435
      TabIndex        =   6
      Top             =   825
      Visible         =   0   'False
      Width           =   4170
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   165
      X2              =   7650
      Y1              =   3915
      Y2              =   3915
   End
   Begin VB.Image Pos27 
      Height          =   750
      Left            =   150
      Picture         =   "frmCustomDeck.frx":2893C
      Stretch         =   -1  'True
      Top             =   3030
      Width           =   255
   End
   Begin VB.Image Pos28 
      Height          =   750
      Left            =   435
      Picture         =   "frmCustomDeck.frx":28F7D
      Stretch         =   -1  'True
      Top             =   3030
      Width           =   255
   End
   Begin VB.Image Pos29 
      Height          =   750
      Left            =   720
      Picture         =   "frmCustomDeck.frx":295D6
      Stretch         =   -1  'True
      Top             =   3030
      Width           =   255
   End
   Begin VB.Image Pos30 
      Height          =   750
      Left            =   1005
      Picture         =   "frmCustomDeck.frx":29C2D
      Stretch         =   -1  'True
      Top             =   3030
      Width           =   255
   End
   Begin VB.Image Pos31 
      Height          =   750
      Left            =   1290
      Picture         =   "frmCustomDeck.frx":2A250
      Stretch         =   -1  'True
      Top             =   3030
      Width           =   255
   End
   Begin VB.Image Pos32 
      Height          =   750
      Left            =   1575
      Picture         =   "frmCustomDeck.frx":2A8A3
      Stretch         =   -1  'True
      Top             =   3030
      Width           =   255
   End
   Begin VB.Image Pos33 
      Height          =   750
      Left            =   1860
      Picture         =   "frmCustomDeck.frx":2AEFB
      Stretch         =   -1  'True
      Top             =   3030
      Width           =   255
   End
   Begin VB.Image Pos34 
      Height          =   750
      Left            =   2145
      Picture         =   "frmCustomDeck.frx":2B53D
      Stretch         =   -1  'True
      Top             =   3030
      Width           =   255
   End
   Begin VB.Image Pos35 
      Height          =   750
      Left            =   2430
      Picture         =   "frmCustomDeck.frx":2BB8E
      Stretch         =   -1  'True
      Top             =   3030
      Width           =   255
   End
   Begin VB.Image Pos36 
      Height          =   750
      Left            =   2715
      Picture         =   "frmCustomDeck.frx":2C1DE
      Stretch         =   -1  'True
      Top             =   3030
      Width           =   255
   End
   Begin VB.Image Pos37 
      Height          =   750
      Left            =   3000
      Picture         =   "frmCustomDeck.frx":2C82E
      Stretch         =   -1  'True
      Top             =   3030
      Width           =   255
   End
   Begin VB.Image Pos38 
      Height          =   750
      Left            =   3285
      Picture         =   "frmCustomDeck.frx":2CE6A
      Stretch         =   -1  'True
      Top             =   3030
      Width           =   255
   End
   Begin VB.Image Pos39 
      Height          =   750
      Left            =   3570
      Picture         =   "frmCustomDeck.frx":2D48F
      Stretch         =   -1  'True
      Top             =   3030
      Width           =   255
   End
   Begin VB.Image Pos40 
      Height          =   750
      Left            =   3990
      Picture         =   "frmCustomDeck.frx":2DADC
      Stretch         =   -1  'True
      Top             =   3030
      Width           =   255
   End
   Begin VB.Image Pos41 
      Height          =   750
      Left            =   4275
      Picture         =   "frmCustomDeck.frx":2E0B4
      Stretch         =   -1  'True
      Top             =   3030
      Width           =   255
   End
   Begin VB.Image Pos42 
      Height          =   750
      Left            =   4560
      Picture         =   "frmCustomDeck.frx":2E681
      Stretch         =   -1  'True
      Top             =   3030
      Width           =   255
   End
   Begin VB.Image Pos43 
      Height          =   750
      Left            =   4845
      Picture         =   "frmCustomDeck.frx":2EC7B
      Stretch         =   -1  'True
      Top             =   3030
      Width           =   255
   End
   Begin VB.Image Pos44 
      Height          =   750
      Left            =   5130
      Picture         =   "frmCustomDeck.frx":2F278
      Stretch         =   -1  'True
      Top             =   3030
      Width           =   255
   End
   Begin VB.Image Pos45 
      Height          =   750
      Left            =   5415
      Picture         =   "frmCustomDeck.frx":2F860
      Stretch         =   -1  'True
      Top             =   3030
      Width           =   255
   End
   Begin VB.Image Pos46 
      Height          =   750
      Left            =   5700
      Picture         =   "frmCustomDeck.frx":2FE4F
      Stretch         =   -1  'True
      Top             =   3030
      Width           =   255
   End
   Begin VB.Image Pos47 
      Height          =   750
      Left            =   5985
      Picture         =   "frmCustomDeck.frx":3044F
      Stretch         =   -1  'True
      Top             =   3030
      Width           =   255
   End
   Begin VB.Image Pos48 
      Height          =   750
      Left            =   6270
      Picture         =   "frmCustomDeck.frx":30A1D
      Stretch         =   -1  'True
      Top             =   3030
      Width           =   255
   End
   Begin VB.Image Pos49 
      Height          =   750
      Left            =   6555
      Picture         =   "frmCustomDeck.frx":31015
      Stretch         =   -1  'True
      Top             =   3030
      Width           =   255
   End
   Begin VB.Image Pos50 
      Height          =   750
      Left            =   6840
      Picture         =   "frmCustomDeck.frx":31612
      Stretch         =   -1  'True
      Top             =   3030
      Width           =   255
   End
   Begin VB.Image Pos51 
      Height          =   750
      Left            =   7125
      Picture         =   "frmCustomDeck.frx":31C4D
      Stretch         =   -1  'True
      Top             =   3030
      Width           =   255
   End
   Begin VB.Image Pos52 
      Height          =   750
      Left            =   7410
      Picture         =   "frmCustomDeck.frx":322A8
      Stretch         =   -1  'True
      Top             =   3030
      Width           =   255
   End
   Begin VB.Image Pos1 
      Height          =   750
      Left            =   150
      Picture         =   "frmCustomDeck.frx":32916
      Stretch         =   -1  'True
      Top             =   1935
      Width           =   255
   End
   Begin VB.Image Pos2 
      Height          =   750
      Left            =   435
      Picture         =   "frmCustomDeck.frx":32D32
      Stretch         =   -1  'True
      Top             =   1935
      Width           =   255
   End
   Begin VB.Image Pos3 
      Height          =   750
      Left            =   720
      Picture         =   "frmCustomDeck.frx":3315F
      Stretch         =   -1  'True
      Top             =   1935
      Width           =   255
   End
   Begin VB.Image Pos4 
      Height          =   750
      Left            =   1005
      Picture         =   "frmCustomDeck.frx":33588
      Stretch         =   -1  'True
      Top             =   1935
      Width           =   255
   End
   Begin VB.Image Pos5 
      Height          =   750
      Left            =   1290
      Picture         =   "frmCustomDeck.frx":339F9
      Stretch         =   -1  'True
      Top             =   1935
      Width           =   255
   End
   Begin VB.Image Pos6 
      Height          =   750
      Left            =   1575
      Picture         =   "frmCustomDeck.frx":33E34
      Stretch         =   -1  'True
      Top             =   1935
      Width           =   255
   End
   Begin VB.Image Pos7 
      Height          =   750
      Left            =   1860
      Picture         =   "frmCustomDeck.frx":3426C
      Stretch         =   -1  'True
      Top             =   1935
      Width           =   255
   End
   Begin VB.Image Pos8 
      Height          =   750
      Left            =   2145
      Picture         =   "frmCustomDeck.frx":3467D
      Stretch         =   -1  'True
      Top             =   1935
      Width           =   255
   End
   Begin VB.Image Pos9 
      Height          =   750
      Left            =   2430
      Picture         =   "frmCustomDeck.frx":34A81
      Stretch         =   -1  'True
      Top             =   1935
      Width           =   255
   End
   Begin VB.Image Pos10 
      Height          =   750
      Left            =   2715
      Picture         =   "frmCustomDeck.frx":34EB5
      Stretch         =   -1  'True
      Top             =   1935
      Width           =   255
   End
   Begin VB.Image Pos11 
      Height          =   750
      Left            =   3000
      Picture         =   "frmCustomDeck.frx":354B9
      Stretch         =   -1  'True
      Top             =   1935
      Width           =   255
   End
   Begin VB.Image Pos12 
      Height          =   750
      Left            =   3285
      Picture         =   "frmCustomDeck.frx":35A5F
      Stretch         =   -1  'True
      Top             =   1935
      Width           =   255
   End
   Begin VB.Image Pos13 
      Height          =   750
      Left            =   3570
      Picture         =   "frmCustomDeck.frx":360A8
      Stretch         =   -1  'True
      Top             =   1935
      Width           =   255
   End
   Begin VB.Image Pos14 
      Height          =   750
      Left            =   3990
      Picture         =   "frmCustomDeck.frx":366E9
      Stretch         =   -1  'True
      Top             =   1935
      Width           =   255
   End
   Begin VB.Image Pos15 
      Height          =   750
      Left            =   4275
      Picture         =   "frmCustomDeck.frx":36CDB
      Stretch         =   -1  'True
      Top             =   1935
      Width           =   255
   End
   Begin VB.Image Pos16 
      Height          =   750
      Left            =   4560
      Picture         =   "frmCustomDeck.frx":37319
      Stretch         =   -1  'True
      Top             =   1935
      Width           =   255
   End
   Begin VB.Image Pos17 
      Height          =   750
      Left            =   4845
      Picture         =   "frmCustomDeck.frx":37955
      Stretch         =   -1  'True
      Top             =   1935
      Width           =   255
   End
   Begin VB.Image Pos18 
      Height          =   750
      Left            =   5130
      Picture         =   "frmCustomDeck.frx":37F4F
      Stretch         =   -1  'True
      Top             =   1935
      Width           =   255
   End
   Begin VB.Image Pos19 
      Height          =   750
      Left            =   5415
      Picture         =   "frmCustomDeck.frx":3857F
      Stretch         =   -1  'True
      Top             =   1935
      Width           =   255
   End
   Begin VB.Image Pos20 
      Height          =   750
      Left            =   5700
      Picture         =   "frmCustomDeck.frx":38BC5
      Stretch         =   -1  'True
      Top             =   1935
      Width           =   255
   End
   Begin VB.Image Pos21 
      Height          =   750
      Left            =   5985
      Picture         =   "frmCustomDeck.frx":391FB
      Stretch         =   -1  'True
      Top             =   1935
      Width           =   255
   End
   Begin VB.Image Pos22 
      Height          =   750
      Left            =   6270
      Picture         =   "frmCustomDeck.frx":39841
      Stretch         =   -1  'True
      Top             =   1935
      Width           =   255
   End
   Begin VB.Image Pos23 
      Height          =   750
      Left            =   6555
      Picture         =   "frmCustomDeck.frx":39E90
      Stretch         =   -1  'True
      Top             =   1935
      Width           =   255
   End
   Begin VB.Image Pos24 
      Height          =   750
      Left            =   6840
      Picture         =   "frmCustomDeck.frx":3A4E2
      Stretch         =   -1  'True
      Top             =   1935
      Width           =   255
   End
   Begin VB.Image Pos25 
      Height          =   750
      Left            =   7125
      Picture         =   "frmCustomDeck.frx":3AB39
      Stretch         =   -1  'True
      Top             =   1935
      Width           =   255
   End
   Begin VB.Image Pos26 
      Height          =   750
      Left            =   7410
      Picture         =   "frmCustomDeck.frx":3B197
      Stretch         =   -1  'True
      Top             =   1935
      Width           =   255
   End
End
Attribute VB_Name = "frmCustomDeck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_DragDrop(Source As Control, X As Single, Y As Single)
    'if a card index is dropped on the Form (and not on a valid Position spot),
    'the item will be moved back to its original position.
    Source.Move OriginalLeft, OriginalTop
    Source.ZOrder
    AdjustStackVals
End Sub
Private Sub AdjustStackValsSpecial()
'positions the StackVal1(#) objects in order
'this subroutine is only called by Transfer Stack New routine
'because I still want to retain the CustomDeck(3,x) values in case additional
'retained work needs to be done.
    StackVal1(1).Move 210, 1680
    StackVal1(2).Move 495, 1680
    StackVal1(3).Move 780, 1680
    StackVal1(4).Move 1065, 1680
    StackVal1(5).Move 1350, 1680
    StackVal1(6).Move 1635, 1680
    StackVal1(7).Move 1920, 1680
    StackVal1(8).Move 2205, 1680
    StackVal1(9).Move 2490, 1680
    StackVal1(10).Move 2775, 1680
    StackVal1(11).Move 3060, 1680
    StackVal1(12).Move 3345, 1680
    StackVal1(13).Move 3630, 1680
    StackVal1(14).Move 4050, 1680
    StackVal1(15).Move 4335, 1680
    StackVal1(16).Move 4620, 1680
    StackVal1(17).Move 4905, 1680
    StackVal1(18).Move 5190, 1680
    StackVal1(19).Move 5475, 1680
    StackVal1(20).Move 5760, 1680
    StackVal1(21).Move 6045, 1680
    StackVal1(22).Move 6330, 1680
    StackVal1(23).Move 6615, 1680
    StackVal1(24).Move 6900, 1680
    StackVal1(25).Move 7185, 1680
    StackVal1(26).Move 7470, 1680
    StackVal1(27).Move 210, 2775
    StackVal1(28).Move 495, 2775
    StackVal1(29).Move 780, 2775
    StackVal1(30).Move 1065, 2775
    StackVal1(31).Move 1350, 2775
    StackVal1(32).Move 1635, 2775
    StackVal1(33).Move 1920, 2775
    StackVal1(34).Move 2205, 2775
    StackVal1(35).Move 2490, 2775
    StackVal1(36).Move 2775, 2775
    StackVal1(37).Move 3060, 2775
    StackVal1(38).Move 3345, 2775
    StackVal1(39).Move 3630, 2775
    StackVal1(40).Move 4050, 2775
    StackVal1(41).Move 4335, 2775
    StackVal1(42).Move 4620, 2775
    StackVal1(43).Move 4905, 2775
    StackVal1(44).Move 5190, 2775
    StackVal1(45).Move 5475, 2775
    StackVal1(46).Move 5760, 2775
    StackVal1(47).Move 6045, 2775
    StackVal1(48).Move 6330, 2775
    StackVal1(49).Move 6615, 2775
    StackVal1(50).Move 6900, 2775
    StackVal1(51).Move 7185, 2775
    StackVal1(52).Move 7470, 2775
End Sub
Private Sub AdjustStackVals()
'position the StackVal1(#) objects correctly
If ImportedCustomDeck = 0 And CreatedStanyonDeck = 0 Then
    StackVal1(1).Move 210, 1680
    StackVal1(2).Move 495, 1680
    StackVal1(3).Move 780, 1680
    StackVal1(4).Move 1065, 1680
    StackVal1(5).Move 1350, 1680
    StackVal1(6).Move 1635, 1680
    StackVal1(7).Move 1920, 1680
    StackVal1(8).Move 2205, 1680
    StackVal1(9).Move 2490, 1680
    StackVal1(10).Move 2775, 1680
    StackVal1(11).Move 3060, 1680
    StackVal1(12).Move 3345, 1680
    StackVal1(13).Move 3630, 1680
    StackVal1(14).Move 4050, 1680
    StackVal1(15).Move 4335, 1680
    StackVal1(16).Move 4620, 1680
    StackVal1(17).Move 4905, 1680
    StackVal1(18).Move 5190, 1680
    StackVal1(19).Move 5475, 1680
    StackVal1(20).Move 5760, 1680
    StackVal1(21).Move 6045, 1680
    StackVal1(22).Move 6330, 1680
    StackVal1(23).Move 6615, 1680
    StackVal1(24).Move 6900, 1680
    StackVal1(25).Move 7185, 1680
    StackVal1(26).Move 7470, 1680
    StackVal1(27).Move 210, 2775
    StackVal1(28).Move 495, 2775
    StackVal1(29).Move 780, 2775
    StackVal1(30).Move 1065, 2775
    StackVal1(31).Move 1350, 2775
    StackVal1(32).Move 1635, 2775
    StackVal1(33).Move 1920, 2775
    StackVal1(34).Move 2205, 2775
    StackVal1(35).Move 2490, 2775
    StackVal1(36).Move 2775, 2775
    StackVal1(37).Move 3060, 2775
    StackVal1(38).Move 3345, 2775
    StackVal1(39).Move 3630, 2775
    StackVal1(40).Move 4050, 2775
    StackVal1(41).Move 4335, 2775
    StackVal1(42).Move 4620, 2775
    StackVal1(43).Move 4905, 2775
    StackVal1(44).Move 5190, 2775
    StackVal1(45).Move 5475, 2775
    StackVal1(46).Move 5760, 2775
    StackVal1(47).Move 6045, 2775
    StackVal1(48).Move 6330, 2775
    StackVal1(49).Move 6615, 2775
    StackVal1(50).Move 6900, 2775
    StackVal1(51).Move 7185, 2775
    StackVal1(52).Move 7470, 2775
Else
    For m% = 1 To 52
        For Each Ctrl In Controls
            If Ctrl.Tag = "Index" Then
                If CustomDeck(2, m%) = Right(Ctrl.Name, Len(Ctrl.Name) - 5) Then
                    StackVal1(CustomDeck(3, m%)).Move Ctrl.Left + 60, Ctrl.Top - 255
                End If
            End If
        Next Ctrl
    Next m%
End If
End Sub

Private Sub Form_Load()
'clear StackVal positions
ImportedCustomDeck = 0
CreatedStanyonDeck = 0
ReorderRetainedStack.Enabled = False
'set the default retained values
For i% = 1 To 52
    CustomDeck(3, i%) = i%
Next i%
'position the StackVal1(#) objects correctly
AdjustStackVals
End Sub

Public Sub Index2H_DblClick()
    Index2H.Move 435, 4305
    StanyonParameters.Visible = False
    AdjustStackVals
End Sub

Public Sub Index3H_DblClick()
    Index3H.Move 720, 4305
    StanyonParameters.Visible = False
    AdjustStackVals
End Sub
Public Sub Index4H_DblClick()
    Index4H.Move 1005, 4305
    StanyonParameters.Visible = False
    AdjustStackVals
End Sub
Public Sub Index5H_DblClick()
    Index5H.Move 1290, 4305
    StanyonParameters.Visible = False
    AdjustStackVals
End Sub
Public Sub Index6H_DblClick()
    Index6H.Move 1575, 4305
    StanyonParameters.Visible = False
    AdjustStackVals
End Sub
Public Sub Index7H_DblClick()
    Index7H.Move 1860, 4305
    StanyonParameters.Visible = False
    AdjustStackVals
End Sub
Public Sub Index8H_DblClick()
    Index8H.Move 2145, 4305
    StanyonParameters.Visible = False
    AdjustStackVals
End Sub
Public Sub Index9H_DblClick()
    Index9H.Move 2430, 4305
    StanyonParameters.Visible = False
    AdjustStackVals
End Sub
Public Sub Index10H_DblClick()
    Index10H.Move 2715, 4305
    StanyonParameters.Visible = False
    AdjustStackVals
End Sub
Public Sub IndexJH_DblClick()
    IndexJH.Move 3000, 4305
    StanyonParameters.Visible = False
    AdjustStackVals
End Sub
Public Sub IndexQH_DblClick()
    IndexQH.Move 3285, 4305
    StanyonParameters.Visible = False
    AdjustStackVals
End Sub
Public Sub IndexKH_DblClick()
    IndexKH.Move 3570, 4305
    StanyonParameters.Visible = False
    AdjustStackVals
End Sub
Public Sub IndexAH_DblClick()
    IndexAH.Move 150, 4305
    StanyonParameters.Visible = False
    AdjustStackVals
End Sub
Public Sub Index2D_DblClick()
    Index2D.Move 4275, 4305
    StanyonParameters.Visible = False
    AdjustStackVals
End Sub
Public Sub Index3D_DblClick()
    Index3D.Move 4560, 4305
    StanyonParameters.Visible = False
    AdjustStackVals
End Sub
Public Sub Index4D_DblClick()
    Index4D.Move 4845, 4305
    StanyonParameters.Visible = False
    AdjustStackVals
End Sub
Public Sub Index5D_DblClick()
    Index5D.Move 5130, 4305
    StanyonParameters.Visible = False
    AdjustStackVals
End Sub
Public Sub Index6D_DblClick()
    Index6D.Move 5415, 4305
    StanyonParameters.Visible = False
    AdjustStackVals
End Sub
Public Sub Index7D_DblClick()
    Index7D.Move 5700, 4305
    StanyonParameters.Visible = False
    AdjustStackVals
End Sub
Public Sub Index8D_DblClick()
    Index8D.Move 5985, 4305
    StanyonParameters.Visible = False
    AdjustStackVals
End Sub
Public Sub Index9D_DblClick()
    Index9D.Move 6270, 4305
    StanyonParameters.Visible = False
    AdjustStackVals
End Sub
Public Sub Index10D_DblClick()
    Index10D.Move 6555, 4305
    StanyonParameters.Visible = False
    AdjustStackVals
End Sub
Public Sub IndexJD_DblClick()
    IndexJD.Move 6840, 4305
    StanyonParameters.Visible = False
    AdjustStackVals
End Sub
Public Sub IndexQD_DblClick()
    IndexQD.Move 7125, 4305
    StanyonParameters.Visible = False
    AdjustStackVals
End Sub
Public Sub IndexKD_DblClick()
    IndexKD.Move 7410, 4305
    StanyonParameters.Visible = False
    AdjustStackVals
End Sub
Public Sub IndexAD_DblClick()
    IndexAD.Move 3990, 4305
    StanyonParameters.Visible = False
    AdjustStackVals
End Sub
Public Sub Index2S_DblClick()
    Index2S.Move 435, 5400
    StanyonParameters.Visible = False
    AdjustStackVals
End Sub
Public Sub Index3S_DblClick()
    Index3S.Move 720, 5400
    StanyonParameters.Visible = False
    AdjustStackVals
End Sub
Public Sub Index4S_DblClick()
    Index4S.Move 1005, 5400
    StanyonParameters.Visible = False
    AdjustStackVals
End Sub
Public Sub Index5S_DblClick()
    Index5S.Move 1290, 5400
    AdjustStackVals
End Sub
Public Sub Index6S_DblClick()
    Index6S.Move 1575, 5400
    StanyonParameters.Visible = False
    AdjustStackVals
End Sub
Public Sub Index7S_DblClick()
    Index7S.Move 1860, 5400
    StanyonParameters.Visible = False
    AdjustStackVals
End Sub
Public Sub Index8S_DblClick()
    Index8S.Move 2145, 5400
    StanyonParameters.Visible = False
    AdjustStackVals
End Sub
Public Sub Index9S_DblClick()
    Index9S.Move 2430, 5400
    StanyonParameters.Visible = False
    AdjustStackVals
End Sub
Public Sub Index10S_DblClick()
    Index10S.Move 2715, 5400
    StanyonParameters.Visible = False
    AdjustStackVals
End Sub
Public Sub IndexJS_DblClick()
    IndexJS.Move 3000, 5400
    StanyonParameters.Visible = False
    AdjustStackVals
End Sub
Public Sub IndexQS_DblClick()
    IndexQS.Move 3285, 5400
    StanyonParameters.Visible = False
    AdjustStackVals
End Sub
Public Sub IndexKS_DblClick()
    IndexKS.Move 3570, 5400
    StanyonParameters.Visible = False
    AdjustStackVals
End Sub
Public Sub IndexAS_DblClick()
    IndexAS.Move 150, 5400
    StanyonParameters.Visible = False
    AdjustStackVals
End Sub
Public Sub Index2C_DblClick()
    Index2C.Move 4275, 5400
    StanyonParameters.Visible = False
    AdjustStackVals
End Sub
Public Sub Index3C_DblClick()
    Index3C.Move 4560, 5400
    StanyonParameters.Visible = False
    AdjustStackVals
End Sub
Public Sub Index4C_DblClick()
    Index4C.Move 4845, 5400
    StanyonParameters.Visible = False
    AdjustStackVals
End Sub
Public Sub Index5C_DblClick()
    Index5C.Move 5130, 5400
    StanyonParameters.Visible = False
    AdjustStackVals
End Sub
Public Sub Index6C_DblClick()
    Index6C.Move 5415, 5400
    StanyonParameters.Visible = False
    AdjustStackVals
End Sub
Public Sub Index7C_DblClick()
    Index7C.Move 5700, 5400
    StanyonParameters.Visible = False
    AdjustStackVals
End Sub
Public Sub Index8C_DblClick()
    Index8C.Move 5985, 5400
    StanyonParameters.Visible = False
    AdjustStackVals
End Sub
Public Sub Index9C_DblClick()
    Index9C.Move 6270, 5400
    StanyonParameters.Visible = False
    AdjustStackVals
End Sub
Public Sub Index10C_DblClick()
    Index10C.Move 6555, 5400
    StanyonParameters.Visible = False
    AdjustStackVals
End Sub
Public Sub IndexJC_DblClick()
    IndexJC.Move 6840, 5400
    StanyonParameters.Visible = False
    AdjustStackVals
End Sub
Public Sub IndexQC_DblClick()
    IndexQC.Move 7125, 5400
    StanyonParameters.Visible = False
    AdjustStackVals
End Sub
Public Sub IndexKC_DblClick()
    IndexKC.Move 7410, 5400
    StanyonParameters.Visible = False
    AdjustStackVals
End Sub
Public Sub IndexAC_DblClick()
    IndexAC.Move 3990, 5400
    StanyonParameters.Visible = False
    AdjustStackVals
End Sub

Public Sub Index2H_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Index2H.Drag 1
    OriginalTop = 4305
    OriginalLeft = 435
    StanyonParameters.Visible = False
End Sub

Public Sub Index3H_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Index3H.Drag 1
    OriginalTop = 4305
    OriginalLeft = 720
    StanyonParameters.Visible = False
End Sub
Public Sub Index4H_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Index4H.Drag 1
    OriginalTop = 4305
    OriginalLeft = 1005
    StanyonParameters.Visible = False
End Sub
Public Sub Index5H_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Index5H.Drag 1
    OriginalTop = 4305
    OriginalLeft = 1290
    StanyonParameters.Visible = False
End Sub
Public Sub Index6H_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Index6H.Drag 1
    OriginalTop = 4305
    OriginalLeft = 1575
    StanyonParameters.Visible = False
End Sub
Public Sub Index7H_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Index7H.Drag 1
    OriginalTop = 4305
    OriginalLeft = 1860
    StanyonParameters.Visible = False
End Sub
Public Sub Index8H_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Index8H.Drag 1
    OriginalTop = 4305
    OriginalLeft = 2145
    StanyonParameters.Visible = False
End Sub
Public Sub Index9H_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Index9H.Drag 1
    OriginalTop = 4305
    OriginalLeft = 2430
    StanyonParameters.Visible = False
End Sub
Public Sub Index10H_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Index10H.Drag 1
    OriginalTop = 4305
    OriginalLeft = 2715
    StanyonParameters.Visible = False
End Sub
Public Sub IndexJH_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    IndexJH.Drag 1
    OriginalTop = 4305
    OriginalLeft = 3000
    StanyonParameters.Visible = False
End Sub
Public Sub IndexQH_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    IndexQH.Drag 1
    OriginalTop = 4305
    OriginalLeft = 3285
    StanyonParameters.Visible = False
End Sub
Public Sub IndexKH_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    IndexKH.Drag 1
    OriginalTop = 4305
    OriginalLeft = 3570
    StanyonParameters.Visible = False
End Sub

Public Sub IndexAH_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    IndexAH.Drag 1
    OriginalTop = 4305
    OriginalLeft = 150
    StanyonParameters.Visible = False
End Sub
Public Sub Index2D_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Index2D.Drag 1
    OriginalTop = 4305
    OriginalLeft = 4275
    StanyonParameters.Visible = False
End Sub
Public Sub Index3D_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Index3D.Drag 1
    OriginalTop = 4305
    OriginalLeft = 4560
    StanyonParameters.Visible = False
End Sub
Public Sub Index4D_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Index4D.Drag 1
    OriginalTop = 4305
    OriginalLeft = 4845
    StanyonParameters.Visible = False
End Sub
Public Sub Index5D_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Index5D.Drag 1
    OriginalTop = 4305
    OriginalLeft = 5130
    StanyonParameters.Visible = False
End Sub
Public Sub Index6D_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Index6D.Drag 1
    OriginalTop = 4305
    OriginalLeft = 5415
    StanyonParameters.Visible = False
End Sub
Public Sub Index7D_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Index7D.Drag 1
    OriginalTop = 4305
    OriginalLeft = 5700
    StanyonParameters.Visible = False
End Sub
Public Sub Index8D_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Index8D.Drag 1
    OriginalTop = 4305
    OriginalLeft = 5985
    StanyonParameters.Visible = False
End Sub
Public Sub Index9D_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Index9D.Drag 1
    OriginalTop = 4305
    OriginalLeft = 6270
    StanyonParameters.Visible = False
End Sub
Public Sub Index10D_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Index10D.Drag 1
    OriginalTop = 4305
    OriginalLeft = 6555
    StanyonParameters.Visible = False
End Sub
Public Sub IndexJD_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    IndexJD.Drag 1
    OriginalTop = 4305
    OriginalLeft = 6840
    StanyonParameters.Visible = False
End Sub
Public Sub IndexQD_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    IndexQD.Drag 1
    OriginalTop = 4305
    OriginalLeft = 7125
    StanyonParameters.Visible = False
End Sub
Public Sub IndexKD_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    IndexKD.Drag 1
    OriginalTop = 4305
    OriginalLeft = 7410
    StanyonParameters.Visible = False
End Sub

Public Sub IndexAD_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    IndexAD.Drag 1
    OriginalTop = 4305
    OriginalLeft = 3990
    StanyonParameters.Visible = False
End Sub
Public Sub Index2S_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Index2S.Drag 1
    OriginalTop = 5400
    OriginalLeft = 435
    StanyonParameters.Visible = False
End Sub
Public Sub Index3S_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Index3S.Drag 1
    OriginalTop = 5400
    OriginalLeft = 720
    StanyonParameters.Visible = False
End Sub
Public Sub Index4S_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Index4S.Drag 1
    OriginalTop = 5400
    OriginalLeft = 1005
    StanyonParameters.Visible = False
End Sub
Public Sub Index5S_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Index5S.Drag 1
    OriginalTop = 5400
    OriginalLeft = 1290
    StanyonParameters.Visible = False
End Sub
Public Sub Index6S_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Index6S.Drag 1
    OriginalTop = 5400
    OriginalLeft = 1575
    StanyonParameters.Visible = False
End Sub
Public Sub Index7S_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Index7S.Drag 1
    OriginalTop = 5400
    OriginalLeft = 1860
    StanyonParameters.Visible = False
End Sub
Public Sub Index8S_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Index8S.Drag 1
    OriginalTop = 5400
    OriginalLeft = 2145
    StanyonParameters.Visible = False
End Sub
Public Sub Index9S_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Index9S.Drag 1
    OriginalTop = 5400
    OriginalLeft = 2430
    StanyonParameters.Visible = False
End Sub
Public Sub Index10S_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Index10S.Drag 1
    OriginalTop = 5400
    OriginalLeft = 2715
    StanyonParameters.Visible = False
End Sub
Public Sub IndexJS_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    IndexJS.Drag 1
    OriginalTop = 5400
    OriginalLeft = 3000
    StanyonParameters.Visible = False
End Sub
Public Sub IndexQS_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    IndexQS.Drag 1
    OriginalTop = 5400
    OriginalLeft = 3285
    StanyonParameters.Visible = False
End Sub
Public Sub IndexKS_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    IndexKS.Drag 1
    OriginalTop = 5400
    OriginalLeft = 3570
    StanyonParameters.Visible = False
End Sub

Public Sub IndexAS_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    IndexAS.Drag 1
    OriginalTop = 5400
    OriginalLeft = 150
    StanyonParameters.Visible = False
End Sub
Public Sub Index2C_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Index2C.Drag 1
    OriginalTop = 5400
    OriginalLeft = 4275
    StanyonParameters.Visible = False
End Sub
Public Sub Index3C_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Index3C.Drag 1
    OriginalTop = 5400
    OriginalLeft = 4560
    StanyonParameters.Visible = False
End Sub
Public Sub Index4C_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Index4C.Drag 1
    OriginalTop = 5400
    OriginalLeft = 4845
    StanyonParameters.Visible = False
End Sub
Public Sub Index5C_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Index5C.Drag 1
    OriginalTop = 5400
    OriginalLeft = 5130
    StanyonParameters.Visible = False
End Sub
Public Sub Index6C_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Index6C.Drag 1
    OriginalTop = 5400
    OriginalLeft = 5415
    StanyonParameters.Visible = False
End Sub
Public Sub Index7C_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Index7C.Drag 1
    OriginalTop = 5400
    OriginalLeft = 5700
    StanyonParameters.Visible = False
End Sub
Public Sub Index8C_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Index8C.Drag 1
    OriginalTop = 5400
    OriginalLeft = 5985
    StanyonParameters.Visible = False
End Sub
Public Sub Index9C_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Index9C.Drag 1
    OriginalTop = 5400
    OriginalLeft = 6270
    StanyonParameters.Visible = False
End Sub
Public Sub Index10C_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Index10C.Drag 1
    OriginalTop = 5400
    OriginalLeft = 6555
    StanyonParameters.Visible = False
End Sub
Public Sub IndexJC_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    IndexJC.Drag 1
    OriginalTop = 5400
    OriginalLeft = 6840
    StanyonParameters.Visible = False
End Sub
Public Sub IndexQC_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    IndexQC.Drag 1
    OriginalTop = 5400
    OriginalLeft = 7125
    StanyonParameters.Visible = False
End Sub
Public Sub IndexKC_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    IndexKC.Drag 1
    OriginalTop = 5400
    OriginalLeft = 7410
    StanyonParameters.Visible = False
End Sub

Public Sub IndexAC_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    IndexAC.Drag 1
    OriginalTop = 5400
    OriginalLeft = 3990
    StanyonParameters.Visible = False
End Sub
Public Sub CheckStanyonParameters()
'this routine checks that each of the first five "CustomDeck(2,x)" entries
'correspond to actual matching card indexes in those positions.
'if each position is correct, a binary unit is added to "StanyonParameterError"
'this error variable reads from left to right:
'1 corresponds to the first position
'2 corresponds to the second position
'4 corresponds to the third position
'8 corresponds to the fourth position
'16 corresponds to the fifth position
'when the sum is anything but 31, there is an error
StanyonParameterError = 0
For i% = 1 To 5
    For Each Ctrl In Controls
        If Ctrl.Tag = "Index" Then
            If CustomDeck(2, i%) = Right(Ctrl.Name, Len(Ctrl.Name) - 5) And _
                Ctrl.Left = (150 + (i% - 1) * 285) And _
                Ctrl.Top = 1935 Then
                    StanyonParameterError = StanyonParameterError + 2 ^ (i% - 1)
            End If
        End If
    Next Ctrl
Next i%
End Sub
Public Sub CheckSuitParameters()
'this subroutine assumes that "CheckStanyonParamerers()" did not generate
'an error. (i.e. the error free sum for StanyonParameterCheck = 31
'this routine checks that the first four cards in the stack contain
'each of the four suits
StanyonSuitError = 0
FFS = Empty
For i% = 1 To 4
    FFS = FFS & Right(CustomDeck(2, i%), 1)
Next i%
StanyonSuitError = InStr(FFS, "C") + InStr(FFS, "H") + _
    InStr(FFS, "S") + InStr(FFS, "D")
' the last statement reports the position of each C, H, S, & D.
' If each suit appears only once, then the sum should equal to 10.
' If the sum is less than 10, then there is an entry error.
End Sub
Public Sub CheckCycleParameter()
'this subroutine assumes that "CheckStanyonParamerers()" did not generate
'an error. (i.e. the error free sum for StanyonParameterCheck = 31
'this routine checks that the value of the first and fifth card are not the same.
'if there is no error, StanyonCycleError=0
'if there is an error, StanyonCycleError=1
StanyonCycleError = 0
    For Each Ctrl In Controls
        If Ctrl.Tag = "Index" Then
            If CustomDeck(2, 1) = Right(Ctrl.Name, Len(Ctrl.Name) - 5) And _
                Ctrl.Left = 150 And _
                Ctrl.Top = 1935 Then
                    c1$ = Left(CustomDeck(2, 1), Len(CustomDeck(2, 1)) - 1)
            End If
            If CustomDeck(2, 5) = Right(Ctrl.Name, Len(Ctrl.Name) - 5) And _
                Ctrl.Left = 1290 And _
                Ctrl.Top = 1935 Then
                    c2$ = Left(CustomDeck(2, 5), Len(CustomDeck(2, 5)) - 1)
            End If
        End If
    Next Ctrl
    If c1$ = c2$ Then
        StanyonCycleError = 1
    Else
        StanyonCycleError = 0
    End If
End Sub
Public Sub CreateStanyonParameters()
For i% = 1 To 4
    r1 = Left(CustomDeck(2, i%), Len(CustomDeck(2, i%)) - 1)
    If r1 = "J" Then
        cv1 = 11
    ElseIf r1 = "Q" Then
        cv1 = 12
    ElseIf r1 = "K" Then
        cv1 = 13
    ElseIf r1 = "A" Then
        cv1 = 1
    Else
        cv1 = Val(r1)
    End If
    r2 = Left(CustomDeck(2, i% + 1), Len(CustomDeck(2, i% + 1)) - 1)
    If r2 = "J" Then
        cv2 = 11
    ElseIf r2 = "Q" Then
        cv2 = 12
    ElseIf r2 = "K" Then
        cv2 = 13
    ElseIf r2 = "A" Then
        cv2 = 1
    Else
        cv2 = Val(r2)
    End If
    If cv2 - cv1 < 0 Then
        NextCardIncrements(i%) = ((cv2 - cv1 - 1 + 13) Mod 13) + 1
        'the Mod function does not handle negative values correctly
    Else
        NextCardIncrements(i%) = ((cv2 - cv1 - 1) Mod 13) + 1
    End If
    SuitOrder(i%) = Right(CustomDeck(2, i%), 1)
Next i%
For j% = 1 To 4
    If NextCardIncrements(j%) > 6 Then
        NextCardIncrementsReport(j%) = NextCardIncrements(j%) - 13
    Else
        NextCardIncrementsReport(j%) = NextCardIncrements(j%)
    End If
Next j%
End Sub

Public Sub CreateStanyonVariationDeck()
StanyonStartingCard = CustomDeck(2, 1)
StanyonVariationDeck(1, 1) = 1
StanyonVariationDeck(2, 1) = StanyonStartingCard
p = Left(StanyonStartingCard, Len(StanyonStartingCard) - 1)
'strip off the card index value as text, but not the suit
q = Right(StanyonStartingCard, 1)
'strip off the suit of the starting card
For j% = 1 To 4
    If SuitOrder(j%) = q Then
        SuitPointerOffset = j%
    End If
Next j%
'figures out where in the suit cycle the starting card starts
If p = "J" Then
    CardValue = 11
ElseIf p = "Q" Then
    CardValue = 12
ElseIf p = "K" Then
    CardValue = 13
ElseIf p = "A" Then
    CardValue = 1
Else
    CardValue = Val(p)
End If
'convert the card value from text to integer
For i% = 2 To 52
    StanyonVariationDeck(1, i%) = i%
    SuitPointer = ((i% - 1 + SuitPointerOffset - 1) Mod 4) + 1
    IncrementPointer = ((i% - 2) Mod 4) + 1
    CardValue = ((CardValue + NextCardIncrements(IncrementPointer) - 1) Mod 13) + 1
    If CardValue = 13 Then
        CardText = "K"
    ElseIf CardValue = 12 Then
        CardText = "Q"
    ElseIf CardValue = 11 Then
        CardText = "J"
    ElseIf CardValue = 1 Then
        CardText = "A"
    Else
        CardText = Right(Str(CardValue), Len(Str(CardValue)) - 1)
        'when a string is converted from a value, an extra space is added
        'to the front for the sign, which needs to be stripped off
    End If
    StanyonVariationDeck(2, i%) = CardText & SuitOrder(SuitPointer)
    'this next segment tracks selections and reversals
Next i%
For m% = 1 To 52
    For z% = 1 To 2
        CustomDeck(z%, m%) = StanyonVariationDeck(z%, m%)
    Next z%
    If ImportedCustomDeck = 0 Then
        CustomDeck(3, m%) = m%
        'new section
        'in case a person creates a Stanyon deck from the New state
        'and then starts to reposition the cards before retaining the Stack Values,
        'establish a phantom CustomDeckAsImported state.
        CustomDeckAsImported(1, m%) = CustomDeck(1, m%)
        CustomDeckAsImported(2, m%) = CustomDeck(2, m%)
        'end new section
    ElseIf ImportedCustomDeck = 1 Then
        For k% = 1 To 52
            If CustomDeckAsImported(2, k%) = StanyonVariationDeck(2, m%) Then
                CustomDeck(3, m%) = CustomDeckAsImported(1, k%)
                CustomDeck(4, m%) = CustomDeckAsImported(4, k%)
                CustomDeck(6, m%) = CustomDeckAsImported(6, k%)
            End If
        Next k%
    End If
Next m%
End Sub
Public Sub ImportCurrentDeck_Click()
For m% = 1 To 52
    CustomDeck(1, m%) = m%
    CustomDeck(2, m%) = Deck(2, m%)
    CustomDeck(3, m%) = Deck(1, m%)
    CustomDeck(4, m%) = Deck(4, m%)
    CustomDeck(6, m%) = Deck(6, m%)
    CustomDeckAsImported(1, m%) = Deck(1, m%)
    CustomDeckAsImported(2, m%) = Deck(2, m%)
    CustomDeckAsImported(4, m%) = Deck(4, m%)
    CustomDeckAsImported(6, m%) = Deck(6, m%)
Next m%
ImportedCustomDeck = 1
CreatedStanyonDeck = 0
ReorderRetainedStack.Enabled = True
StanyonParameters.Visible = False
PositionStanyonVariationIndexes
'the prior call positions the card indexes to the custom deck space.  It does
'not have to be a Stanyon variation deck
AdjustStackVals
End Sub


Private Sub ReorderRetainedStack_Click()
'this procedure is only called from an enabled button state
ResetCustomStackRetain_Click
For i% = 1 To 13
    For Each Ctrl In Controls
        If Ctrl.Tag = "Index" Then
            If CustomDeck(2, i%) = Right(Ctrl.Name, Len(Ctrl.Name) - 5) Then
                Ctrl.Move (150 + (i% - 1) * 285), 1935
            End If
        End If
    Next Ctrl
Next i%
For i% = 14 To 26
    For Each Ctrl In Controls
        If Ctrl.Tag = "Index" Then
            If CustomDeck(2, i%) = Right(Ctrl.Name, Len(Ctrl.Name) - 5) Then
                Ctrl.Move (3990 + (i% - 14) * 285), 1935
            End If
        End If
    Next Ctrl
Next i%
For i% = 27 To 39
    For Each Ctrl In Controls
        If Ctrl.Tag = "Index" Then
            If CustomDeck(2, i%) = Right(Ctrl.Name, Len(Ctrl.Name) - 5) Then
                Ctrl.Move (150 + (i% - 27) * 285), 3030
            End If
        End If
    Next Ctrl
Next i%
For i% = 40 To 52
    For Each Ctrl In Controls
        If Ctrl.Tag = "Index" Then
            If CustomDeck(2, i%) = Right(Ctrl.Name, Len(Ctrl.Name) - 5) Then
                Ctrl.Move (3990 + (i% - 40) * 285), 3030
            End If
        End If
    Next Ctrl
Next i%
AdjustStackVals
End Sub

Public Sub ResetCustomStackRetain_Click()
'this sub will not lose the original position values associated with the cards.
'new section to fix the retained anomaly
Dim pCardsLeftLow As Boolean
pCardsLeftLow = False
For Each Ctrl In Controls
    If Ctrl.Tag = "Index" Then
        If Ctrl.Top > 4000 Then
            pCardsLeftLow = True
        End If
    End If
Next Ctrl
If pCardsLeftLow And ImportedCustomDeck = 0 And CreatedStanyonDeck = 0 Then
    MsgBox ("Since you are starting from a New Custom Deck state," & Chr(13) & _
        "you must complete positioning all the cards above the dividing line" & Chr(13) _
        & "before you can Retain the Stack Values of your custom stack." & Chr(13) & Chr(13) & _
        "You may do a partial Reset: Retain only after Importing a stack" & Chr(13) & _
        "or after Creating a Stanyon Variation deck.")
    Exit Sub
ElseIf ImportedCustomDeck = 0 Then
'ElseIf ImportedCustomDeck = 0 And CreatedStanyonDeck = 0 Then
    For m% = 1 To 52
        CustomDeck(3, m%) = CustomDeck(1, m%)
        CustomDeckAsImported(1, m%) = CustomDeck(1, m%)
        CustomDeckAsImported(2, m%) = CustomDeck(2, m%)
    Next m%
End If
'end of new section
    IndexAH.Move 150, 4305
    Index2H.Move 435, 4305
    Index3H.Move 720, 4305
    Index4H.Move 1005, 4305
    Index5H.Move 1290, 4305
    Index6H.Move 1575, 4305
    Index7H.Move 1860, 4305
    Index8H.Move 2145, 4305
    Index9H.Move 2430, 4305
    Index10H.Move 2715, 4305
    IndexJH.Move 3000, 4305
    IndexQH.Move 3285, 4305
    IndexKH.Move 3570, 4305
    IndexAD.Move 3990, 4305
    Index2D.Move 4275, 4305
    Index3D.Move 4560, 4305
    Index4D.Move 4845, 4305
    Index5D.Move 5130, 4305
    Index6D.Move 5415, 4305
    Index7D.Move 5700, 4305
    Index8D.Move 5985, 4305
    Index9D.Move 6270, 4305
    Index10D.Move 6555, 4305
    IndexJD.Move 6840, 4305
    IndexQD.Move 7125, 4305
    IndexKD.Move 7410, 4305
    IndexAS.Move 150, 5400
    Index2S.Move 435, 5400
    Index3S.Move 720, 5400
    Index4S.Move 1005, 5400
    Index5S.Move 1290, 5400
    Index6S.Move 1575, 5400
    Index7S.Move 1860, 5400
    Index8S.Move 2145, 5400
    Index9S.Move 2430, 5400
    Index10S.Move 2715, 5400
    IndexJS.Move 3000, 5400
    IndexQS.Move 3285, 5400
    IndexKS.Move 3570, 5400
    IndexAC.Move 3990, 5400
    Index2C.Move 4275, 5400
    Index3C.Move 4560, 5400
    Index4C.Move 4845, 5400
    Index5C.Move 5130, 5400
    Index6C.Move 5415, 5400
    Index7C.Move 5700, 5400
    Index8C.Move 5985, 5400
    Index9C.Move 6270, 5400
    Index10C.Move 6555, 5400
    IndexJC.Move 6840, 5400
    IndexQC.Move 7125, 5400
    IndexKC.Move 7410, 5400
    'new section
    For i% = 1 To 52
        CustomDeck(1, i%) = CustomDeckAsImported(1, i%)
        CustomDeck(2, i%) = CustomDeckAsImported(2, i%)
        CustomDeck(3, i%) = CustomDeckAsImported(1, i%)
    Next i%
    'end of new section
    StanyonParameters.Visible = False
    ImportedCustomDeck = 1
    CreatedStanyonDeck = 0
    ReorderRetainedStack.Enabled = True
    AdjustStackVals
End Sub
Public Sub Pos1_DragDrop(Source As Control, X As Single, Y As Single)
    Source.Move 150, 1935
    CustomDeck(1, 1) = 1
    CustomDeck(2, 1) = Right(Source.Name, Len(Source.Name) - 5)
    For i% = 1 To 52
        If CustomDeckAsImported(2, i%) = Right(Source.Name, Len(Source.Name) - 5) Then
            CustomDeck(3, 1) = CustomDeckAsImported(1, i%)
            CustomDeck(4, 1) = CustomDeckAsImported(4, i%)
            CustomDeck(6, 1) = CustomDeckAsImported(6, i%)
        End If
    Next i%
    AdjustStackVals
End Sub

Public Sub Pos2_DragDrop(Source As Control, X As Single, Y As Single)
    Source.Move 435, 1935
    CustomDeck(1, 2) = 2
    CustomDeck(2, 2) = Right(Source.Name, Len(Source.Name) - 5)
    For i% = 1 To 52
        If CustomDeckAsImported(2, i%) = Right(Source.Name, Len(Source.Name) - 5) Then
            CustomDeck(3, 2) = CustomDeckAsImported(1, i%)
            CustomDeck(4, 2) = CustomDeckAsImported(4, i%)
            CustomDeck(6, 2) = CustomDeckAsImported(6, i%)
        End If
    Next i%
    AdjustStackVals
End Sub
Public Sub Pos3_DragDrop(Source As Control, X As Single, Y As Single)
    Source.Move 720, 1935
    CustomDeck(1, 3) = 3
    CustomDeck(2, 3) = Right(Source.Name, Len(Source.Name) - 5)
    For i% = 1 To 52
        If CustomDeckAsImported(2, i%) = Right(Source.Name, Len(Source.Name) - 5) Then
            CustomDeck(3, 3) = CustomDeckAsImported(1, i%)
            CustomDeck(4, 3) = CustomDeckAsImported(4, i%)
            CustomDeck(6, 3) = CustomDeckAsImported(6, i%)
        End If
    Next i%
    AdjustStackVals
End Sub
Public Sub Pos4_DragDrop(Source As Control, X As Single, Y As Single)
    Source.Move 1005, 1935
    CustomDeck(1, 4) = 4
    CustomDeck(2, 4) = Right(Source.Name, Len(Source.Name) - 5)
    For i% = 1 To 52
        If CustomDeckAsImported(2, i%) = Right(Source.Name, Len(Source.Name) - 5) Then
            CustomDeck(3, 4) = CustomDeckAsImported(1, i%)
            CustomDeck(4, 4) = CustomDeckAsImported(4, i%)
            CustomDeck(6, 4) = CustomDeckAsImported(6, i%)
        End If
    Next i%
    AdjustStackVals
End Sub
Public Sub Pos5_DragDrop(Source As Control, X As Single, Y As Single)
    Source.Move 1290, 1935
    CustomDeck(1, 5) = 5
    CustomDeck(2, 5) = Right(Source.Name, Len(Source.Name) - 5)
    For i% = 1 To 52
        If CustomDeckAsImported(2, i%) = Right(Source.Name, Len(Source.Name) - 5) Then
            CustomDeck(3, 5) = CustomDeckAsImported(1, i%)
            CustomDeck(4, 5) = CustomDeckAsImported(4, i%)
            CustomDeck(6, 5) = CustomDeckAsImported(6, i%)
        End If
    Next i%
    AdjustStackVals
End Sub
Public Sub Pos6_DragDrop(Source As Control, X As Single, Y As Single)
    Source.Move 1575, 1935
    CustomDeck(1, 6) = 6
    CustomDeck(2, 6) = Right(Source.Name, Len(Source.Name) - 5)
    For i% = 1 To 52
        If CustomDeckAsImported(2, i%) = Right(Source.Name, Len(Source.Name) - 5) Then
            CustomDeck(3, 6) = CustomDeckAsImported(1, i%)
            CustomDeck(4, 6) = CustomDeckAsImported(4, i%)
            CustomDeck(6, 6) = CustomDeckAsImported(6, i%)
        End If
    Next i%
    AdjustStackVals
End Sub
Public Sub Pos7_DragDrop(Source As Control, X As Single, Y As Single)
    Source.Move 1860, 1935
    CustomDeck(1, 7) = 7
    CustomDeck(2, 7) = Right(Source.Name, Len(Source.Name) - 5)
    For i% = 1 To 52
        If CustomDeckAsImported(2, i%) = Right(Source.Name, Len(Source.Name) - 5) Then
            CustomDeck(3, 7) = CustomDeckAsImported(1, i%)
            CustomDeck(4, 7) = CustomDeckAsImported(4, i%)
            CustomDeck(6, 7) = CustomDeckAsImported(6, i%)
        End If
    Next i%
    AdjustStackVals
End Sub
Public Sub Pos8_DragDrop(Source As Control, X As Single, Y As Single)
    Source.Move 2145, 1935
    CustomDeck(1, 8) = 8
    CustomDeck(2, 8) = Right(Source.Name, Len(Source.Name) - 5)
    For i% = 1 To 52
        If CustomDeckAsImported(2, i%) = Right(Source.Name, Len(Source.Name) - 5) Then
            CustomDeck(3, 8) = CustomDeckAsImported(1, i%)
            CustomDeck(4, 8) = CustomDeckAsImported(4, i%)
            CustomDeck(6, 8) = CustomDeckAsImported(6, i%)
        End If
    Next i%
    AdjustStackVals
End Sub
Public Sub Pos9_DragDrop(Source As Control, X As Single, Y As Single)
    Source.Move 2430, 1935
    CustomDeck(1, 9) = 9
    CustomDeck(2, 9) = Right(Source.Name, Len(Source.Name) - 5)
    For i% = 1 To 52
        If CustomDeckAsImported(2, i%) = Right(Source.Name, Len(Source.Name) - 5) Then
            CustomDeck(3, 9) = CustomDeckAsImported(1, i%)
            CustomDeck(4, 9) = CustomDeckAsImported(4, i%)
            CustomDeck(6, 9) = CustomDeckAsImported(6, i%)
        End If
    Next i%
    AdjustStackVals
End Sub
Public Sub Pos10_DragDrop(Source As Control, X As Single, Y As Single)
    Source.Move 2715, 1935
    CustomDeck(1, 10) = 10
    CustomDeck(2, 10) = Right(Source.Name, Len(Source.Name) - 5)
    For i% = 1 To 52
        If CustomDeckAsImported(2, i%) = Right(Source.Name, Len(Source.Name) - 5) Then
            CustomDeck(3, 10) = CustomDeckAsImported(1, i%)
            CustomDeck(4, 10) = CustomDeckAsImported(4, i%)
            CustomDeck(6, 10) = CustomDeckAsImported(6, i%)
        End If
    Next i%
    AdjustStackVals
End Sub
Public Sub Pos11_DragDrop(Source As Control, X As Single, Y As Single)
    Source.Move 3000, 1935
    CustomDeck(1, 11) = 11
    CustomDeck(2, 11) = Right(Source.Name, Len(Source.Name) - 5)
    For i% = 1 To 52
        If CustomDeckAsImported(2, i%) = Right(Source.Name, Len(Source.Name) - 5) Then
            CustomDeck(3, 11) = CustomDeckAsImported(1, i%)
            CustomDeck(4, 11) = CustomDeckAsImported(4, i%)
            CustomDeck(6, 11) = CustomDeckAsImported(6, i%)
        End If
    Next i%
    AdjustStackVals
End Sub
Public Sub Pos12_DragDrop(Source As Control, X As Single, Y As Single)
    Source.Move 3285, 1935
    CustomDeck(1, 12) = 12
    CustomDeck(2, 12) = Right(Source.Name, Len(Source.Name) - 5)
    For i% = 1 To 52
        If CustomDeckAsImported(2, i%) = Right(Source.Name, Len(Source.Name) - 5) Then
            CustomDeck(3, 12) = CustomDeckAsImported(1, i%)
            CustomDeck(4, 12) = CustomDeckAsImported(4, i%)
            CustomDeck(6, 12) = CustomDeckAsImported(6, i%)
        End If
    Next i%
    AdjustStackVals
End Sub
Public Sub Pos13_DragDrop(Source As Control, X As Single, Y As Single)
    Source.Move 3570, 1935
    CustomDeck(1, 13) = 13
    CustomDeck(2, 13) = Right(Source.Name, Len(Source.Name) - 5)
    For i% = 1 To 52
        If CustomDeckAsImported(2, i%) = Right(Source.Name, Len(Source.Name) - 5) Then
            CustomDeck(3, 13) = CustomDeckAsImported(1, i%)
            CustomDeck(4, 13) = CustomDeckAsImported(4, i%)
            CustomDeck(6, 13) = CustomDeckAsImported(6, i%)
        End If
    Next i%
    AdjustStackVals
End Sub
Public Sub Pos14_DragDrop(Source As Control, X As Single, Y As Single)
    Source.Move 3990, 1935
    CustomDeck(1, 14) = 14
    CustomDeck(2, 14) = Right(Source.Name, Len(Source.Name) - 5)
    For i% = 1 To 52
        If CustomDeckAsImported(2, i%) = Right(Source.Name, Len(Source.Name) - 5) Then
            CustomDeck(3, 14) = CustomDeckAsImported(1, i%)
            CustomDeck(4, 14) = CustomDeckAsImported(4, i%)
            CustomDeck(6, 14) = CustomDeckAsImported(6, i%)
        End If
    Next i%
    AdjustStackVals
End Sub
Public Sub Pos15_DragDrop(Source As Control, X As Single, Y As Single)
    Source.Move 4275, 1935
    CustomDeck(1, 15) = 15
    CustomDeck(2, 15) = Right(Source.Name, Len(Source.Name) - 5)
    For i% = 1 To 52
        If CustomDeckAsImported(2, i%) = Right(Source.Name, Len(Source.Name) - 5) Then
            CustomDeck(3, 15) = CustomDeckAsImported(1, i%)
            CustomDeck(4, 15) = CustomDeckAsImported(4, i%)
            CustomDeck(6, 15) = CustomDeckAsImported(6, i%)
        End If
    Next i%
    AdjustStackVals
End Sub
Public Sub Pos16_DragDrop(Source As Control, X As Single, Y As Single)
    Source.Move 4560, 1935
    CustomDeck(1, 16) = 16
    CustomDeck(2, 16) = Right(Source.Name, Len(Source.Name) - 5)
    For i% = 1 To 52
        If CustomDeckAsImported(2, i%) = Right(Source.Name, Len(Source.Name) - 5) Then
            CustomDeck(3, 16) = CustomDeckAsImported(1, i%)
            CustomDeck(4, 16) = CustomDeckAsImported(4, i%)
            CustomDeck(6, 16) = CustomDeckAsImported(6, i%)
        End If
    Next i%
    AdjustStackVals
End Sub
Public Sub Pos17_DragDrop(Source As Control, X As Single, Y As Single)
    Source.Move 4845, 1935
    CustomDeck(1, 17) = 17
    CustomDeck(2, 17) = Right(Source.Name, Len(Source.Name) - 5)
    For i% = 1 To 52
        If CustomDeckAsImported(2, i%) = Right(Source.Name, Len(Source.Name) - 5) Then
            CustomDeck(3, 17) = CustomDeckAsImported(1, i%)
            CustomDeck(4, 17) = CustomDeckAsImported(4, i%)
            CustomDeck(6, 17) = CustomDeckAsImported(6, i%)
        End If
    Next i%
    AdjustStackVals
End Sub
Public Sub Pos18_DragDrop(Source As Control, X As Single, Y As Single)
    Source.Move 5130, 1935
    CustomDeck(1, 18) = 18
    CustomDeck(2, 18) = Right(Source.Name, Len(Source.Name) - 5)
    For i% = 1 To 52
        If CustomDeckAsImported(2, i%) = Right(Source.Name, Len(Source.Name) - 5) Then
            CustomDeck(3, 18) = CustomDeckAsImported(1, i%)
            CustomDeck(4, 18) = CustomDeckAsImported(4, i%)
            CustomDeck(6, 18) = CustomDeckAsImported(6, i%)
        End If
    Next i%
    AdjustStackVals
End Sub
Public Sub Pos19_DragDrop(Source As Control, X As Single, Y As Single)
    Source.Move 5415, 1935
    CustomDeck(1, 19) = 19
    CustomDeck(2, 19) = Right(Source.Name, Len(Source.Name) - 5)
    For i% = 1 To 52
        If CustomDeckAsImported(2, i%) = Right(Source.Name, Len(Source.Name) - 5) Then
            CustomDeck(3, 19) = CustomDeckAsImported(1, i%)
            CustomDeck(4, 19) = CustomDeckAsImported(4, i%)
            CustomDeck(6, 19) = CustomDeckAsImported(6, i%)
        End If
    Next i%
    AdjustStackVals
End Sub
Public Sub Pos20_DragDrop(Source As Control, X As Single, Y As Single)
    Source.Move 5700, 1935
    CustomDeck(1, 20) = 20
    CustomDeck(2, 20) = Right(Source.Name, Len(Source.Name) - 5)
    For i% = 1 To 52
        If CustomDeckAsImported(2, i%) = Right(Source.Name, Len(Source.Name) - 5) Then
            CustomDeck(3, 20) = CustomDeckAsImported(1, i%)
            CustomDeck(4, 20) = CustomDeckAsImported(4, i%)
            CustomDeck(6, 20) = CustomDeckAsImported(6, i%)
        End If
    Next i%
    AdjustStackVals
End Sub
Public Sub Pos21_DragDrop(Source As Control, X As Single, Y As Single)
    Source.Move 5985, 1935
    CustomDeck(1, 21) = 21
    CustomDeck(2, 21) = Right(Source.Name, Len(Source.Name) - 5)
    For i% = 1 To 52
        If CustomDeckAsImported(2, i%) = Right(Source.Name, Len(Source.Name) - 5) Then
            CustomDeck(3, 21) = CustomDeckAsImported(1, i%)
            CustomDeck(4, 21) = CustomDeckAsImported(4, i%)
            CustomDeck(6, 21) = CustomDeckAsImported(6, i%)
        End If
    Next i%
    AdjustStackVals
End Sub
Public Sub Pos22_DragDrop(Source As Control, X As Single, Y As Single)
    Source.Move 6270, 1935
    CustomDeck(1, 22) = 22
    CustomDeck(2, 22) = Right(Source.Name, Len(Source.Name) - 5)
    For i% = 1 To 52
        If CustomDeckAsImported(2, i%) = Right(Source.Name, Len(Source.Name) - 5) Then
            CustomDeck(3, 22) = CustomDeckAsImported(1, i%)
            CustomDeck(4, 22) = CustomDeckAsImported(4, i%)
            CustomDeck(6, 22) = CustomDeckAsImported(6, i%)
        End If
    Next i%
    AdjustStackVals
End Sub
Public Sub Pos23_DragDrop(Source As Control, X As Single, Y As Single)
    Source.Move 6555, 1935
    CustomDeck(1, 23) = 23
    CustomDeck(2, 23) = Right(Source.Name, Len(Source.Name) - 5)
    For i% = 1 To 52
        If CustomDeckAsImported(2, i%) = Right(Source.Name, Len(Source.Name) - 5) Then
            CustomDeck(3, 23) = CustomDeckAsImported(1, i%)
            CustomDeck(4, 23) = CustomDeckAsImported(4, i%)
            CustomDeck(6, 23) = CustomDeckAsImported(6, i%)
        End If
    Next i%
    AdjustStackVals
End Sub
Public Sub Pos24_DragDrop(Source As Control, X As Single, Y As Single)
    Source.Move 6840, 1935
    CustomDeck(1, 24) = 24
    CustomDeck(2, 24) = Right(Source.Name, Len(Source.Name) - 5)
    For i% = 1 To 52
        If CustomDeckAsImported(2, i%) = Right(Source.Name, Len(Source.Name) - 5) Then
            CustomDeck(3, 24) = CustomDeckAsImported(1, i%)
            CustomDeck(4, 24) = CustomDeckAsImported(4, i%)
            CustomDeck(6, 24) = CustomDeckAsImported(6, i%)
        End If
    Next i%
    AdjustStackVals
End Sub
Public Sub Pos25_DragDrop(Source As Control, X As Single, Y As Single)
    Source.Move 7125, 1935
    CustomDeck(1, 25) = 25
    CustomDeck(2, 25) = Right(Source.Name, Len(Source.Name) - 5)
    For i% = 1 To 52
        If CustomDeckAsImported(2, i%) = Right(Source.Name, Len(Source.Name) - 5) Then
            CustomDeck(3, 25) = CustomDeckAsImported(1, i%)
            CustomDeck(4, 25) = CustomDeckAsImported(4, i%)
            CustomDeck(6, 25) = CustomDeckAsImported(6, i%)
        End If
    Next i%
    AdjustStackVals
End Sub
Public Sub Pos26_DragDrop(Source As Control, X As Single, Y As Single)
    Source.Move 7410, 1935
    CustomDeck(1, 26) = 26
    CustomDeck(2, 26) = Right(Source.Name, Len(Source.Name) - 5)
    For i% = 1 To 52
        If CustomDeckAsImported(2, i%) = Right(Source.Name, Len(Source.Name) - 5) Then
            CustomDeck(3, 26) = CustomDeckAsImported(1, i%)
            CustomDeck(4, 26) = CustomDeckAsImported(4, i%)
            CustomDeck(6, 26) = CustomDeckAsImported(6, i%)
        End If
    Next i%
    AdjustStackVals
End Sub
Public Sub Pos27_DragDrop(Source As Control, X As Single, Y As Single)
    Source.Move 150, 3030
    CustomDeck(1, 27) = 27
    CustomDeck(2, 27) = Right(Source.Name, Len(Source.Name) - 5)
    For i% = 1 To 52
        If CustomDeckAsImported(2, i%) = Right(Source.Name, Len(Source.Name) - 5) Then
            CustomDeck(3, 27) = CustomDeckAsImported(1, i%)
            CustomDeck(4, 27) = CustomDeckAsImported(4, i%)
            CustomDeck(6, 27) = CustomDeckAsImported(6, i%)
        End If
    Next i%
    AdjustStackVals
End Sub

Public Sub Pos28_DragDrop(Source As Control, X As Single, Y As Single)
    Source.Move 435, 3030
    CustomDeck(1, 28) = 28
    CustomDeck(2, 28) = Right(Source.Name, Len(Source.Name) - 5)
    For i% = 1 To 52
        If CustomDeckAsImported(2, i%) = Right(Source.Name, Len(Source.Name) - 5) Then
            CustomDeck(3, 28) = CustomDeckAsImported(1, i%)
            CustomDeck(4, 28) = CustomDeckAsImported(4, i%)
            CustomDeck(6, 28) = CustomDeckAsImported(6, i%)
        End If
    Next i%
    AdjustStackVals
End Sub
Public Sub Pos29_DragDrop(Source As Control, X As Single, Y As Single)
    Source.Move 720, 3030
    CustomDeck(1, 29) = 29
    CustomDeck(2, 29) = Right(Source.Name, Len(Source.Name) - 5)
    For i% = 1 To 52
        If CustomDeckAsImported(2, i%) = Right(Source.Name, Len(Source.Name) - 5) Then
            CustomDeck(3, 29) = CustomDeckAsImported(1, i%)
            CustomDeck(4, 29) = CustomDeckAsImported(4, i%)
            CustomDeck(6, 29) = CustomDeckAsImported(6, i%)
        End If
    Next i%
    AdjustStackVals
End Sub
Public Sub Pos30_DragDrop(Source As Control, X As Single, Y As Single)
    Source.Move 1005, 3030
    CustomDeck(1, 30) = 30
    CustomDeck(2, 30) = Right(Source.Name, Len(Source.Name) - 5)
    For i% = 1 To 52
        If CustomDeckAsImported(2, i%) = Right(Source.Name, Len(Source.Name) - 5) Then
            CustomDeck(3, 30) = CustomDeckAsImported(1, i%)
            CustomDeck(4, 30) = CustomDeckAsImported(4, i%)
            CustomDeck(6, 30) = CustomDeckAsImported(6, i%)
        End If
    Next i%
    AdjustStackVals
End Sub
Public Sub Pos31_DragDrop(Source As Control, X As Single, Y As Single)
    Source.Move 1290, 3030
    CustomDeck(1, 31) = 31
    CustomDeck(2, 31) = Right(Source.Name, Len(Source.Name) - 5)
    For i% = 1 To 52
        If CustomDeckAsImported(2, i%) = Right(Source.Name, Len(Source.Name) - 5) Then
            CustomDeck(3, 31) = CustomDeckAsImported(1, i%)
            CustomDeck(4, 31) = CustomDeckAsImported(4, i%)
            CustomDeck(6, 31) = CustomDeckAsImported(6, i%)
        End If
    Next i%
    AdjustStackVals
End Sub
Public Sub Pos32_DragDrop(Source As Control, X As Single, Y As Single)
    Source.Move 1575, 3030
    CustomDeck(1, 32) = 32
    CustomDeck(2, 32) = Right(Source.Name, Len(Source.Name) - 5)
    For i% = 1 To 52
        If CustomDeckAsImported(2, i%) = Right(Source.Name, Len(Source.Name) - 5) Then
            CustomDeck(3, 32) = CustomDeckAsImported(1, i%)
            CustomDeck(4, 32) = CustomDeckAsImported(4, i%)
            CustomDeck(6, 32) = CustomDeckAsImported(6, i%)
        End If
    Next i%
    AdjustStackVals
End Sub
Public Sub Pos33_DragDrop(Source As Control, X As Single, Y As Single)
    Source.Move 1860, 3030
    CustomDeck(1, 33) = 33
    CustomDeck(2, 33) = Right(Source.Name, Len(Source.Name) - 5)
    For i% = 1 To 52
        If CustomDeckAsImported(2, i%) = Right(Source.Name, Len(Source.Name) - 5) Then
            CustomDeck(3, 33) = CustomDeckAsImported(1, i%)
            CustomDeck(4, 33) = CustomDeckAsImported(4, i%)
            CustomDeck(6, 33) = CustomDeckAsImported(6, i%)
        End If
    Next i%
    AdjustStackVals
End Sub
Public Sub Pos34_DragDrop(Source As Control, X As Single, Y As Single)
    Source.Move 2145, 3030
    CustomDeck(1, 34) = 34
    CustomDeck(2, 34) = Right(Source.Name, Len(Source.Name) - 5)
    For i% = 1 To 52
        If CustomDeckAsImported(2, i%) = Right(Source.Name, Len(Source.Name) - 5) Then
            CustomDeck(3, 34) = CustomDeckAsImported(1, i%)
            CustomDeck(4, 34) = CustomDeckAsImported(4, i%)
            CustomDeck(6, 34) = CustomDeckAsImported(6, i%)
        End If
    Next i%
    AdjustStackVals
End Sub
Public Sub Pos35_DragDrop(Source As Control, X As Single, Y As Single)
    Source.Move 2430, 3030
    CustomDeck(1, 35) = 35
    CustomDeck(2, 35) = Right(Source.Name, Len(Source.Name) - 5)
    For i% = 1 To 52
        If CustomDeckAsImported(2, i%) = Right(Source.Name, Len(Source.Name) - 5) Then
            CustomDeck(3, 35) = CustomDeckAsImported(1, i%)
            CustomDeck(4, 35) = CustomDeckAsImported(4, i%)
            CustomDeck(6, 35) = CustomDeckAsImported(6, i%)
        End If
    Next i%
    AdjustStackVals
End Sub
Public Sub Pos36_DragDrop(Source As Control, X As Single, Y As Single)
    Source.Move 2715, 3030
    CustomDeck(1, 36) = 36
    CustomDeck(2, 36) = Right(Source.Name, Len(Source.Name) - 5)
    For i% = 1 To 52
        If CustomDeckAsImported(2, i%) = Right(Source.Name, Len(Source.Name) - 5) Then
            CustomDeck(3, 36) = CustomDeckAsImported(1, i%)
            CustomDeck(4, 36) = CustomDeckAsImported(4, i%)
            CustomDeck(6, 36) = CustomDeckAsImported(6, i%)
        End If
    Next i%
    AdjustStackVals
End Sub
Public Sub Pos37_DragDrop(Source As Control, X As Single, Y As Single)
    Source.Move 3000, 3030
    CustomDeck(1, 37) = 37
    CustomDeck(2, 37) = Right(Source.Name, Len(Source.Name) - 5)
    For i% = 1 To 52
        If CustomDeckAsImported(2, i%) = Right(Source.Name, Len(Source.Name) - 5) Then
            CustomDeck(3, 37) = CustomDeckAsImported(1, i%)
            CustomDeck(4, 37) = CustomDeckAsImported(4, i%)
            CustomDeck(6, 37) = CustomDeckAsImported(6, i%)
        End If
    Next i%
    AdjustStackVals
End Sub
Public Sub Pos38_DragDrop(Source As Control, X As Single, Y As Single)
    Source.Move 3285, 3030
    CustomDeck(1, 38) = 38
    CustomDeck(2, 38) = Right(Source.Name, Len(Source.Name) - 5)
    For i% = 1 To 52
        If CustomDeckAsImported(2, i%) = Right(Source.Name, Len(Source.Name) - 5) Then
            CustomDeck(3, 38) = CustomDeckAsImported(1, i%)
            CustomDeck(4, 38) = CustomDeckAsImported(4, i%)
            CustomDeck(6, 38) = CustomDeckAsImported(6, i%)
        End If
    Next i%
    AdjustStackVals
End Sub
Public Sub Pos39_DragDrop(Source As Control, X As Single, Y As Single)
    Source.Move 3570, 3030
    CustomDeck(1, 39) = 39
    CustomDeck(2, 39) = Right(Source.Name, Len(Source.Name) - 5)
    For i% = 1 To 52
        If CustomDeckAsImported(2, i%) = Right(Source.Name, Len(Source.Name) - 5) Then
            CustomDeck(3, 39) = CustomDeckAsImported(1, i%)
            CustomDeck(4, 39) = CustomDeckAsImported(4, i%)
            CustomDeck(6, 39) = CustomDeckAsImported(6, i%)
        End If
    Next i%
    AdjustStackVals
End Sub
Public Sub Pos40_DragDrop(Source As Control, X As Single, Y As Single)
    Source.Move 3990, 3030
    CustomDeck(1, 40) = 40
    CustomDeck(2, 40) = Right(Source.Name, Len(Source.Name) - 5)
    For i% = 1 To 52
        If CustomDeckAsImported(2, i%) = Right(Source.Name, Len(Source.Name) - 5) Then
            CustomDeck(3, 40) = CustomDeckAsImported(1, i%)
            CustomDeck(4, 40) = CustomDeckAsImported(4, i%)
            CustomDeck(6, 40) = CustomDeckAsImported(6, i%)
        End If
    Next i%
    AdjustStackVals
End Sub
Public Sub Pos41_DragDrop(Source As Control, X As Single, Y As Single)
    Source.Move 4275, 3030
    CustomDeck(1, 41) = 41
    CustomDeck(2, 41) = Right(Source.Name, Len(Source.Name) - 5)
    For i% = 1 To 52
        If CustomDeckAsImported(2, i%) = Right(Source.Name, Len(Source.Name) - 5) Then
            CustomDeck(3, 41) = CustomDeckAsImported(1, i%)
            CustomDeck(4, 41) = CustomDeckAsImported(4, i%)
            CustomDeck(6, 41) = CustomDeckAsImported(6, i%)
        End If
    Next i%
    AdjustStackVals
End Sub
Public Sub Pos42_DragDrop(Source As Control, X As Single, Y As Single)
    Source.Move 4560, 3030
    CustomDeck(1, 42) = 42
    CustomDeck(2, 42) = Right(Source.Name, Len(Source.Name) - 5)
    For i% = 1 To 52
        If CustomDeckAsImported(2, i%) = Right(Source.Name, Len(Source.Name) - 5) Then
            CustomDeck(3, 42) = CustomDeckAsImported(1, i%)
            CustomDeck(4, 42) = CustomDeckAsImported(4, i%)
            CustomDeck(6, 42) = CustomDeckAsImported(6, i%)
        End If
    Next i%
    AdjustStackVals
End Sub
Public Sub Pos43_DragDrop(Source As Control, X As Single, Y As Single)
    Source.Move 4845, 3030
    CustomDeck(1, 43) = 43
    CustomDeck(2, 43) = Right(Source.Name, Len(Source.Name) - 5)
    For i% = 1 To 52
        If CustomDeckAsImported(2, i%) = Right(Source.Name, Len(Source.Name) - 5) Then
            CustomDeck(3, 43) = CustomDeckAsImported(1, i%)
            CustomDeck(4, 43) = CustomDeckAsImported(4, i%)
            CustomDeck(6, 43) = CustomDeckAsImported(6, i%)
        End If
    Next i%
    AdjustStackVals
End Sub
Public Sub Pos44_DragDrop(Source As Control, X As Single, Y As Single)
    Source.Move 5130, 3030
    CustomDeck(1, 44) = 44
    CustomDeck(2, 44) = Right(Source.Name, Len(Source.Name) - 5)
    For i% = 1 To 52
        If CustomDeckAsImported(2, i%) = Right(Source.Name, Len(Source.Name) - 5) Then
            CustomDeck(3, 44) = CustomDeckAsImported(1, i%)
            CustomDeck(4, 44) = CustomDeckAsImported(4, i%)
            CustomDeck(6, 44) = CustomDeckAsImported(6, i%)
        End If
    Next i%
    AdjustStackVals
End Sub
Public Sub Pos45_DragDrop(Source As Control, X As Single, Y As Single)
    Source.Move 5415, 3030
    CustomDeck(1, 45) = 45
    CustomDeck(2, 45) = Right(Source.Name, Len(Source.Name) - 5)
    For i% = 1 To 52
        If CustomDeckAsImported(2, i%) = Right(Source.Name, Len(Source.Name) - 5) Then
            CustomDeck(3, 45) = CustomDeckAsImported(1, i%)
            CustomDeck(4, 45) = CustomDeckAsImported(4, i%)
            CustomDeck(6, 45) = CustomDeckAsImported(6, i%)
        End If
    Next i%
    AdjustStackVals
End Sub
Public Sub Pos46_DragDrop(Source As Control, X As Single, Y As Single)
    Source.Move 5700, 3030
    CustomDeck(1, 46) = 46
    CustomDeck(2, 46) = Right(Source.Name, Len(Source.Name) - 5)
    For i% = 1 To 52
        If CustomDeckAsImported(2, i%) = Right(Source.Name, Len(Source.Name) - 5) Then
            CustomDeck(3, 46) = CustomDeckAsImported(1, i%)
            CustomDeck(4, 46) = CustomDeckAsImported(4, i%)
            CustomDeck(6, 46) = CustomDeckAsImported(6, i%)
        End If
    Next i%
    AdjustStackVals
End Sub
Public Sub Pos47_DragDrop(Source As Control, X As Single, Y As Single)
    Source.Move 5985, 3030
    CustomDeck(1, 47) = 47
    CustomDeck(2, 47) = Right(Source.Name, Len(Source.Name) - 5)
    For i% = 1 To 52
        If CustomDeckAsImported(2, i%) = Right(Source.Name, Len(Source.Name) - 5) Then
            CustomDeck(3, 47) = CustomDeckAsImported(1, i%)
            CustomDeck(4, 47) = CustomDeckAsImported(4, i%)
            CustomDeck(6, 47) = CustomDeckAsImported(6, i%)
        End If
    Next i%
    AdjustStackVals
End Sub
Public Sub Pos48_DragDrop(Source As Control, X As Single, Y As Single)
    Source.Move 6270, 3030
    CustomDeck(1, 48) = 48
    CustomDeck(2, 48) = Right(Source.Name, Len(Source.Name) - 5)
    For i% = 1 To 52
        If CustomDeckAsImported(2, i%) = Right(Source.Name, Len(Source.Name) - 5) Then
            CustomDeck(3, 48) = CustomDeckAsImported(1, i%)
            CustomDeck(4, 48) = CustomDeckAsImported(4, i%)
            CustomDeck(6, 48) = CustomDeckAsImported(6, i%)
        End If
    Next i%
    AdjustStackVals
End Sub
Public Sub Pos49_DragDrop(Source As Control, X As Single, Y As Single)
    Source.Move 6555, 3030
    CustomDeck(1, 49) = 49
    CustomDeck(2, 49) = Right(Source.Name, Len(Source.Name) - 5)
    For i% = 1 To 52
        If CustomDeckAsImported(2, i%) = Right(Source.Name, Len(Source.Name) - 5) Then
            CustomDeck(3, 49) = CustomDeckAsImported(1, i%)
            CustomDeck(4, 49) = CustomDeckAsImported(4, i%)
            CustomDeck(6, 49) = CustomDeckAsImported(6, i%)
        End If
    Next i%
    AdjustStackVals
End Sub
Public Sub Pos50_DragDrop(Source As Control, X As Single, Y As Single)
    Source.Move 6840, 3030
    CustomDeck(1, 50) = 50
    CustomDeck(2, 50) = Right(Source.Name, Len(Source.Name) - 5)
    For i% = 1 To 52
        If CustomDeckAsImported(2, i%) = Right(Source.Name, Len(Source.Name) - 5) Then
            CustomDeck(3, 50) = CustomDeckAsImported(1, i%)
            CustomDeck(4, 50) = CustomDeckAsImported(4, i%)
            CustomDeck(6, 50) = CustomDeckAsImported(6, i%)
        End If
    Next i%
    AdjustStackVals
End Sub
Public Sub Pos51_DragDrop(Source As Control, X As Single, Y As Single)
    Source.Move 7125, 3030
    CustomDeck(1, 51) = 51
    CustomDeck(2, 51) = Right(Source.Name, Len(Source.Name) - 5)
    For i% = 1 To 52
        If CustomDeckAsImported(2, i%) = Right(Source.Name, Len(Source.Name) - 5) Then
            CustomDeck(3, 51) = CustomDeckAsImported(1, i%)
            CustomDeck(4, 51) = CustomDeckAsImported(4, i%)
            CustomDeck(6, 51) = CustomDeckAsImported(6, i%)
        End If
    Next i%
    AdjustStackVals
End Sub
Public Sub Pos52_DragDrop(Source As Control, X As Single, Y As Single)
    Source.Move 7410, 3030
    CustomDeck(1, 52) = 52
    CustomDeck(2, 52) = Right(Source.Name, Len(Source.Name) - 5)
    For i% = 1 To 52
        If CustomDeckAsImported(2, i%) = Right(Source.Name, Len(Source.Name) - 5) Then
            CustomDeck(3, 52) = CustomDeckAsImported(1, i%)
            CustomDeck(4, 52) = CustomDeckAsImported(4, i%)
            CustomDeck(6, 52) = CustomDeckAsImported(6, i%)
        End If
    Next i%
    AdjustStackVals
End Sub
Public Sub PositionStanyonVariationIndexes()
For i% = 1 To 13
    For Each Ctrl In Controls
        If Ctrl.Tag = "Index" Then
            If CustomDeck(2, i%) = Right(Ctrl.Name, Len(Ctrl.Name) - 5) Then
                Ctrl.Move (150 + (i% - 1) * 285), 1935
            End If
        End If
    Next Ctrl
Next i%
For i% = 14 To 26
    For Each Ctrl In Controls
        If Ctrl.Tag = "Index" Then
            If CustomDeck(2, i%) = Right(Ctrl.Name, Len(Ctrl.Name) - 5) Then
                Ctrl.Move (3990 + (i% - 14) * 285), 1935
            End If
        End If
    Next Ctrl
Next i%
For i% = 27 To 39
    For Each Ctrl In Controls
        If Ctrl.Tag = "Index" Then
            If CustomDeck(2, i%) = Right(Ctrl.Name, Len(Ctrl.Name) - 5) Then
                Ctrl.Move (150 + (i% - 27) * 285), 3030
            End If
        End If
    Next Ctrl
Next i%
For i% = 40 To 52
    For Each Ctrl In Controls
        If Ctrl.Tag = "Index" Then
            If CustomDeck(2, i%) = Right(Ctrl.Name, Len(Ctrl.Name) - 5) Then
                Ctrl.Move (3990 + (i% - 40) * 285), 3030
            End If
        End If
    Next Ctrl
Next i%
End Sub
Public Sub ResetCustomStack_Click()
    IndexAH.Move 150, 4305
    Index2H.Move 435, 4305
    Index3H.Move 720, 4305
    Index4H.Move 1005, 4305
    Index5H.Move 1290, 4305
    Index6H.Move 1575, 4305
    Index7H.Move 1860, 4305
    Index8H.Move 2145, 4305
    Index9H.Move 2430, 4305
    Index10H.Move 2715, 4305
    IndexJH.Move 3000, 4305
    IndexQH.Move 3285, 4305
    IndexKH.Move 3570, 4305
    IndexAD.Move 3990, 4305
    Index2D.Move 4275, 4305
    Index3D.Move 4560, 4305
    Index4D.Move 4845, 4305
    Index5D.Move 5130, 4305
    Index6D.Move 5415, 4305
    Index7D.Move 5700, 4305
    Index8D.Move 5985, 4305
    Index9D.Move 6270, 4305
    Index10D.Move 6555, 4305
    IndexJD.Move 6840, 4305
    IndexQD.Move 7125, 4305
    IndexKD.Move 7410, 4305
    IndexAS.Move 150, 5400
    Index2S.Move 435, 5400
    Index3S.Move 720, 5400
    Index4S.Move 1005, 5400
    Index5S.Move 1290, 5400
    Index6S.Move 1575, 5400
    Index7S.Move 1860, 5400
    Index8S.Move 2145, 5400
    Index9S.Move 2430, 5400
    Index10S.Move 2715, 5400
    IndexJS.Move 3000, 5400
    IndexQS.Move 3285, 5400
    IndexKS.Move 3570, 5400
    IndexAC.Move 3990, 5400
    Index2C.Move 4275, 5400
    Index3C.Move 4560, 5400
    Index4C.Move 4845, 5400
    Index5C.Move 5130, 5400
    Index6C.Move 5415, 5400
    Index7C.Move 5700, 5400
    Index8C.Move 5985, 5400
    Index9C.Move 6270, 5400
    Index10C.Move 6555, 5400
    IndexJC.Move 6840, 5400
    IndexQC.Move 7125, 5400
    IndexKC.Move 7410, 5400
    For i% = 1 To 52
        CustomDeck(1, i%) = Empty
        CustomDeck(2, i%) = Empty
        CustomDeckAsImported(1, i%) = Empty
        CustomDeckAsImported(2, i%) = Empty
        CustomDeckAsImported(4, i%) = Empty
        CustomDeckAsImported(6, i%) = Empty
        CustomDeck(3, i%) = i%
    Next i%
    ImportedCustomDeck = 0
    CreatedStanyonDeck = 0
    ReorderRetainedStack.Enabled = False
    StanyonParameters.Visible = False
    AdjustStackVals
End Sub
Public Sub StanyonSpecialDeck_Click()
CheckStanyonParameters
If StanyonParameterError < 31 Then
    MsgBox "You must place cards in each of the first five positions."
    Exit Sub
End If
CheckSuitParameters
If StanyonSuitError < 10 Then
    MsgBox "The first four cards must contain one each of the four suits."
    Exit Sub
End If
CheckCycleParameter
If StanyonCycleError = 1 Then
    MsgBox "The first & fifth cards must not be the same value."
    Exit Sub
End If
CreateStanyonParameters
CreateStanyonVariationDeck
PositionStanyonVariationIndexes
Text1 = "1st Card: " & StanyonStartingCard
Text2 = "Suit Order: "
text3 = "Card Increments: "
For i% = 1 To 4
    Text2 = Text2 & SuitOrder(i%) & " "
    text3 = text3 & NextCardIncrementsReport(i%) & " "
Next i%
StanyonParameters.Caption = Text1 & "   " & Text2 & "   " & text3
StanyonParameters.Visible = True
CreatedStanyonDeck = 1
'new section
'need to create a phanton Imported Deck condition
ImportedCustomDeck = 1
'end new section
AdjustStackVals
End Sub

Public Sub TransferCustomStack_Click()
TransferStatus = 0
For Each Ctrl In Controls
    If Ctrl.Tag = "Index" Then
        If Ctrl.Top > 4000 Then
            TransferStatus = 1
        End If
    End If
Next Ctrl
If TransferStatus = 1 Then
    MsgBox "You must position all the cards above the line" & Chr(13) _
        & "before you can transfer your custom stack."
    Exit Sub
End If
frmStackView.ClearSelections_Click
For i% = 1 To 52
    Deck(1, i%) = CustomDeck(1, i%)
    Deck(2, i%) = CustomDeck(2, i%)
    Deck(4, i%) = CustomDeck(4, i%)
    Deck(6, i%) = CustomDeck(6, i%)
Next i%
For k% = 1 To 52
    For m% = 1 To 52
        If Deck(1, m%) = k% Then
            TestOriginalDeck(1, k%) = Deck(1, m%)
            TestOriginalDeck(2, k%) = Deck(2, m%)
        End If
    Next m%
Next k%
DeckCount = 52
frmStackView.ShowCards
AdjustStackValsSpecial
End Sub

Public Sub TransferCustomStackRetain_Click()
TransferStatus = 0
For Each Ctrl In Controls
    If Ctrl.Tag = "Index" Then
        If Ctrl.Top > 4000 Then
            TransferStatus = 1
        End If
    End If
Next Ctrl
If TransferStatus = 1 Then
    MsgBox "You must position all the cards above the line" & Chr(13) _
        & "before you can transfer your custom stack."
    Exit Sub
End If
If ImportedCustomDeck = 0 And CreatedStanyonDeck = 0 Then
    MsgBox ("Since you are starting from a New Custom Deck state," & Chr(13) & _
        "you must use the Transfer: New button to transfer the Custom Deck." & Chr(13) & Chr(13) & _
        "You may do a Transfer: Retain only after Importing a stack" & Chr(13) & _
        "or after Creating a Stanyon Variation deck." & Chr(13) & Chr(13) & _
        "(At this stage, you may also press the Reset: Retain button, and your" & Chr(13) & _
        "Custom Deck will behave like an imported deck.)")
    Exit Sub
End If
frmStackView.ClearSelections_Click
For i% = 1 To 52
    Deck(1, i%) = CustomDeck(3, i%)
    Deck(2, i%) = CustomDeck(2, i%)
    Deck(4, i%) = CustomDeck(4, i%)
    Deck(6, i%) = CustomDeck(6, i%)
Next i%
For k% = 1 To 52
    For m% = 1 To 52
        If CustomDeckAsImported(1, m%) = k% Then
            TestOriginalDeck(1, k%) = CustomDeckAsImported(1, m%)
            TestOriginalDeck(2, k%) = CustomDeckAsImported(2, m%)
        End If
    Next m%
Next k%
DeckCount = 52
frmStackView.ShowCards
AdjustStackVals
End Sub
Public Sub frmCustomDeck_DragDrop(Source As Control, X As Single, Y As Single)
    Source.Move OriginalLeft, OriginalTop
    Source.ZOrder
End Sub
Private Sub Form_Unload(Cancel As Integer)
    frmMain.mnuCustomDeck.Checked = False
End Sub

