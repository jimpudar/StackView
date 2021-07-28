VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4635
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   8340
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   8340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraMainFrame 
      Height          =   4590
      Left            =   45
      TabIndex        =   0
      Top             =   -15
      Width           =   8265
      Begin VB.PictureBox picLogo 
         AutoSize        =   -1  'True
         Height          =   3990
         Left            =   180
         Picture         =   "frmSplash.frx":1CCA
         ScaleHeight     =   3930
         ScaleWidth      =   3405
         TabIndex        =   1
         Top             =   315
         Width           =   3465
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         Caption         =   "StackView"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   870
         Left            =   3990
         TabIndex        =   5
         Tag             =   "Product"
         Top             =   390
         Width           =   3225
      End
      Begin VB.Label lblVersion 
         AutoSize        =   -1  'True
         Caption         =   "Version 5.0.0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3990
         TabIndex        =   4
         Tag             =   "Version"
         Top             =   1365
         Width           =   1590
      End
      Begin VB.Label lblCompany 
         Caption         =   "Contact:   nick@stackview.com"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3990
         TabIndex        =   3
         Tag             =   "Company"
         Top             =   2925
         Width           =   3690
      End
      Begin VB.Label lblCopyright 
         Caption         =   "Copyright 2003, 2004, 2005, 2006 by  Nick Pudar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   3990
         TabIndex        =   2
         Tag             =   "Copyright"
         Top             =   1875
         Width           =   3930
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblProductName.Caption = App.Title
    lblCopyright.Caption = "Copyright © 2003, 2004, 2005, 2006 by Nick Pudar"
End Sub

