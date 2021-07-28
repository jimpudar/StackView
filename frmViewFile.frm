VERSION 5.00
Begin VB.Form frmViewFile 
   Caption         =   "View File: "
   ClientHeight    =   8400
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   5580
   Icon            =   "frmViewFile.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8400
   ScaleMode       =   0  'User
   ScaleWidth      =   5700
   Begin VB.TextBox FileTextBox 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8280
      Left            =   60
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   60
      Width           =   5490
   End
End
Attribute VB_Name = "frmViewFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Me.Width = 5500
    Me.Height = 8400
    FormResize
End Sub

Private Sub FormResize()
    On Error Resume Next
    FileTextBox.Move 50, 50, Me.ScaleWidth - 75, Me.ScaleHeight - 75
End Sub

Private Sub Form_Resize()
    FormResize
End Sub
