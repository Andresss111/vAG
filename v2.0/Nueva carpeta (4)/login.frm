VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "Form5"
   ClientHeight    =   2730
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7710
   LinkTopic       =   "Form5"
   Picture         =   "login.frx":0000
   ScaleHeight     =   2730
   ScaleWidth      =   7710
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   2775
      Left            =   3960
      ScaleHeight     =   2715
      ScaleWidth      =   3675
      TabIndex        =   0
      Top             =   0
      Width           =   3735
      Begin VB.CommandButton Command2 
         Caption         =   "Cliente"
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   1800
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Administrador"
         Height          =   375
         Left            =   2160
         TabIndex        =   1
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Image Image2 
         Height          =   780
         Left            =   2400
         Picture         =   "login.frx":4E0A
         Top             =   840
         Width           =   780
      End
      Begin VB.Image Image1 
         Height          =   780
         Left            =   600
         Picture         =   "login.frx":5CC2
         Top             =   840
         Width           =   780
      End
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Form7.Show
    Me.Hide
End Sub

Private Sub Command2_Click()
    Form1.Show
    Form1.Label8.Caption = 0
    Me.Hide
End Sub
