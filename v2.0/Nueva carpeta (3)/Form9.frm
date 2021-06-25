VERSION 5.00
Begin VB.Form Form9 
   Caption         =   "Form9"
   ClientHeight    =   2505
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5145
   LinkTopic       =   "Form9"
   ScaleHeight     =   2505
   ScaleWidth      =   5145
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Añadir producto"
      Height          =   495
      Left            =   2760
      TabIndex        =   1
      Top             =   1560
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ver Productos"
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   1560
      Width           =   1815
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Form1.Show
End Sub

Private Sub Command2_Click()
    Form6.Show
End Sub
