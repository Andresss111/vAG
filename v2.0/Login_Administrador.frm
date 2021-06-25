VERSION 5.00
Begin VB.Form Form7 
   Caption         =   "Form7"
   ClientHeight    =   5235
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5835
   LinkTopic       =   "Form7"
   Picture         =   "Login_Administrador.frx":0000
   ScaleHeight     =   5235
   ScaleWidth      =   5835
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdsal 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   2400
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Login_Administrador.frx":134F3
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4560
      Width           =   1335
   End
   Begin VB.CommandButton cmding 
      Height          =   255
      Left            =   480
      Picture         =   "Login_Administrador.frx":1402C
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4560
      Width           =   1215
   End
   Begin VB.TextBox txtcon 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1440
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   3720
      Width           =   2415
   End
   Begin VB.TextBox txtusu 
      Height          =   285
      Left            =   1440
      TabIndex        =   3
      Top             =   3120
      Width           =   2415
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Contraseña:"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   3720
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario:"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   3120
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Bienvenido"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   2040
      Width           =   1815
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmding_Click()
    With TU
        If .State = 1 Then .Close
        .Open "select * from Login_Ad where [Usuario]like '" & txtusu.Text & "'", base, adOpenStatic, adLockBatchOptimistic
        If .EOF Or .BOF Then MsgBox "El usuario no existe", vbCritical: Exit Sub
        .Find "Usuario='" & txtusu.Text & "'"
        If !Contraseña = txtcon.Text Then Form9.Show: Form7.Hide: Form7.Label2.Caption = "F" Else MsgBox "El usuario y contraseña no coinciden", vbCritical: Exit Sub
    End With
End Sub

Private Sub cmdsal_Click()
    Form5.Show
    Form7.Hide
End Sub

