VERSION 5.00
Begin VB.Form Form7 
   Caption         =   "Form7"
   ClientHeight    =   3360
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4920
   LinkTopic       =   "Form7"
   ScaleHeight     =   3360
   ScaleWidth      =   4920
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdsal 
      Caption         =   "Salir"
      Height          =   255
      Left            =   2280
      TabIndex        =   6
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton cmding 
      Caption         =   "Ingresar"
      Height          =   255
      Left            =   720
      TabIndex        =   5
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox txtcon 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   1440
      Width           =   2415
   End
   Begin VB.TextBox txtusu 
      Height          =   285
      Left            =   1320
      TabIndex        =   3
      Top             =   960
      Width           =   2415
   End
   Begin VB.Label Label3 
      Caption         =   "Contraseña:"
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Usuario:"
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Bienvenido"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Top             =   240
      Width           =   1215
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
        If !Contraseña = txtcon.Text Then Form9.Show Else MsgBox "El usuario y contraseña no coinciden", vbCritical: Exit Sub
        
    End With
End Sub
