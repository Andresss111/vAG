VERSION 5.00
Begin VB.Form Form8 
   Caption         =   "Form8"
   ClientHeight    =   4365
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6270
   LinkTopic       =   "Form8"
   Picture         =   "Clientes.frx":0000
   ScaleHeight     =   4365
   ScaleWidth      =   6270
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdgu 
      Caption         =   "Guardar"
      Height          =   375
      Left            =   4920
      TabIndex        =   11
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox txtema 
      Height          =   285
      Left            =   3000
      TabIndex        =   5
      Top             =   2760
      Width           =   2055
   End
   Begin VB.TextBox txttel 
      Height          =   285
      Left            =   3000
      TabIndex        =   4
      Top             =   2280
      Width           =   1935
   End
   Begin VB.TextBox txtdir 
      Height          =   285
      Left            =   3000
      TabIndex        =   3
      Top             =   1800
      Width           =   2175
   End
   Begin VB.TextBox txtruc 
      Height          =   285
      Left            =   3000
      TabIndex        =   2
      Top             =   1320
      Width           =   2175
   End
   Begin VB.TextBox txtnomc 
      Height          =   285
      Left            =   3000
      TabIndex        =   1
      Top             =   840
      Width           =   2415
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Email:"
      Height          =   255
      Left            =   1560
      TabIndex        =   10
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Telefono:"
      Height          =   255
      Left            =   1560
      TabIndex        =   9
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Dirección:"
      Height          =   255
      Left            =   1560
      TabIndex        =   8
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "RUC:"
      Height          =   255
      Left            =   1560
      TabIndex        =   7
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre:"
      Height          =   255
      Left            =   1560
      TabIndex        =   6
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdgu_Click()
    CTC
    If txtruc.Text = "" Or txtnomc.Text = "" Or txttel.Text = "" Or txtdir.Text = "" Then MsgBox "Por favor rellenar los campos requeridos": Exit Sub
    With Clientes
    
    
        .AddNew
        !Id_C = txtruc.Text
        !Nombre = txtnomc.Text
        !Celular = txttel.Text
        !Dirección = txtdir.Text
        !Email = txtema.Text
        .UpdateBatch
   
        
    End With
    Form8.Hide
End Sub

