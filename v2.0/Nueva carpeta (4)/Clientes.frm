VERSION 5.00
Begin VB.Form Form8 
   Caption         =   "Form8"
   ClientHeight    =   4365
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5925
   LinkTopic       =   "Form8"
   Picture         =   "Clientes.frx":0000
   ScaleHeight     =   4365
   ScaleWidth      =   5925
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Modificar"
      Height          =   375
      Left            =   3000
      TabIndex        =   12
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton cmdgu 
      Caption         =   "Guardar"
      Height          =   375
      Left            =   4440
      TabIndex        =   11
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox txtema 
      Height          =   285
      Left            =   3000
      TabIndex        =   5
      Top             =   2760
      Width           =   2415
   End
   Begin VB.TextBox txttel 
      Height          =   285
      Left            =   3000
      TabIndex        =   4
      Top             =   2280
      Width           =   1455
   End
   Begin VB.TextBox txtdir 
      Height          =   285
      Left            =   3000
      TabIndex        =   3
      Top             =   1800
      Width           =   2415
   End
   Begin VB.TextBox txtruc 
      Height          =   285
      Left            =   3000
      TabIndex        =   2
      Top             =   1320
      Width           =   1455
   End
   Begin VB.TextBox txtnomc 
      Height          =   285
      Left            =   3000
      TabIndex        =   1
      Top             =   840
      Width           =   2415
   End
   Begin VB.Label Label7 
      Caption         =   "Label7"
      Enabled         =   0   'False
      Height          =   495
      Left            =   240
      TabIndex        =   13
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Email:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
    If Label7.Caption <> "F" Then
        CTC
        If txtruc.Text = "" Or txtnomc.Text = "" Or txttel.Text = "" Or txtdir.Text = "" Then MsgBox "Por favor rellenar los campos requeridos": Exit Sub
        With Clientes
            x = Label7.Caption
            .Find "Id_C='" & x & "'"
            !Id_C = txtruc.Text
            !Nombre = txtnomc.Text
            !Celular = txttel.Text
            !Dirección = txtdir.Text
            !Email = txtema.Text
            .UpdateBatch
        End With
        Form8.Hide
    Else
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
    End If
End Sub

Private Sub Command1_Click()
    If MsgBox("Esta seguro de editar el registro", vbYesNo) = vbNo Then Exit Sub
    txtruc.SetFocus
    txtnomc.Text = ""
    txttel.Text = ""
    txtdir.Text = ""
    txtema.Text = ""
    txtnomc.Enabled = False
    txttel.Enabled = False
    txtdir.Enabled = False
    txtema.Enabled = False
    Label7.Caption = "T"
End Sub

Private Sub txtruc_Change()
    If Label7.Caption = "F" Then Exit Sub
    CTP
    With Clientes
        x = txtruc.Text
        If .State = 1 Then .Close
        .Open "select * from Cliente where [Id_C]like '" & x & "'", base, adOpenStatic, adLockBatchOptimistic
        If .EOF Or .BOF Then Exit Sub
        .Find "Id_C = '" & x & "'"
        If .EOF Or .BOF Then Exit Sub
        txtnomc.Text = !Nombre
        txtdir.Text = !Dirección
        txttel.Text = !Celular
        txtema.Text = !Email
        Label7.Caption = !Id_C
        txtnomc.Enabled = True
        txttel.Enabled = True
        txtdir.Enabled = True
        txtema.Enabled = True
    End With
End Sub
