VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   5235
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11370
   LinkTopic       =   "Form4"
   ScaleHeight     =   5235
   ScaleWidth      =   11370
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   8880
      TabIndex        =   10
      Top             =   1800
      Width           =   495
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Option2"
      Height          =   255
      Left            =   9840
      TabIndex        =   8
      Top             =   1080
      Width           =   255
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H000000C0&
      Caption         =   "Option1"
      Height          =   255
      Left            =   9240
      MaskColor       =   &H00FF0000&
      TabIndex        =   7
      Top             =   1080
      Width           =   255
   End
   Begin VB.ComboBox lstSeleccionar 
      Height          =   315
      ItemData        =   "Espesificaciones.frx":0000
      Left            =   6960
      List            =   "Espesificaciones.frx":000D
      TabIndex        =   6
      Top             =   1800
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Agregar a Carrito"
      Height          =   495
      Left            =   7920
      MaskColor       =   &H00FFC0FF&
      TabIndex        =   0
      Top             =   3960
      UseMaskColor    =   -1  'True
      Width           =   2895
   End
   Begin VB.Label Label6 
      Height          =   255
      Left            =   9840
      TabIndex        =   11
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label VF 
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   600
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   4095
      Left            =   2880
      Stretch         =   -1  'True
      Top             =   360
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "Camiseta deportiva estanpada"
      BeginProperty Font 
         Name            =   "Transformers Movie"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6960
      TabIndex        =   5
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "US$15"
      Height          =   375
      Left            =   9720
      TabIndex        =   4
      Top             =   240
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0080FF80&
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   7080
      Shape           =   2  'Oval
      Top             =   1200
      Width           =   255
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   7560
      Shape           =   2  'Oval
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label Label3 
      Caption         =   "Color"
      Height          =   255
      Left            =   8280
      TabIndex        =   3
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   $"Espesificaciones.frx":001A
      Height          =   615
      Left            =   6960
      TabIndex        =   2
      Top             =   2880
      Width           =   4215
   End
   Begin VB.Label Label5 
      Caption         =   "Descripción"
      Height          =   255
      Left            =   6960
      TabIndex        =   1
      Top             =   2400
      Width           =   2175
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    If VF.Caption = 0 Then Exit Sub
    If Text1.Text = "" Then Exit Sub
    CTEMP
    With Temp
        .AddNew
        !Id_P_FK = Form1.Label5.Caption
        !Descripción = Label4.Caption
        !Talla = lstSeleccionar.Text
        !Cantidad = Text1.Text
        !Precio = Label2.Caption
        !Total = Label6.Caption
        .UpdateBatch
    End With
    Form4.Hide
End Sub

Private Sub lstSeleccionar_LostFocus()
    VF.Caption = 1
    CTP
    If lstSeleccionar = "S" Then
        With TP
            x = Form1.Label5.Caption
            .Find "Id_Producto='" & x & "'"
            If Val(Trim(!Talla_S)) = 0 Then MsgBox "NO EXISTE EN STOCK": VF.Caption = 0
        End With
    End If
    If lstSeleccionar = "M" Then
        With TP
            x = Form1.Label5.Caption
            .Find "Id_Producto='" & x & "'"
            If Val(Trim(!Talla_M)) = 0 Then MsgBox "NO EXISTE EN STOCK": VF.Caption = 0
        End With
    End If
    If lstSeleccionar = "G" Then
        With TP
            x = Form1.Label5.Caption
            .Find "Id_Producto='" & x & "'"
            If Val(Trim(!Talla_G)) = 0 Then MsgBox "NO EXISTE EN STOCK": VF.Caption = 0
        End With
    End If
    Text1.Enabled = True
    Text1.SetFocus
End Sub

Private Sub Option1_Click()
Option2.Visible = False
End Sub

Private Sub Text1_Change()
    If Text1.Text = "" Then Exit Sub
    CTP
    With TP
        x = Form1.Label5.Caption
        .Find "Id_Producto='" & x & "'"
        If lstSeleccionar = "S" Then If Text1.Text > Val(Trim(!Talla_S)) Then MsgBox "Supera el stock": Text1.Text = "": Exit Sub
        If lstSeleccionar = "M" Then If Text1.Text > Val(Trim(!Talla_M)) Then MsgBox "Supera el stock": Text1.Text = "": Exit Sub
        If lstSeleccionar = "G" Then If Text1.Text > Val(Trim(!Talla_G)) Then MsgBox "Supera el stock": Text1.Text = "": Exit Sub
    End With
    Label6.Caption = Val(Text1.Text) * Val(Label2.Caption)
End Sub
