VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   5235
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9165
   LinkTopic       =   "Form4"
   ScaleHeight     =   5235
   ScaleWidth      =   9165
   StartUpPosition =   3  'Windows Default
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   315
      Left            =   4440
      TabIndex        =   11
      Top             =   1920
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "S"
      Text            =   "Seleccionar..."
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   6360
      TabIndex        =   9
      Top             =   1920
      Width           =   495
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Option2"
      Height          =   255
      Left            =   7320
      TabIndex        =   7
      Top             =   1200
      Width           =   255
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H000000C0&
      Caption         =   "Option1"
      Height          =   255
      Left            =   6720
      MaskColor       =   &H00FF0000&
      TabIndex        =   6
      Top             =   1200
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Agregar a Carrito"
      Enabled         =   0   'False
      Height          =   495
      Left            =   5400
      MaskColor       =   &H00FFC0FF&
      TabIndex        =   0
      Top             =   4080
      UseMaskColor    =   -1  'True
      Width           =   2895
   End
   Begin VB.Label Label6 
      Height          =   255
      Left            =   7320
      TabIndex        =   10
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label VF 
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   600
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   4095
      Left            =   360
      Stretch         =   -1  'True
      Top             =   480
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
      Left            =   4440
      TabIndex        =   5
      Top             =   360
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "US$15"
      Height          =   375
      Left            =   7200
      TabIndex        =   4
      Top             =   360
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0080FF80&
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   4560
      Shape           =   2  'Oval
      Top             =   1320
      Width           =   255
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   5040
      Shape           =   2  'Oval
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label Label3 
      Caption         =   "Color"
      Height          =   255
      Left            =   5760
      TabIndex        =   3
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   $"Espesificaciones.frx":0000
      Height          =   615
      Left            =   4440
      TabIndex        =   2
      Top             =   3000
      Width           =   4215
   End
   Begin VB.Label Label5 
      Caption         =   "Descripción"
      Height          =   255
      Left            =   4440
      TabIndex        =   1
      Top             =   2520
      Width           =   2175
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    If Text1.Text = "" Then Exit Sub
    CTEMP
    With Temp
        .AddNew
        !Id_P_FK = Form1.Label5.Caption
        !Descripción = Label4.Caption
        !Talla = DataCombo1.BoundText
        !Cantidad = Text1.Text
        !Precio = Label2.Caption
        !Total = Label6.Caption
        .UpdateBatch
    End With
    Form4.Hide
End Sub

Private Sub DataCombo1_Click(Area As Integer)
    CTP
    With TP
        x = Form1.Label5.Caption
        .Find "Id_Producto='" & x & "'"
        If DataCombo1.BoundText = "S" Then If Val(Trim(!Talla_S)) = 0 Then MsgBox "NO EXISTE EN STOCK": DataCombo1.Text = "Seleccionar...": Exit Sub
        If DataCombo1.BoundText = "M" Then If Val(Trim(!Talla_M)) = 0 Then MsgBox "NO EXISTE EN STOCK": DataCombo1.Text = "Seleccionar...": Exit Sub
        If DataCombo1.BoundText = "G" Then If Val(Trim(!Talla_G)) = 0 Then MsgBox "NO EXISTE EN STOCK": DataCombo1.Text = "Seleccionar...": Exit Sub
    End With
    If DataCombo1.Text = "Seleccionar..." Then Exit Sub
    Command1.Enabled = True
    Text1.Enabled = True
    Text1.SetFocus
End Sub

Private Sub Form_Load()
    CTabla1
    Set DataCombo1.RowSource = Tabla1
    DataCombo1.BoundColumn = "Campo1"
    DataCombo1.ListField = "Campo1"
End Sub

Private Sub Text1_Change()
    If Text1.Text = "" Then Exit Sub
    CTP
    With TP
        x = Form1.Label5.Caption
        .Find "Id_Producto='" & x & "'"
        If DataCombo1.BoundText = "S" Then If Text1.Text > Val(Trim(!Talla_S)) Then MsgBox "Supera el stock": Text1.Text = "": Exit Sub
        If DataCombo1.BoundText = "M" Then If Text1.Text > Val(Trim(!Talla_M)) Then MsgBox "Supera el stock": Text1.Text = "": Exit Sub
        If DataCombo1.BoundText = "G" Then If Text1.Text > Val(Trim(!Talla_G)) Then MsgBox "Supera el stock": Text1.Text = "": Exit Sub
    End With
    Label6.Caption = Val(Text1.Text) * Val(Label2.Caption)
End Sub
