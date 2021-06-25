VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   5235
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9165
   LinkTopic       =   "Form4"
   Picture         =   "Espesificaciones.frx":0000
   ScaleHeight     =   5235
   ScaleWidth      =   9165
   StartUpPosition =   3  'Windows Default
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   315
      Left            =   4440
      TabIndex        =   8
      Top             =   1320
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "S"
      Text            =   "Seleccionar..."
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   5280
      TabIndex        =   6
      Top             =   2280
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Enabled         =   0   'False
      Height          =   495
      Left            =   5640
      MaskColor       =   &H00FFC0FF&
      Picture         =   "Espesificaciones.frx":21C0A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4320
      UseMaskColor    =   -1  'True
      Width           =   2895
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Total:"
      Height          =   255
      Left            =   6120
      TabIndex        =   14
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4440
      TabIndex        =   13
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Stock:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6240
      TabIndex        =   12
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Precio:"
      Height          =   255
      Left            =   4440
      TabIndex        =   11
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label Label8 
      Caption         =   "Label8"
      Height          =   255
      Left            =   6840
      TabIndex        =   10
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Talla"
      Height          =   255
      Left            =   4440
      TabIndex        =   9
      Top             =   960
      Width           =   615
   End
   Begin VB.Label Label6 
      Height          =   255
      Left            =   6720
      TabIndex        =   7
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label VF 
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   4095
      Left            =   600
      Stretch         =   -1  'True
      Top             =   480
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Camiseta deportiva estanpada"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   4440
      TabIndex        =   4
      Top             =   360
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "US$15"
      Height          =   255
      Left            =   5040
      TabIndex        =   3
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00C0FFFF&
      Height          =   615
      Left            =   4440
      TabIndex        =   2
      Top             =   3000
      Width           =   4215
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Descripción"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4440
      TabIndex        =   1
      Top             =   2760
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
    If DataCombo1.Text = "Seleccionar..." Then Exit Sub
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
    Form1.Command4.Enabled = True
    Form4.Hide
End Sub

Private Sub DataCombo1_Change()
    CTP
    With TP
        X = Form1.Label5.Caption
        .Find "Id_Producto='" & X & "'"
        If DataCombo1.BoundText = "S" Then If Val(Trim(!Talla_S)) = 0 Then MsgBox "NO EXISTE EN STOCK": DataCombo1.Text = "Seleccionar...": Exit Sub Else Label8.Caption = !Talla_S
        If DataCombo1.BoundText = "M" Then If Val(Trim(!Talla_M)) = 0 Then MsgBox "NO EXISTE EN STOCK": DataCombo1.Text = "Seleccionar...": Exit Sub Else Label8.Caption = !Talla_M
        If DataCombo1.BoundText = "G" Then If Val(Trim(!Talla_G)) = 0 Then MsgBox "NO EXISTE EN STOCK": DataCombo1.Text = "Seleccionar...": Exit Sub Else Label8.Caption = !Talla_G
    End With
    If DataCombo1.Text = "Seleccionar..." Then Exit Sub
    Command1.Enabled = True
    Text1.Enabled = True
    Text1.SetFocus
End Sub

Private Sub DataCombo1_Click(Area As Integer)
    CTabla1
    Set DataCombo1.RowSource = Tabla1
    DataCombo1.BoundColumn = "Campo1"
    DataCombo1.ListField = "Campo1"
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
        X = Form1.Label5.Caption
        .Find "Id_Producto='" & X & "'"
        If DataCombo1.BoundText = "S" Then If Text1.Text > Val(Trim(!Talla_S)) Then MsgBox "Supera el stock": Text1.Text = "": Exit Sub
        If DataCombo1.BoundText = "M" Then If Text1.Text > Val(Trim(!Talla_M)) Then MsgBox "Supera el stock": Text1.Text = "": Exit Sub
        If DataCombo1.BoundText = "G" Then If Text1.Text > Val(Trim(!Talla_G)) Then MsgBox "Supera el stock": Text1.Text = "": Exit Sub
    End With
    Label6.Caption = Val(Text1.Text) * Val(Label2.Caption)
End Sub

