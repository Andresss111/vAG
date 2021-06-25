VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Form6 
   Caption         =   "Form6"
   ClientHeight    =   5250
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8400
   LinkTopic       =   "Form6"
   Picture         =   "Tipo de Producto.frx":0000
   ScaleHeight     =   5250
   ScaleWidth      =   8400
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1680
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   315
      Left            =   5400
      TabIndex        =   18
      Top             =   2520
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   556
      _Version        =   393216
      Text            =   "Seleccionar..."
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   5520
      TabIndex        =   17
      Top             =   3960
      Width           =   615
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Index           =   2
      Left            =   6960
      TabIndex        =   15
      Top             =   3360
      Width           =   375
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Index           =   1
      Left            =   6240
      TabIndex        =   13
      Top             =   3360
      Width           =   375
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Index           =   0
      Left            =   5520
      TabIndex        =   11
      Top             =   3360
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Borrar"
      Height          =   375
      Left            =   1440
      TabIndex        =   9
      Top             =   4320
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Modificar"
      Height          =   375
      Left            =   2760
      TabIndex        =   8
      Top             =   4320
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Guardar"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   4320
      Width           =   975
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   5400
      TabIndex        =   6
      Top             =   1920
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   5400
      TabIndex        =   5
      Top             =   1440
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   5400
      TabIndex        =   4
      Top             =   960
      Width           =   2415
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Talla:"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   22
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label ID 
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label10 
      Caption         =   "Label10"
      Height          =   255
      Left            =   6960
      TabIndex        =   20
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   2295
      Left            =   1200
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label Label9 
      Caption         =   "No_Picture.jpg"
      Height          =   255
      Left            =   5760
      TabIndex        =   19
      Top             =   120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "G"
      Height          =   255
      Left            =   7080
      TabIndex        =   16
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "M"
      Height          =   255
      Left            =   6360
      TabIndex        =   14
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "S"
      Height          =   255
      Left            =   5640
      TabIndex        =   12
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo de Producto"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3480
      TabIndex        =   10
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Precio"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   3
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Descripcion"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   2
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Etiqueta"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Productos"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
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
      Width           =   3615
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    If DataCombo1.Text = "Seleccionar..." Then MsgBox "Seleccione un tipo de producto": Exit Sub
    CTP
    With TP
        .AddNew
        !Etiqueta = Trim(Text1.Text)
        !Descripcion = Trim(Text2.Text)
        !Precio = Trim(Text3.Text)
        !Talla_S = Trim(Text4(0).Text)
        !Talla_M = Trim(Text4(1).Text)
        !Talla_G = Trim(Text4(2).Text)
        !Cantidad = Trim(Text7.Text)
        !URL = Trim(Label9.Caption)
        !Id_TP_FK = Trim(Label10.Caption)
        .UpdateBatch
        .MoveLast
        ID.Caption = !Id_Producto
    End With
    Command1.Enabled = False
End Sub

Private Sub Command2_Click()
    If DataCombo1.Text = "Seleccionar..." Then MsgBox "Seleccione un tipo de producto": Exit Sub
    'If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4(0).Text = "" Or Text4(1).Text = "" Or Text4(2).Text = "" Or Text7.Text = "" Then MsgBox "Rellene los campos": Exit Sub
    x = ID.Caption
    CTP
    With TP
        .Find "Id_Producto='" & ID.Caption & "'"
        !Etiqueta = Trim(Text1.Text)
        !Descripcion = Trim(Text2.Text)
        !Precio = Trim(Text3.Text)
        !Talla_S = Trim(Text4(0).Text)
        !Talla_M = Trim(Text4(1).Text)
        !Talla_G = Trim(Text4(2).Text)
        !Cantidad = Trim(Text7.Text)
        !URL = Trim(Label9.Caption)
        !Id_TP_FK = Trim(Label10.Caption)
        .UpdateBatch
    End With
    
End Sub

Private Sub Command3_Click()
    CTP
    With TP
        .Find "Id_Producto='" & ID.Caption & "'"
        '.Delete 'psdt duda existencial si le borro o no por lo del detalle fact
        '.UpdateBatch
    End With
End Sub

Private Sub DataCombo1_Change()
    CTTP
    With TTP
        .Find "Descripción='" & Trim(DataCombo1.BoundText) & "'"
        Label10.Caption = !Id_TP
    End With
    Set DataCombo1.RowSource = TTP
    DataCombo1.BoundColumn = "Descripción"
    DataCombo1.ListField = "Descripción"
End Sub

Private Sub Form_Load()
    CTTP
    Set DataCombo1.RowSource = TTP
    DataCombo1.BoundColumn = "Descripción"
    DataCombo1.ListField = "Descripción"
    Image1.Picture = LoadPicture(App.Path & "\img\" & Label9.Caption)
End Sub

Private Sub Image1_Click()
    CommonDialog1.DialogTitle = "Selecciona un archivo"
    CommonDialog1.Filter = "Archivo |*.jpg"
    CommonDialog1.ShowOpen
    If CommonDialog1.FileName <> "" Then
        b = CommonDialog1.FileName
        Image1.Picture = LoadPicture(b)
        Label9.Caption = CommonDialog1.FileTitle
    Else
        MsgBox "No se encontro ningun archivo", vbInformation, "Error"
    End If
End Sub

Private Sub Text4_Change(Index As Integer)
    Text7.Text = Val(Text4(0).Text) + Val(Text4(1).Text) + Val(Text4(2).Text)
    If KeyAscii = "13" Then If Index = "0" Then Text4(1).SetFocus
End Sub
