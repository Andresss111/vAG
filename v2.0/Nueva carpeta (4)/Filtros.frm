VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   4335
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3525
   LinkTopic       =   "Form3"
   ScaleHeight     =   4335
   ScaleWidth      =   3525
   StartUpPosition =   3  'Windows Default
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   315
      Left            =   240
      TabIndex        =   6
      Top             =   1320
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   556
      _Version        =   393216
      Text            =   "Seleccionar..."
   End
   Begin VB.OptionButton Option7 
      Caption         =   "G"
      Height          =   255
      Left            =   2040
      TabIndex        =   5
      Top             =   3120
      Width           =   615
   End
   Begin VB.OptionButton Option6 
      Caption         =   "M"
      Height          =   255
      Left            =   1080
      TabIndex        =   4
      Top             =   3120
      Width           =   615
   End
   Begin VB.OptionButton Option5 
      Caption         =   "S"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   255
      Left            =   2760
      TabIndex        =   7
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "Talla"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Categoria"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Filtros"
      Height          =   255
      Left            =   1440
      TabIndex        =   0
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub DataCombo1_Change()
    CTTP
    With TTP
        .Find "Descripción='" & Trim(DataCombo1.BoundText) & "'"
        Label4.Caption = !Id_TP
    End With
    CTP
    With TP
        If .State = 1 Then .Close
        x = Label4.Caption
        .Open "select * from Producto where [Id_TP_FK]like '" & x & "'", base, adOpenStatic, adLockBatchOptimistic
        Form1.invicible
        For i = 0 To 6
            If .EOF Or .BOF Then Exit Sub
            If i = 0 Then
                .MoveFirst
            Else
                .MoveNext
            End If
            If .EOF Or .BOF Then Exit Sub
            If Trim(!URL) = "" Then
                Form1.Image1(i).Picture = LoadPicture("C:\Proyecto\final\img\nimg.jpg")
            Else
                y = App.Path
                Form1.Image1(i).Picture = LoadPicture(y & "\img\" & Trim(!URL))
            End If
            Form1.Label4(i).Caption = !Etiqueta
            Form1.Label6(i).Caption = !Id_Producto
            Form1.Image1(i).Visible = True
            Form1.Label4(i).Visible = True
        Next i
        Form1.Label7.Caption = !Id_Producto
    End With
    CTTP
    Set DataCombo1.RowSource = TTP
    DataCombo1.BoundColumn = "Descripción"
    DataCombo1.ListField = "Descripción"
End Sub

Private Sub DataCombo1_Click(Area As Integer)
    CTTP
    Set DataCombo1.RowSource = TTP
    DataCombo1.BoundColumn = "Descripción"
    DataCombo1.ListField = "Descripción"
End Sub

Private Sub Form_Load()
    CTTP
    Set DataCombo1.RowSource = TTP
    DataCombo1.BoundColumn = "Descripción"
    DataCombo1.ListField = "Descripción"
End Sub

Sub bus()
    With TP
        If .State = 1 Then .Close
        x = Label4.Caption
        .Open "select * from Producto where [Id_TP_FK]like '" & x & "'", base, adOpenStatic, adLockBatchOptimistic
    End With
End Sub
