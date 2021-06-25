VERSION 5.00
Begin VB.Form Form11 
   Caption         =   "Form11"
   ClientHeight    =   6330
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9450
   LinkTopic       =   "Form11"
   ScaleHeight     =   6330
   ScaleWidth      =   9450
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command1"
      Height          =   375
      Left            =   8880
      TabIndex        =   37
      Top             =   5280
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   8880
      TabIndex        =   36
      Top             =   480
      Width           =   375
   End
   Begin VB.Label Labeln 
      Caption         =   "Label5"
      Height          =   255
      Index           =   7
      Left            =   0
      TabIndex        =   47
      Top             =   5280
      Width           =   255
   End
   Begin VB.Label Labeln 
      Caption         =   "Label5"
      Height          =   255
      Index           =   6
      Left            =   0
      TabIndex        =   46
      Top             =   4680
      Width           =   255
   End
   Begin VB.Label Labeln 
      Caption         =   "Label5"
      Height          =   255
      Index           =   5
      Left            =   0
      TabIndex        =   45
      Top             =   4080
      Width           =   255
   End
   Begin VB.Label Labeln 
      Caption         =   "Label5"
      Height          =   255
      Index           =   4
      Left            =   0
      TabIndex        =   44
      Top             =   3480
      Width           =   255
   End
   Begin VB.Label Labeln 
      Caption         =   "Label5"
      Height          =   255
      Index           =   3
      Left            =   0
      TabIndex        =   43
      Top             =   2880
      Width           =   255
   End
   Begin VB.Label Labeln 
      Caption         =   "Label5"
      Height          =   255
      Index           =   2
      Left            =   0
      TabIndex        =   42
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Labeln 
      Caption         =   "Label5"
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   41
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label Labeln 
      Caption         =   "Label5"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   40
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label ind 
      Caption         =   "Label5"
      Height          =   255
      Left            =   9000
      TabIndex        =   39
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label id 
      Caption         =   "Label5"
      Height          =   255
      Left            =   9000
      TabIndex        =   38
      Top             =   960
      Width           =   255
   End
   Begin VB.Label Labeln3 
      Alignment       =   2  'Center
      Caption         =   "Observación"
      Height          =   375
      Index           =   7
      Left            =   4680
      TabIndex        =   35
      Top             =   5280
      Width           =   3855
   End
   Begin VB.Label Labeln2 
      Caption         =   "Talla del Producto"
      Height          =   375
      Index           =   7
      Left            =   3720
      TabIndex        =   34
      Top             =   5280
      Width           =   855
   End
   Begin VB.Label Labeln1 
      Caption         =   "Nombre del Producto"
      Height          =   375
      Index           =   7
      Left            =   1560
      TabIndex        =   33
      Top             =   5280
      Width           =   2055
   End
   Begin VB.Label Labeln0 
      Caption         =   "Codigo del Producto"
      Height          =   375
      Index           =   7
      Left            =   360
      TabIndex        =   32
      Top             =   5280
      Width           =   1095
   End
   Begin VB.Label Labeln3 
      Alignment       =   2  'Center
      Caption         =   "Observación"
      Height          =   375
      Index           =   6
      Left            =   4680
      TabIndex        =   31
      Top             =   4680
      Width           =   3855
   End
   Begin VB.Label Labeln2 
      Caption         =   "Talla del Producto"
      Height          =   375
      Index           =   6
      Left            =   3720
      TabIndex        =   30
      Top             =   4680
      Width           =   855
   End
   Begin VB.Label Labeln1 
      Caption         =   "Nombre del Producto"
      Height          =   375
      Index           =   6
      Left            =   1560
      TabIndex        =   29
      Top             =   4680
      Width           =   2055
   End
   Begin VB.Label Labeln0 
      Caption         =   "Codigo del Producto"
      Height          =   375
      Index           =   6
      Left            =   360
      TabIndex        =   28
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Label Labeln3 
      Alignment       =   2  'Center
      Caption         =   "Observación"
      Height          =   375
      Index           =   5
      Left            =   4680
      TabIndex        =   27
      Top             =   4080
      Width           =   3855
   End
   Begin VB.Label Labeln2 
      Caption         =   "Talla del Producto"
      Height          =   375
      Index           =   5
      Left            =   3720
      TabIndex        =   26
      Top             =   4080
      Width           =   855
   End
   Begin VB.Label Labeln1 
      Caption         =   "Nombre del Producto"
      Height          =   375
      Index           =   5
      Left            =   1560
      TabIndex        =   25
      Top             =   4080
      Width           =   2055
   End
   Begin VB.Label Labeln0 
      Caption         =   "Codigo del Producto"
      Height          =   375
      Index           =   5
      Left            =   360
      TabIndex        =   24
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label Labeln3 
      Alignment       =   2  'Center
      Caption         =   "Observación"
      Height          =   375
      Index           =   4
      Left            =   4680
      TabIndex        =   23
      Top             =   3480
      Width           =   3855
   End
   Begin VB.Label Labeln2 
      Caption         =   "Talla del Producto"
      Height          =   375
      Index           =   4
      Left            =   3720
      TabIndex        =   22
      Top             =   3480
      Width           =   855
   End
   Begin VB.Label Labeln1 
      Caption         =   "Nombre del Producto"
      Height          =   375
      Index           =   4
      Left            =   1560
      TabIndex        =   21
      Top             =   3480
      Width           =   2055
   End
   Begin VB.Label Labeln0 
      Caption         =   "Codigo del Producto"
      Height          =   375
      Index           =   4
      Left            =   360
      TabIndex        =   20
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Label Labeln3 
      Alignment       =   2  'Center
      Caption         =   "Observación"
      Height          =   375
      Index           =   3
      Left            =   4680
      TabIndex        =   19
      Top             =   2880
      Width           =   3855
   End
   Begin VB.Label Labeln2 
      Caption         =   "Talla del Producto"
      Height          =   375
      Index           =   3
      Left            =   3720
      TabIndex        =   18
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label Labeln1 
      Caption         =   "Nombre del Producto"
      Height          =   375
      Index           =   3
      Left            =   1560
      TabIndex        =   17
      Top             =   2880
      Width           =   2055
   End
   Begin VB.Label Labeln0 
      Caption         =   "Codigo del Producto"
      Height          =   375
      Index           =   3
      Left            =   360
      TabIndex        =   16
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label Labeln3 
      Alignment       =   2  'Center
      Caption         =   "Observación"
      Height          =   375
      Index           =   2
      Left            =   4680
      TabIndex        =   15
      Top             =   2280
      Width           =   3855
   End
   Begin VB.Label Labeln2 
      Caption         =   "Talla del Producto"
      Height          =   375
      Index           =   2
      Left            =   3720
      TabIndex        =   14
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label Labeln1 
      Caption         =   "Nombre del Producto"
      Height          =   375
      Index           =   2
      Left            =   1560
      TabIndex        =   13
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Label Labeln0 
      Caption         =   "Codigo del Producto"
      Height          =   375
      Index           =   2
      Left            =   360
      TabIndex        =   12
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Labeln3 
      Alignment       =   2  'Center
      Caption         =   "Observación"
      Height          =   375
      Index           =   1
      Left            =   4680
      TabIndex        =   11
      Top             =   1680
      Width           =   3855
   End
   Begin VB.Label Labeln2 
      Caption         =   "Talla del Producto"
      Height          =   375
      Index           =   1
      Left            =   3720
      TabIndex        =   10
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Labeln1 
      Caption         =   "Nombre del Producto"
      Height          =   375
      Index           =   1
      Left            =   1560
      TabIndex        =   9
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Label Labeln0 
      Caption         =   "Codigo del Producto"
      Height          =   375
      Index           =   1
      Left            =   360
      TabIndex        =   8
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Labeln3 
      Alignment       =   2  'Center
      Caption         =   "Observación"
      Height          =   375
      Index           =   0
      Left            =   4680
      TabIndex        =   7
      Top             =   1080
      Width           =   3855
   End
   Begin VB.Label Labeln2 
      Caption         =   "Talla del Producto"
      Height          =   375
      Index           =   0
      Left            =   3720
      TabIndex        =   6
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Labeln1 
      Caption         =   "Nombre del Producto"
      Height          =   375
      Index           =   0
      Left            =   1560
      TabIndex        =   5
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label Labeln0 
      Caption         =   "Codigo del Producto"
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   4
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Observación"
      Height          =   375
      Left            =   4680
      TabIndex        =   3
      Top             =   480
      Width           =   3855
   End
   Begin VB.Label Label3 
      Caption         =   "Talla del Producto"
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Nombre del Producto"
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Codigo del Producto"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   1095
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    CNF
    With NF
        invicible
        For i = 0 To 7
            If .EOF Or .BOF Then Exit Sub
            If i = 0 Then
                .MoveFirst
            Else
                .MoveNext
            End If
            If .EOF Or .BOF Then Exit Sub
            Labeln(i).Caption = !Id_NF
            Labeln0(i).Caption = !Id_P
            Labeln1(i).Caption = !N_P
            Labeln2(i).Caption = !T_P
            Labeln3(i).Caption = !Observacion
            Labeln0(i).Visible = True
            Labeln1(i).Visible = True
            Labeln2(i).Visible = True
            Labeln3(i).Visible = True
            If !V = "F" Then
                Labeln0(i).BackColor = &HFF&
                Labeln1(i).BackColor = &HFF&
                Labeln2(i).BackColor = &HFF&
                Labeln3(i).BackColor = &HFF&
            Else
                Labeln0(i).BackColor = &H8000000F
                Labeln1(i).BackColor = &H8000000F
                Labeln2(i).BackColor = &H8000000F
                Labeln3(i).BackColor = &H8000000F
            End If
            id.Caption = !Id_NF
        Next i
    End With
End Sub

Sub invicible()
    For i = 0 To 7
        Labeln0(i).Visible = False
        Labeln1(i).Visible = False
        Labeln2(i).Visible = False
        Labeln3(i).Visible = False
    Next i
End Sub

Sub BS()
    x = ind.Caption
    y = x
    x = Labeln0(x).Caption
    Labeln0(y).BackColor = &H8000000F
    Labeln1(y).BackColor = &H8000000F
    Labeln2(y).BackColor = &H8000000F
    Labeln3(y).BackColor = &H8000000F
    y = Labeln(y).Caption
    CNF
    With NF
        .Find "Id_NF='" & y & "'"
        !V = "V"
        .UpdateBatch
    End With
    CTP
    With TP
        .Find "Id_Producto='" & x & "'"
            Form6.Show
            If Trim(!URL) = "" Then
                Form6.Image1.Picture = LoadPicture("& App.Path &\img\df.jpg")
            Else
                y = App.Path
                Form6.Image1.Picture = LoadPicture(y & "\img\" & Trim(!URL))
            End If
            Form6.Text1.Text = Trim(!Etiqueta)
            Form6.Text2.Text = Trim(!Descripcion)
            Form6.Text3.Text = Trim(!Precio)
            Form6.Command1.Enabled = False
            Form6.Command3.Visible = False
            Form6.Command2.Enabled = True
            Form6.Text4(0).Text = Trim(!Talla_S)
            Form6.Text4(1).Text = Trim(!Talla_M)
            Form6.Text4(2).Text = Trim(!Talla_G)
            Form6.id.Caption = Trim(!Id_Producto)
            Form6.Label9.Caption = Trim(!URL)
            Form6.NFP.Caption = "T"
    End With
End Sub

Private Sub Labeln0_Click(Index As Integer)
    ind.Caption = Index
    BS
End Sub

Private Sub Labeln1_Click(Index As Integer)
    ind.Caption = Index
    BS
End Sub

Private Sub Labeln2_Click(Index As Integer)
    ind.Caption = Index
    BS
End Sub

Private Sub Labeln3_Click(Index As Integer)
    ind.Caption = Index
    BS
End Sub
