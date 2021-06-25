VERSION 5.00
Begin VB.Form Form9 
   Caption         =   "Form9"
   ClientHeight    =   6930
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11595
   LinkTopic       =   "Form9"
   ScaleHeight     =   6930
   ScaleWidth      =   11595
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   735
      Left            =   720
      TabIndex        =   17
      Top             =   4440
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   10200
      TabIndex        =   16
      Top             =   5400
      Width           =   375
   End
   Begin VB.Label posicion 
      Caption         =   "Label1"
      Height          =   375
      Left            =   10560
      TabIndex        =   15
      Top             =   4320
      Width           =   255
   End
   Begin VB.Label codigo 
      Caption         =   "Label1"
      Height          =   135
      Index           =   6
      Left            =   9480
      TabIndex        =   14
      Top             =   4560
      Width           =   135
   End
   Begin VB.Label Label 
      Caption         =   "Label1"
      Height          =   375
      Index           =   6
      Left            =   7680
      TabIndex        =   13
      Top             =   6240
      Width           =   1455
   End
   Begin VB.Image Imagen 
      Height          =   2175
      Index           =   6
      Left            =   7680
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Label codigo 
      Caption         =   "Label1"
      Height          =   135
      Index           =   5
      Left            =   6960
      TabIndex        =   12
      Top             =   4560
      Width           =   135
   End
   Begin VB.Label Label 
      Caption         =   "Label1"
      Height          =   375
      Index           =   5
      Left            =   5160
      TabIndex        =   11
      Top             =   6240
      Width           =   1455
   End
   Begin VB.Image Imagen 
      Height          =   2175
      Index           =   5
      Left            =   5160
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Label codigo 
      Caption         =   "Label1"
      Height          =   135
      Index           =   4
      Left            =   4440
      TabIndex        =   10
      Top             =   4560
      Width           =   135
   End
   Begin VB.Label Label 
      Caption         =   "Label1"
      Height          =   375
      Index           =   4
      Left            =   2640
      TabIndex        =   9
      Top             =   6240
      Width           =   1455
   End
   Begin VB.Image Imagen 
      Height          =   2175
      Index           =   4
      Left            =   2640
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Label codigo 
      Caption         =   "Label1"
      Height          =   135
      Index           =   3
      Left            =   10800
      TabIndex        =   8
      Top             =   1440
      Width           =   135
   End
   Begin VB.Label Label 
      Caption         =   "Label1"
      Height          =   375
      Index           =   3
      Left            =   9000
      TabIndex        =   7
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Image Imagen 
      Height          =   2175
      Index           =   3
      Left            =   9000
      Stretch         =   -1  'True
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label codigo 
      Caption         =   "Label1"
      Height          =   135
      Index           =   2
      Left            =   8280
      TabIndex        =   6
      Top             =   1440
      Width           =   135
   End
   Begin VB.Label Label 
      Caption         =   "Label1"
      Height          =   375
      Index           =   2
      Left            =   6480
      TabIndex        =   5
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Image Imagen 
      Height          =   2175
      Index           =   2
      Left            =   6480
      Stretch         =   -1  'True
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label codigo 
      Caption         =   "Label1"
      Height          =   135
      Index           =   1
      Left            =   5760
      TabIndex        =   4
      Top             =   1440
      Width           =   135
   End
   Begin VB.Label Label 
      Caption         =   "Label1"
      Height          =   375
      Index           =   1
      Left            =   3960
      TabIndex        =   3
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Image Imagen 
      Height          =   2175
      Index           =   1
      Left            =   3960
      Stretch         =   -1  'True
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label codigo 
      Caption         =   "Label1"
      Height          =   135
      Index           =   0
      Left            =   3240
      TabIndex        =   2
      Top             =   1440
      Width           =   135
   End
   Begin VB.Label Label5 
      Caption         =   "Label1"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   495
   End
   Begin VB.Label Label 
      Caption         =   "Label1"
      Height          =   375
      Index           =   0
      Left            =   1440
      TabIndex        =   0
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Image Imagen 
      Height          =   2175
      Index           =   0
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   720
      Width           =   1575
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub invicible()
    For i = 0 To 6
        Imagen(i).Visible = False
        Label(i).Visible = False
    Next i
End Sub

Private Sub Command1_Click()
    CTP
    With TP
        .Find "Id_Producto='" & posicion.Caption & "'"
        invicible
        For i = 0 To 6
            If .EOF Or .BOF Then Exit Sub
                .MoveNext
            If .EOF Or .BOF Then Exit Sub
            If Trim(!URL) = "" Then
                Imagen(i).Picture = LoadPicture("C:\Proyecto\final\img\nimg.jpg")
            Else
                Imagen(i).Picture = LoadPicture(Trim(!URL))
            End If
            Label(i).Caption = !Etiqueta
            codigo(i).Caption = !Id_Producto
            Imagen(i).Visible = True
            Label(i).Visible = True
        Next i
        posicion.Caption = !Id_Producto
    End With
End Sub

Private Sub Command2_Click()
    copiar
End Sub

Private Sub Form_Load()
    CTEMP
    With Temp
        .MoveFirst
        For i = 1 To .RecordCount
            .Delete
            .MoveNext
        Next i
    End With
    Form9.Hide
End Sub
