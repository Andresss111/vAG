VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16155
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   9000
   ScaleWidth      =   16155
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Height          =   615
      Left            =   840
      Picture         =   "Form1.frx":1841E
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   720
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      DisabledPicture =   "Form1.frx":186F1
      DownPicture     =   "Form1.frx":18FB3
      Height          =   855
      Left            =   480
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":19875
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   7800
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Height          =   495
      Left            =   13680
      Picture         =   "Form1.frx":1A137
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   6480
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Height          =   495
      Left            =   2160
      Picture         =   "Form1.frx":1A657
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   6480
      Width           =   495
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   240
      Top             =   6240
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   $"Form1.frx":1A915
      OLEDBString     =   $"Form1.frx":1A9B3
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      DisabledPicture =   "Form1.frx":1AA51
      DownPicture     =   "Form1.frx":1BFC8
      DragIcon        =   "Form1.frx":1C56D
      Height          =   495
      Left            =   13800
      MaskColor       =   &H0080FFFF&
      Picture         =   "Form1.frx":4F8DB
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label8 
      Caption         =   "Label8"
      Height          =   255
      Left            =   240
      TabIndex        =   23
      Top             =   6720
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label7 
      Height          =   255
      Left            =   240
      TabIndex        =   21
      Top             =   7440
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label6 
      Height          =   255
      Index           =   6
      Left            =   12720
      TabIndex        =   18
      Top             =   6720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label6 
      Height          =   255
      Index           =   5
      Left            =   9120
      TabIndex        =   17
      Top             =   6720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label6 
      Height          =   255
      Index           =   4
      Left            =   5640
      TabIndex        =   16
      Top             =   6720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label6 
      Height          =   255
      Index           =   3
      Left            =   12480
      TabIndex        =   15
      Top             =   1680
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label6 
      Height          =   255
      Index           =   2
      Left            =   9120
      TabIndex        =   14
      Top             =   1680
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label6 
      Height          =   255
      Index           =   1
      Left            =   5760
      TabIndex        =   13
      Top             =   1680
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label6 
      Height          =   255
      Index           =   0
      Left            =   2520
      TabIndex        =   12
      Top             =   1680
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   7080
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Index           =   6
      Left            =   10440
      TabIndex        =   10
      Top             =   8400
      Width           =   2175
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   5
      Left            =   6840
      TabIndex        =   9
      Top             =   8400
      Width           =   2175
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Index           =   4
      Left            =   3360
      TabIndex        =   8
      Top             =   8400
      Width           =   2175
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   3
      Left            =   11640
      TabIndex        =   7
      Top             =   4920
      Width           =   2175
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Index           =   2
      Left            =   8280
      TabIndex        =   6
      Top             =   4920
      Width           =   2175
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Index           =   1
      Left            =   4800
      TabIndex        =   5
      Top             =   4800
      Width           =   2175
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Index           =   0
      Left            =   1680
      TabIndex        =   4
      Top             =   4920
      Width           =   2175
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   2655
      Index           =   6
      Left            =   10320
      Stretch         =   -1  'True
      Top             =   5520
      Width           =   2295
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   2655
      Index           =   5
      Left            =   6720
      Stretch         =   -1  'True
      Top             =   5520
      Width           =   2295
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   2655
      Index           =   4
      Left            =   3240
      Stretch         =   -1  'True
      Top             =   5520
      Width           =   2295
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   2655
      Index           =   3
      Left            =   11520
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   2655
      Index           =   2
      Left            =   8160
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   2655
      Index           =   1
      Left            =   4800
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   2655
      Index           =   0
      Left            =   1560
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "El deporte con estilo"
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   5520
      TabIndex        =   3
      Top             =   960
      Width           =   4095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "FAIS"
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   1095
      Left            =   5040
      TabIndex        =   2
      Top             =   360
      Width           =   5055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Filtros"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   12600
      TabIndex        =   0
      Top             =   1200
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Form3.Show
End Sub

Private Sub Command2_Click()
    invicible
    CTP
    If Form3.DataCombo1.Text <> "Seleccionar..." Then Form3.bus
    With TP
        .Find "Id_Producto='" & Label7.Caption & "'"
        x = 6
        For i = 0 To 6
            If .EOF Or .BOF Then Exit Sub
                .MovePrevious
            If .EOF Or .BOF Then Exit Sub
            If Trim(!URL) = "" Then
                Image1(x).Picture = LoadPicture("& App.Path &\img\df.jpg")
            Else
                Y = App.Path
                Image1(x).Picture = LoadPicture(Y & "\img\" & Trim(!URL))
            End If
            Label4(x).Caption = !Etiqueta
            Label6(x).Caption = !Id_Producto
            Image1(x).Visible = True
            Label4(x).Visible = True
            x = x - 1
        Next i
        Label7.Caption = !Id_Producto
    End With
End Sub

Private Sub Command3_Click()
    invicible
    CTP
    If Form3.DataCombo1.Text <> "Seleccionar..." Then Form3.bus
    With TP
        x = Label7.Caption
        .Find "Id_Producto='" & x & "'"
        For i = 0 To 6
            If .EOF Or .BOF Then Exit Sub
            .MoveNext
            If .EOF Or .BOF Then Exit Sub
            If Trim(!URL) = "" Then
                Image1(i).Picture = LoadPicture("& App.Path &\img\df.jpg")
            Else
                Y = App.Path
                Image1(i).Picture = LoadPicture(Y & "\img\" & Trim(!URL))
            End If
            Label4(i).Caption = !Etiqueta
            Label6(i).Caption = !Id_Producto
            Image1(i).Visible = True
            Label4(i).Visible = True
        Next i
        Label7.Caption = !Id_Producto
    End With
End Sub

Private Sub Command4_Click()
    Form2.Show
    Form2.inicio
End Sub

Private Sub Command5_Click()
    If Label8.Caption = 1 Then Form9.Show Else Form5.Show
    Form3.Hide
    Form1.Hide
End Sub

Private Sub Form_Load()
    inicio
End Sub

Sub inicio()
    Command4.Enabled = False
    CTP
    With TP
        invicible
        For i = 0 To 6
            If .EOF Or .BOF Then Exit Sub
            If i = 0 Then
                .MoveFirst
            Else
                .MoveNext
            End If
            If .EOF Or .BOF Then Exit Sub
            If Trim(!URL) = "" Then
                Image1(i).Picture = LoadPicture("C:\Proyecto\final\img\nimg.jpg")
            Else
                Y = App.Path
                Image1(i).Picture = LoadPicture(Y & "\img\" & Trim(!URL))
            End If
            Label4(i).Caption = !Etiqueta
            Label6(i).Caption = !Id_Producto
            Image1(i).Visible = True
            Label4(i).Visible = True
        Next i
        Label7.Caption = !Id_Producto
    End With
    CTEMP
    With Temp
        If .EOF Or .BOF Then Exit Sub
        x = .RecordCount
    End With
    For i = 1 To x
        With Temp
            .Delete
            .MoveNext
            .UpdateBatch
        End With
    Next i
End Sub

Sub antiguo()
invicible
    CTP
    With TP
        For i = 0 To 6
            If .EOF Or .BOF Then Exit Sub
            If i = 0 Then
                .MoveFirst
            Else
                .MoveNext
            End If
            If .EOF Or .BOF Then Exit Sub
            If Trim(!URL) = "" Then
                Image1(i).Picture = LoadPicture("& App.Path &\img\df.jpg")
            Else
                Image1(i).Picture = LoadPicture(Trim(!URL))
            End If
            Label4(i).Caption = !Etiqueta
            Label6(i).Caption = !Id_Producto
            Image1(i).Visible = True
            Label4(i).Visible = True
        Next i
        Label7.Caption = !Id_Producto
    End With
End Sub

Sub invicible()
    For i = 0 To 6
        Image1(i).Visible = False
        Label4(i).Visible = False
    Next i
End Sub

Private Sub Image1_Click(Index As Integer)
    Form3.Hide
    Label5.Caption = Index
    x = Label5.Caption
    Label5.Caption = Label6(x).Caption
    x = Label5.Caption
    CTP
    With TP
        .Find "Id_Producto='" & x & "'"
        If Label8.Caption = 1 Then
            Form6.Show
            If Trim(!URL) = "" Then
                Form6.Image1.Picture = LoadPicture("& App.Path &\img\df.jpg")
            Else
                Y = App.Path
                Form6.Image1.Picture = LoadPicture(Y & "\img\" & Trim(!URL))
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
            Form6.ID.Caption = Trim(!Id_Producto)
            Form6.Label9.Caption = Trim(!URL)
            Form6.NFP.Caption = "F"
        Else
            Form4.Show
            If Trim(!URL) = "" Then
                Form4.Image2.Picture = LoadPicture("& App.Path &\img\df.jpg")
            Else
                Y = App.Path
                Form4.Image2.Picture = LoadPicture(Y & "\img\" & Trim(!URL))
            End If
            Form4.Label1.Caption = Trim(!Etiqueta)
            Form4.Label4.Caption = Trim(!Descripcion)
            Form4.Label2.Caption = Trim(!Precio)
            Form4.Text1.Enabled = False
            Form4.Label6.Caption = ""
            Form4.Label8.Caption = ""
            Form4.Text1.Text = ""
            Form4.DataCombo1.Text = "Seleccionar..."
            Form4.Command1.Enabled = False
        End If
    End With
End Sub


