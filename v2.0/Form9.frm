VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form9 
   Caption         =   "Form9"
   ClientHeight    =   6465
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12360
   LinkTopic       =   "Form9"
   Picture         =   "Form9.frx":0000
   ScaleHeight     =   6465
   ScaleWidth      =   12360
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command8 
      Caption         =   "Crear Reporte"
      Enabled         =   0   'False
      Height          =   495
      Left            =   7920
      TabIndex        =   15
      Top             =   3720
      Width           =   1815
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   11520
      Top             =   1440
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   360
      Top             =   480
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
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
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc2"
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
   Begin VB.CommandButton Command7 
      Caption         =   "Ver Facturas False"
      Height          =   495
      Left            =   7920
      TabIndex        =   9
      Top             =   2280
      Width           =   1815
   End
   Begin VB.CommandButton Command5 
      Height          =   615
      Left            =   360
      Picture         =   "Form9.frx":4305A
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5520
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Buscar Fatura"
      Height          =   495
      Left            =   10200
      TabIndex        =   5
      Top             =   5760
      Width           =   1815
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   360
      Top             =   120
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
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
      Connect         =   ""
      OLEDBString     =   ""
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
   Begin VB.CommandButton Command3 
      Caption         =   "Cancelar Factura"
      Enabled         =   0   'False
      Height          =   495
      Left            =   7920
      TabIndex        =   3
      Top             =   3000
      Width           =   1815
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2655
      Left            =   1320
      TabIndex        =   2
      Top             =   1920
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   4683
      _Version        =   393216
      Enabled         =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   12298
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   12298
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Añadir producto"
      Height          =   495
      Left            =   8040
      TabIndex        =   1
      Top             =   5760
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ver Productos"
      Height          =   495
      Left            =   5880
      TabIndex        =   0
      Top             =   5760
      Width           =   1815
   End
   Begin VB.CommandButton Command6 
      Caption         =   "NT"
      Height          =   615
      Left            =   11040
      TabIndex        =   13
      Top             =   240
      Width           =   615
   End
   Begin VB.Label NNF 
      BackStyle       =   0  'Transparent
      Caption         =   "Label6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   11520
      TabIndex        =   14
      Top             =   840
      Width           =   375
   End
   Begin VB.Label rep 
      Caption         =   "rep"
      Height          =   495
      Left            =   8520
      TabIndex        =   12
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      Height          =   375
      Left            =   5160
      TabIndex        =   11
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image Image3 
      Height          =   840
      Left            =   8520
      Picture         =   "Form9.frx":4332D
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   960
   End
   Begin VB.Image Image2 
      Height          =   840
      Left            =   10560
      Picture         =   "Form9.frx":4A433
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   960
   End
   Begin VB.Image Image1 
      Height          =   840
      Left            =   6360
      Picture         =   "Form9.frx":4B99C
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   960
   End
   Begin VB.Label Label4 
      Caption         =   "0"
      Height          =   255
      Left            =   3720
      TabIndex        =   10
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   255
      Left            =   1680
      TabIndex        =   7
      Top             =   480
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "F"
      Height          =   375
      Left            =   2880
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   1800
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Form1.Show
    Form1.inicio
    Form1.Label8.Caption = 1
    Form9.Hide
End Sub

Private Sub Command2_Click()
    Form6.Show
    Form6.Command3.Visible = True
    Form6.Command2.Enabled = False
    Form6.Command1.Enabled = True
    Form6.Image1.Picture = LoadPicture(App.Path & "\img\No_Picture.jpg")
    Form9.Hide
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    CFact
    With Fact
        .Find "Id_F='" & Label1.Caption & "'"
        !Valido = "False"
        .UpdateBatch
    End With
    CDFact
        With DFact
        If .State = 1 Then .Close
        .Open "select * from Detalle_Factura where [Id_F]like '" & Label1.Caption & "'", base, adOpenStatic, adLockBatchOptimistic
            For i = 1 To .RecordCount
                If .EOF Or .BOF Then Exit Sub
                a = !Id_P_FK
                b = !Talla
                c = !Cantidad
                CTP
                With TP
                    .Find "Id_Producto='" & a & "'"
                    If b = "S" Then !Talla_S = Val(!Talla_S) + Val(c)
                    If b = "M" Then !Talla_M = Val(!Talla_M) + Val(c)
                    If b = "G" Then !Talla_G = Val(!Talla_G) + Val(c)
                    .UpdateBatch
                End With
                .MoveNext
            Next i
        End With
        Label5.Caption = "T"
    If Label2.Caption = "F" Then carga3
    If Label2.Caption = "T" Then carga2
End Sub

Private Sub Command4_Click()
    Label2.Caption = "F"
    Form10.Show
    Form10.Text1.Text = ""
End Sub

Private Sub Command5_Click()
    Form7.Show
    Form9.Hide
End Sub

Private Sub Command6_Click()
    Form11.Show
End Sub

Private Sub Command7_Click()
    If Command7.Caption = "Ver Facturas False" Then
        With Fact
            If .State = 1 Then .Close
            y = "False"
            .Open "select * from Factura where [Valido]like '" & y & "'", base, adOpenStatic, adLockBatchOptimistic
        End With
        Set DataGrid1.DataSource = Fact
        DataGrid1.Columns(6).Width = 0
        Command7.Caption = "Ver Facturas True"
        Command3.Enabled = False
    Else
        carga3
        Command7.Caption = "Ver Facturas False"
    End If
End Sub

Private Sub Command8_Click()
    Command8.Enabled = False
    Set DataReport2.DataSource = Fact
    y = App.Path
    'DataReport2.Sections("Sección4").Controls("Image1").Picture = LoadPicture(y & "\img\logo.jpg")
    DataReport2.Show
End Sub

Private Sub DataGrid1_Click()
    If DataGrid1.ApproxCount < 1 Then Exit Sub
    If Command7.Caption = "Ver Facturas False" Then Command3.Enabled = True
    If Label2.Caption = "T" Then
        With Fact
            If !Valido = "False" Then Command3.Enabled = False
            Label1.Caption = !Id_F
        End With
    Else
        If Label5.Caption <> "T" Then
            With Adodc1.Recordset
                If !Valido = "False" Then Command3.Enabled = False
                Label1.Caption = !Id_F
            End With
        Else
            With Fact
                If !Valido = "False" Then Command3.Enabled = False
                Label1.Caption = !Id_F
            End With
        End If
    End If
End Sub

Sub carga()
    Adodc1.CursorLocation = adUseClient
    Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\base\base.mdb;Persist Security Info=False"
    x = "True"
    Adodc1.RecordSource = "select * from Factura where [Valido]like '" & x & "'"
    Set DataGrid1.DataSource = Adodc1
    DataGrid1.Columns(6).Width = 0
End Sub

Sub carga2()
    With Fact
        If .State = 1 Then .Close
        x = Label3.Caption
        y = "True"
        .Open "select * from Factura where [Id_C]like '" & x & "' and [Valido]like '" & y & "'", base, adOpenStatic, adLockBatchOptimistic
    End With
    Set DataGrid1.DataSource = Fact
    DataGrid1.Columns(6).Width = 0
End Sub

Sub carga3()
    With Fact
        If .State = 1 Then .Close
        y = "True"
        .Open "select * from Factura where [Valido]like '" & y & "'", base, adOpenStatic, adLockBatchOptimistic
    End With
    Set DataGrid1.DataSource = Fact
    DataGrid1.Columns(6).Width = 0
End Sub

Sub f()
    With DataReport1
        .Sections("Sección4").Controls("Etiqueta1").Caption = Label1.Caption
    End With
    CFact
    With Fact
        .Find "Id_F='" & Label1.Caption & "'"
        rep.Caption = !Id_C
        DataReport1.Sections("Sección3").Controls("Etiqueta12").Caption = !Subtotal
        DataReport1.Sections("Sección3").Controls("Etiqueta13").Caption = !IVA
        DataReport1.Sections("Sección3").Controls("Etiqueta14").Caption = !Total
        If !Valido = "True" Then DataReport1.Sections("Sección4").Controls("Etiqueta18").Caption = "Vigente"
    End With
    CTC
    With Clientes
        .Find "Id_C='" & rep.Caption & "'"
        rep.Caption = !Nombre
        DataReport1.Sections("Sección4").Controls("Etiqueta11").Caption = rep.Caption
    End With
    With DFact
        If .State = 1 Then .Close
        y = Label1.Caption
        .Open "select * from Detalle_Factura where [Id_F]like '" & y & "'", base, adOpenStatic, adLockBatchOptimistic
    End With
    Set DataReport1.DataSource = DFact
    DataReport1.Show
End Sub

Private Sub Form_Load()
    carga
End Sub

Private Sub Timer1_Timer()
    CNF
    With NF
        If .State = 1 Then .Close
        x = "F"
        .Open "select * from NF where [V]like '" & x & "'", base, adOpenStatic, adLockBatchOptimistic
        NNF.Caption = .RecordCount
    End With
End Sub
