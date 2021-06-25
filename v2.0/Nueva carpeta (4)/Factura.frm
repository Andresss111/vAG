VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form2 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form2"
   ClientHeight    =   7230
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8070
   LinkTopic       =   "Form2"
   ScaleHeight     =   7230
   ScaleWidth      =   8070
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   480
      Top             =   6840
      Width           =   1455
      _ExtentX        =   2566
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
   Begin VB.PictureBox Picture1 
      Height          =   6255
      Left            =   600
      ScaleHeight     =   6195
      ScaleWidth      =   7155
      TabIndex        =   0
      Top             =   360
      Width           =   7215
      Begin VB.CommandButton Command1 
         Caption         =   "Guardar"
         Height          =   375
         Left            =   2520
         TabIndex        =   20
         Top             =   5640
         Width           =   1335
      End
      Begin VB.CommandButton cmdcli 
         Caption         =   "Nuevo Cliente"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   5640
         Width           =   1575
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   2175
         Left            =   240
         TabIndex        =   18
         Top             =   2520
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   3836
         _Version        =   393216
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
      Begin VB.TextBox Text3 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5760
         TabIndex        =   17
         Text            =   "12%"
         Top             =   5160
         Width           =   735
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5760
         TabIndex        =   16
         Top             =   4800
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5760
         TabIndex        =   15
         Top             =   5520
         Width           =   735
      End
      Begin VB.TextBox txttel 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5280
         TabIndex        =   11
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox txtdir 
         Enabled         =   0   'False
         Height          =   285
         Left            =   840
         TabIndex        =   9
         Top             =   2040
         Width           =   3615
      End
      Begin VB.TextBox txtruc 
         Height          =   285
         Left            =   4800
         TabIndex        =   7
         Top             =   1560
         Width           =   1695
      End
      Begin VB.TextBox txtnom 
         Enabled         =   0   'False
         Height          =   285
         Left            =   840
         TabIndex        =   5
         Top             =   1560
         Width           =   3255
      End
      Begin VB.Label Label11 
         Height          =   375
         Left            =   360
         TabIndex        =   21
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Subtotal:"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4800
         TabIndex        =   14
         Top             =   4800
         Width           =   735
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Total:"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4800
         TabIndex        =   13
         Top             =   5520
         Width           =   735
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "IVA:"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4800
         TabIndex        =   12
         Top             =   5160
         Width           =   735
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Teléfono:"
         BeginProperty Font 
            Name            =   "Niagara Solid"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4560
         TabIndex        =   10
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Dirección:"
         BeginProperty Font 
            Name            =   "Niagara Solid"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "RUC:"
         BeginProperty Font 
            Name            =   "Niagara Solid"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4200
         TabIndex        =   6
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Nombre:"
         BeginProperty Font 
            Name            =   "Niagara Solid"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Obispo Alberto Ordoñez Crespo Esquina, Cdla Católica"
         BeginProperty Font 
            Name            =   "Niagara Solid"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         TabIndex        =   3
         Top             =   840
         Width           =   4935
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   """El deporte con estilo"""
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   375
         Left            =   3240
         TabIndex        =   2
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "FAIS"
         BeginProperty Font 
            Name            =   "Poor Richard"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   495
         Left            =   2160
         TabIndex        =   1
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Label Label13 
      Caption         =   "Label13"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   3720
      Width           =   255
   End
   Begin VB.Label bloq 
      Caption         =   "Label13"
      Height          =   495
      Left            =   120
      TabIndex        =   23
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label Label12 
      Caption         =   "Label12"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   1440
      Width           =   255
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdcli_Click()
    Form8.Show
    Form8.txtdir = ""
    Form8.txtema = ""
    Form8.txtnomc = ""
    Form8.txtruc = ""
    Form8.txttel = ""
    Form8.txtnomc.SetFocus
    Form8.txtnomc.Enabled = True
    Form8.txttel.Enabled = True
    Form8.txtdir.Enabled = True
    Form8.txtema.Enabled = True
    Form8.Label7.Caption = "F"
End Sub

Private Sub Command1_Click()
    If txtnom.Text = "" Then MsgBox "Rellene los datos del cliente", vbCritical: Exit Sub
    If DataGrid1.ApproxCount < 1 Then MsgBox "Añada productos a su compra", vbCritical: Exit Sub
    CFact
    With Fact
        .AddNew
        !Id_C = txtruc.Text
        !Fecha = Date
        !Subtotal = Text2.Text
        !IVA = CDbl(Text3.Text)
        !Total = CDbl(Text1.Text)
        !Valido = "True"
        .UpdateBatch
        Label11.Caption = !Id_F
    End With
    CTEMP
    With Temp
        x = .RecordCount
    End With
    For i = 1 To x
        CTEMP
        With Temp
            If i = 1 Then
                .MoveFirst
            Else
                .Find "Id='" & Label12.Caption & "'"
                .MoveNext
            End If
            a = !Id_P_FK
            b = !Descripción
            c = !Talla
            d = !Cantidad
            e = !Precio
            f = !Total
            Label12.Caption = !ID
        End With
        Label13.Caption = c
        CTP
        With TP
            If Label13.Caption = "S" Then !Talla_S = Val(!Talla_S) - Val(d)
            If Label13.Caption = "M" Then !Talla_M = Val(!Talla_M) - Val(d)
            If Label13.Caption = "G" Then !Talla_G = Val(!Talla_G) - Val(d)
            .UpdateBatch
        End With
        CDFact
        With DFact
            .AddNew
            !Id_P_FK = a
            !Id_F = Label11.Caption
            !Descripción = b
            !Talla = c
            !Cantidad = d
            !Precio = e
            !Total = f
            .UpdateBatch
        End With
    Next i
    CTEMP
    Set DataGrid1.DataSource = Temp
    With Temp
        x = .RecordCount
    End With
    For i = 1 To x
        With Temp
            .Delete
            .MoveNext
            .UpdateBatch
        End With
    Next i
    Form1.Show
    Form2.Hide
End Sub

Private Sub Command2_Click()
    CTEMP
    Set DataGrid1.DataSource = Temp
    With Temp
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

Private Sub Form_Load()
    'Adodc1.CursorLocation = adUseClient
    'Adodc1."Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\base\base.mdb;Persist Security Info=False"
    'Adodc1.RecordSource = "select * from Temp"
    CFact
    With Fact
        If .EOF Or .BOF Then
            Label11.Caption = "1"
        Else
            .MoveLast
            Label11.Caption = Val(!Id_F) + 1
        End If
    End With
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    txtnom.Text = ""
    txtdir.Text = ""
    txtruc.Text = ""
    txttel.Text = ""
    bloq.Caption = "0"
    CTEMP
    Set DataGrid1.DataSource = Temp
End Sub

Private Sub txtruc_Change()
    If bloq.Caption = "1" Then Exit Sub
    CTP
    With Clientes
        x = txtruc.Text
        If .State = 1 Then .Close
        .Open "select * from Cliente where [Id_C]like '" & x & "'", base, adOpenStatic, adLockBatchOptimistic
        If .EOF Or .BOF Then Exit Sub
        .Find "Id_C = '" & x & "'"
        If .EOF Or .BOF Then Exit Sub
        txtnom.Text = !Nombre
        txtdir.Text = !Dirección
        txttel.Text = !Celular
        bloq.Caption = "1"
    End With
    CTEMP
    With Temp
        For i = 1 To .RecordCount
            If i = 1 Then
                .MoveFirst
            Else
                .MoveNext
            End If
            Text2.Text = Val(Text2.Text) + Val(!Total)
        Next i
    End With
    Text3.Text = CDbl(Text2.Text) * 0.12
    Text1.Text = CDbl(Text2.Text) + CDbl(Text3.Text)
    CTEMP
    Set DataGrid1.DataSource = Temp
End Sub
