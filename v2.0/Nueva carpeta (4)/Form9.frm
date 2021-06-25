VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form9 
   Caption         =   "Form9"
   ClientHeight    =   3165
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9150
   LinkTopic       =   "Form9"
   ScaleHeight     =   3165
   ScaleWidth      =   9150
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Buscar Fatura"
      Height          =   495
      Left            =   6840
      TabIndex        =   5
      Top             =   1680
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
      Left            =   6840
      TabIndex        =   3
      Top             =   2280
      Width           =   1815
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2295
      Left            =   480
      TabIndex        =   2
      Top             =   480
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   4048
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
      Left            =   6840
      TabIndex        =   1
      Top             =   1080
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ver Productos"
      Height          =   495
      Left            =   6840
      TabIndex        =   0
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   255
      Left            =   3840
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
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
    Form1.Label8.Caption = 1
End Sub

Private Sub Command2_Click()
    Form6.Show
    Form6.Command3.Enabled = False
    Form6.Command2.Enabled = False
End Sub

Private Sub Command3_Click()
    CFact
    With Fact
        .Find "Id_F='" & Label1.Caption & "'"
        !Valido = "False"
        .UpdateBatch
        If .State = 1 Then .Close
        x = Label3.Caption
        y = "True"
        .Open "select * from Factura where [Id_C]like '" & x & "' and [Valido]like '" & y & "'", base, adOpenStatic, adLockBatchOptimistic
    End With
    If Label2.Caption = "T" Then Set DataGrid1.DataSource = Fact Else Set DataGrid1.DataSource = Adodc1
    Command3.Enabled = False
End Sub

Private Sub Command4_Click()
    Label2.Caption = "F"
    Form10.Show
    Form10.Text1.Text = ""
    If Label2.Caption = "T" Then Set DataGrid1.DataSource = Fact Else Set DataGrid1.DataSource = Adodc1
End Sub

Private Sub DataGrid1_Click()
    Command3.Enabled = True
    With Adodc1.Recordset
        Label1.Caption = !Id_F
    End With
    If Label2.Caption = "T" Then Set DataGrid1.DataSource = Fact Else Set DataGrid1.DataSource = Adodc1
End Sub

Sub carga()
    Adodc1.CursorLocation = adUseClient
    Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\base\base.mdb;Persist Security Info=False"
    x = "True"
    Adodc1.RecordSource = "select * from Factura where [Valido]like '" & x & "'"
    Set DataGrid1.DataSource = Adodc1
End Sub

Private Sub Form_Load()
    carga
End Sub
