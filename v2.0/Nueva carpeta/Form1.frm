VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16155
   LinkTopic       =   "Form1"
   ScaleHeight     =   9000
   ScaleWidth      =   16155
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   960
      Top             =   6120
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
      Connect         =   $"Form1.frx":0000
      OLEDBString     =   $"Form1.frx":009E
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
   Begin VB.HScrollBar HScroll1 
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   7680
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      DisabledPicture =   "Form1.frx":013C
      DownPicture     =   "Form1.frx":16B3
      DragIcon        =   "Form1.frx":1C58
      Height          =   495
      Left            =   13800
      MaskColor       =   &H0080FFFF&
      Picture         =   "Form1.frx":34FC6
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   375
      Index           =   6
      Left            =   10440
      TabIndex        =   11
      Top             =   8160
      Width           =   2175
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   375
      Index           =   5
      Left            =   6840
      TabIndex        =   10
      Top             =   8160
      Width           =   2175
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   375
      Index           =   4
      Left            =   3360
      TabIndex        =   9
      Top             =   8160
      Width           =   2175
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   375
      Index           =   3
      Left            =   11640
      TabIndex        =   8
      Top             =   4800
      Width           =   2175
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   375
      Index           =   2
      Left            =   8160
      TabIndex        =   7
      Top             =   4800
      Width           =   2175
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   375
      Index           =   1
      Left            =   4800
      TabIndex        =   6
      Top             =   4800
      Width           =   2175
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   375
      Index           =   0
      Left            =   1680
      TabIndex        =   5
      Top             =   4800
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   2655
      Index           =   6
      Left            =   10320
      Stretch         =   -1  'True
      Top             =   5400
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   2655
      Index           =   5
      Left            =   6720
      Stretch         =   -1  'True
      Top             =   5400
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   2655
      Index           =   4
      Left            =   3240
      Stretch         =   -1  'True
      Top             =   5400
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   2655
      Index           =   3
      Left            =   11520
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   2655
      Index           =   2
      Left            =   8160
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   2655
      Index           =   1
      Left            =   4800
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Image Image1 
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
      Left            =   6360
      TabIndex        =   3
      Top             =   1080
      Width           =   4095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
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
      Left            =   5880
      TabIndex        =   2
      Top             =   360
      Width           =   5055
   End
   Begin VB.Label Label1 
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
Private Sub Form_Load()
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
            Image1(i).Visible = True
            Label4(i).Visible = True
        Next i
    End With
End Sub

Sub invicible()
    For i = 0 To 6
        Image1(i).Visible = False
        Label4(i).Visible = False
    Next i
End Sub



Private Sub Image1_Click(Index As Integer)
    Label5.Caption = Index
    
End Sub

