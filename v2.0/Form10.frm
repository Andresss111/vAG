VERSION 5.00
Begin VB.Form Form10 
   Caption         =   "Form10"
   ClientHeight    =   2520
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8175
   LinkTopic       =   "Form10"
   ScaleHeight     =   2520
   ScaleWidth      =   8175
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Todos"
      Height          =   375
      Left            =   6720
      TabIndex        =   12
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Buscar"
      Height          =   255
      Left            =   1080
      TabIndex        =   5
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   1680
      TabIndex        =   4
      Top             =   1440
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Buscar"
      Height          =   255
      Left            =   4800
      TabIndex        =   3
      Top             =   2040
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   6480
      TabIndex        =   2
      Top             =   1440
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   4080
      TabIndex        =   1
      Top             =   1440
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      Top             =   930
      Width           =   1455
   End
   Begin VB.Label Label7 
      Caption         =   "Estado de factura"
      Height          =   255
      Left            =   5040
      TabIndex        =   13
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label6 
      Caption         =   "Hasta:"
      Height          =   255
      Left            =   5640
      TabIndex        =   11
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label Label5 
      Caption         =   "Desde:"
      Height          =   255
      Left            =   3240
      TabIndex        =   10
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "Busqueda por fecha aplicar formato xx/xx/xxxx"
      Height          =   255
      Left            =   3240
      TabIndex        =   9
      Top             =   960
      Width           =   4455
   End
   Begin VB.Label Label3 
      Caption         =   "Total mayor a:"
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "RUC:"
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   960
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Buscar por:"
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    If Text2.Text = "" Or Text3.Text = "" Then MsgBox "Rellenar el cuadro de busqueda": Text2.SetFocus: Exit Sub
    z = Command3.Caption
    If Command3.Caption = "Todos" Then z = "%%"
    With Fact
        x = "#" & Text2.Text & "#"
        If .State = 1 Then .Close
        .Open "select * from Factura WHERE ((Factura.[Fecha])= " & x & ")", base, adOpenStatic, adLockBatchOptimistic
        If .EOF Or .BOF Then MsgBox "Fecha (Desde) no existente": Exit Sub
        y = "#" & Text3.Text & "#"
        If .State = 1 Then .Close
        .Open "select * from Factura where ((Factura.[Fecha])= " & y & ")", base, adOpenStatic, adLockBatchOptimistic
        If .EOF Or .BOF Then MsgBox "Fecha (Hasta) no existente": Exit Sub
        If .State = 1 Then .Close
        .Open "select * from Factura WHERE ((Factura.[Fecha])>= " & x & ") AND ((Factura.[Fecha])<= " & y & ") AND [Valido]like '" & z & "'", base, adOpenStatic, adLockBatchOptimistic
    End With
    Set Form9.DataGrid1.DataSource = Fact
End Sub

Private Sub Command2_Click()
    If Text4.Text = "" Then MsgBox "Rellenar el cuadro de busqueda": Text4.SetFocus: Exit Sub
    z = Command3.Caption
    If Command3.Caption = "Todos" Then z = "%%"
    With Fact
        x = Text4.Text
        If .State = 1 Then .Close
        .Open "select * from Factura WHERE ((Factura.Total)> " & x & ") AND [Valido]like '" & z & "'", base, adOpenStatic, adLockBatchOptimistic
        If .EOF Or .BOF Then MsgBox "No hay factura con un mayor total": Exit Sub
    End With
    Set Form9.DataGrid1.DataSource = Fact
End Sub

Private Sub Command3_Click()
    If Command3.Caption = "Todos" Then Command3.Caption = "True": Exit Sub
    If Command3.Caption = "True" Then Command3.Caption = "False": Exit Sub
    If Command3.Caption = "False" Then Command3.Caption = "Todos": Exit Sub
End Sub

Private Sub Text1_Change()
    z = Command3.Caption
    If Command3.Caption = "Todos" Then z = "%%"
    CFact
    With Fact
        x = Text1.Text
        If .State = 1 Then .Close
        .Open "select * from Factura where [Id_C]like '" & x & "' and [Valido]like '" & z & "'", base, adOpenStatic, adLockBatchOptimistic
        If .EOF Or .BOF Then Exit Sub
        Form9.Label3.Caption = x
        Me.Hide
    End With
    Form9.Label2 = "T"
    Form9.Label4.Caption = "1"
    Set Form9.DataGrid1.DataSource = Fact
End Sub
