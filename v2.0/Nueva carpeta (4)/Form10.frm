VERSION 5.00
Begin VB.Form Form10 
   Caption         =   "Form10"
   ClientHeight    =   1155
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form10"
   ScaleHeight     =   1155
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   3015
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Text1_Change()
    CFact
    With Fact
        x = Text1.Text
        If .State = 1 Then .Close
        y = "True"
        .Open "select * from Factura where [Id_C]like '" & x & "' and [Valido]like '" & y & "'", base, adOpenStatic, adLockBatchOptimistic
        If .EOF Or .BOF Then Exit Sub
        Form9.Label3.Caption = x
        Me.Hide
    End With
    Form9.Label2 = "T"
    Set Form9.DataGrid1.DataSource = Fact
End Sub
