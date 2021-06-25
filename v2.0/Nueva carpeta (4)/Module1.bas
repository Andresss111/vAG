Attribute VB_Name = "Module1"
Global base As New ADODB.Connection
Global TP As New Recordset
Global Clientes As New Recordset
Global Temp As New Recordset
Global ITP As New Recordset
Global Fact As New Recordset
Global DFact As New Recordset
Global TTP As New Recordset
Global TU As New Recordset
Global Tabla1 As New Recordset

Sub main()
    With base
        .CursorLocation = adUseClient
        .Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\base\base.mdb;Persist Security Info=False"
        Form5.Show
    End With
End Sub

Sub CTP()
    With TP
        If .State = 1 Then .Close
        .Open "select * from Producto", base, adOpenStatic, adLockBatchOptimistic
    End With
End Sub

Sub CTEMP()
    With Temp
        If .State = 1 Then .Close
        .Open "select * from Temp", base, adOpenStatic, adLockBatchOptimistic
    End With
End Sub
Sub CTC()
    With Clientes
        If .State = 1 Then .Close
        .Open "select * from Cliente", base, adOpenStatic, adLockBatchOptimistic
    End With
End Sub

Sub CITP()
    With ITP
        If .State = 1 Then .Close
        .Open "select * from IProducto", base, adOpenStatic, adLockBatchOptimistic
    End With
End Sub

Sub CFact()
    With Fact
        If .State = 1 Then .Close
        .Open "select * from Factura", base, adOpenStatic, adLockBatchOptimistic
    End With
End Sub

Sub CDFact()
    With DFact
        If .State = 1 Then .Close
        .Open "select * from Detalle_Factura", base, adOpenStatic, adLockBatchOptimistic
    End With
End Sub

Sub CTTP()
    With TTP
        If .State = 1 Then .Close
        .Open "select * from Tipo_de_Producto", base, adOpenStatic, adLockBatchOptimistic
    End With
End Sub

Sub CTabla1()
    With Tabla1
        If .State = 1 Then .Close
        .Open "select * from Tabla1", base, adOpenStatic, adLockBatchOptimistic
    End With
End Sub

Sub CTU()
    With TU
        If .State = 1 Then .Close
        .Open "select * from Login_Ad", base, adOpenStatic, adLockBatchOptimistic
    End With
End Sub

