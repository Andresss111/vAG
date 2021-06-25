Attribute VB_Name = "Module1"
Global base As New ADODB.Connection
Global TP As New Recordset
Global Clientes As New Recordset
Global Temp As New Recordset
Global ITP As New Recordset
Global Fact As New Recordset
Global DFact As New Recordset

Sub main()
    With base
        .CursorLocation = adUseClient
        .Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\base\base.mdb;Persist Security Info=False"
        Form1.Show
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

Sub copiar()
    Dim x, a, b, d As String
    Dim e As Integer
    e = 1
    CTP
    With TP
        d = .RecordCount
    End With
        For i = 1 To d
            CTP
            With TP
                If i = 1 Then
                    .MoveFirst
                Else
                    For x = 1 To e
                    .MoveNext
                    Next x
                End If
                If .EOF Or .BOF Then
                Else
                    a = !Id_Producto
                    b = !Etiqueta
                    c = !URL
                    e = e + 1
                End If
            End With
            CITP
            With ITP
                .AddNew
                !Id_Producto = a
                !Etiqueta = b
                !URL = c
                .UpdateBatch
            End With
        Next i
End Sub
