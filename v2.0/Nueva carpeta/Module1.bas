Attribute VB_Name = "Module1"
Global base As New ADODB.Connection
Global TP As New Recordset

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
