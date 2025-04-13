'Subrutina que selecciona una sugerencia del campo de resultados y lo ubica en el campo de registro:
Private Sub ListBox1Sugerencias_Click()
    
    Me.TextBuscar.Value = Me.ListBox1Sugerencias.Value
    Me.ListBox1Sugerencias.Clear ' Limpia el ListBox

End Sub


'Subrutina que busca el criterio digitado en el campo de registro:
Private Sub TextBuscar_Change()

    Dim textoBuscar As String
    Dim celda As Range
    Dim hojaDatos As Worksheet

    textoBuscar = Me.TextBuscar.Value
    Set hojaDatos = ThisWorkbook.Sheets("Tablas")

    Me.ListBox1Sugerencias.Clear 'Limpia las sugerencias anteriores

    If textoBuscar <> "" Then
        For Each celda In hojaDatos.Range("A1:F" & hojaDatos.Cells(hojaDatos.Rows.Count, "A").End(xlUp).Row)
            If UCase(celda.Value) Like UCase(textoBuscar & "*") Then
                Me.ListBox1Sugerencias.AddItem celda.Value
            End If
        Next celda
    End If

End Sub


'Subrutina de inicialización:
Private Sub RegistroProducto_Initialize()

Me.ListBox1Sugerencias.ColumnCount = 6 'Cantidad de columnas que se visualizan.
Me.ListBox1Sugerencias.ColumnWidths = "100;100;100;100;100;100" 'Tamaño de las columnas visualizadas.
Me.ListBox1Sugerencias.RowSource = "Tabla" 'Origen.

End Sub


Private Sub TextBuscar_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    If KeyCode = vbKeyReturn Then
        ' Si el usuario presiona Enter, toma la primera sugerencia (si hay alguna)
        If Me.ListBox1Sugerencias.ListCount > 0 Then
            Me.TextBuscar.Value = Me.ListBox1Sugerencias.List(0)
            Me.ListBox1Sugerencias.Clear
        End If
    End If
    
End Sub
