VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RegistroProducto 
   Caption         =   "M�dulo de Ventas"
   ClientHeight    =   7275
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15195
   OleObjectBlob   =   "RegistroProducto.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "RegistroProducto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

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

'Subrutina de inicializaci�n:

Private Sub RegistroProducto_Initialize()

Me.ListBox1Sugerencias.ColumnCount = 6 'Cantidad de columnas que se visualizan.
Me.ListBox1Sugerencias.ColumnWidths = "100;100;100;100;100;100" 'Tama�o de las columnas visualizadas.
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
