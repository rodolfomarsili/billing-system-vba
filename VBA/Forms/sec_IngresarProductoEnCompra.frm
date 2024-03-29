VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} sec_IngresarProductoEnCompra 
   ClientHeight    =   7140
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   12300
   OleObjectBlob   =   "sec_IngresarProductoEnCompra.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "sec_IngresarProductoEnCompra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CommandButton_IngresarProducto_Click()

Dim NuevaExistencia As Long
Dim i As Integer
Dim a As Integer
Dim b As Integer
Dim Cantidad As Long
Dim Costo As Single
Dim CostoUnitario As Single
Dim Importe As Single
Dim IngresarUnidades As Byte
Dim ModificarCosto As Byte
Dim FilaAModificar As Integer
Dim Codigo As String

    
    'Se obtiene el largo del listado de productos
    a = FormularioAnterior.ListBox_Listado.ListCount
        
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Verificaciones'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Inicializar
        
    ' Se verifica que la casilla de cantidad no este vacia
    If (TextBox_Cantidad.Value = Empty) Then
        MsgBox "Ingresa la cantidad a facturar", , "Compras"
        Exit Sub
    End If
    
    'Nueva existencia
    NuevaExistencia = Val(Label_Existencia.Caption) + TextBox_Cantidad.Value
    
    If Val(TextBox_Costo) = 0 Then
        MsgBox "Debes modificar el precio de este producto antes de agregarlo a la factura", , "Compras"
        Exit Sub
    End If
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Modificar Costo del Producto'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        If Not Val(TextBox_Costo.Text) = Val(Label_Costo.Caption) Then

                CostoUnitario = Val(TextBox_Costo.Text) / Val(Label_UnidadesPorBulto.Caption)
                ListBox_ListadoProductos.List(ListBox_ListadoProductos.ListIndex, 2) = Format(CostoUnitario, "0.0000")
                
                ModificarCosto = MsgBox("Has modificado el costo de este produto, �Deseas guardar el cambio?", vbYesNo + vbExclamation, "Compras")
                If ModificarCosto = vbYes Then


                        FilaAModificar = ObtenerFila(HojaInventario, ListBox_ListadoProductos.List(ListBox_ListadoProductos.ListIndex, 0), ColumnaCodigo)
                        'FilaAModificarCliente = ObtenerFila(HojaBaseClientes, TextBox_Codigo.Text, ColumnaCodigoCliente)
                        HojaInventario.Cells(FilaAModificar, ColumnaCostoBulto) = TextBox_Costo.Value

                End If
        End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Llenado de Formulario'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
     'Verificacion de item repetido
     If a > 0 Then
        For i = 0 To a - 1
        
            Codigo = FormularioAnterior.ListBox_Listado.List(i, 1)
            If ListBox_ListadoProductos.List(ListBox_ListadoProductos.ListIndex, 0) = Codigo Then
            
                
            IngresarUnidades = MsgBox("Ya has ingresado este articulo, �Deseas a�adir mas unidades a la factura?", vbYesNo + vbExclamation, "Compras")
            If IngresarUnidades = vbNo Then Exit Sub
            
            'Se suma la nueva cantidad ingresada a la que ya existe previamente en el listado
            Cantidad = Val(FormularioAnterior.ListBox_Listado.List(i, 3)) + Val(TextBox_Cantidad)
            Costo = Val(ListBox_ListadoProductos.List(ListBox_ListadoProductos.ListIndex, 2))
            Importe = Cantidad * Costo
            NuevaExistencia = Val(FormularioAnterior.ListBox_Listado.List(i, 0)) + Val(TextBox_Cantidad)
            
            
            FormularioAnterior.ListBox_Listado.List(i, 0) = NuevaExistencia
            FormularioAnterior.ListBox_Listado.List(i, 3) = Cantidad
            FormularioAnterior.ListBox_Listado.List(i, 4) = Format(Costo, "0.0000")
            FormularioAnterior.ListBox_Listado.List(i, 5) = Format(Importe, "0.00")
            
            ActualizarSubTotal
        
            Unload Me
            Exit Sub
        
        
            End If
        Next i
    End If
    
    'Se a�ade un nuevo item con las columnas respectivas
        FormularioAnterior.ListBox_Listado.AddItem
        FormularioAnterior.ListBox_Listado.List(a, 1) = ListBox_ListadoProductos.List(ListBox_ListadoProductos.ListIndex, 0) 'Codigo
        FormularioAnterior.ListBox_Listado.List(a, 2) = ListBox_ListadoProductos.List(ListBox_ListadoProductos.ListIndex, 1) 'Producto
        FormularioAnterior.ListBox_Listado.List(a, 3) = TextBox_Cantidad
        
        Costo = Val(ListBox_ListadoProductos.List(ListBox_ListadoProductos.ListIndex, 2)) 'Costo
        FormularioAnterior.ListBox_Listado.List(a, 4) = Format(Costo, "0.0000")
        
        FormularioAnterior.ListBox_Listado.List(a, 5) = Format(Val(TextBox_Cantidad.Text) * Costo, "0.00")
        FormularioAnterior.ListBox_Listado.List(a, 0) = NuevaExistencia
        
        
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Limpieza de Formulario'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        ActualizarSubTotal
        
        Unload Me

    
End Sub


Private Sub ListBox_ListadoProductos_Click()

Dim i As Integer
Dim a As Integer
Dim FilaDeCodigo As Integer

    Inicializar
    
    a = FormularioAnterior.ListBox_Listado.ListCount
    
    On Error Resume Next
    
    FilaDeCodigo = ObtenerFila(HojaInventario, ListBox_ListadoProductos.List(ListBox_ListadoProductos.ListIndex, 0), ColumnaCodigo)

    Label_Producto.Caption = HojaInventario.Cells(FilaDeCodigo, ColumnaProducto)
    Label_Presentacion.Caption = HojaInventario.Cells(FilaDeCodigo, ColumnaPresentacion)
    TextBox_Costo = Format(HojaInventario.Cells(FilaDeCodigo, ColumnaCostoBulto), "0.00")
    Label_Costo = TextBox_Costo
    Label_UnidadesPorBulto.Caption = HojaInventario.Cells(FilaDeCodigo, ColumnaUnidadesPorBulto)
    Label_Existencia.Caption = HojaInventario.Cells(FilaDeCodigo, ColumnaExistencia)
    
    'Verificacion de nueva existencia
    If a > 0 Then
        For i = 0 To a - 1
            If ListBox_ListadoProductos.List(ListBox_ListadoProductos.ListIndex, 0) = FormularioAnterior.ListBox_Listado.List(i, 1) Then
                Label_NuevaExistencia.Visible = True
                Label_NuevaExistencia.Caption = "Nueva Existencia: " & FormularioAnterior.ListBox_Listado.List(i, 0)
                Exit For
            Else
                Label_NuevaExistencia.Visible = False
            End If
        Next i
    End If
    
    'Cambiar color a la etiqueta de existencia
    If Val(Label_Existencia) = 0 Then
        Label_Existencia.ForeColor = &HFF&
    Else
        Label_Existencia.ForeColor = &H80000012
    End If
    
   
    
End Sub

Private Sub TextBox_Costo_Enter()
        With TextBox_Costo
                .SetFocus
                .SelStart = 0
                .SelLength = Len(.Text)
         End With
End Sub


Private Sub ListBox_ListadoProductos_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
        With TextBox_Costo
                .SetFocus
                .SelStart = 0
                .SelLength = Len(.Text)
         End With
End Sub

Private Sub ListBox_ListadoProductos_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then Unload Me
End Sub

Private Sub TextBox_Buscar_Change()
    FiltrarProductoEnListBox
End Sub

Private Sub TextBox_Cantidad_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    If KeyAscii = 27 Then Unload Me
    
    If Not (KeyAscii >= 48 And KeyAscii <= 57) Then
        KeyAscii = 0
    End If
End Sub


Private Sub TextBox_Costo_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
Dim UbicacionPunto As Integer
    If KeyAscii = 27 Then Unload Me
    UbicacionPunto = InStr(TextBox_Costo.Text, ".")
    If (KeyAscii = 46 And UbicacionPunto > 0) Then
        KeyAscii = 0
    End If
    If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46) Then
        KeyAscii = 0
    End If
End Sub

Private Sub UserForm_Initialize()

Dim i As Integer
Dim a As Integer
        
    Inicializar
    
    Label_NuevaExistencia.Visible = False
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        ' Formato del listado
        ListBox_ListadoProductos.ColumnCount = 3
        ListBox_ListadoProductos.ColumnWidths = "100 pt; 298 pt; 60 pt"
    
        a = 0
        'Lenado del listado
        For i = 2 To UltimaFilaInventario
            ListBox_ListadoProductos.AddItem
            ListBox_ListadoProductos.List(a, 0) = HojaInventario.Cells(i, ColumnaCodigo)
            ListBox_ListadoProductos.List(a, 1) = HojaInventario.Cells(i, ColumnaProducto)
            ListBox_ListadoProductos.List(a, 2) = Format(HojaInventario.Cells(i, ColumnaCostoUnidad), "0.0000")
            
            a = a + 1
        Next i
    
     
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    
End Sub


Private Sub UserForm_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then Unload Me
End Sub

Sub FiltrarProductoEnListBox()

Dim i As Integer
Dim a As Integer
Dim Producto As String
Dim Codigo As String

    Inicializar
    
    ListBox_ListadoProductos.Clear
    
    If TextBox_Buscar = "" Then
        
        a = 0
        For i = 2 To UltimaFilaInventario
        
            ListBox_ListadoProductos.AddItem
            ListBox_ListadoProductos.List(a, 0) = HojaInventario.Cells(i, ColumnaCodigo)
            ListBox_ListadoProductos.List(a, 1) = HojaInventario.Cells(i, ColumnaProducto)
            ListBox_ListadoProductos.List(a, 2) = Format(HojaInventario.Cells(i, ColumnaCostoUnidad), "0.0000")
            
            a = a + 1
            
        Next i
        
        Exit Sub
        
    End If
    
    For i = 2 To UltimaFilaInventario
        Producto = HojaInventario.Cells(i, ColumnaProducto)
        Codigo = HojaInventario.Cells(i, ColumnaCodigo)
        
        If UCase(Producto) Like "*" & UCase(TextBox_Buscar.Value) & "*" Then
        
            ListBox_ListadoProductos.AddItem
            ListBox_ListadoProductos.List(a, 0) = HojaInventario.Cells(i, ColumnaCodigo)
            ListBox_ListadoProductos.List(a, 1) = HojaInventario.Cells(i, ColumnaProducto)
            ListBox_ListadoProductos.List(a, 2) = Format(HojaInventario.Cells(i, ColumnaCostoUnidad), "0.0000")
            
            a = a + 1
            
        'He a�adido todo este fragmento para que me busque al mismo tiempo por codigo.
        ElseIf Codigo Like "*" & UCase(TextBox_Buscar.Value) & "*" Then
        
            ListBox_ListadoProductos.AddItem
            ListBox_ListadoProductos.List(a, 0) = HojaInventario.Cells(i, ColumnaCodigo)
            ListBox_ListadoProductos.List(a, 1) = HojaInventario.Cells(i, ColumnaProducto)
            ListBox_ListadoProductos.List(a, 2) = Format(HojaInventario.Cells(i, ColumnaCostoUnidad), "0.0000")
            
            a = a + 1
            
        End If
        
    Next i


End Sub
