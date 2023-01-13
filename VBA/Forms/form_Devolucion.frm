VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form_Devolucion 
   Caption         =   "Devolucion"
   ClientHeight    =   7770
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   18120
   OleObjectBlob   =   "form_Devolucion.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "form_Devolucion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CheckBox_AnadirACredito_Click()
    If CheckBox_AnadirACredito.Value = True Then
        Frame_Caja.Enabled = False
        ComboBox_Caja.Enabled = False
        Label_SaldoCaja.Enabled = False
    Else
        Frame_Caja.Enabled = True
        ComboBox_Caja.Enabled = True
        Label_SaldoCaja.Enabled = True
    End If
End Sub

Private Sub ComboBox_Caja_Change()

    Label_AsteriscoCaja.Visible = False
    
    'Actualizar monto de caja en pantalla
    ActualizarSaldoCajaEnPantalla Label_SaldoCaja, ComboBox_Caja.Text
    
End Sub


Private Sub ComboBox_TipoDeTransaccionID_Change()

Dim FilaCorrelativo As Byte
    
    Inicializar
    
    FilaCorrelativo = ObtenerFila(HojaCorrelativos, ComboBox_TipoDeTransaccionID.Text, ColumnaLeyenda)
    Label_TipoDeTransaccion.Caption = HojaCorrelativos.Cells(FilaCorrelativo, ColumnaPrefijo)
    
    If Not Label_TipoDeTransaccion.Caption = "VTA-CTD" Then
        CheckBox_AnadirACredito.Value = False
        CheckBox_AnadirACredito.Enabled = False
        
        Frame_Caja.Enabled = False
        ComboBox_Caja.Enabled = False
        Label_SaldoCaja.Enabled = False
    Else
        CheckBox_AnadirACredito.Enabled = True
        
        Frame_Caja.Enabled = True
        ComboBox_Caja.Enabled = True
        Label_SaldoCaja.Enabled = True
    End If
    
End Sub

Private Sub CommandButton_Aceptar_Click()

Dim Codigo As String
Dim Devolviendo As Integer
Dim Producto As String
Dim Cantidad As Long
Dim Comentario As String
Dim Precio As Single
            
Dim Procesar As Byte
Dim FilaCliente As Integer
Dim FilaDevolucion As Long
Dim a As Long
Dim i As Long
Dim j As Long
Dim FilaCaja As Byte
Dim IDResponsable As String
Dim IDCaja As String
Dim IDCliente As String
Dim Fecha As Date

  
    Fecha = TextBox_Dia & "/" & TextBox_Mes & "/" & TextBox_Ano

    Inicializar
    
    Label_AsteriscoTipoDeTransaccion.Visible = False
    Label_AsteriscoID1.Visible = False
    Label_AsteriscoID2.Visible = False
    Label_AsteriscoTipoDeTransaccion.Visible = False
    Label_AsteriscoCaja.Visible = False
    
    FilaCaja = ObtenerFila(HojaCajas, ComboBox_Caja.Text, ColumnaIDCaja)
    
    a = ListBox_Devoluciones.ListCount
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Verificaciones''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If ComboBox_TipoDeTransaccionID.ListIndex = -1 Then
        Label_AsteriscoTipoDeTransaccion.Visible = True
        MsgBox "Selecciona un tipo de transaccion valida", , "Devolucion"
        Exit Sub
    End If
    
    If ComboBox_TipoDeTransaccionID = Empty Or TextBox_ID1 = Empty Or TextBox_ID2 = Empty Then
        Label_AsteriscoTipoDeTransaccion.Visible = True
        Label_AsteriscoID1.Visible = True
        Label_AsteriscoID2.Visible = True
        MsgBox "Debes rellenar todos los campos de transaccion", , "Devolucion"
        Exit Sub
    End If
    
    If a = 0 Then
        MsgBox "No hay productos agegados a la lista de devolucion", , "Devolucion"
        Label_Cantidad_Auxiliar = Empty
        Exit Sub
    End If
    
    If FilaCaja = 0 Then
        Label_AsteriscoCaja.Visible = True
        MsgBox "Selecciona una caja valida", , "Devolucion"
        Exit Sub
    End If
    
    FilaCliente = ObtenerFila(HojaClientes, Label_Cliente.Caption, ColumnaNombreCliente)
    
    IDResponsable = HojaCajas.Cells(FilaCaja, ColumnaIDResponsableCaja)
    IDCliente = HojaClientes.Cells(FilaCliente, ColumnaIDCliente)
    
    'Ultima verificacion antes de procesar la factura
    Procesar = MsgBox("¿Seguro que deseas procesar esta transaccion?", vbYesNo + vbExclamation, "Devolucion")
    If Procesar = vbNo Then Exit Sub
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        
        '''''''''''''''''''''''''''''''''''''''''''''''''''Devolucion de contado''''''''''''''''''''''''''''''''''''''''''''''''''
        
        If Label_TipoDeTransaccion.Caption = "VTA-CTD" Then
        
            'Añadir el dinero a caja o al credito del cliente segun se seleccione
            If CheckBox_AnadirACredito.Value = True Then
                
                HojaClientes.Cells(FilaCliente, ColumnaSaldoCreditoCliente) = HojaClientes.Cells(FilaCliente, ColumnaSaldoCreditoCliente) - Val(Label_Importe.Caption)
            
                Comentario = "[" & Label_TipoDeTransaccion.Caption & "-" & TextBox_ID1.Text & "-" & TextBox_ID2.Text & "]"
                Comentario = Comentario + Chr(13) + "[" & "Monto abonado al credito del cliente" & "]"
                
                
                        'Establecer el numero de unidades devueltas
                        For i = a - 1 To 0 Step -1
                            
                            Codigo = ListBox_Devoluciones.List(i, 0)
                            Producto = ListBox_Devoluciones.List(i, 1)
                            Precio = Val(ListBox_Devoluciones.List(i, 2))
                            Cantidad = Val(ListBox_Devoluciones.List(i, 3))
                            
                            IncluirEnHistorial Frame_Correlativo, Label_CorrelativoPrefijo, Fecha, Codigo, Producto, ComboBox_Caja.Text, Cantidad, Comentario, IDCliente, IDResponsable, Precio
                        
                        Next i
                
            Else
            
                Comentario = "[" & Label_TipoDeTransaccion.Caption & "-" & TextBox_ID1.Text & "-" & TextBox_ID2.Text & "]"
                
                        'Establecer el numero de unidades devueltas
                        For i = a - 1 To 0 Step -1
                            
                            Codigo = ListBox_Devoluciones.List(i, 0)
                            Producto = ListBox_Devoluciones.List(i, 1)
                            Precio = Val(ListBox_Devoluciones.List(i, 2))
                            Cantidad = Val(ListBox_Devoluciones.List(i, 3))
                            
                            IncluirEnHistorial Frame_Correlativo, Label_CorrelativoPrefijo, Fecha, Codigo, Producto, ComboBox_Caja.Text, Cantidad, Comentario, IDCliente, IDResponsable, Precio, False
                        
                        Next i
                
            End If
        
        End If
        
        
        '''''''''''''''''''''''''''''''''''''''''''''''''''Devolucion de credito''''''''''''''''''''''''''''''''''''''''''''''''''
        
        If Label_TipoDeTransaccion.Caption = "VTA-CDT" Then
        
            HojaClientes.Cells(FilaCliente, ColumnaSaldoCreditoCliente) = HojaClientes.Cells(FilaCliente, ColumnaSaldoCreditoCliente) - Val(Label_Importe.Caption)
                
            Comentario = "[" & Label_TipoDeTransaccion.Caption & "-" & TextBox_ID1.Text & "-" & TextBox_ID2.Text & "]"
            
                'Establecer el numero de unidades devueltas
                        For i = a - 1 To 0 Step -1
                            
                            Codigo = ListBox_Devoluciones.List(i, 0)
                            Producto = ListBox_Devoluciones.List(i, 1)
                            Precio = 0
                            Cantidad = Val(ListBox_Devoluciones.List(i, 3))
                            
                            IncluirEnHistorial Frame_Correlativo, Label_CorrelativoPrefijo, Fecha, Codigo, Producto, ComboBox_Caja.Text, Cantidad, Comentario, IDCliente, IDResponsable, Precio
                        
                        Next i
        End If
        
        
        '''''''''''''''''''''''''''''''''''''''''''''''''''Devolucion de consignacion''''''''''''''''''''''''''''''''''''''''''''''''''
        
        If Label_TipoDeTransaccion.Caption = "VTA-CSN" Then
        
        'Eliminacion de las existencias en el inventario de consignacion del cliente seleccionado
        For i = 0 To a - 1
            For j = 2 To UltimaFilaInventario
                If ListBox_Devoluciones.List(i, 0) = LibroClientes.Sheets(IDCliente).Cells(j, ColumnaCodigoCliente).Text Then
                    
                    LibroClientes.Sheets(IDCliente).Cells(j, ColumnaExistenciaCliente) = LibroClientes.Sheets(IDCliente).Cells(j, ColumnaExistenciaCliente) - Val(ListBox_Devoluciones.List(i, 3))
                    Exit For
                    Exit For

                End If
            Next j
        Next i
                
            Comentario = "[" & Label_TipoDeTransaccion.Caption & "-" & TextBox_ID1.Text & "-" & TextBox_ID2.Text & "]"
            
                'Establecer el numero de unidades devueltas
                        For i = a - 1 To 0 Step -1
                            
                            Codigo = ListBox_Devoluciones.List(i, 0)
                            Producto = ListBox_Devoluciones.List(i, 1)
                            Precio = 0
                            Cantidad = Val(ListBox_Devoluciones.List(i, 3))
                            
                            IncluirEnHistorial Frame_Correlativo, Label_CorrelativoPrefijo, Fecha, Codigo, Producto, ComboBox_Caja.Text, Cantidad, Comentario, IDCliente, IDResponsable, Precio
                        
                        Next i
        End If
        
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Devolver''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        'Añadir los productos devueltos al inventario
        For i = 0 To a - 1
            For j = 2 To UltimaFilaInventario
                If Val(ListBox_Devoluciones.List(i, 0)) = HojaInventario.Cells(j, ColumnaCodigo) Then
                    HojaInventario.Cells(j, ColumnaExistencia) = HojaInventario.Cells(j, ColumnaExistencia) + Val(ListBox_Devoluciones.List(i, 3))
                    Exit For
                End If
            Next j
        Next i
        
        
        
        'Establecer el numero de unidades devueltas
        For i = 0 To a - 1
            
            Codigo = ListBox_Devoluciones.List(i, 0)
            Devolviendo = Val(ListBox_Devoluciones.List(i, 3))
            
            FilaDevolucion = ObtenerFilaProductoDeTransaccion(Label_TipoDeTransaccion.Caption, TextBox_ID1.Text, TextBox_ID2.Text, Codigo)
            HojaHistorial.Cells(FilaDevolucion, ColumnaDevueltoHistorial) = HojaHistorial.Cells(FilaDevolucion, ColumnaDevueltoHistorial) + Devolviendo
            
        Next i

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Limpieza de formulario''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        'Actualizar saldo de caja en pantalla
        ActualizarSaldoCajaEnPantalla Label_SaldoCaja, ComboBox_Caja.Text
        
        ActualizarCorrelativo Frame_Correlativo, Label_CorrelativoPrefijo
        
        ActualizarCorrelativoEnPantalla Frame_Correlativo, Label_CorrelativoPrefijo
        
        LimpiarFormulario
        
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        
        ActualizarDashboard
        
        MsgBox "Devolucion realizada exitosamente", , "Devolucion"

End Sub


Private Sub CommandButton_Buscar_Click()

Dim i As Long
Dim a As Long
Dim FilaCliente As Integer
Dim FilaResponsable As Byte
    
    Inicializar
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Verificaciones''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Ocultar todos los asteriscos
    Label_AsteriscoTipoDeTransaccion.Visible = False
    Label_AsteriscoID1.Visible = False
    Label_AsteriscoID2.Visible = False
    
    'Limpieza de tabla de historial temporal
    If Not HojaHistorialTemporal.Range("A1") = Empty Then
        HojaHistorialTemporal.Range("A1").CurrentRegion.Delete
    End If
    
    'Limpieza de las tablas de filtros
    Workbooks("Historial.xlsm").Sheets("Inicio").Range("A101:G101").Clear
    Workbooks("Historial.xlsm").Sheets("Inicio").Range("A111:G111").Clear
    
    If Not ComboBox_TipoDeTransaccionID.ListIndex > -1 Then
            MsgBox "Debes seleccionar una opcion valida para el Tipo de Transaccion ", , "Devolucion"
            Exit Sub
    End If
    
    If (ComboBox_TipoDeTransaccionID = Empty Or TextBox_ID1 = Empty Or TextBox_ID2 = Empty) Then
        Label_AsteriscoTipoDeTransaccion.Visible = True
        Label_AsteriscoID1.Visible = True
        Label_AsteriscoID2.Visible = True
        MsgBox "Debes rellenar TODOS los campos del recuadro de transaccion", , "Devolucion"
        Exit Sub
    End If
    
        
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
   
    'Limpiar lista de devolucion
    ListBox_Devoluciones.Clear
    Label_Importe = Empty
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Tabla Historial Temporal''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    'Filtro Tipo de Transaccion
    Workbooks("Historial.xlsm").Sheets("Inicio").Range("A111") = Label_TipoDeTransaccion.Caption
    'Filtro ID1
    Workbooks("Historial.xlsm").Sheets("Inicio").Range("B111") = TextBox_ID1
    'Filtro ID2
    Workbooks("Historial.xlsm").Sheets("Inicio").Range("C111") = TextBox_ID2
        
    'LLenado de tabla de historial temporal
    FiltrarHistorialPorTransaccion

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Formulario Historial Temporal''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Actualizar ultima fila del historial temporal
    UltimaFilaHistorialTemporal = ObtenerUltimaFila(HojaHistorialTemporal, ColumnaFechaHistorial)
    
     'Limpieza de tabla de historial en el formulario
    ListBox_Historial.Clear
        
    'Llenado de tabla de historial en el formulario
    a = 0
    For i = 2 To UltimaFilaHistorialTemporal
        
            ListBox_Historial.AddItem
            ListBox_Historial.List(a, 0) = HojaHistorialTemporal.Cells(i, ColumnaCodigoHistorial)
            ListBox_Historial.List(a, 1) = HojaHistorialTemporal.Cells(i, ColumnaProductoHistorial)
            ListBox_Historial.List(a, 2) = HojaHistorialTemporal.Cells(i, ColumnaCantidadHistorial)
            ListBox_Historial.List(a, 3) = Format(HojaHistorialTemporal.Cells(i, ColumnaPrecioHistorial), "0.00")
            ListBox_Historial.List(a, 4) = Format(HojaHistorialTemporal.Cells(i, ColumnaImporteHistorial), "0.00")
            ListBox_Historial.List(a, 5) = HojaHistorialTemporal.Cells(i, ColumnaDevueltoHistorial)
            ListBox_Historial.List(a, 6) = HojaHistorialTemporal.Cells(i, ColumnaCantidadHistorial) - HojaHistorialTemporal.Cells(i, ColumnaDevueltoHistorial)
            
    
            a = a + 1
            
    Next i
    
    If UltimaFilaHistorialTemporal > 1 Then
    
            Label_Fecha.Caption = HojaHistorialTemporal.Cells(UltimaFilaHistorialTemporal, ColumnaFechaHistorial)
            Label_Hora.Caption = FormatDateTime(HojaHistorialTemporal.Cells(UltimaFilaHistorialTemporal, ColumnaHoraHistorial), vbShortTime)
            Label_Caja.Caption = HojaHistorialTemporal.Cells(UltimaFilaHistorialTemporal, ColumnaIDCajaHistorial)
            
            FilaCliente = ObtenerFila(HojaClientes, HojaHistorialTemporal.Cells(UltimaFilaHistorialTemporal, ColumnaIDClienteHistorial), ColumnaIDCliente)
            Label_Cliente.Caption = HojaClientes.Cells(FilaCliente, ColumnaNombreCliente)
            
            FilaResponsable = ObtenerFila(HojaUsuarios, HojaHistorialTemporal.Cells(UltimaFilaHistorialTemporal, ColumnaIDResponsableHistorial), ColumnaIDUsuario)
            Label_Responsable.Caption = HojaUsuarios.Cells(FilaResponsable, ColumnaNombreUsuario)
    End If
            
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    ''''''''''''''''''''''''''''''''''''''''''''''''''''

        
End Sub

Private Sub CommandButton_Anadir_Click()
    Anadir
End Sub



Private Sub ListBox_Historial_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Anadir
End Sub

Sub Anadir()

Dim a As Integer
Dim i As Integer
Dim Restante As Long
Dim NuevoImporteHistorial As Single
Dim NuevoImporteDevoluciones As Single
    
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Verificaciones'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'No hacer nada si no hay algun item seleccionado en la lista de consignaciones
    If ListBox_Historial.ListIndex = -1 Then Exit Sub
    
    'Mostrar el formaulario de añadir cantidad
    sec_Cantidad.Show
    
    'Si la cantidad a mover es 0, no se hace nada
    If Val(Label_Cantidad_Auxiliar) = 0 Then Exit Sub
    
    'Verificacion de si existe el articulo agregado en la lista PorPagar
    a = ListBox_Devoluciones.ListCount
    If a > 0 Then
        For i = 0 To a - 1
            If ListBox_Historial.List(ListBox_Historial.ListIndex, 0) = ListBox_Devoluciones.List(i, 0) Then
                MsgBox "Ya has ingresado este articulo, eliminalo de la lista de devoluciones y vuelve a agregarlo", , "Devolucion"
                Label_Cantidad_Auxiliar = Empty
                Exit Sub
            End If
        Next i
    End If
    
    'Calculo del restante
    Restante = ListBox_Historial.List(ListBox_Historial.ListIndex, 6) - Val(Label_Cantidad_Auxiliar)
    
    If Restante < 0 Then
        MsgBox "No puedes devolver mas unidades de las que fueron vendidas", , "Devolucion"
        Label_Cantidad_Auxiliar = Empty
        Exit Sub
    End If


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Llenar listado PorPagar'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    
    
        a = ListBox_Devoluciones.ListCount
        ListBox_Devoluciones.AddItem
        ListBox_Devoluciones.List(a, 0) = ListBox_Historial.List(ListBox_Historial.ListIndex, 0)                                                                'Codigo
        ListBox_Devoluciones.List(a, 1) = ListBox_Historial.List(ListBox_Historial.ListIndex, 1)                                                                'Producto
        ListBox_Devoluciones.List(a, 2) = ListBox_Historial.List(ListBox_Historial.ListIndex, 3)                                                                'Precio
        ListBox_Devoluciones.List(a, 3) = Val(Label_Cantidad_Auxiliar)                                                                                          'Devolviendo
        ListBox_Devoluciones.List(a, 4) = Format(Val(ListBox_Historial.List(ListBox_Historial.ListIndex, 3)) * Val(Label_Cantidad_Auxiliar), "0.00")            'Importe
        
        'Estos campos no se ven en la tabla de devoluciones
        ListBox_Devoluciones.List(a, 5) = ListBox_Historial.List(ListBox_Historial.ListIndex, 2)                                                                'Cantidad
        ListBox_Devoluciones.List(a, 6) = ListBox_Historial.List(ListBox_Historial.ListIndex, 5)                                                                'Devuelto
        
    If Restante = 0 Then 'Se ejecuta cuado se añaden TODAS las existencias del item seleccionado
    
        ListBox_Historial.RemoveItem (ListBox_Historial.ListIndex)
        ListBox_Historial.ListIndex = -1
    
    End If
     
        
    If Restante > 0 Then 'Se ejecuta cuado se añaden ALGUNAS las existencias del item seleccionado
    
        ListBox_Historial.List(ListBox_Historial.ListIndex, 6) = Restante        'Restante historial
        ListBox_Historial.ListIndex = -1
      
    End If
    
    Label_Cantidad_Auxiliar.Caption = Empty
    
    ActualizarImporte
    
End Sub

Private Sub CommandButton_Quitar_Click()
    Quitar
End Sub

Private Sub ListBox_Devoluciones_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Quitar
End Sub

Sub Quitar()

Dim i As Integer
Dim a As Integer

    'No hacer nada si no hay algun item seleccionado en la lista de pagados
    If ListBox_Devoluciones.ListIndex = -1 Then Exit Sub
    
    'Verificacion de si existe el articulo agregado en la lista de consignaciones
    a = ListBox_Historial.ListCount
    
    If a > 0 Then
        For i = 0 To a - 1
            If ListBox_Devoluciones.List(ListBox_Devoluciones.ListIndex, 0) = ListBox_Historial.List(i, 0) Then 'Si existe
            
                ListBox_Historial.List(i, 6) = Val(ListBox_Historial.List(i, 6)) + Val(ListBox_Devoluciones.List(ListBox_Devoluciones.ListIndex, 3))
                
                ListBox_Devoluciones.RemoveItem (ListBox_Devoluciones.ListIndex)
                ActualizarImporte
                Exit Sub
                
            End If
        Next i
    End If
    
    'No existe en el listado de consignaciones
    ListBox_Historial.AddItem
    ListBox_Historial.List(a, 0) = ListBox_Devoluciones.List(ListBox_Devoluciones.ListIndex, 0)                                                                                             'Codigo
    ListBox_Historial.List(a, 1) = ListBox_Devoluciones.List(ListBox_Devoluciones.ListIndex, 1)                                                                                             'Producto
    ListBox_Historial.List(a, 2) = ListBox_Devoluciones.List(ListBox_Devoluciones.ListIndex, 5)                                                                                             'Cantidad
    ListBox_Historial.List(a, 3) = ListBox_Devoluciones.List(ListBox_Devoluciones.ListIndex, 2)                                                                                             'Precio
    
    ListBox_Historial.List(a, 4) = Format(Val(ListBox_Devoluciones.List(ListBox_Devoluciones.ListIndex, 2)) * Val(ListBox_Devoluciones.List(ListBox_Devoluciones.ListIndex, 5)), "0.00")    'Importe
    ListBox_Historial.List(a, 5) = ListBox_Devoluciones.List(ListBox_Devoluciones.ListIndex, 6)                                                                                             'Devuelto
    ListBox_Historial.List(a, 6) = Val(ListBox_Devoluciones.List(ListBox_Devoluciones.ListIndex, 3))                                                                                        'Restante
    
    ListBox_Devoluciones.RemoveItem (ListBox_Devoluciones.ListIndex)
    
    ActualizarImporte
    
End Sub

Private Sub CommandButton_Cancelar_Click()
    Unload Me
End Sub

Private Sub TextBox_ID1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
Dim Campo As Object
    Set Campo = TextBox_ID1
    Select Case Len(Campo)
        Case 0, 1, 2, 3: If Not (KeyAscii >= 48 And KeyAscii <= 57) Then KeyAscii = 0
        Case 4: KeyAscii = 0
    End Select
End Sub

Private Sub TextBox_ID2_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
Dim Campo As Object
    Set Campo = TextBox_ID2
    Select Case Len(Campo)
        Case 0, 1, 2, 3: If Not (KeyAscii >= 48 And KeyAscii <= 57) Then KeyAscii = 0
        Case 4: KeyAscii = 0
    End Select
End Sub

Private Sub UserForm_Initialize()

Dim i As Byte

    Inicializar
    
    TextBox_Dia = Day(Date)
    TextBox_Mes = Month(Date)
    TextBox_Ano = Year(Date)
    
    LimpiarFormulario
    
    Label_CorrelativoPrefijo.Caption = "Devolucion"
    
            'Establecer propiedades del listado del historial
            ListBox_Historial.ColumnCount = 7
            ListBox_Historial.ColumnWidths = "100 pt; 298 pt; 50 pt; 60 pt; 60 pt; 60 pt; 60 pt"
            
            'Establecer propiedades del listado de devolucion
            ListBox_Devoluciones.ColumnCount = 7
            ListBox_Devoluciones.ColumnWidths = "100 pt; 298 pt; 60 pt; 60 pt; 60 pt; 50 pt; 60 pt"
            
    
    'Llenado combobox Tipo de Transaccion
    ComboBox_TipoDeTransaccionID.AddItem ("Venta de Contado")
    ComboBox_TipoDeTransaccionID.AddItem ("Venta a Credito")
    ComboBox_TipoDeTransaccionID.AddItem ("Venta a Consignacion")
    
    'Llenado del ComboBox de cajas
    For i = 2 To UltimaFilaCajas
         If (Mid(HojaCajas.Cells(i, ColumnaIDCaja), 1, 3) = "USD") Then ComboBox_Caja.AddItem (HojaCajas.Cells(i, ColumnaIDCaja))
    Next i
    
    ComboBox_Caja.Text = "USD-DEIBYS"
    
    TextBox_ID1.Text = "0001"
    
    Set FormularioAnterior = Me
    
    ActualizarCorrelativoEnPantalla Frame_Correlativo, Label_CorrelativoPrefijo
    
End Sub

Private Sub UserForm_Terminate()
     Set FormularioAnterior = Nothing
End Sub

Sub ActualizarImporte()
Dim a As Integer
Dim i As Integer
Dim Importe As Single

    a = ListBox_Devoluciones.ListCount
    Importe = 0
    If a > 0 Then
        For i = 0 To a - 1
            Importe = Importe + Val(ListBox_Devoluciones.List(i, 4))
        Next i
    End If
    
    Label_Importe.Caption = Format(Importe, "0.000")
    
End Sub

Sub LimpiarFormulario()
    
    Label_AsteriscoTipoDeTransaccion.Visible = False
    Label_AsteriscoID1.Visible = False
    Label_AsteriscoID2.Visible = False
    Label_AsteriscoCaja.Visible = False
    Label_Caja = Empty
    Label_Cliente = Empty
    Label_Responsable = Empty
    Label_Fecha = Empty
    Label_Hora = Empty
    
    TextBox_ID2 = Empty
    Label_Cantidad_Auxiliar = Empty
    
    ListBox_Historial.Clear
    ListBox_Devoluciones.Clear
    Label_Importe = Empty
    CheckBox_AnadirACredito.Value = False
    
End Sub
