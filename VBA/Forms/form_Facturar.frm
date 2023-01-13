VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form_Facturar 
   Caption         =   "Facturar"
   ClientHeight    =   9810.001
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   12096
   OleObjectBlob   =   "form_Facturar.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "form_Facturar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ComboBox_Caja_Change()

    Label_AsteriscoCaja.Visible = False
    
    'Actualizar monto de caja en pantalla
    ActualizarSaldoCajaEnPantalla Label_SaldoCaja, ComboBox_Caja.Text
    
End Sub

Private Sub CommandButton_Facturar_Click()

Dim Cod As Variant
Dim NuevaExistencia As Long
Dim i As Integer
Dim j As Integer
Dim a As Integer
Dim Cantidad As Long
Dim CodigoInventario As String
Dim Codigo As String
Dim Producto As String
Dim Precio As Single
Dim IngresarCliente As Byte
Dim ProcesarFactura As Byte
Dim ProcederFactura As Byte
Dim FilaCaja As Byte
Dim FilaDeCliente As Integer
Dim IDResponsable As String
Dim LimiteCredito As Single
Dim SaldoCredito As Single
Dim NuevoSaldoCredito As Single
Dim Fecha As Date
  
    Fecha = TextBox_Dia & "/" & TextBox_Mes & "/" & TextBox_Ano

    Inicializar
    
    'Ocultar todos los asteriscos
    Label_AsteriscoCaja.Visible = False
    Label_AsteriscoFormaDePago.Visible = False
    
    a = ListBox_Listado.ListCount
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Verificaciones'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    'Verificacion de que existan productos ingresados en la factura
    If a = 0 Then
        MsgBox "No hay productos agregados a la factura", , "Facturar"
        Exit Sub
    End If

    FilaDeCliente = ObtenerFila(HojaClientes, TextBox_IDCliente.Text, ColumnaIDCliente)
    
    'Procedimiento a ejecutar si el cliente no existe.
    If FilaDeCliente = 0 Then
        IngresarCliente = MsgBox("El cliente ingresado no existe. ¿Deseas abrir el formulario de registro?", vbYesNo + vbExclamation, "Facturar")
        If IngresarCliente = vbYes Then
            'Limpiar formulario de cliente en factura
            TextBox_NombreCliente = Empty
            TextBox_IDCliente = Empty
            TextBox_DireccionCliente = Empty
            TextBox_TelefonoCliente = Empty
            'Ingresar nuevo cliente
            form_RegistrarCliente.Show
            Exit Sub
        Else
            Exit Sub
        End If
    End If

    'Verificacion de credito habilitado
    If (HojaClientes.Cells(FilaDeCliente, ColumnaCreditoCliente) = False And ComboBox_FormaDePago = "Credito") Then
        MsgBox "Los creditos estan deshabilitados para este cliente", , "Facturar"
        Exit Sub
    End If
    
    'Verificacion de consignaciones habilitadas
    If (HojaClientes.Cells(FilaDeCliente, ColumnaConsignacionCliente) = False And ComboBox_FormaDePago = "Consignacion") Then
        MsgBox "Las consignaciones estan deshabilitadas para este cliente", , "Facturar"
        Exit Sub
    End If
    
    'Verificacion de comentario para cliente "Contado"
    If (TextBox_IDCliente.Text = "V-00000000" And TextBox_Comentario.Text = Empty) Then
        MsgBox "Agrega un comentario a esta transaccion para tener una referencia futura", , "Facturar"
        Exit Sub
    End If
    
    'Verificacion de caja seleccionada
    If ObtenerFila(HojaCajas, ComboBox_Caja.Text, ColumnaIDCaja) = 0 Then
        Label_AsteriscoCaja.Visible = True
        MsgBox "Selecciona una Caja valida", , "Facturar"
        Exit Sub
    End If
    
    IDResponsable = HojaCajas.Cells(ObtenerFila(HojaCajas, ComboBox_Caja.Text, ColumnaIDCaja), ColumnaIDResponsableCaja)
    
    'Verificacion de forma de pago seleccionada
    If Not (ComboBox_FormaDePago.Text = "Contado" Or ComboBox_FormaDePago.Text = "Credito" Or ComboBox_FormaDePago.Text = "Consignacion") Then
        Label_AsteriscoFormaDePago.Visible = True
        MsgBox "Selecciona una Forma de Pago valida", , "Facturar"
        Exit Sub
    End If
    
    'Ultima verificacion antes de procesar la factura
    ProcesarFactura = MsgBox("¿Seguro que deseas procesar esta factura?", vbYesNo + vbExclamation, "Facturar")
    If ProcesarFactura = vbNo Then Exit Sub
    
    
        

    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Factura de Contado'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    'Procedimiento a ejecutar cuando la compra es de contado.
    If ComboBox_FormaDePago = "Contado" Then
        
        'Eliminacion de las existencias en el inventario para cada producto de la factura
        For i = 0 To a - 1
        
            Cod = Val(ListBox_Listado.List(i, 1))
            NuevaExistencia = Val(ListBox_Listado.List(i, 0))
            
            ModificarExistenciaInventario Cod, NuevaExistencia
        
        Next i
        
        'Incluir en el hitorial
        For i = a - 1 To 0 Step -1
            
            Codigo = ListBox_Listado.List(i, 1)
            Producto = ListBox_Listado.List(i, 2)
            Cantidad = ListBox_Listado.List(i, 3)
            Precio = ListBox_Listado.List(i, 4)
            NuevaExistencia = Val(ListBox_Listado.List(i, 0))
            
            IncluirEnHistorial Frame_Correlativo, Label_CorrelativoPrefijo, Fecha, Codigo, Producto, ComboBox_Caja.Text, Cantidad, TextBox_Comentario.Text, TextBox_IDCliente.Text, IDResponsable, Precio, True, NuevaExistencia
        
        Next i
        
    
    End If
    
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Factura de Credito'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    'Procedimiento a ejecutar cuando la compra es a credito.
    If ComboBox_FormaDePago = "Credito" Then
    
            
        ProcesarFactura = MsgBox("¿Seguro que deseas procesar esta factura de CREDITO?", vbYesNo + vbExclamation, "Facturar")
        If ProcesarFactura = vbNo Then Exit Sub
        
        LimiteCredito = HojaClientes.Cells(FilaDeCliente, ColumnaLimiteCreditoCliente)
        SaldoCredito = HojaClientes.Cells(FilaDeCliente, ColumnaSaldoCreditoCliente)
        NuevoSaldoCredito = Val(TextBox_Total) + SaldoCredito

        'Procedimiento a ejecutar si la compra supera el limite de credito establecido para ese cliente
        If NuevoSaldoCredito > LimiteCredito Then
            ProcederFactura = MsgBox("El monto de la factura supera el limite de credito establecido para este cliente, ¿Deseas proceder con la factura?", vbYesNo + vbExclamation, "Facturar")
            If ProcederFactura = vbNo Then Exit Sub
        End If

        'Eliminacion de las existencias en el inventario para cada producto de la factura
        For i = 0 To a - 1
        
            Cod = Val(ListBox_Listado.List(i, 1))
            NuevaExistencia = Val(ListBox_Listado.List(i, 0))
            
            ModificarExistenciaInventario Cod, NuevaExistencia
        
        Next i
        
        'Aumentar saldo de credito
        IngresarSaldoACredito TextBox_IDCliente.Text, Val(TextBox_Total)
        
        'Incluir en el hitorial
        For i = a - 1 To 0 Step -1
            
            Codigo = ListBox_Listado.List(i, 1)
            Producto = ListBox_Listado.List(i, 2)
            Cantidad = ListBox_Listado.List(i, 3)
            Precio = ListBox_Listado.List(i, 4)
            NuevaExistencia = Val(ListBox_Listado.List(i, 0))
            
            IncluirEnHistorial Frame_Correlativo, Label_CorrelativoPrefijo, Fecha, Codigo, Producto, "USD", Cantidad, TextBox_Comentario.Text, TextBox_IDCliente.Text, IDResponsable, Precio, , NuevaExistencia
        
        Next i
    
    
    End If
    
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Factura de Consignacion'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    
        'Procedimiento a ejecutar cuando la compra es a consignacion.
        If ComboBox_FormaDePago = "Consignacion" Then
    
        ProcesarFactura = MsgBox("¿Seguro que deseas procesar esta factura a CONSIGNACION?", vbYesNo + vbExclamation, "Facturar")
        If ProcesarFactura = vbNo Then Exit Sub
        
        'Ingreso de las existencias en el inventario de consignacion del cliente seleccionado
        For i = 0 To a - 1
            For j = 2 To UltimaFilaInventario
                If ListBox_Listado.List(i, 1) = LibroClientes.Sheets(TextBox_IDCliente.Text).Cells(j, ColumnaCodigoCliente).Text Then
                    
                    LibroClientes.Sheets(TextBox_IDCliente.Text).Cells(j, ColumnaExistenciaCliente) = LibroClientes.Sheets(TextBox_IDCliente.Text).Cells(j, ColumnaExistenciaCliente) + Val(ListBox_Listado.List(i, 3))
                    Exit For
                    Exit For

                End If
            Next j
        Next i

        'Eliminacion de las existencias en el inventario para cada producto de la factura
        For i = 0 To a - 1
        
            Cod = Val(ListBox_Listado.List(i, 1))
            NuevaExistencia = Val(ListBox_Listado.List(i, 0))
            
            ModificarExistenciaInventario Cod, NuevaExistencia
        
        Next i
        
        'Incluir en el hitorial
        For i = a - 1 To 0 Step -1
            
            Codigo = ListBox_Listado.List(i, 1)
            Producto = ListBox_Listado.List(i, 2)
            Cantidad = ListBox_Listado.List(i, 3)
            Precio = ListBox_Listado.List(i, 4)
            NuevaExistencia = Val(ListBox_Listado.List(i, 0))
            
            IncluirEnHistorial Frame_Correlativo, Label_CorrelativoPrefijo, Fecha, Codigo, Producto, "USD", Cantidad, TextBox_Comentario.Text, TextBox_IDCliente.Text, IDResponsable, Precio, , NuevaExistencia
        
        Next i
    
    End If
       
   
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Limpieza de Formulario'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
    'Actualizar correlativo
    ActualizarCorrelativo Frame_Correlativo, Label_CorrelativoPrefijo
    
    'Actualizar correlativo en pantalla
    ActualizarCorrelativoEnPantalla Frame_Correlativo, Label_CorrelativoPrefijo
    
    'Actualizar monto de caja en pantalla
    ActualizarSaldoCajaEnPantalla Label_SaldoCaja, ComboBox_Caja.Text
    
    'Limpieza del formulario de factura
    TextBox_NombreCliente = Empty
    TextBox_IDCliente = Empty
    TextBox_DireccionCliente = Empty
    TextBox_TelefonoCliente = Empty
    ListBox_Listado.Clear
    TextBox_Comentario = Empty
    TextBox_SubTotal = Empty
    TextBox_Descuento = 0
    TextBox_Total = Empty
    
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    ActualizarDashboard
    
    MsgBox "Factura realizada exitosamente", , "Facturar"
    
End Sub

Private Sub ComboBox_FormaDePago_Change()

    Label_AsteriscoFormaDePago.Visible = False
    
    'Actualizar correlativo en pantalla
    Label_CorrelativoPrefijo.Caption = ComboBox_FormaDePago.Text
    ActualizarCorrelativoEnPantalla Frame_Correlativo, Label_CorrelativoPrefijo
    
End Sub

Private Sub CommandButton_SeleccionarCliente_Click()
    sec_ListadoDeClientes.Show
End Sub

Private Sub TextBox_Dia_Change()
Dim Campo As Object
    Set Campo = TextBox_Dia
    If Len(Campo) = 2 Then TextBox_Mes.SetFocus
End Sub

Private Sub TextBox_Mes_Change()
Dim Campo As Object
    Set Campo = TextBox_Mes
    If Len(Campo) = 2 Then TextBox_Ano.SetFocus
End Sub

Private Sub TextBox_Dia_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
Dim Campo As Object
    Set Campo = TextBox_Dia
    Select Case Len(Campo)
        Case 0: If Not (KeyAscii >= 48 And KeyAscii <= 51) Then KeyAscii = 0
        Case 1: If Not (KeyAscii >= 48 And KeyAscii <= 57) Then KeyAscii = 0
        Case 2: KeyAscii = 0
    End Select
End Sub

Private Sub TextBox_Mes_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
Dim Campo As Object
    Set Campo = TextBox_Mes
    Select Case Len(Campo)
        Case 0: If Not (KeyAscii >= 48 And KeyAscii <= 51) Then KeyAscii = 0
        Case 1: If Not (KeyAscii >= 48 And KeyAscii <= 57) Then KeyAscii = 0
        Case 2: KeyAscii = 0
    End Select
End Sub

Private Sub TextBox_Ano_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
Dim Campo As Object
    Set Campo = TextBox_Ano
    Select Case Len(Campo)
        Case 0: If Not (KeyAscii = 50) Then KeyAscii = 0
        Case 1: If Not (KeyAscii = 48) Then KeyAscii = 0
        Case 2, 3: If Not (KeyAscii >= 48 And KeyAscii <= 57) Then KeyAscii = 0
        Case 4: KeyAscii = 0
    End Select
End Sub

Private Sub TextBox_IDCliente_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    
    If (Len(TextBox_IDCliente) = 0 And (KeyAscii = 69 Or KeyAscii = 101)) Then TextBox_IDCliente = "E-"
    If (Len(TextBox_IDCliente) = 0 And (KeyAscii = 71 Or KeyAscii = 103)) Then TextBox_IDCliente = "G-"
    If (Len(TextBox_IDCliente) = 0 And (KeyAscii = 74 Or KeyAscii = 106)) Then TextBox_IDCliente = "J-"
    If (Len(TextBox_IDCliente) = 0 And (KeyAscii = 80 Or KeyAscii = 112)) Then TextBox_IDCliente = "P-"
    If (Len(TextBox_IDCliente) = 0 And (KeyAscii = 86 Or KeyAscii = 118)) Then TextBox_IDCliente = "V-"
    
    If Not (KeyAscii >= 48 And KeyAscii <= 57) Then KeyAscii = 0
    
    If Len(TextBox_IDCliente) = 10 Then KeyAscii = 0
        
End Sub

Private Sub TextBox_IDCliente_AfterUpdate()

Dim FilaDeCliente As Integer

    Inicializar
    
    On Error Resume Next
    
    FilaDeCliente = ObtenerFila(HojaClientes, TextBox_IDCliente.Text, ColumnaIDCliente)
    
    TextBox_NombreCliente = HojaClientes.Cells(FilaDeCliente, ColumnaNombreCliente)
    TextBox_DireccionCliente = HojaClientes.Cells(FilaDeCliente, ColumnaDireccionCliente)
    TextBox_TelefonoCliente = HojaClientes.Cells(FilaDeCliente, ColumnaTelefonoCliente)
    
End Sub

Private Sub CommandButton_EliminarItem_Click()

    'Si el listado de productos no esta vacio, se elmina el item elegido, de no elegirse ninguno se van eliminando uno a uno
    If (ListBox_Listado.ListIndex >= 0) Then
        ListBox_Listado.RemoveItem (ListBox_Listado.ListIndex)
        ActualizarSubTotal
    End If

End Sub

Private Sub CommandButton_IngresarItem_Click()
    sec_IngresarProductoEnFactura.Show
End Sub

Private Sub TextBox_Descuento_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
Dim UbicacionPunto As Integer
    UbicacionPunto = InStr(TextBox_Descuento.Text, ".")
    If (KeyAscii = 46 And UbicacionPunto > 0) Then
        KeyAscii = 0
    End If
    If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46) Then
        KeyAscii = 0
    End If
End Sub

Private Sub TextBox_Descuento_AfterUpdate()

    On Error Resume Next
    
    If TextBox_Descuento.Value = Empty Then
        TextBox_Descuento = 0
        ActualizarSubTotal
    Else
        ActualizarSubTotal
    End If
    
End Sub


Private Sub UserForm_Initialize()

Dim i As Integer
Dim FilaCaja As Byte
        
    Inicializar
    
    TextBox_Dia = Day(Date)
    TextBox_Mes = Month(Date)
    TextBox_Ano = Year(Date)
    
    Label_AsteriscoCaja.Visible = False
    Label_AsteriscoFormaDePago.Visible = False
    
    CommandButton_IngresarItem.Picture = LoadPicture(ThisWorkbook.Path & "\Resources\images\mas.jpg")
    CommandButton_EliminarItem.Picture = LoadPicture(ThisWorkbook.Path & "\Resources\images\menos.jpg")
    
    CommandButton_SeleccionarCliente.Picture = LoadPicture(ThisWorkbook.Path & "\Resources\images\cliente.jpg")
    
    ' Formato del listado
    ListBox_Listado.ColumnCount = 6
    ListBox_Listado.ColumnWidths = "60 pt; 100 pt; 298 pt; 50 pt; 60 pt; 70 pt"
    
    ' Se establece el valor del descuento en 0%
    TextBox_Descuento.Value = 0
    
    'Llenado del ComboBox de cajas
    For i = 2 To UltimaFilaCajas
         If (Mid(HojaCajas.Cells(i, ColumnaIDCaja), 1, 3) = "USD") Then ComboBox_Caja.AddItem (HojaCajas.Cells(i, ColumnaIDCaja))
    Next i
    
    ComboBox_Caja.Text = "USD-DEIBYS"

    ' Llenado de ComboBox de forma de pago
    ComboBox_FormaDePago.AddItem ("Contado")
    ComboBox_FormaDePago.AddItem ("Credito")
    ComboBox_FormaDePago.AddItem ("Consignacion")

    Set FormularioAnterior = Me
    
    ComboBox_FormaDePago.Text = "Contado"

End Sub

Private Sub UserForm_Terminate()
    
    Set FormularioAnterior = Nothing
     
End Sub
