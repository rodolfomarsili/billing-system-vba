VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form_Pagos 
   Caption         =   "Pagos"
   ClientHeight    =   10710
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   13032
   OleObjectBlob   =   "form_Pagos.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "form_Pagos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Creditos'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub PagarCredito()

Dim FilaCliente As Integer
Dim FilaCajaCredito As Byte
Dim ProcesarAbono As Byte
Dim Fecha As Date
Dim IDResponsable As String
  
    Fecha = TextBox_Dia & "/" & TextBox_Mes & "/" & TextBox_Ano

    Inicializar
    
    Label_AsteriscoMontoAbonado.Visible = False
    Label_AsteriscoCajaCredito.Visible = False
    
    FilaCliente = ObtenerFila(HojaClientes, TextBox_IDCliente.Text, ColumnaIDCliente)
    FilaCajaCredito = ObtenerFila(HojaCajas, ComboBox_CajaCredito.Text, ColumnaIDCaja)
    
    IDResponsable = HojaCajas.Cells(FilaCajaCredito, ColumnaIDResponsableCaja)
''''''''''''''''''''''Verificaciones'''''''''''''''''''''''''''''

    If FilaCliente = 0 Then
        MsgBox "Debes seleccionar un cliente valido para realizar esta operacion", , "Pago de credito"
        Exit Sub
    End If
    
    If FilaCajaCredito = 0 Then
        Label_AsteriscoCajaCredito.Visible = True
        MsgBox "Selecciona una caja valida", , "Pago de credito"
        Exit Sub
    End If
    
    If Val(TextBox_MontoAbonado) <= 0 Then
        Label_AsteriscoMontoAbonado.Visible = True
        MsgBox "Ingresa el monto a abonar", , "Pago de credito"
        Exit Sub
    End If
    
    'Ultima verificacion antes de procesar el abono de credito
    ProcesarAbono = MsgBox("¿Seguro que deseas procesar esta operacion?", vbYesNo + vbExclamation, "Pago de credito")
    If ProcesarAbono = vbNo Then Exit Sub

''''''''''''''''''''''''''Modificar Credito'''''''''''''''''''''''''''''''''''''''
    'Establecer el nuevo saldo de credito para el cliente
    HojaClientes.Cells(FilaCliente, ColumnaSaldoCreditoCliente) = Val(Label_SaldoCreditoRestanteCliente)

    'Incluir en el hitorial
    IncluirEnHistorial Frame_Correlativo, Label_CorrelativoPrefijo, Fecha, , , ComboBox_CajaCredito.Text, , "Abono: " & TextBox_MontoAbonado.Text & " $", TextBox_IDCliente.Text, IDResponsable, Val(TextBox_MontoAbonado), True
    
''''''''''''''''''''''''Limpieza del Formulario'''''''''''''''''''''''''''''''''''''
        
    'Actualizar correlativo
    ActualizarCorrelativo Frame_Correlativo, Label_CorrelativoPrefijo
    
    'Limpieza del formulario
    Label_SaldoCreditoCliente.Caption = HojaClientes.Cells(FilaCliente, ColumnaSaldoCreditoCliente)
    TextBox_MontoAbonado = Empty
    
    Label_Pagado.Caption = "True"
    
    ActualizarDashboard
    
    MsgBox "Pago realizado exitosamente", , "Pago de credito"
    
End Sub



Private Sub TextBox_MontoAbonado_Change()
    
    Label_AsteriscoMontoAbonado.Visible = False
    If Label_SaldoCreditoCliente <> Empty Then Label_SaldoCreditoRestanteCliente = Format(Val(Label_SaldoCreditoCliente) - Val(TextBox_MontoAbonado), "0.00")
    If TextBox_MontoAbonado = Empty Then Label_SaldoCreditoRestanteCliente = Empty
    
End Sub

Private Sub TextBox_MontoAbonado_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

Dim UbicacionPunto As Integer

    UbicacionPunto = InStr(TextBox_MontoAbonado.Text, ".")
    
    If (KeyAscii = 46 And UbicacionPunto > 0) Then
        KeyAscii = 0
    End If
    
    If Not (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 46) Then
        KeyAscii = 0
    End If
    
End Sub

Private Sub ComboBox_CajaCredito_Change()
    
    Label_AsteriscoCajaCredito.Visible = False
    
    'Actualizar monto de caja en pantalla
    ActualizarSaldoCajaEnPantalla Label_SaldoCajaCredito, ComboBox_CajaCredito.Text
    
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Consignaciones'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub PagarConsignacion()

Dim FilaCajaConsignacion As Byte
Dim FilaCliente As Byte
Dim ProcesarAbono As Byte
Dim IDResponsable As String
Dim a As Integer
Dim i As Integer
Dim j As Integer
Dim Cantidad As Long
Dim Codigo As String
Dim Producto As String
Dim Precio As Single
Dim NuevaExistencia As Long
Dim Fecha As Date

  
    Fecha = TextBox_Dia & "/" & TextBox_Mes & "/" & TextBox_Ano

    Inicializar
    
    Label_AsteriscoCajaConsignacion.Visible = False
    
    FilaCliente = ObtenerFila(HojaClientes, TextBox_IDCliente.Text, ColumnaIDCliente)
    FilaCajaConsignacion = ObtenerFila(HojaCajas, ComboBox_CajaConsignacion.Text, ColumnaIDCaja)
    
    IDResponsable = HojaCajas.Cells(FilaCajaConsignacion, ColumnaIDResponsableCaja)
    
    a = ListBox_PorPagar.ListCount
'''''''''''''''''''''''''''''Verificaciones''''''''''''''''''''''''''''''''''''
    If FilaCliente = 0 Then
        MsgBox "Debes seleccionar un cliente valido para realizar esta operacion", , "Pago de consignacion"
        Exit Sub
    End If
    
    If a = 0 Then
        MsgBox "No hay productos añadidos a la lista de pago", , "Pago de consignacion"
        Exit Sub
    End If
    
    If FilaCajaConsignacion = 0 Then
        Label_AsteriscoCajaConsignacion.Visible = True
        MsgBox "Selecciona una caja valida", , "Pago de consignacion"
        Exit Sub
    End If
    
    'Ultima verificacion antes de procesar el abono de credito
    ProcesarAbono = MsgBox("¿Seguro que deseas procesar esta operacion?", vbYesNo + vbExclamation, "Pago de consignacion")
    If ProcesarAbono = vbNo Then Exit Sub

      
    

    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    
''''''''''''''''''''''''''Modificar Listado de Consignacion''''''''''''''''''''''''''''''''''
    
    'Eliminacion de las existencias en el inventario de consignacion del cliente seleccionado
        For i = 0 To a - 1
            For j = 2 To UltimaFilaInventario
                If (Val(ListBox_PorPagar.List(i, 1)) = LibroClientes.Sheets(TextBox_IDCliente.Text).Cells(j, ColumnaCodigoCliente)) Then
                    LibroClientes.Sheets(TextBox_IDCliente.Text).Cells(j, ColumnaExistenciaCliente) = Val(ListBox_PorPagar.List(i, 0))
                    Exit For
                End If
            Next j
        Next i
''''''''''''''''''''''''''Limpieza del Formulario'''''''''''''''''''''''''''''''''''
    
    'Incluir en el hitorial
    For i = a - 1 To 0 Step -1
        
        Codigo = ListBox_PorPagar.List(i, 1)
        Producto = ListBox_PorPagar.List(i, 2)
        Cantidad = ListBox_PorPagar.List(i, 3)
        Precio = ListBox_PorPagar.List(i, 4)
        NuevaExistencia = Val(ListBox_PorPagar.List(i, 0))
        IncluirEnHistorial Frame_Correlativo, Label_CorrelativoPrefijo, Fecha, Codigo, Producto, ComboBox_CajaConsignacion.Text, Cantidad, , TextBox_IDCliente.Text, IDResponsable, Precio, True, NuevaExistencia
    
    Next i
    
    'Actualizar correlativo
    ActualizarCorrelativo Frame_Correlativo, Label_CorrelativoPrefijo
    
    'Limpiar formulario
    ListBox_PorPagar.Clear
    Label_Importe = Empty
    
    Label_Pagado.Caption = "True"
    
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    ActualizarDashboard
    
    MsgBox "Pago realizado exitosamente", , "Pago de consignacion"
    
End Sub

Private Sub CommandButton_Anadir_Click()
    Anadir
End Sub

Private Sub CommandButton_Quitar_Click()
    Quitar
End Sub

Private Sub ListBox_Consignaciones_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Anadir
End Sub

Private Sub ListBox_PorPagar_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Quitar
End Sub

Private Sub CommandButton_AnadirTodo_Click()
Dim a As Integer
Dim b As Integer
Dim i As Integer
Dim j As Integer

    a = ListBox_Consignaciones.ListCount
    b = ListBox_PorPagar.ListCount
    
    If b > 0 Then
        MsgBox "Esta opcion solo esta disponible si no has agregado articulos al listado de pago", , "Pago de consignacion"
        Exit Sub
    End If
    
    j = a - 1
    For i = 0 To a - 1

        ListBox_PorPagar.AddItem
        ListBox_PorPagar.List(i, 1) = ListBox_Consignaciones.List(j, 0)
        ListBox_PorPagar.List(i, 2) = ListBox_Consignaciones.List(j, 1)
        ListBox_PorPagar.List(i, 3) = ListBox_Consignaciones.List(j, 2)
        ListBox_PorPagar.List(i, 4) = ListBox_Consignaciones.List(j, 3)
        ListBox_PorPagar.List(i, 5) = ListBox_Consignaciones.List(j, 4)
        
        ListBox_Consignaciones.RemoveItem (j)
        
        j = j - 1
    Next i
    
    ActualizarImporte
    
End Sub

Private Sub CommandButton_QuitarTodo_Click()
Dim a As Integer
Dim b As Integer
Dim i As Integer
Dim j As Integer

    a = ListBox_Consignaciones.ListCount
    b = ListBox_PorPagar.ListCount
    
    If a > 0 Then
        MsgBox "Esta opcion solo esta disponible si no has agregado articulos al listado de pago", , "Pago de consignacion"
        Exit Sub
    End If
    
    i = b - 1
    For j = 0 To b - 1

        ListBox_Consignaciones.AddItem
        ListBox_Consignaciones.List(j, 0) = ListBox_PorPagar.List(i, 1)
        ListBox_Consignaciones.List(j, 1) = ListBox_PorPagar.List(i, 2)
        ListBox_Consignaciones.List(j, 2) = ListBox_PorPagar.List(i, 3)
        ListBox_Consignaciones.List(j, 3) = ListBox_PorPagar.List(i, 4)
        ListBox_Consignaciones.List(j, 4) = ListBox_PorPagar.List(i, 5)
        
        ListBox_PorPagar.RemoveItem (i)
        
        i = i - 1
    Next j
    
    ActualizarImporte
    
End Sub

Private Sub ComboBox_CajaConsignacion_Change()

    Label_AsteriscoCajaConsignacion.Visible = False
    
    'Actualizar monto de caja en pantalla
    ActualizarSaldoCajaEnPantalla Label_SaldoCajaConsignacion, ComboBox_CajaConsignacion.Text
    
End Sub

Private Sub TextBox_IDCliente_Change()

Dim FilaDeCliente As Integer
Dim i As Integer
Dim a As Integer

    Inicializar
    
    On Error Resume Next
    
    

    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    ListBox_Consignaciones.Clear
    a = 0
    'Ingreso de las existencias en el inventario de consignacion del cliente seleccionado
    For i = 2 To UltimaFilaInventario
        If LibroClientes.Sheets(TextBox_IDCliente.Text).Cells(i, ColumnaExistenciaCliente) <> 0 Then
            
            ListBox_Consignaciones.AddItem
            ListBox_Consignaciones.List(a, 0) = LibroClientes.Sheets(TextBox_IDCliente.Text).Cells(i, ColumnaCodigoCliente)
            ListBox_Consignaciones.List(a, 1) = LibroClientes.Sheets(TextBox_IDCliente.Text).Cells(i, ColumnaProductoCliente)
            ListBox_Consignaciones.List(a, 2) = LibroClientes.Sheets(TextBox_IDCliente.Text).Cells(i, ColumnaExistenciaCliente)
            ListBox_Consignaciones.List(a, 3) = Format(LibroClientes.Sheets(TextBox_IDCliente.Text).Cells(i, ColumnaPrecioUnitarioCliente), "0.0000")
            ListBox_Consignaciones.List(a, 4) = Format(LibroClientes.Sheets(TextBox_IDCliente.Text).Cells(i, ColumnaImporteCliente), "0.0000")
            
            a = a + 1
        End If
    Next i
    
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    
End Sub

Sub ActualizarImporte()
Dim a As Integer
Dim i As Integer
Dim Importe As Single

    a = ListBox_PorPagar.ListCount
    Importe = 0
    If a > 0 Then
        For i = 0 To a - 1
            Importe = Importe + Val(ListBox_PorPagar.List(i, 5))
        Next i
    End If
    
    Label_Importe.Caption = Format(Importe, "0,0.0000")
    
End Sub

Sub Anadir()

Dim a As Integer
Dim i As Integer
Dim Restante As Long
Dim NuevoImporteConsignaciones As Single
Dim NuevoImportePorPagar As Single
    
'''''''''''''''''''''''''''''''''''Verificaciones'''''''''''''''''''''''''''''''''''
    'No hacer nada si no hay algun item seleccionado en la lista de consignaciones
    If ListBox_Consignaciones.ListIndex = -1 Then Exit Sub
    
    'Mostrar el formaulario de añadir cantidad
    sec_Cantidad.Show
    
    'Si la cantidad a mover es 0, no se hace nada
    If Val(Label_Cantidad_Auxiliar) = 0 Then Exit Sub
    
    'Verificacion de si existe el articulo agregado en la lista PorPagar
    a = ListBox_PorPagar.ListCount
    If a > 0 Then
        For i = 0 To a - 1
            If ListBox_Consignaciones.List(ListBox_Consignaciones.ListIndex, 0) = ListBox_PorPagar.List(i, 1) Then
                MsgBox "Ya has ingresado este articulo, eliminalo de la lista de pagos y vuelve a agregarlo", , "Pago consignacion"
                Label_Cantidad_Auxiliar = Empty
                Exit Sub
            End If
        Next i
    End If
    
    'Calculo del restante
    Restante = ListBox_Consignaciones.List(ListBox_Consignaciones.ListIndex, 2) - Val(Label_Cantidad_Auxiliar)
    
    If Restante < 0 Then
        MsgBox "Debes ingresar una cantidad menor o igual a la que se encuentra consignada", , "Pago consignacion"
        Label_Cantidad_Auxiliar = Empty
        Exit Sub
    End If


''''''''''''''''''''''''''Llenar listado PorPagar'''''''''''''''''''''''''''''''''''

    If Restante = 0 Then 'Se ejecuta cuado se añaden TODAS las existencias del item seleccionado
    
        a = ListBox_PorPagar.ListCount
        ListBox_PorPagar.AddItem
        ListBox_PorPagar.List(a, 0) = Restante
        ListBox_PorPagar.List(a, 1) = ListBox_Consignaciones.List(ListBox_Consignaciones.ListIndex, 0)  'Codigo
        ListBox_PorPagar.List(a, 2) = ListBox_Consignaciones.List(ListBox_Consignaciones.ListIndex, 1)  'Producto
        ListBox_PorPagar.List(a, 3) = ListBox_Consignaciones.List(ListBox_Consignaciones.ListIndex, 2)  'Cantidad
        ListBox_PorPagar.List(a, 4) = ListBox_Consignaciones.List(ListBox_Consignaciones.ListIndex, 3)  'Precio
        ListBox_PorPagar.List(a, 5) = ListBox_Consignaciones.List(ListBox_Consignaciones.ListIndex, 4)  'Importe
        
    
        ListBox_Consignaciones.RemoveItem (ListBox_Consignaciones.ListIndex)
        ListBox_Consignaciones.ListIndex = -1
    
    End If
     
    If Restante > 0 Then 'Se ejecuta cuado se añaden ALGUNAS las existencias del item seleccionado
    
        a = ListBox_PorPagar.ListCount
        ListBox_PorPagar.AddItem
        ListBox_PorPagar.List(a, 0) = Restante
        ListBox_PorPagar.List(a, 1) = ListBox_Consignaciones.List(ListBox_Consignaciones.ListIndex, 0)  'Codigo
        ListBox_PorPagar.List(a, 2) = ListBox_Consignaciones.List(ListBox_Consignaciones.ListIndex, 1)  'Producto
        ListBox_PorPagar.List(a, 3) = Val(ListBox_Consignaciones.List(ListBox_Consignaciones.ListIndex, 2)) - Restante  'Cantidad
        ListBox_PorPagar.List(a, 4) = ListBox_Consignaciones.List(ListBox_Consignaciones.ListIndex, 3)  'Precio
        NuevoImportePorPagar = (Val(ListBox_Consignaciones.List(ListBox_Consignaciones.ListIndex, 2)) - Restante) * Val(ListBox_Consignaciones.List(ListBox_Consignaciones.ListIndex, 3))
        ListBox_PorPagar.List(a, 5) = Format(NuevoImportePorPagar, "0.0000")  'Importe
        
        ListBox_Consignaciones.List(ListBox_Consignaciones.ListIndex, 2) = Restante 'Cantidad
        NuevoImporteConsignaciones = Restante * Val(ListBox_Consignaciones.List(ListBox_Consignaciones.ListIndex, 3))
        ListBox_Consignaciones.List(ListBox_Consignaciones.ListIndex, 4) = Format(NuevoImporteConsignaciones, "0.0000") 'Importe
        ListBox_Consignaciones.ListIndex = -1
      
    End If
    
    Label_Cantidad_Auxiliar.Caption = Empty
    
    ActualizarImporte
    
End Sub


Sub Quitar()

Dim i As Integer
Dim a As Integer
Dim NuevoImporteConsignaciones As Single

    'No hacer nada si no hay algun item seleccionado en la lista de pagados
    If ListBox_PorPagar.ListIndex = -1 Then Exit Sub
    
    'Verificacion de si existe el articulo agregado en la lista de consignaciones
    a = ListBox_Consignaciones.ListCount
    
    If a > 0 Then
        For i = 0 To a - 1
            If ListBox_PorPagar.List(ListBox_PorPagar.ListIndex, 1) = ListBox_Consignaciones.List(i, 0) Then 'Si existe
            
                ListBox_Consignaciones.List(i, 2) = Val(ListBox_Consignaciones.List(i, 2)) + Val(ListBox_PorPagar.List(ListBox_PorPagar.ListIndex, 3))
                ListBox_Consignaciones.List(i, 4) = Format(Val(ListBox_Consignaciones.List(i, 4)) + Val(ListBox_PorPagar.List(ListBox_PorPagar.ListIndex, 5)), "0.0000")
                
                ListBox_PorPagar.RemoveItem (ListBox_PorPagar.ListIndex)
                ActualizarImporte
                
                Exit Sub
            End If
        Next i
    End If
    
    'No existe en el listado de consignaciones
    ListBox_Consignaciones.AddItem
    ListBox_Consignaciones.List(a, 0) = ListBox_PorPagar.List(ListBox_PorPagar.ListIndex, 1)
    ListBox_Consignaciones.List(a, 1) = ListBox_PorPagar.List(ListBox_PorPagar.ListIndex, 2)
    ListBox_Consignaciones.List(a, 2) = ListBox_PorPagar.List(ListBox_PorPagar.ListIndex, 3)
    ListBox_Consignaciones.List(a, 3) = ListBox_PorPagar.List(ListBox_PorPagar.ListIndex, 4)
    ListBox_Consignaciones.List(a, 4) = ListBox_PorPagar.List(ListBox_PorPagar.ListIndex, 5)
    
    ListBox_PorPagar.RemoveItem (ListBox_PorPagar.ListIndex)
    
    
    ActualizarImporte
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Comunes'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

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

Private Sub CommandButton_Pagar_Click()

    Inicializar
    
    If MultiPage_Pagos.Pages(MultiPage_Pagos.SelectedItem.Index).Caption = "Creditos" Then
        PagarCredito
    Else
        PagarConsignacion
    End If
    
    If Label_Pagado.Caption = "True" Then
    
        'Actualizar el saldo de la caja en la pantalla
        ActualizarSaldoCajaEnPantalla Label_SaldoCajaCredito, ComboBox_CajaCredito.Text
        ActualizarSaldoCajaEnPantalla Label_SaldoCajaConsignacion, ComboBox_CajaConsignacion.Text
        
        'Actualizar correlativo en pantalla
        ActualizarCorrelativoEnPantalla Frame_Correlativo, Label_CorrelativoPrefijo
        
    End If
    
    Label_Pagado.Caption = "False"
    
End Sub

Private Sub CommandButton_Cancelar_Click()
    Unload Me
End Sub

Private Sub CommandButton_SeleccionarCliente_Click()
    sec_ListadoDeClientes.Show
End Sub

Private Sub MultiPage_Pagos_Change()

    If MultiPage_Pagos.Pages(MultiPage_Pagos.SelectedItem.Index).Caption = "Creditos" Then
        Label_CorrelativoPrefijo.Caption = "Credito"
    Else
        Label_CorrelativoPrefijo.Caption = "Consignacion"
    End If
    
    'Actualizar correlativo en pantalla
    ActualizarCorrelativoEnPantalla Frame_Correlativo, Label_CorrelativoPrefijo
    
End Sub
    
Private Sub UserForm_Click()
    ListBox_Consignaciones.ListIndex = -1
    ListBox_PorPagar.ListIndex = -1
End Sub

Private Sub UserForm_Initialize()

Dim i As Integer

    Inicializar
    
    TextBox_Dia = Day(Date)
    TextBox_Mes = Month(Date)
    TextBox_Ano = Year(Date)
    
    Label_AsteriscoMontoAbonado.Visible = False
    Label_AsteriscoCajaCredito.Visible = False
    Label_AsteriscoCajaConsignacion.Visible = False
    
    Label_CorrelativoPrefijo.Caption = "Consignacion"
    
    CommandButton_SeleccionarCliente.Picture = LoadPicture(ThisWorkbook.Path & "\Resources\images\cliente.jpg")
    
    'Cargar un background blanco para cada pagina del control multipage
    For i = 0 To MultiPage_Pagos.Count - 1
        MultiPage_Pagos.Pages(i).Picture = LoadPicture(ThisWorkbook.Path & RutaImages & "background.jpg")
        MultiPage_Pagos.Pages(i).PictureSizeMode = fmPictureSizeModeStretch
    Next i
    
    'Formato del listado de consignaciones
    ListBox_Consignaciones.ColumnCount = 5
    ListBox_Consignaciones.ColumnWidths = "112 pt; 298 pt; 60 pt; 80 pt; 80 pt"

    'Formato del listado de pagos
    ListBox_PorPagar.ColumnCount = 6
    ListBox_PorPagar.ColumnWidths = "60 pt; 112 pt; 298 pt; 60 pt; 80 pt; 80 pt"
    
    'Llenado del ComboBox de cajas
    For i = 2 To UltimaFilaCajas
        If (Mid(HojaCajas.Cells(i, ColumnaIDCaja), 1, 3) = "USD") Then
            ComboBox_CajaCredito.AddItem (HojaCajas.Cells(i, ColumnaIDCaja))
            ComboBox_CajaConsignacion.AddItem (HojaCajas.Cells(i, ColumnaIDCaja))
        End If
    Next i
    
    ComboBox_CajaCredito.Text = "USD-DEIBYS"
    ComboBox_CajaConsignacion.Text = "USD-DEIBYS"
    'Actualizar el saldo de la caja en la pantalla
    ActualizarSaldoCajaEnPantalla Label_SaldoCajaCredito, ComboBox_CajaCredito.Text
    ActualizarSaldoCajaEnPantalla Label_SaldoCajaConsignacion, ComboBox_CajaConsignacion.Text
    
    Set FormularioAnterior = Me
    
    'Actualizar correlativo en pantalla
    ActualizarCorrelativoEnPantalla Frame_Correlativo, Label_CorrelativoPrefijo
    
End Sub

Private Sub UserForm_Terminate()
    Set FormularioAnterior = Nothing
End Sub
