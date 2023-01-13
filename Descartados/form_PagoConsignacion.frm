VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form_PagoConsignacion 
   Caption         =   "Pago de Consignacion"
   ClientHeight    =   8610.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8850.001
   OleObjectBlob   =   "form_PagoConsignacion.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "form_PagoConsignacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ComboBox_Caja_Change()

Dim FilaCaja As Byte
    
    On Error Resume Next
    
    Label_AsteriscoCaja.Visible = False
    
    FilaCaja = ObtenerFila(HojaCajas, ComboBox_Caja.Text, ColumnaIDCaja)
    Label_SaldoCaja.Caption = "Saldo Actual: " & Format(HojaCajas.Cells(FilaCaja, ColumnaSaldoCaja), "0.00") & " $"
    
End Sub

Private Sub CommandButton_Anadir_Click()
    Anadir
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
        ListBox_PorPagar.List(i, 0) = ListBox_Consignaciones.List(j, 0)
        ListBox_PorPagar.List(i, 1) = ListBox_Consignaciones.List(j, 1)
        ListBox_PorPagar.List(i, 2) = ListBox_Consignaciones.List(j, 2)
        ListBox_PorPagar.List(i, 3) = ListBox_Consignaciones.List(j, 3)
        ListBox_PorPagar.List(i, 4) = ListBox_Consignaciones.List(j, 4)
        
        ListBox_Consignaciones.RemoveItem (j)
        
        j = j - 1
    Next i
    
    ActualizarImporte
    
End Sub

Private Sub CommandButton_LimpiarFormulario_Click()

Dim Limpiar As Byte

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Verificaciones'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Ultima verificacion antes de procesar el abono de credito
    Limpiar = MsgBox("¿Seguro que deseas limpiar todo el formulario?", vbYesNo + vbExclamation, "Pago de consignacion")
    If Limpiar = vbNo Then Exit Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Limpieza de Formulario'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Limpieza del formulario de factura
    TextBox_NombreCliente = Empty
    TextBox_IDCliente = Empty
    TextBox_DireccionCliente = Empty
    TextBox_TelefonoCliente = Empty
    Label_Importe = Empty
    ListBox_Consignaciones.Clear
    ListBox_PorPagar.Clear
    
    
End Sub

Private Sub CommandButton_Quitar_Click()
    Quitar
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
        ListBox_Consignaciones.List(j, 0) = ListBox_PorPagar.List(i, 0)
        ListBox_Consignaciones.List(j, 1) = ListBox_PorPagar.List(i, 1)
        ListBox_Consignaciones.List(j, 2) = ListBox_PorPagar.List(i, 2)
        ListBox_Consignaciones.List(j, 3) = ListBox_PorPagar.List(i, 3)
        ListBox_Consignaciones.List(j, 4) = ListBox_PorPagar.List(i, 4)
        
        ListBox_PorPagar.RemoveItem (i)
        
        i = i - 1
    Next j
    
    ActualizarImporte
    
End Sub

Private Sub CommandButton_Pagar_Click()

Dim FilaCaja As Byte
Dim FilaCliente As Byte
Dim ProcesarAbono As Byte
Dim Fecha As String
Dim IDResponsable As String
Dim a As Integer
Dim i As Integer
Dim j As Integer
Dim Cantidad As Long
Dim Codigo As String
Dim Producto As String
Dim Precio As Single

  
    Fecha = TextBox_Mes & "/" & TextBox_Dia & "/" & TextBox_Ano

    Inicializar
    
    Label_AsteriscoCaja.Visible = False
    
    FilaCliente = ObtenerFila(HojaClientes, TextBox_IDCliente.Text, ColumnaIDCliente)
    FilaCaja = ObtenerFila(HojaCajas, ComboBox_Caja.Text, ColumnaIDCaja)
    
    IDResponsable = HojaCajas.Cells(FilaCaja, ColumnaIDResponsableCaja)
    
    a = ListBox_PorPagar.ListCount
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Verificaciones'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If FilaCliente = 0 Then
        MsgBox "Debes seleccionar un cliente valido para realizar esta operacion", , "Pago de consignacion"
        Exit Sub
    End If
    
    If a = 0 Then
        MsgBox "No hay productos añadidos a la lista de pago", , "Pago de consignacion"
        Exit Sub
    End If
    
    If FilaCaja = 0 Then
        Label_AsteriscoCaja.Visible = True
        MsgBox "Selecciona una caja valida", , "Pago de consignacion"
        Exit Sub
    End If
    
    'Ultima verificacion antes de procesar el abono de credito
    ProcesarAbono = MsgBox("¿Seguro que deseas procesar esta operacion?", vbYesNo + vbExclamation, "Pago de consignacion")
    If ProcesarAbono = vbNo Then Exit Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Modificar Listado de Consignacion''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    'Eliminacion de las existencias en el inventario de consignacion del cliente seleccionado
        For i = 0 To a - 1
            For j = 3 To UltimaFilaInventario
                If (Val(ListBox_PorPagar.List(i, 0)) = LibroClientes.Sheets(TextBox_IDCliente.Text).Cells(j, ColumnaCodigoCliente)) Then
                    LibroClientes.Sheets(TextBox_IDCliente.Text).Cells(j, ColumnaExistenciaCliente) = LibroClientes.Sheets(TextBox_IDCliente.Text).Cells(j, ColumnaExistenciaCliente) - Val(ListBox_PorPagar.List(i, 2))
                    Exit For
                End If
            Next j
        Next i

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Limpieza del Formulario'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    'Incluir en el hitorial
    For i = 0 To a - 1
        
        Codigo = ListBox_PorPagar.List(i, 0)
        Producto = ListBox_PorPagar.List(i, 1)
        Cantidad = ListBox_PorPagar.List(i, 2)
        Precio = ListBox_PorPagar.List(i, 3)
        IncluirEnHistorial Fecha, Codigo, Producto, Label_CorrelativoPrefijo.Caption, ComboBox_Caja, Cantidad, , TextBox_IDCliente.Text, IDResponsable, , Precio
    
    Next i
    
    'Abonar sal do a caja correspondiente
    AbonarSaldoACaja ComboBox_Caja, Val(Label_Importe)
    
    'Actualizar saldo en hoja de clientes
    HojaClientes.Cells(FilaCliente, ColumnaSaldoConsignacionCliente) = LibroClientes.Sheets(TextBox_IDCliente.Text).Cells(1, ColumnaImporteTotalCliente)
    
    'Actualizar el saldo de la caja en la pantalla
    Label_SaldoCaja.Caption = "Saldo Actual: " & Format(HojaCajas.Cells(FilaCaja, ColumnaSaldoCaja), "0.00") & " $"
    
    'Limpieza del formulario de factura
    TextBox_NombreCliente = Empty
    TextBox_IDCliente = Empty
    TextBox_DireccionCliente = Empty
    TextBox_TelefonoCliente = Empty
    Label_Importe = Empty
    ListBox_Consignaciones.Clear
    ListBox_PorPagar.Clear
    
    'Actualizar correlativo
    ActualizarCorrelativo Label_CorrelativoPrefijo.Caption
    
    ActualizarCorrelativoEnPantalla Label_CorrelativoPrefijo.Caption

    MsgBox "Pago realizado exitosamente", , "Pago de consignacion"
    
End Sub


Private Sub CommandButton_Cancelar_Click()
    Unload Me
End Sub



Private Sub CommandButton_SeleccionarCliente_Click()
    sec_ListadoDeClientes.Show
End Sub


Private Sub ListBox_Consignaciones_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Anadir
End Sub


Private Sub ListBox_PorPagar_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
        Quitar
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


Private Sub TextBox_IDCliente_Change()

Dim FilaDeCliente As Integer
Dim i As Integer
Dim a As Integer

    Inicializar
    
    On Error Resume Next
    
    ListBox_Consignaciones.Clear
    a = 0
    'Ingreso de las existencias en el inventario de consignacion del cliente seleccionado
    For i = 3 To UltimaFilaInventario
        If LibroClientes.Sheets(TextBox_IDCliente.Text).Cells(i, ColumnaExistenciaCliente) <> 0 Then
            
            ListBox_Consignaciones.AddItem
            ListBox_Consignaciones.List(a, 0) = LibroClientes.Sheets(TextBox_IDCliente.Text).Cells(i, ColumnaCodigoCliente)
            ListBox_Consignaciones.List(a, 1) = LibroClientes.Sheets(TextBox_IDCliente.Text).Cells(i, ColumnaProductoCliente)
            ListBox_Consignaciones.List(a, 2) = LibroClientes.Sheets(TextBox_IDCliente.Text).Cells(i, ColumnaExistenciaCliente)
            ListBox_Consignaciones.List(a, 3) = Format(LibroClientes.Sheets(TextBox_IDCliente.Text).Cells(i, ColumnaPrecioUnitarioCliente), "0.000")
            ListBox_Consignaciones.List(a, 4) = Format(LibroClientes.Sheets(TextBox_IDCliente.Text).Cells(i, ColumnaImporteCliente), "0.000")
            
            a = a + 1
        End If
    Next i
    

    
End Sub

Private Sub UserForm_Initialize()

Dim i As Integer
Dim FilaCaja As Byte
Dim RutaDeAcceso As String
    
    Inicializar
    
    TextBox_Dia = Day(Date)
    TextBox_Mes = Month(Date)
    TextBox_Ano = Year(Date)
    
    RutaDeAcceso = "\Resources\images\"
    
    Label_AsteriscoCaja.Visible = False
    
    Label_CorrelativoPrefijo.Caption = "P-Consignacion"
    
    CommandButton_SeleccionarCliente.Picture = LoadPicture(ThisWorkbook.Path & RutaDeAcceso & "cliente.jpg")
    
    'Formato del listado de consignaciones
    ListBox_Consignaciones.ColumnCount = 5
    ListBox_Consignaciones.ColumnWidths = "60 pt; 150 pt; 50 pt; 60 pt; 60 pt"

    'Formato del listado de pagos
    ListBox_PorPagar.ColumnCount = 5
    ListBox_PorPagar.ColumnWidths = "60 pt; 150 pt; 50 pt; 60 pt; 60 pt"
    
    'Llenado del ComboBox de cajas
    For i = 3 To UltimaFilaCajas
         If (Mid(HojaCajas.Cells(i, ColumnaIDCaja), 1, 3) = "USD") Then ComboBox_Caja.AddItem (HojaCajas.Cells(i, ColumnaIDCaja))
    Next i
    ComboBox_Caja.Text = "USD-DEIBYS"
    FilaCaja = ObtenerFila(HojaCajas, ComboBox_Caja.Text, ColumnaIDCaja)
    Label_SaldoCaja.Caption = "Saldo Actual: " & Format(HojaCajas.Cells(FilaCaja, ColumnaSaldoCaja), "0.00") & " $"
    
    
    Set FormularioAnterior = Me
        
    ActualizarCorrelativoEnPantalla Label_CorrelativoPrefijo.Caption
    
End Sub

Private Sub UserForm_Terminate()
    Set FormularioAnterior = Nothing
End Sub

Sub ActualizarImporte()
Dim a As Integer
Dim i As Integer
Dim Importe As Single

    a = ListBox_PorPagar.ListCount
    Importe = 0
    If a > 0 Then
        For i = 0 To a - 1
            Importe = Importe + Val(ListBox_PorPagar.List(i, 4))
        Next i
    End If
    
    Label_Importe.Caption = Format(Importe, "0.000")
    
End Sub

Sub Anadir()

Dim a As Integer
Dim i As Integer
Dim Restante As Long
Dim NuevoImporteConsignaciones As Single
Dim NuevoImportePorPagar As Single
    
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Verificaciones'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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
            If ListBox_Consignaciones.List(ListBox_Consignaciones.ListIndex, 0) = ListBox_PorPagar.List(i, 0) Then
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


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Llenar listado PorPagar'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    If Restante = 0 Then 'Se ejecuta cuado se añaden TODAS las existencias del item seleccionado
    
        a = ListBox_PorPagar.ListCount
        ListBox_PorPagar.AddItem
        ListBox_PorPagar.List(a, 0) = ListBox_Consignaciones.List(ListBox_Consignaciones.ListIndex, 0)  'Codigo
        ListBox_PorPagar.List(a, 1) = ListBox_Consignaciones.List(ListBox_Consignaciones.ListIndex, 1)  'Producto
        ListBox_PorPagar.List(a, 2) = ListBox_Consignaciones.List(ListBox_Consignaciones.ListIndex, 2)  'Cantidad
        ListBox_PorPagar.List(a, 3) = ListBox_Consignaciones.List(ListBox_Consignaciones.ListIndex, 3)  'Precio
        ListBox_PorPagar.List(a, 4) = ListBox_Consignaciones.List(ListBox_Consignaciones.ListIndex, 4)  'Importe
        
    
        ListBox_Consignaciones.RemoveItem (ListBox_Consignaciones.ListIndex)
        ListBox_Consignaciones.ListIndex = -1
    
    End If
     
    If Restante > 0 Then 'Se ejecuta cuado se añaden ALGUNAS las existencias del item seleccionado
    
        a = ListBox_PorPagar.ListCount
        ListBox_PorPagar.AddItem
        ListBox_PorPagar.List(a, 0) = ListBox_Consignaciones.List(ListBox_Consignaciones.ListIndex, 0)  'Codigo
        ListBox_PorPagar.List(a, 1) = ListBox_Consignaciones.List(ListBox_Consignaciones.ListIndex, 1)  'Producto
        ListBox_PorPagar.List(a, 2) = Val(ListBox_Consignaciones.List(ListBox_Consignaciones.ListIndex, 2)) - Restante  'Cantidad
        ListBox_PorPagar.List(a, 3) = ListBox_Consignaciones.List(ListBox_Consignaciones.ListIndex, 3)  'Precio
        NuevoImportePorPagar = (Val(ListBox_Consignaciones.List(ListBox_Consignaciones.ListIndex, 2)) - Restante) * Val(ListBox_Consignaciones.List(ListBox_Consignaciones.ListIndex, 3))
        ListBox_PorPagar.List(a, 4) = Format(NuevoImportePorPagar, "0.000")  'Importe
        
        ListBox_Consignaciones.List(ListBox_Consignaciones.ListIndex, 2) = Restante 'Cantidad
        NuevoImporteConsignaciones = Restante * Val(ListBox_Consignaciones.List(ListBox_Consignaciones.ListIndex, 3))
        ListBox_Consignaciones.List(ListBox_Consignaciones.ListIndex, 4) = Format(NuevoImporteConsignaciones, "0.000") 'Importe
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
            If ListBox_PorPagar.List(ListBox_PorPagar.ListIndex, 0) = ListBox_Consignaciones.List(i, 0) Then 'Si existe
            
                ListBox_Consignaciones.List(i, 2) = Val(ListBox_Consignaciones.List(i, 2)) + Val(ListBox_PorPagar.List(ListBox_PorPagar.ListIndex, 2))
                ListBox_Consignaciones.List(i, 4) = Format(Val(ListBox_Consignaciones.List(i, 4)) + Val(ListBox_PorPagar.List(ListBox_PorPagar.ListIndex, 4)), "0.000")
                
                ListBox_PorPagar.RemoveItem (ListBox_PorPagar.ListIndex)
                ActualizarImporte
                
                Exit Sub
            End If
        Next i
    End If
    
    'No existe en el listado de consignaciones
    ListBox_Consignaciones.AddItem
    ListBox_Consignaciones.List(a, 0) = ListBox_PorPagar.List(ListBox_PorPagar.ListIndex, 0)
    ListBox_Consignaciones.List(a, 1) = ListBox_PorPagar.List(ListBox_PorPagar.ListIndex, 1)
    ListBox_Consignaciones.List(a, 2) = ListBox_PorPagar.List(ListBox_PorPagar.ListIndex, 2)
    ListBox_Consignaciones.List(a, 3) = ListBox_PorPagar.List(ListBox_PorPagar.ListIndex, 3)
    ListBox_Consignaciones.List(a, 4) = ListBox_PorPagar.List(ListBox_PorPagar.ListIndex, 4)
    
    ListBox_PorPagar.RemoveItem (ListBox_PorPagar.ListIndex)
    
    
    ActualizarImporte
    
End Sub
