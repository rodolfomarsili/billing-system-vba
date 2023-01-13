VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form_Historial 
   Caption         =   "Historial"
   ClientHeight    =   11055
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   16176
   OleObjectBlob   =   "form_Historial.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "form_Historial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CheckBox_Correlativo_Change()
    
    Label_AsteriscoTipoDeTransaccion.Visible = False
    Label_AsteriscoID1.Visible = False
    Label_AsteriscoID2.Visible = False
    Label_AsteriscoDesdeFecha.Visible = False
    Label_AsteriscoHastaFecha.Visible = False
    
    If CheckBox_Correlativo.Value = True Then
        
        TextBox_ID1.Text = "0001"
        
        ComboBox_TipoDeTransaccionID.Enabled = True
        TextBox_ID1.Enabled = True
        TextBox_ID2.Enabled = True
        
        TextBox_DesdeFechaDia.Enabled = False
        TextBox_DesdeFechaMes.Enabled = False
        TextBox_DesdeFechaAno.Enabled = False
        TextBox_HastaFechaDia.Enabled = False
        TextBox_HastaFechaMes.Enabled = False
        TextBox_HastaFechaAno.Enabled = False
        ComboBox_TipoDeTransaccion.Enabled = False
        ComboBox_Cliente.Enabled = False
        ComboBox_Responsable.Enabled = False
        ComboBox_Caja.Enabled = False
        ComboBox_Producto.Enabled = False
        CheckBox_TipoDeTransaccion.Enabled = False
        CheckBox_Cliente.Enabled = False
        CheckBox_Responsable.Enabled = False
        CheckBox_Caja.Enabled = False
        CheckBox_Producto.Enabled = False
        
    Else
        
        TextBox_ID1.Text = Empty
    
        ComboBox_TipoDeTransaccionID.Enabled = False
        TextBox_ID1.Enabled = False
        TextBox_ID2.Enabled = False
        
        TextBox_DesdeFechaDia.Enabled = True
        TextBox_DesdeFechaMes.Enabled = True
        TextBox_DesdeFechaAno.Enabled = True
        TextBox_HastaFechaDia.Enabled = True
        TextBox_HastaFechaMes.Enabled = True
        TextBox_HastaFechaAno.Enabled = True
        ComboBox_TipoDeTransaccion.Enabled = True
        ComboBox_Cliente.Enabled = True
        ComboBox_Responsable.Enabled = True
        ComboBox_Caja.Enabled = True
        ComboBox_Producto.Enabled = True
        CheckBox_TipoDeTransaccion.Enabled = True
        CheckBox_Cliente.Enabled = True
        CheckBox_Responsable.Enabled = True
        CheckBox_Caja.Enabled = True
        CheckBox_Producto.Enabled = True
        
    End If
    
End Sub


Private Sub CheckBox_TipoDeTransaccion_Change()
        LimpiarFiltros
End Sub

Private Sub CheckBox_Cliente_Change()
        LimpiarFiltros
End Sub

Private Sub CheckBox_Producto_Change()
        LimpiarFiltros
End Sub

Private Sub CheckBox_Responsable_Change()
        LimpiarFiltros
End Sub

Private Sub CheckBox_Caja_Change()
        LimpiarFiltros
End Sub


Private Sub CommandButton_Buscar_Click()

Dim FilaCliente As Integer
Dim FilaResponsable As Byte
Dim FilaCorrelativo As Byte
Dim FilaCodigo As Integer
Dim i As Long
Dim a As Long
Dim UltimaColumnaHistorial As String
Dim ImporteTotal As Single

    Inicializar
    
    'Ocultar todos los asteriscos
    Label_AsteriscoDesdeFecha.Visible = False
    Label_AsteriscoHastaFecha.Visible = False
    Label_AsteriscoTipoDeTransaccion.Visible = False
    Label_AsteriscoID1.Visible = False
    Label_AsteriscoID2.Visible = False
    
    LimpiarFiltros
    TextBox_Comentario = Empty
    
    If Mid(ComboBox_TipoDeTransaccion.Text, 1, 6) = "Compra" Then
        Label7.Caption = "R$"
    Else
        Label7.Caption = "$"
    End If
    
    
    'Limpieza de tabla de historial temporal
    If Not HojaHistorialTemporal.Range("A1") = Empty Then
        HojaHistorialTemporal.Range("A1").CurrentRegion.Delete
    End If
    
    'Limpieza de las tablas de filtros
    Workbooks("Historial.xlsm").Sheets("Inicio").Range("A101:G101").Clear
    Workbooks("Historial.xlsm").Sheets("Inicio").Range("A111:G111").Clear
    
        
            

    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Busqueda Por Correlativo'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If CheckBox_Correlativo.Value = True Then
        
        If Not ComboBox_TipoDeTransaccionID.ListIndex > -1 Then
            MsgBox "Debes seleccionar una opcion valida para el Tipo de Transaccion "
            Exit Sub
        End If
    
        If (ComboBox_TipoDeTransaccionID = Empty Or TextBox_ID1 = Empty Or TextBox_ID2 = Empty) Then
            Label_AsteriscoTipoDeTransaccion.Visible = True
            Label_AsteriscoID1.Visible = True
            Label_AsteriscoID2.Visible = True
            MsgBox "Debes rellenar TODOS los campos del recuadro de transaccion", , "Historial"
            Exit Sub
        End If
        
        'Filtro Tipo de Transaccion
        FilaCorrelativo = ObtenerFila(HojaCorrelativos, ComboBox_TipoDeTransaccionID.Text, ColumnaLeyenda)
        Workbooks("Historial.xlsm").Sheets("Inicio").Range("A111") = HojaCorrelativos.Cells(FilaCorrelativo, ColumnaPrefijo)
        'Filtro ID1
        Workbooks("Historial.xlsm").Sheets("Inicio").Range("B111") = TextBox_ID1
        'Filtro ID2
        Workbooks("Historial.xlsm").Sheets("Inicio").Range("C111") = TextBox_ID2
        
        'LLenado de tabla de historial temporal
        FiltrarHistorialPorTransaccion
    
    End If
      
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Busqueda Por Filtros'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    If CheckBox_Correlativo.Value = False Then
    
        If (TextBox_DesdeFechaDia = Empty Or TextBox_DesdeFechaMes = Empty Or TextBox_DesdeFechaAno = Empty Or TextBox_HastaFechaDia = Empty Or TextBox_HastaFechaMes = Empty Or TextBox_HastaFechaAno = Empty) Then
            Label_AsteriscoDesdeFecha.Visible = True
            Label_AsteriscoHastaFecha.Visible = True
            MsgBox "Debes rellenar los campos de fecha", , "Historial"
            Exit Sub
        End If
    
        'Establecimiento de los filtros
        
        'Filtro Desde Fecha
        Workbooks("Historial.xlsm").Sheets("Inicio").Range("A101") = ">=" & TextBox_DesdeFechaMes & "/" & TextBox_DesdeFechaDia & "/" & TextBox_DesdeFechaAno
        'Filtro Hasta Fecha
        Workbooks("Historial.xlsm").Sheets("Inicio").Range("B101") = "<=" & TextBox_HastaFechaMes & "/" & TextBox_HastaFechaDia & "/" & TextBox_HastaFechaAno
        'Filtro Tipo de Transaccion
        If (CheckBox_TipoDeTransaccion.Value And ComboBox_TipoDeTransaccion <> "") = True Then
            If (ComboBox_TipoDeTransaccion = "Ventas") Then
                Workbooks("Historial.xlsm").Sheets("Inicio").Range("C101") = "*" & "VTA" & "*"
            Else
                FilaCorrelativo = ObtenerFila(HojaCorrelativos, ComboBox_TipoDeTransaccion.Text, ColumnaLeyenda)
                Workbooks("Historial.xlsm").Sheets("Inicio").Range("C101") = HojaCorrelativos.Cells(FilaCorrelativo, ColumnaPrefijo)
            End If
        Else
            Workbooks("Historial.xlsm").Sheets("Inicio").Range("C101") = Empty
        End If
        'Filtro Cliente
        If (CheckBox_Cliente.Value And ComboBox_Cliente <> "") = True Then
            FilaCliente = ObtenerFila(HojaClientes, ComboBox_Cliente.Text, ColumnaNombreCliente)
            Workbooks("Historial.xlsm").Sheets("Inicio").Range("D101") = HojaClientes.Cells(FilaCliente, ColumnaIDCliente)
        Else
            Workbooks("Historial.xlsm").Sheets("Inicio").Range("D101") = Empty
        End If
        'Filtro Producto
        If (CheckBox_Producto.Value And ComboBox_Producto <> "") = True Then
            FilaCodigo = ObtenerFila(HojaInventario, ComboBox_Producto.Text, ColumnaProducto)
            Workbooks("Historial.xlsm").Sheets("Inicio").Range("E101") = HojaInventario.Cells(FilaCodigo, ColumnaCodigo)
        Else
            Workbooks("Historial.xlsm").Sheets("Inicio").Range("E101") = Empty
        End If
        'Filtro Responsable
        If (CheckBox_Responsable.Value And ComboBox_Responsable <> "") = True Then
            FilaResponsable = ObtenerFila(HojaUsuarios, ComboBox_Responsable.Text, ColumnaNombreUsuario)
            Workbooks("Historial.xlsm").Sheets("Inicio").Range("F101") = HojaUsuarios.Cells(FilaResponsable, ColumnaIDUsuario)
        Else
            Workbooks("Historial.xlsm").Sheets("Inicio").Range("F101") = Empty
        End If
        'Filtro Tipo de Caja
        If (CheckBox_Caja.Value And ComboBox_Caja <> "") = True Then
            Workbooks("Historial.xlsm").Sheets("Inicio").Range("G101") = ComboBox_Caja
        Else
            Workbooks("Historial.xlsm").Sheets("Inicio").Range("G101") = Empty
        End If
        
        
        'LLenado de tabla de historial temporal
        FiltrarHistorialPorFecha
        
    End If
        
        
    ''''''''''''''''''''''''''''''''''''''''''''''''''''Llenado de Historial de Formulario'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
            
        UltimaColumnaHistorial = ObtenerLetraDeColumna(ObtenerUltimaColumna(HojaHistorial, 1))
        
        'Actualizar ultima fila del historial temporal
        UltimaFilaHistorialTemporal = ObtenerUltimaFila(HojaHistorialTemporal, ColumnaFechaHistorial)
        
        'Limpieza de tabla de historial en el formulario
        ListBox_Historial.Clear
        ListBox_Comentarios.Clear
        ListBox_Devoluciones.Clear
        
        'Llenado de tabla de historial en el formulario
        ImporteTotal = 0
        a = 0
        For i = 2 To UltimaFilaHistorialTemporal
            If i = 2 Then
                    ListBox_Historial.AddItem
                    ListBox_Historial.List(a, 0) = HojaHistorialTemporal.Cells(i, ColumnaFechaHistorial)
                    If Not HojaHistorialTemporal.Cells(i, ColumnaTipoDeTransaccionHistorial) = Empty Then ListBox_Historial.List(a, 1) = HojaHistorialTemporal.Cells(i, ColumnaTipoDeTransaccionHistorial) & " - " & HojaHistorialTemporal.Cells(i, ColumnaID1Historial) & " - " & HojaHistorialTemporal.Cells(i, ColumnaID2Historial)
                    ListBox_Historial.List(a, 2) = HojaHistorialTemporal.Cells(i, ColumnaProductoHistorial)
                    ListBox_Historial.List(a, 3) = HojaHistorialTemporal.Cells(i, ColumnaIDCajaHistorial)
                    ListBox_Historial.List(a, 4) = HojaHistorialTemporal.Cells(i, ColumnaCantidadHistorial)
                    ListBox_Historial.List(a, 5) = HojaHistorialTemporal.Cells(i, ColumnaNuevaExistenciaHistorial)
                    ListBox_Historial.List(a, 6) = Format(HojaHistorialTemporal.Cells(i, ColumnaCostoHistorial), "0.0000")
                    ListBox_Historial.List(a, 7) = Format(HojaHistorialTemporal.Cells(i, ColumnaPrecioHistorial), "0.0000")
                    ListBox_Historial.List(a, 8) = Format(HojaHistorialTemporal.Cells(i, ColumnaImporteHistorial), "0.00")
        
                    ListBox_Comentarios.AddItem
                    ListBox_Comentarios.List(a, 0) = HojaHistorialTemporal.Cells(i, ColumnaDescripcionHistorial)
                    
                    ListBox_Devoluciones.AddItem
                    ListBox_Devoluciones.List(a, 0) = HojaHistorialTemporal.Cells(i, ColumnaDevueltoHistorial)
                        
                    ImporteTotal = ImporteTotal + Val(HojaHistorialTemporal.Cells(i, ColumnaImporteHistorial))
                    a = a + 1
            Else
                    'Si la concatenacion TIPO DE DE TRANSACCION, ID1 E ID2, de la fila actual del historial temporal, resultaser igual a la de la fila
                    'anterior, entonces se muestran todas las columnas del historial menos la de tipo de transaccion. En caso contrario, se muestra todo.
                    
                    If HojaHistorialTemporal.Cells(i, ColumnaTipoDeTransaccionHistorial) & " - " & HojaHistorialTemporal.Cells(i, ColumnaID1Historial) & " - " & HojaHistorialTemporal.Cells(i, ColumnaID2Historial) = HojaHistorialTemporal.Cells(i - 1, ColumnaTipoDeTransaccionHistorial) & " - " & HojaHistorialTemporal.Cells(i - 1, ColumnaID1Historial) & " - " & HojaHistorialTemporal.Cells(i - 1, ColumnaID2Historial) Then
                            ListBox_Historial.AddItem
                            'ListBox_Historial.List(a, 0) = HojaHistorialTemporal.Cells(i, ColumnaFechaHistorial)
                            ListBox_Historial.List(a, 2) = HojaHistorialTemporal.Cells(i, ColumnaProductoHistorial)
                            ListBox_Historial.List(a, 3) = HojaHistorialTemporal.Cells(i, ColumnaIDCajaHistorial)
                            ListBox_Historial.List(a, 4) = HojaHistorialTemporal.Cells(i, ColumnaCantidadHistorial)
                            ListBox_Historial.List(a, 5) = HojaHistorialTemporal.Cells(i, ColumnaNuevaExistenciaHistorial)
                            ListBox_Historial.List(a, 6) = Format(HojaHistorialTemporal.Cells(i, ColumnaCostoHistorial), "0.0000")
                            ListBox_Historial.List(a, 7) = Format(HojaHistorialTemporal.Cells(i, ColumnaPrecioHistorial), "0.0000")
                            ListBox_Historial.List(a, 8) = Format(HojaHistorialTemporal.Cells(i, ColumnaImporteHistorial), "0.00")
                
                            ListBox_Comentarios.AddItem
                            ListBox_Comentarios.List(a, 0) = HojaHistorialTemporal.Cells(i, ColumnaDescripcionHistorial)
                            
                            ListBox_Devoluciones.AddItem
                            ListBox_Devoluciones.List(a, 0) = HojaHistorialTemporal.Cells(i, ColumnaDevueltoHistorial)
                
                            ImporteTotal = ImporteTotal + Val(HojaHistorialTemporal.Cells(i, ColumnaImporteHistorial))
                            a = a + 1
                    Else
                            ListBox_Historial.AddItem
                            ListBox_Historial.List(a, 0) = HojaHistorialTemporal.Cells(i, ColumnaFechaHistorial)
                            If Not HojaHistorialTemporal.Cells(i, ColumnaTipoDeTransaccionHistorial) = Empty Then ListBox_Historial.List(a, 1) = HojaHistorialTemporal.Cells(i, ColumnaTipoDeTransaccionHistorial) & " - " & HojaHistorialTemporal.Cells(i, ColumnaID1Historial) & " - " & HojaHistorialTemporal.Cells(i, ColumnaID2Historial)
                            ListBox_Historial.List(a, 2) = HojaHistorialTemporal.Cells(i, ColumnaProductoHistorial)
                            ListBox_Historial.List(a, 3) = HojaHistorialTemporal.Cells(i, ColumnaIDCajaHistorial)
                            ListBox_Historial.List(a, 4) = HojaHistorialTemporal.Cells(i, ColumnaCantidadHistorial)
                            ListBox_Historial.List(a, 5) = HojaHistorialTemporal.Cells(i, ColumnaNuevaExistenciaHistorial)
                            ListBox_Historial.List(a, 6) = Format(HojaHistorialTemporal.Cells(i, ColumnaCostoHistorial), "0.0000")
                            ListBox_Historial.List(a, 7) = Format(HojaHistorialTemporal.Cells(i, ColumnaPrecioHistorial), "0.0000")
                            ListBox_Historial.List(a, 8) = Format(HojaHistorialTemporal.Cells(i, ColumnaImporteHistorial), "0.00")
                
                            ListBox_Comentarios.AddItem
                            ListBox_Comentarios.List(a, 0) = HojaHistorialTemporal.Cells(i, ColumnaDescripcionHistorial)
                            
                            ListBox_Devoluciones.AddItem
                            ListBox_Devoluciones.List(a, 0) = HojaHistorialTemporal.Cells(i, ColumnaDevueltoHistorial)
                
                            ImporteTotal = ImporteTotal + Val(HojaHistorialTemporal.Cells(i, ColumnaImporteHistorial))
                            a = a + 1
                    End If
                
            End If
        Next i
        
        
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
  
        Label_ImporteTotal.Caption = Format(ImporteTotal, "0.0000")


End Sub

Private Sub CommandButton_Cancelar_Click()
    Unload Me
End Sub


Private Sub ListBox_Historial_Click()
   
Dim FilaIDCliente As Integer
Dim FilaIDResponsable As Integer
Dim FilaIDCorrelativo As Byte
Dim FilaHistorial As Integer

    On Error Resume Next
    
    FilaHistorial = ListBox_Historial.ListIndex + 2
    
    LimpiarFiltros
    
    TextBox_Comentario = Empty
    
    FilaIDCorrelativo = ObtenerFila(HojaCorrelativos, HojaHistorialTemporal.Cells(FilaHistorial, ColumnaTipoDeTransaccionHistorial), ColumnaPrefijo)
    FilaIDCliente = ObtenerFila(HojaClientes, HojaHistorialTemporal.Cells(FilaHistorial, ColumnaIDClienteHistorial), ColumnaIDCliente)
    FilaIDResponsable = ObtenerFila(HojaUsuarios, HojaHistorialTemporal.Cells(FilaHistorial, ColumnaIDResponsableHistorial), ColumnaIDUsuario)

    ComboBox_TipoDeTransaccion.Text = HojaCorrelativos.Cells(FilaIDCorrelativo, ColumnaLeyenda)
    ComboBox_Cliente.Text = HojaClientes.Cells(FilaIDCliente, ColumnaNombreCliente)
    ComboBox_Responsable.Text = HojaUsuarios.Cells(FilaIDResponsable, ColumnaNombreUsuario)
    TextBox_Comentario = ListBox_Comentarios.List(ListBox_Historial.ListIndex, 0)
    
    If Val(ListBox_Devoluciones.List(ListBox_Historial.ListIndex, 0)) > 0 Then
        Label_Devuelto_Titulo.Visible = True
        Label_Devuelto.Visible = True
        Label_Devuelto = ListBox_Devoluciones.List(ListBox_Historial.ListIndex, 0)
    Else
        Label_Devuelto_Titulo.Visible = False
        Label_Devuelto.Visible = False
        Label_Devuelto = 0
    End If
    
End Sub

Private Sub TextBox_DesdeFechaDia_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
Dim Campo As Object
    Set Campo = TextBox_DesdeFechaDia
    Select Case Len(Campo)
        Case 0: If Not (KeyAscii >= 48 And KeyAscii <= 51) Then KeyAscii = 0
        Case 1: If Not (KeyAscii >= 48 And KeyAscii <= 57) Then KeyAscii = 0
        Case 2: KeyAscii = 0
    End Select
End Sub

Private Sub TextBox_DesdeFechaMes_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
Dim Campo As Object
    Set Campo = TextBox_DesdeFechaMes
    Select Case Len(Campo)
        Case 0: If Not (KeyAscii >= 48 And KeyAscii <= 51) Then KeyAscii = 0
        Case 1: If Not (KeyAscii >= 48 And KeyAscii <= 57) Then KeyAscii = 0
        Case 2: KeyAscii = 0
    End Select
End Sub

Private Sub TextBox_DesdeFechaAno_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
Dim Campo As Object
    Set Campo = TextBox_DesdeFechaAno
    Select Case Len(Campo)
        Case 0: If Not (KeyAscii = 50) Then KeyAscii = 0
        Case 1: If Not (KeyAscii = 48) Then KeyAscii = 0
        Case 2, 3: If Not (KeyAscii >= 48 And KeyAscii <= 57) Then KeyAscii = 0
        Case 4: KeyAscii = 0
    End Select
End Sub

Private Sub TextBox_HastaFechaDia_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
Dim Campo As Object
    Set Campo = TextBox_HastaFechaDia
    Select Case Len(Campo)
        Case 0: If Not (KeyAscii >= 48 And KeyAscii <= 51) Then KeyAscii = 0
        Case 1: If Not (KeyAscii >= 48 And KeyAscii <= 57) Then KeyAscii = 0
        Case 2: KeyAscii = 0
    End Select
End Sub

Private Sub TextBox_HastaFechaMes_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
Dim Campo As Object
    Set Campo = TextBox_HastaFechaMes
    Select Case Len(Campo)
        Case 0: If Not (KeyAscii >= 48 And KeyAscii <= 51) Then KeyAscii = 0
        Case 1: If Not (KeyAscii >= 48 And KeyAscii <= 57) Then KeyAscii = 0
        Case 2: KeyAscii = 0
    End Select
End Sub

Private Sub TextBox_HastaFechaAno_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
Dim Campo As Object
    Set Campo = TextBox_HastaFechaAno
    Select Case Len(Campo)
        Case 0: If Not (KeyAscii = 50) Then KeyAscii = 0
        Case 1: If Not (KeyAscii = 48) Then KeyAscii = 0
        Case 2, 3: If Not (KeyAscii >= 48 And KeyAscii <= 57) Then KeyAscii = 0
        Case 4: KeyAscii = 0
    End Select
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

Dim i As Integer
    
    Inicializar
    
    TextBox_DesdeFechaDia = Day(Date)
    TextBox_DesdeFechaMes = Month(Date)
    TextBox_DesdeFechaAno = Year(Date)
    
    TextBox_HastaFechaDia = Day(Date)
    TextBox_HastaFechaMes = Month(Date)
    TextBox_HastaFechaAno = Year(Date)
    
    'Establecer propiedades del listado
    ListBox_Historial.List = HojaHistorialTemporal.Range("A1:I1").Value 'Esta linea es para engañar a vba y dimensionar la lista con mas de 10 columnas
    ListBox_Historial.Clear 'Esta linea va en conjunto con la anterior
    ListBox_Historial.ColumnCount = 10
    ListBox_Comentarios.ColumnCount = 1
    ListBox_Historial.ColumnWidths = "70 pt; 120 pt; 190 pt; 70 pt; 60 pt; 85 pt; 60 pt; 60 pt; 60 pt"
    
    'Ocultar asteriscos de error
    Label_AsteriscoTipoDeTransaccion.Visible = False
    Label_AsteriscoID1.Visible = False
    Label_AsteriscoID2.Visible = False
    Label_AsteriscoDesdeFecha.Visible = False
    Label_AsteriscoHastaFecha.Visible = False
            
    'Deshabilitar la busqueda por correlativo
    CheckBox_Correlativo.Value = False
    ComboBox_TipoDeTransaccionID.Enabled = False
    TextBox_ID1.Enabled = False
    TextBox_ID2.Enabled = False
    
    'Llenado combobox Tipo de Transaccion
        ComboBox_TipoDeTransaccion.AddItem ("Ventas")
    For i = 2 To UltimaFilaCorrelativos
        ComboBox_TipoDeTransaccionID.AddItem (HojaCorrelativos.Cells(i, ColumnaLeyenda))
        ComboBox_TipoDeTransaccion.AddItem (HojaCorrelativos.Cells(i, ColumnaLeyenda))
    Next i
        
        
    'Llenado combobox Cliente
    For i = 2 To UltimaFilaClientes
        ComboBox_Cliente.AddItem (HojaClientes.Cells(i, ColumnaNombreCliente))
    Next i
    
    'Llenado combobox Responsable
    For i = 2 To UltimaFilaUsuarios
        ComboBox_Responsable.AddItem (HojaUsuarios.Cells(i, ColumnaNombreUsuario))
    Next i
    
    'Llenado combobox Caja
    For i = 2 To UltimaFilaCajas
        If (Mid(HojaCajas.Cells(i, ColumnaIDCaja), 1, 3) = "USD") Then ComboBox_Caja.AddItem (HojaCajas.Cells(i, ColumnaIDCaja))
    Next i
    
    'Llenado combobox Producto
    For i = 2 To UltimaFilaInventario
        ComboBox_Producto.AddItem (HojaInventario.Cells(i, ColumnaProducto))
    Next i
    
    Label_Devuelto_Titulo.Visible = False
    Label_Devuelto.Visible = False
    Label_Devuelto = 0
    
    LimpiarFiltros
    
    
    
    
End Sub

Private Sub UserForm_Terminate()

    If Not HojaHistorialTemporal.Range("A1") = Empty Then HojaHistorialTemporal.Range("A1").CurrentRegion.Delete
    
    Workbooks("Historial.xlsm").Sheets("Inicio").Range("A101:G101").Clear
    Workbooks("Historial.xlsm").Sheets("Inicio").Range("A111:G111").Clear
    
End Sub

Private Sub LimpiarFiltros()

        'Filtro Tipo de Transaccion
        If CheckBox_TipoDeTransaccion.Value = False Then
                ComboBox_TipoDeTransaccion = Empty
                ComboBox_TipoDeTransaccion.Locked = True
        Else
                ComboBox_TipoDeTransaccion.Locked = False
        End If
        
        'Filtro Cliente
        If CheckBox_Cliente.Value = False Then
                ComboBox_Cliente = Empty
                ComboBox_Cliente.Locked = True
        Else
                ComboBox_Cliente.Locked = False
        End If
        
        'Filtro Producto
        If CheckBox_Producto.Value = False Then
                ComboBox_Producto = Empty
                ComboBox_Producto.Locked = True
        Else
                ComboBox_Producto.Locked = False
        End If
        
        'Filtro Responsable
        If CheckBox_Responsable.Value = False Then
                ComboBox_Responsable = Empty
                ComboBox_Responsable.Locked = True
        Else
                ComboBox_Responsable.Locked = False
        End If
        
        'Filtro Tipo de Caja
        If CheckBox_Caja.Value = False Then
                ComboBox_Caja = Empty
                ComboBox_Caja.Locked = True
        Else
                ComboBox_Caja.Locked = False
        End If
        
End Sub
