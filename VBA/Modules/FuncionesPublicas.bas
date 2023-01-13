Attribute VB_Name = "FuncionesPublicas"
Option Explicit

Sub BloquearAcceso()
    
    Inicializar
    
    'Mostrar Inicio
    HojaInicio.Visible = xlSheetVisible
    'Ocultar Dashboard
    HojaDashboard.Visible = xlSheetHidden
    
    HojaGestion.Range("B2") = Empty
    HojaGestion.Range("B3") = Empty
    HojaGestion.Range("B4") = Empty
    HojaGestion.Range("B5") = "Bloqueado"
    
    
End Sub

Sub BotonMenu()
    
    Inicializar
    
    If HojaGestion.Range("B5") = "Bloqueado" Then
        Load form_Login
        form_Login.Show
    Else
        Load form_Menu
        form_Menu.Show
    End If
    
End Sub

Sub SeleccionarTextoTextBox(TextBox As Object)
    
    With TextBox
        .SetFocus
        .SelStart
        .SelLength = Len(.Text)
    End With
    
End Sub

Function ObtenerColumnaDeTabla(HojaDeTrabajo As Worksheet, FilaCabeceraDeTabla As Integer, ColumnaDeTablaBuscada As String) As Byte

Dim i As Integer

    For i = 1 To HojaDeTrabajo.Cells(FilaCabeceraDeTabla, Columns.Count).End(xlToLeft).Column
    
        If (ColumnaDeTablaBuscada = HojaDeTrabajo.Cells(FilaCabeceraDeTabla, i)) Then
            ObtenerColumnaDeTabla = i
            Exit For
        End If
        
    Next i
    
End Function


Function ObtenerUltimaFila(HojaDeTrabajo As Worksheet, ColumnaDeFilaBuscada As Byte) As Long

    'Ultima fila de la columna de fila buscada de la hoja HojaDeTrabajo
    ObtenerUltimaFila = HojaDeTrabajo.Cells(Rows.Count, ColumnaDeFilaBuscada).End(xlUp).Row

End Function


Function ObtenerUltimaColumna(HojaDeTrabajo As Worksheet, FilaDeColumnaBuscada As Integer) As Byte

    'Ultima columna, en forma de numero, de la fia de columna buscada de la hoja HojaDeTrabajo
    ObtenerUltimaColumna = HojaDeTrabajo.Cells(FilaDeColumnaBuscada, Columns.Count).End(xlToLeft).Column

End Function


Function ObtenerLetraDeColumna(NumeroDeColumnaBuscada As Byte) As String

Dim Direccion As String
Dim DireccionSinFila As String
Dim LetraDeColumna As String

Rem Aca se consigue la direccion de la columna en base al numero de la misma
    Direccion = ActiveSheet.Cells(1, NumeroDeColumnaBuscada).Address(RowAbsolute:=False)
Rem Aca se elimina el numero que denota la fila que inevitablemente esta al usar la funcion address
    DireccionSinFila = Replace(Direccion, 1, "")
Rem Aca se elimina el simbolo $ que inevitablemente esta al usar la funcion address
    LetraDeColumna = Replace(DireccionSinFila, "$", "")

    ObtenerLetraDeColumna = LetraDeColumna

End Function

Function ObtenerFila(HojaDeTrabajo As Worksheet, ValorBuscado As String, ColumnaDeFilaBuscada As Byte) As Long

Dim i As Long
Dim UltimaFila As Long

    ObtenerFila = 0
    
    UltimaFila = ObtenerUltimaFila(HojaDeTrabajo, ColumnaDeFilaBuscada)
    
    For i = 1 To UltimaFila
        If HojaDeTrabajo.Cells(i, ColumnaDeFilaBuscada) = ValorBuscado Then
            ObtenerFila = i
            Exit For
        End If
    Next i
    
    
        

End Function

Function ObtenerFila2(HojaDeTrabajo As Worksheet, ValorBuscado As String, ColumnaDeFilaBuscada As Byte) As Long

Dim LetraDeColumna As String
Dim UltimaFila As Long

Dim RangoDeBusqueda As Range

    LetraDeColumna = ObtenerLetraDeColumna(ColumnaDeFilaBuscada)
    UltimaFila = ObtenerUltimaFila(HojaDeTrabajo, ColumnaDeFilaBuscada)
    
    Set RangoDeBusqueda = HojaDeTrabajo.Range(LetraDeColumna & "1:" & LetraDeColumna & UltimaFila)
    
    ObtenerFila2 = WorksheetFunction.Match(ValorBuscado, RangoDeBusqueda, 0)

End Function

Sub ModificarExistenciaInventario(Codigo As Variant, NuevaExistencia As Long)

Dim LetraDeColumna As String
Dim FilaCodigo As Integer
Dim RangoDeBusqueda As Range
    
    'Se obtiene la letra de la columna donde se encuentran los codigos de los productos
    LetraDeColumna = ObtenerLetraDeColumna(ColumnaCodigo)
    'Se establece el rango de busqueda desde el inicio de la columna hasta el final de la misma
    Set RangoDeBusqueda = HojaInventario.Range(LetraDeColumna & "1:" & LetraDeColumna & UltimaFilaInventario)
    'Se obtiene la fila del producto a modificar
    FilaCodigo = WorksheetFunction.Match(Codigo, RangoDeBusqueda, 0)
    
    'Se modifica la existencia en el inventario
    HojaInventario.Cells(FilaCodigo, ColumnaExistencia) = NuevaExistencia
    
End Sub

Function ObtenerFilaProductoDeTransaccion(TipoDeTransaccion As String, ID1Transaccion As String, ID2Transaccion As String, CodigoDeProducto As String) As Long

Dim LetraDeColumnaAuxiliar As String
Dim ValorBuscado As String
Dim RangoDeBusqueda As Range
    
    Inicializar
    
    ValorBuscado = TipoDeTransaccion & "|" & ID1Transaccion & "|" & ID2Transaccion & "|" & CodigoDeProducto
    
    LetraDeColumnaAuxiliar = ObtenerLetraDeColumna(ColumnaDEVAUXHistorial)
    Set RangoDeBusqueda = HojaHistorial.Range(LetraDeColumnaAuxiliar & "1:" & LetraDeColumnaAuxiliar & UltimaFilaHistorial)
    
    ObtenerFilaProductoDeTransaccion = WorksheetFunction.Match(ValorBuscado, RangoDeBusqueda, 0)
    
End Function

Sub ReordenarInventario()

Dim Rango As String
Dim LetraColumna As String

    Inicializar
    
    LetraColumna = ObtenerLetraDeColumna(ColumnaProducto)
    Rango = LetraColumna & "2:" & LetraColumna & UltimaFilaInventario
    
    HojaInventario.ListObjects("Inventario").Sort. _
        SortFields.Clear
    HojaInventario.ListObjects("Inventario").Sort. _
        SortFields.Add Key:=Range(Rango), SortOn:= _
        xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With HojaInventario.ListObjects("Inventario").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

End Sub

Sub ReordenarInventarioClientes(HojaDeTrabajo As Worksheet)

    HojaDeTrabajo.ListObjects(1).Sort.SortFields.Clear
    
    HojaDeTrabajo.ListObjects(1).Sort.SortFields.Add _
    Key:=HojaDeTrabajo.ListObjects(1).ListColumns(2).Range, _
    SortOn:=xlSortOnValues, _
    Order:=xlAscending, _
    DataOption:=xlSortNormal
    
    With HojaDeTrabajo.ListObjects(1).Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

End Sub

Sub ReordenarClientes()

Dim Rango As String
Dim LetraColumna As String

    Inicializar
    
    LetraColumna = ObtenerLetraDeColumna(ColumnaNombreCliente)
    Rango = LetraColumna & "2:" & LetraColumna & UltimaFilaClientes
    
    HojaClientes.ListObjects("Clientes").Sort. _
        SortFields.Clear
    HojaClientes.ListObjects("Clientes").Sort. _
        SortFields.Add Key:=Range(Rango), SortOn:= _
        xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With HojaClientes.ListObjects("Clientes").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

End Sub

Sub ReordenarHistorial()

Dim Rango As String
Dim LetraColumna As String

    Inicializar
    
    LetraColumna = ObtenerLetraDeColumna(ColumnaFechaHistorial)
    Rango = LetraColumna & "2:" & LetraColumna & UltimaFilaHistorial

    HojaHistorial.ListObjects("Historia").Sort.SortFields.Clear
        
    HojaHistorial.ListObjects("Historia").Sort.SortFields.Add _
         Key:=Range(Rango), SortOn:=xlSortOnValues, Order:= _
        xlDescending, DataOption:=xlSortNormal
    With HojaHistorial.ListObjects("Historia").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
End Sub


Sub ActualizarSaldoCajaEnPantalla(LabelSaldoDeCaja As MSForms.Label, IDCaja As String)

Dim FilaCaja As Byte

    On Error Resume Next
    
    FilaCaja = ObtenerFila(HojaCajas, IDCaja, ColumnaIDCaja)
        Select Case Mid(IDCaja, 1, 3)
            Case "USD": LabelSaldoDeCaja.Caption = "Saldo Actual: " & Format(HojaCajas.Cells(FilaCaja, ColumnaSaldoCaja), "0,0.00") & " $"
            Case "BRL": LabelSaldoDeCaja.Caption = "Saldo Actual: " & Format(HojaCajas.Cells(FilaCaja, ColumnaSaldoCaja), "0,0.00") & " R$"
            Case "VES": LabelSaldoDeCaja.Caption = "Saldo Actual: " & Format(HojaCajas.Cells(FilaCaja, ColumnaSaldoCaja), "0,0.00") & " Bs"
        End Select
        
End Sub


Sub ActualizarCorrelativo(FrameCorrelativo As MSForms.Frame, LabelPrefijoCorrelativo As MSForms.Label)

Dim FilaDeCorrelativo As Byte
Dim Correlativo As String
Dim ID1Correlativo As String
Dim ID2Correlativo As String

    Inicializar
    
    Correlativo = ObtenerCorrelativo(FrameCorrelativo, LabelPrefijoCorrelativo)

    FilaDeCorrelativo = ObtenerFila(HojaCorrelativos, Correlativo, ColumnaPrefijo)
    
    ID1Correlativo = HojaCorrelativos.Cells(FilaDeCorrelativo, ColumnaID1)
    ID2Correlativo = HojaCorrelativos.Cells(FilaDeCorrelativo, ColumnaID2)
    
    
    'Aumento del numero del correlativo
    If Val(ID2Correlativo) < 10000 Then
        
        ID2Correlativo = Val(ID2Correlativo) + 1
        
        If (Val(ID2Correlativo) < 10) Then ID2Correlativo = "000" & ID2Correlativo
        If (Val(ID2Correlativo) >= 10 And Val(ID2Correlativo) < 100) Then ID2Correlativo = "00" & ID2Correlativo
        If (Val(ID2Correlativo) >= 100 And Val(ID2Correlativo) < 1000) Then ID2Correlativo = "0" & ID2Correlativo
    
    Else
        
        ID1Correlativo = Val(ID1Correlativo) + 1

        If (Val(ID1Correlativo) < 10) Then ID1Correlativo = "000" & ID1Correlativo
        If (Val(ID1Correlativo) >= 10 And Val(ID1Correlativo) < 100) Then ID1Correlativo = "00" & ID1Correlativo
        If (Val(ID1Correlativo) >= 100 And Val(ID1Correlativo) < 1000) Then ID1Correlativo = "0" & ID1Correlativo
                
        ID2Correlativo = "0001"
            
    End If
    
    'Se escribe el nuevo correlativo aumentado en la hoja de correlativos
    HojaCorrelativos.Cells(FilaDeCorrelativo, ColumnaID1) = ID1Correlativo
    HojaCorrelativos.Cells(FilaDeCorrelativo, ColumnaID2) = ID2Correlativo
    
End Sub

Function ObtenerCorrelativo(FrameCorrelativo As MSForms.Frame, LabelPrefijoCorrelativo As MSForms.Label) As String

Dim Correlativo As String

    Select Case FrameCorrelativo.Caption
    
        Case "Compra":
                        If LabelPrefijoCorrelativo.Caption = "Contado" Then Correlativo = "COM-CTD"
                        If LabelPrefijoCorrelativo.Caption = "Credito" Then Correlativo = "COM-CDT"
                        If LabelPrefijoCorrelativo.Caption = "Consignacion" Then Correlativo = "COM-CSN"
        
        Case "Venta":
                        If LabelPrefijoCorrelativo.Caption = "Contado" Then Correlativo = "VTA-CTD"
                        If LabelPrefijoCorrelativo.Caption = "Credito" Then Correlativo = "VTA-CDT"
                        If LabelPrefijoCorrelativo.Caption = "Consignacion" Then Correlativo = "VTA-CSN"
        Case "Pago":
                        If LabelPrefijoCorrelativo.Caption = "Credito" Then Correlativo = "PGO-CDT"
                        If LabelPrefijoCorrelativo.Caption = "Consignacion" Then Correlativo = "PGO-CSN"
        Case "Modificacion":
                        Correlativo = "MOD"
        Case "Registro":
                        Correlativo = "REG"
        Case "Devolucion":
                        Correlativo = "DEV"
        Case "Prestamo":
                        Correlativo = "PMO"
        Case "Caja":
                        Correlativo = "MTO-CJA"
        Case "Extras":
                        Correlativo = "EXT"
        Case "Inventario":
                        If LabelPrefijoCorrelativo.Caption = "Recargo" Then Correlativo = "REC-INV"
                        If LabelPrefijoCorrelativo.Caption = "Descargo" Then Correlativo = "DES-INV"
                        
    End Select
    
    ObtenerCorrelativo = Correlativo
    
End Function


Sub ActualizarCorrelativoEnPantalla(FrameCorrelativo As MSForms.Frame, LabelPrefijoCorrelativo As MSForms.Label)

Dim FilaDeCorrelativo As Byte
Dim Correlativo As String

    Inicializar
    
    Correlativo = ObtenerCorrelativo(FrameCorrelativo, LabelPrefijoCorrelativo)
    
    'Actualizar correlativo en pantalla
    FilaDeCorrelativo = ObtenerFila(HojaCorrelativos, Correlativo, ColumnaPrefijo)
    
    'Se establece el correlativo en la factura
    FormularioAnterior.Label_CorrelativoID1.Caption = HojaCorrelativos.Cells(FilaDeCorrelativo, ColumnaID1)
    FormularioAnterior.Label_CorrelativoID2.Caption = HojaCorrelativos.Cells(FilaDeCorrelativo, ColumnaID2)
    
End Sub


Sub IncluirEnHistorial(FrameCorrelativo As MSForms.Frame, LabelPrefijoCorrelativo As MSForms.Label, Fecha As Date, Optional Codigo As String, Optional Producto As String, Optional IDCaja As String, Optional Cantidad As Long, Optional Descripcion As String, Optional IDCliente As String, Optional IDResponsable As String, Optional Monto As Single, Optional AbonarACaja As Boolean, Optional NuevaExistencia As Long)

Dim FilaDeCorrelativo As Byte
Dim FilaCaja As Byte
Dim FilaCliente As Byte
Dim Hora As String
Dim Importe As Single
Dim ID1Correlativo As String
Dim ID2Correlativo As String
Dim Correlativo As String


    Const NuevaFila As Byte = 2
    
    Inicializar
    
    If Cantidad = Empty Then Cantidad = 1
    
        ' Se inserta una fila entera en en principio de la tabla de historia
    HojaHistorial.Range(NuevaFila & ":" & NuevaFila).EntireRow.Insert
    
    Correlativo = ObtenerCorrelativo(FrameCorrelativo, LabelPrefijoCorrelativo)
    
    FilaDeCorrelativo = ObtenerFila(HojaCorrelativos, Correlativo, ColumnaPrefijo)
    
    ID1Correlativo = HojaCorrelativos.Cells(FilaDeCorrelativo, ColumnaID1)
    ID2Correlativo = HojaCorrelativos.Cells(FilaDeCorrelativo, ColumnaID2)
    
    Hora = Hour(Now) & ":" & Minute(Now)
    
    'Ingresar cada argumento en la columna correspondiente
    HojaHistorial.Cells(NuevaFila, ColumnaFechaHistorial) = Fecha
    HojaHistorial.Cells(NuevaFila, ColumnaHoraHistorial) = Hora
    HojaHistorial.Cells(NuevaFila, ColumnaTipoDeTransaccionHistorial) = Correlativo
    HojaHistorial.Cells(NuevaFila, ColumnaID1Historial) = ID1Correlativo
    HojaHistorial.Cells(NuevaFila, ColumnaID2Historial) = ID2Correlativo
        
        If IDCaja <> Empty And IDCaja <> "USD" And IDCaja <> "BRL" And IDCaja <> "VES" Then

                HojaHistorial.Cells(NuevaFila, ColumnaIDCajaHistorial) = IDCaja

                FilaCaja = ObtenerFila(HojaCajas, IDCaja, ColumnaIDCaja)

                HojaHistorial.Cells(NuevaFila, ColumnaSaldoAnteriorCajaHistorial) = HojaCajas.Cells(FilaCaja, ColumnaSaldoCaja).Value

                    If AbonarACaja = True Then
                            HojaHistorial.Cells(NuevaFila, ColumnaMontoCajaHistorial) = Cantidad * Monto
                            HojaHistorial.Cells(NuevaFila, ColumnaSaldoNuevoCajaHistorial) = HojaCajas.Cells(FilaCaja, ColumnaSaldoCaja) + (Cantidad * Monto)
                                'Ingreso del dinero en la caja correspondiente
                                AbonarSaldoACaja IDCaja, Cantidad * Monto
                    ElseIf AbonarACaja = False Then
                            HojaHistorial.Cells(NuevaFila, ColumnaMontoCajaHistorial) = 0 - (Cantidad * Monto)
                            HojaHistorial.Cells(NuevaFila, ColumnaSaldoNuevoCajaHistorial) = HojaCajas.Cells(FilaCaja, ColumnaSaldoCaja) - (Cantidad * Monto)
                                'Descuento del dinero en la caja correspondiente
                                DescontarSaldoDeCaja IDCaja, Cantidad * Monto
                    Else
                            HojaHistorial.Cells(NuevaFila, ColumnaMontoCajaHistorial) = 0
                            HojaHistorial.Cells(NuevaFila, ColumnaSaldoNuevoCajaHistorial) = HojaCajas.Cells(FilaCaja, ColumnaSaldoCaja)

                    End If
        End If
        
    HojaHistorial.Cells(NuevaFila, ColumnaDescripcionHistorial) = Descripcion
    HojaHistorial.Cells(NuevaFila, ColumnaIDClienteHistorial) = IDCliente
    HojaHistorial.Cells(NuevaFila, ColumnaIDResponsableHistorial) = IDResponsable
    HojaHistorial.Cells(NuevaFila, ColumnaCodigoHistorial) = Codigo
    HojaHistorial.Cells(NuevaFila, ColumnaProductoHistorial) = Producto
    HojaHistorial.Cells(NuevaFila, ColumnaCantidadHistorial) = Cantidad
    HojaHistorial.Cells(NuevaFila, ColumnaNuevaExistenciaHistorial) = NuevaExistencia
    
    'Registrar ultima transaccion en dashboard
    FilaCliente = ObtenerFila(HojaClientes, IDCliente, ColumnaIDCliente)
    
    HojaCorrelativos.Cells(2, 6) = Correlativo
    HojaCorrelativos.Cells(2, 7) = Fecha
    If FilaCliente > 0 Then
        HojaCorrelativos.Cells(2, 8) = HojaClientes.Cells(FilaCliente, ColumnaNombreCliente)
    Else
        HojaCorrelativos.Cells(2, 8) = Empty
    End If
    
    
    'Formato especifico para cada columna
    HojaHistorial.Cells(NuevaFila, ColumnaFechaHistorial).NumberFormat = "d/m/yyyy" 'Fecha
    HojaHistorial.Cells(NuevaFila, ColumnaHoraHistorial).NumberFormat = "[$-x-systime]h:mm:ss AM/PM" 'Hora
    HojaHistorial.Cells(NuevaFila, ColumnaTipoDeTransaccionHistorial).NumberFormat = "General" 'General
    HojaHistorial.Cells(NuevaFila, ColumnaID1Historial).NumberFormat = "@" 'Texto
    HojaHistorial.Cells(NuevaFila, ColumnaID2Historial).NumberFormat = "@" 'Texto
    HojaHistorial.Cells(NuevaFila, ColumnaIDCajaHistorial).NumberFormat = "General" 'General
    HojaHistorial.Cells(NuevaFila, ColumnaDescripcionHistorial).NumberFormat = "General" 'General
    HojaHistorial.Cells(NuevaFila, ColumnaIDClienteHistorial).NumberFormat = "General" 'General
    HojaHistorial.Cells(NuevaFila, ColumnaIDResponsableHistorial).NumberFormat = "General" 'General
    HojaHistorial.Cells(NuevaFila, ColumnaCodigoHistorial).NumberFormat = "0"
    HojaHistorial.Cells(NuevaFila, ColumnaProductoHistorial).NumberFormat = "General" 'General
    HojaHistorial.Cells(NuevaFila, ColumnaCantidadHistorial).NumberFormat = "General" 'General
    
    Select Case (Left(IDCaja, 3))
        
        Case "USD":
                If Not Correlativo = "MTO-CJA" Then
                        HojaHistorial.Cells(NuevaFila, ColumnaPrecioHistorial) = Monto
                        HojaHistorial.Cells(NuevaFila, ColumnaPrecioHistorial).NumberFormat = "_-[$$] * #,##0.0000_-;[$$] * -#,##0.0000_-;_-[$$] * ""-""??_-;_-@_-" 'Formato $
                        Importe = Cantidad * Monto
                        HojaHistorial.Cells(NuevaFila, ColumnaImporteHistorial) = Importe
                        HojaHistorial.Cells(NuevaFila, ColumnaImporteHistorial).NumberFormat = "_-[$$] * #,##0.0000_-;[$$] * -#,##0.0000_-;_-[$$] * ""-""??_-;_-@_-" 'Formato $
                End If
                        HojaHistorial.Cells(NuevaFila, ColumnaSaldoAnteriorCajaHistorial).NumberFormat = "_-[$$] * #,##0.0000_-;[$$] * -#,##0.0000_-;_-[$$] * ""-""??_-;_-@_-" 'Formato $
                        HojaHistorial.Cells(NuevaFila, ColumnaMontoCajaHistorial).NumberFormat = "_-[$$] * #,##0.0000_-;[$$] * -#,##0.0000_-;_-[$$] * ""-""??_-;_-@_-" 'Formato $
                        HojaHistorial.Cells(NuevaFila, ColumnaSaldoNuevoCajaHistorial).NumberFormat = "_-[$$] * #,##0.0000_-;[$$] * -#,##0.0000_-;_-[$$] * ""-""??_-;_-@_-" 'Formato $
'
        Case "BRL":
                If Not Correlativo = "MTO-CJA" Then
                        HojaHistorial.Cells(NuevaFila, ColumnaCostoHistorial) = Monto
                        HojaHistorial.Cells(NuevaFila, ColumnaCostoHistorial).NumberFormat = "_-[$R$] * #,##0.00_-;[$R$] * -#,##0.00_-;_-[$R$] * ""-""??_-;_-@_-" 'Formato R$
                        Importe = Cantidad * Monto
                        HojaHistorial.Cells(NuevaFila, ColumnaImporteHistorial) = Importe
                        HojaHistorial.Cells(NuevaFila, ColumnaImporteHistorial).NumberFormat = "_-[$R$] * #,##0.00_-;[$R$] * -#,##0.00_-;_-[$R$] * ""-""??_-;_-@_-" 'Formato R$
                End If
                        HojaHistorial.Cells(NuevaFila, ColumnaSaldoAnteriorCajaHistorial).NumberFormat = "_-[$R$] * #,##0.00_-;[$R$] * -#,##0.00_-;_-[$R$] * ""-""??_-;_-@_-" 'Formato R$
                        HojaHistorial.Cells(NuevaFila, ColumnaMontoCajaHistorial).NumberFormat = "_-[$R$] * #,##0.00_-;[$R$] * -#,##0.00_-;_-[$R$] * ""-""??_-;_-@_-" 'Formato R$
                        HojaHistorial.Cells(NuevaFila, ColumnaSaldoNuevoCajaHistorial).NumberFormat = "_-[$R$] * #,##0.00_-;[$R$] * -#,##0.00_-;_-[$R$] * ""-""??_-;_-@_-" 'Formato R$
        
        Case "VES":
                        HojaHistorial.Cells(NuevaFila, ColumnaSaldoAnteriorCajaHistorial).NumberFormat = "_-[$Bs] * #,##0.00_-;[$Bs] * -#,##0.00_-;_-[$Bs] * ""-""??_-;_-@_-" 'Formato Bs
                        HojaHistorial.Cells(NuevaFila, ColumnaMontoCajaHistorial).NumberFormat = "_-[$Bs] * #,##0.00_-;[$Bs] * -#,##0.00_-;_-[$Bs] * ""-""??_-;_-@_-" 'Formato Bs
                        HojaHistorial.Cells(NuevaFila, ColumnaSaldoNuevoCajaHistorial).NumberFormat = "_-[$Bs] * #,##0.00_-;[$Bs] * -#,##0.00_-;_-[$Bs] * ""-""??_-;_-@_-" 'Formato Bs
    
    End Select
    
    ReordenarHistorial
    
End Sub

Sub IngresarSaldoACredito(IDCliente As String, MontoAIngresar As Single)

Dim FilaSaldoCredito As Integer
    
    Inicializar
    
    FilaSaldoCredito = ObtenerFila(HojaClientes, IDCliente, ColumnaIDCliente)
    
    HojaClientes.Cells(FilaSaldoCredito, ColumnaSaldoCreditoCliente) = HojaClientes.Cells(FilaSaldoCredito, ColumnaSaldoCreditoCliente) + MontoAIngresar
    
End Sub

Sub AbonarSaldoACaja(IDCaja As String, MontoAbonado As Single)

Dim FilaCaja As Integer
    
    Inicializar
    
    FilaCaja = ObtenerFila(HojaCajas, IDCaja, ColumnaIDCaja)
    HojaCajas.Cells(FilaCaja, ColumnaSaldoCaja) = HojaCajas.Cells(FilaCaja, ColumnaSaldoCaja) + MontoAbonado
    
End Sub

Sub DescontarSaldoDeCaja(IDCaja As String, MontoADescontar As Single)

Dim FilaCaja As Integer
    
    Inicializar
    
    FilaCaja = ObtenerFila(HojaCajas, IDCaja, ColumnaIDCaja)
    HojaCajas.Cells(FilaCaja, ColumnaSaldoCaja) = HojaCajas.Cells(FilaCaja, ColumnaSaldoCaja) - MontoADescontar
    
End Sub

Sub FiltrarHistorialPorFecha()

Dim Rango As Range
    
    Inicializar

    Set Rango = HojaHistorial.Range("A1").CurrentRegion
    
    Rango.AdvancedFilter Action:=xlFilterCopy, CriteriaRange:= _
    Workbooks("Historial.xlsm").Sheets("Inicio").ListObjects("Filtro").Range, CopyToRange:=HojaHistorialTemporal.Range("A1"), Unique:=True

End Sub

Sub FiltrarHistorialPorTransaccion()

Dim Rango As Range
    
    Inicializar

    Set Rango = HojaHistorial.Range("A1").CurrentRegion
    
    Rango.AdvancedFilter Action:=xlFilterCopy, CriteriaRange:= _
    Workbooks("Historial.xlsm").Sheets("Inicio").ListObjects("FiltroTransaccion").Range, CopyToRange:=HojaHistorialTemporal.Range("A1"), Unique:=True

End Sub


Sub ActualizarSubTotal()
    
Dim SubTotal As Single
Dim Descuento As Single
Dim Total As Single
Dim i As Integer

    'Se suman los precios subtotales de los distintos productos en uno solo
        SubTotal = 0
        Descuento = Val(FormularioAnterior.TextBox_Descuento) / 100
        Total = 0
        For i = 0 To FormularioAnterior.ListBox_Listado.ListCount - 1
            SubTotal = SubTotal + Val(FormularioAnterior.ListBox_Listado.List(i, 5))
        Next i
        
        Total = SubTotal - (Descuento * SubTotal)
        
        FormularioAnterior.TextBox_SubTotal = Format(SubTotal, "0.00")
        FormularioAnterior.TextBox_Total = Format(Total, "0.00")
        
        
End Sub


Sub AbrirDependencias()

    On Error Resume Next
    
    Inicializar
    
    Workbooks.Open (ThisWorkbook.Path & RutaBooks & "Base de datos.xlsm")
    Workbooks.Open (ThisWorkbook.Path & RutaBooks & "Clientes.xlsm")
    Workbooks.Open (ThisWorkbook.Path & RutaBooks & "Historial.xlsm")
    Workbooks.Open (ThisWorkbook.Path & RutaBooks & "Mobile.xlsx")

End Sub

Sub CerrarDependencias(Optional Guardar As Boolean)
      
    On Error Resume Next
    
    Workbooks("Base de datos.xlsm").Close (Guardar)
    Workbooks("Clientes.xlsm").Close (Guardar)
    Workbooks("Historial.xlsm").Close (Guardar)
    Workbooks("Mobile.xlsx").Close (Guardar)

End Sub


Sub MostrarDependencias()
    
    On Error Resume Next

    Application.Windows("Base de datos.xlsm").Visible = True
    Application.Windows("Clientes.xlsm").Visible = True
    Application.Windows("Historial.xlsm").Visible = True
    Application.Windows("Mobile.xlsx").Visible = True
    
End Sub

Sub OcultarDependencias()
    
    On Error Resume Next
    
    Application.Windows("Base de datos.xlsm").Visible = False
    Application.Windows("Clientes.xlsm").Visible = False
    Application.Windows("Historial.xlsm").Visible = False
    Application.Windows("Mobile.xlsx").Visible = False
    
End Sub

Sub GuardarDependencias()
      
    On Error Resume Next
    
    Workbooks("Base de datos.xlsm").Save
    Workbooks("Clientes.xlsm").Save
    Workbooks("Historial.xlsm").Save

End Sub

Sub OcultarHojas()

Dim Libro As Workbook
Dim Hoja As Worksheet
    
    On Error Resume Next
    
    For Each Libro In Workbooks
        For Each Hoja In Libro.Worksheets
                
                If Not Libro.Name = "Mobile.xlsx" Then
                
                        If Hoja.Name = "Inicio" Or Hoja.Name = "Dashboard" Then
                            Hoja.Visible = xlSheetVisible
                        Else
                            Hoja.Visible = xlSheetHidden
                        End If
                        
                End If
            
        Next Hoja
    Next Libro
    
End Sub

Sub MostrarHojas()

Dim Libro As Workbook
Dim Hoja As Worksheet
    
'    On Error Resume Next
    
    For Each Libro In Workbooks
        For Each Hoja In Libro.Worksheets
        
                If Not Libro.Name = "Mobile.xlsx" Then
                        
                        If Not Hoja.Name = "Inicio" Then Hoja.Visible = xlSheetVisible
                        
                End If
                
        Next Hoja
    Next Libro
    
End Sub

Sub GuardarMobile()
        

    Application.DisplayAlerts = False
        Workbooks("Mobile.xlsx").Worksheets("Dashboard").Delete
        Workbooks("Mobile.xlsx").Worksheets("Inventario").Delete
    Application.DisplayAlerts = True
    
    Application.Windows("Mobile.xlsx").Visible = True
        Workbooks("Gestion.xlsm").Sheets("Dashboard").Copy After:=Workbooks("Mobile.xlsx").Sheets(1)
        Workbooks("Base de datos.xlsm").Sheets("Inventario").Copy After:=Workbooks("Mobile.xlsx").Sheets(2)
        Workbooks("Mobile.xlsx").Worksheets("Dashboard").Select
    Application.Windows("Mobile.xlsx").Visible = False
    
    
    
End Sub


Sub GitSave()
    
    DeleteAndMake
    ExportModules
    
End Sub

Sub DeleteAndMake()
        
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim parentFolder As String: parentFolder = ThisWorkbook.Path & "\VBA"
    Dim childB As String: childB = parentFolder & "\VBA-Code_By_Modules"
        
    On Error Resume Next
    fso.DeleteFolder parentFolder
    On Error GoTo 0
    
    MkDir parentFolder
    MkDir childB
    
End Sub


Sub ExportModules()
       
    Dim pathToExport As String: pathToExport = ThisWorkbook.Path & "\VBA\VBA-Code_By_Modules\"
    
    If Dir(pathToExport) <> "" Then
        Kill pathToExport & "*.*"
    End If
     
    Dim wkb As Workbook: Set wkb = Excel.Workbooks(ThisWorkbook.Name)
    
    Dim unitsCount As Long
    Dim filePath As String
    Dim component As VBIDE.VBComponent
    Dim tryExport As Boolean

    For Each component In wkb.VBProject.VBComponents
        tryExport = True
        filePath = component.Name
       
        Select Case component.Type
            Case vbext_ct_ClassModule
                filePath = filePath & ".cls"
            Case vbext_ct_MSForm
                filePath = filePath & ".frm"
            Case vbext_ct_StdModule
                filePath = filePath & ".bas"
            Case vbext_ct_Document
                tryExport = False
        End Select
        
        If tryExport Then
            Debug.Print unitsCount & " exporting " & filePath
            component.Export pathToExport & "\" & filePath
        End If
    Next

    Debug.Print "Exported at " & pathToExport
    
End Sub

