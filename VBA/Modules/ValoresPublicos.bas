Attribute VB_Name = "ValoresPublicos"

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Gestion'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Hojas
Public HojaInicio As Worksheet
Public HojaDashboard As Worksheet
Public HojaUsuarios As Worksheet
Public HojaGestion As Worksheet
    'Columnas de la hoja de usuarios
    Public ColumnaIDUsuario As Byte
    Public ColumnaNombreUsuario As Byte
    Public ColumnaUsuario As Byte
        'Ultima fila de la hoja de usuarios
        Public UltimaFilaUsuarios As Byte
   
   
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Base de Datos'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Hojas
Public HojaCorrelativos As Worksheet
Public HojaInventario As Worksheet
Public HojaClientes As Worksheet
Public HojaCajas As Worksheet
    'Columnas de la hoja de correlativos
    Public ColumnaPrefijo As Byte
    Public ColumnaLeyenda As Byte
    Public ColumnaID1 As Byte
    Public ColumnaID2 As Byte
        'Ultima fila de la hoja de correlativos
        Public UltimaFilaCorrelativos As Byte
    'Columnas de la hoja de inventario
    Public ColumnaCodigo As Byte
    Public ColumnaProducto As Byte
    Public ColumnaExistencia As Byte
    Public ColumnaPresentacion As Byte
    Public ColumnaUnidadesPorBulto As Byte
    Public ColumnaCostoBulto As Byte
    Public ColumnaCostoUnidad As Byte
    Public ColumnaPrecioBulto As Byte
    Public ColumnaPrecioUnidad As Byte
    Public ColumnaImporteCosto As Byte
    Public ColumnaImportePrecio As Byte
        'Ultima columna y fila del inventario
        Public UltimaFilaInventario As Integer
        Public UltimaColumnaInventario As String
    'Columnas de la hoja de clientes
    Public ColumnaIDCliente As Byte
    Public ColumnaNombreCliente As Byte
    Public ColumnaDireccionCliente As Byte
    Public ColumnaTelefonoCliente As Byte
    Public ColumnaCreditoCliente As Byte
    Public ColumnaConsignacionCliente As Byte
    Public ColumnaLimiteCreditoCliente As Byte
    Public ColumnaSaldoCreditoCliente As Byte
    Public ColumnaSaldoConsignacionCliente As Byte
    Public ColumnaPrestamoUSDCliente As Byte
    Public ColumnaPrestamoBRLCliente As Byte
    Public ColumnaPrestamoVESCliente As Byte
    Public ColumnaDeudaTotalCliente As Byte
        'Ultima fila del listado de clientes
        Public UltimaFilaClientes As Integer
    'Columnas de la hoja de cajas
    Public ColumnaIDResponsableCaja As Byte
    Public ColumnaIDCaja As Byte
    Public ColumnaSaldoCaja As Byte
        'Ultima fila de la hoja de cajas
        Public UltimaFilaCajas As Byte


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Clientes'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Hojas
Public LibroClientes As Workbook
Public HojaBaseClientes As Worksheet
    'Columnas de la hoja base del libro de clientes
    Public ColumnaCodigoCliente As Byte
    Public ColumnaProductoCliente As Byte
    Public ColumnaUnidadesPorBultoCliente As Byte
    Public ColumnaPrecioBultoCliente As Byte
    Public ColumnaPrecioUnitarioCliente As Byte
    Public ColumnaExistenciaCliente As Byte
    Public ColumnaImporteCliente As Byte
    Public ColumnaImporteTotalCliente As Byte


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Historial'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Hojas
Public HojaHistorial As Worksheet
Public HojaHistorialTemporal As Worksheet
    'Columnas de la hoja de historial
    Public ColumnaFechaHistorial As Byte
    Public ColumnaHoraHistorial As Byte
    Public ColumnaDevueltoHistorial As Byte
    Public ColumnaDEVAUXHistorial As Byte
    Public ColumnaTipoDeTransaccionHistorial As Byte
    Public ColumnaID1Historial As Byte
    Public ColumnaID2Historial As Byte
    Public ColumnaIDCajaHistorial As Byte
    Public ColumnaSaldoAnteriorCajaHistorial As Byte
    Public ColumnaMontoCajaHistorial As Byte
    Public ColumnaSaldoNuevoCajaHistorial As Byte
    Public ColumnaDescripcionHistorial As Byte
    Public ColumnaIDClienteHistorial As Byte
    Public ColumnaIDResponsableHistorial As Byte
    Public ColumnaCodigoHistorial As Byte
    Public ColumnaProductoHistorial As Byte
    Public ColumnaCantidadHistorial As Byte
    Public ColumnaNuevaExistenciaHistorial As Byte
    Public ColumnaCostoHistorial As Byte
    Public ColumnaPrecioHistorial As Byte
    Public ColumnaImporteHistorial As Byte
        'Ultima fila del historial
        Public UltimaFilaHistorial As Long
        'Ultima fila del historial temporal
        Public UltimaFilaHistorialTemporal As Long


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Otros'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Public FormularioAnterior As UserForm
Public RutaImages As String
Public RutaBooks As String

Public Sub Inicializar()

    RutaImages = "\Resources\images\"
    RutaBooks = "\Resources\books\"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Gestion'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    'HojasSet
    Set HojaInicio = Workbooks("Gestion.xlsm").Sheets("Inicio")
    Set HojaDashboard = Workbooks("Gestion.xlsm").Sheets("Dashboard")
    Set HojaUsuarios = Workbooks("Gestion.xlsm").Sheets("Usuarios")
    Set HojaGestion = Workbooks("Gestion.xlsm").Sheets("Gestion Interna")
        'Columnas de la hoja de usuarios
        ColumnaIDUsuario = ObtenerColumnaDeTabla(HojaUsuarios, 1, "ID")
        ColumnaNombreUsuario = ObtenerColumnaDeTabla(HojaUsuarios, 1, "Nombre")
        ColumnaUsuario = ObtenerColumnaDeTabla(HojaUsuarios, 1, "Usuario")
            'Ultima fila de la hoja de usuarios
            UltimaFilaUsuarios = ObtenerUltimaFila(HojaUsuarios, ColumnaIDUsuario)
            


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Base de Datos'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    'Hojas
    Set HojaCorrelativos = Workbooks("Base de datos.xlsm").Sheets("Correlativos")
    Set HojaInventario = Workbooks("Base de datos.xlsm").Sheets("Inventario")
    Set HojaClientes = Workbooks("Base de datos.xlsm").Sheets("Clientes")
    Set HojaCajas = Workbooks("Base de datos.xlsm").Sheets("Cajas")
        'Columnas de la hoja de correlativos
        ColumnaPrefijo = ObtenerColumnaDeTabla(HojaCorrelativos, 1, "Prefijo")
        ColumnaLeyenda = ObtenerColumnaDeTabla(HojaCorrelativos, 1, "Leyenda")
        ColumnaID1 = ObtenerColumnaDeTabla(HojaCorrelativos, 1, "ID-1")
        ColumnaID2 = ObtenerColumnaDeTabla(HojaCorrelativos, 1, "ID-2")
            'Ultima fila de la hoja de correlativos
            UltimaFilaCorrelativos = ObtenerUltimaFila(HojaCorrelativos, ColumnaPrefijo)
        'Columnas de la hoja de inventario
        ColumnaCodigo = ObtenerColumnaDeTabla(HojaInventario, 1, "Codigo")
        ColumnaProducto = ObtenerColumnaDeTabla(HojaInventario, 1, "Producto")
        ColumnaExistencia = ObtenerColumnaDeTabla(HojaInventario, 1, "Existencia")
        ColumnaPresentacion = ObtenerColumnaDeTabla(HojaInventario, 1, "Presentacion por unidad")
        ColumnaUnidadesPorBulto = ObtenerColumnaDeTabla(HojaInventario, 1, "Cantidad de unidades por bulto")
        ColumnaCostoBulto = ObtenerColumnaDeTabla(HojaInventario, 1, "Costo por bulto (R$)")
        ColumnaCostoUnidad = ObtenerColumnaDeTabla(HojaInventario, 1, "Costo")
        ColumnaPrecioBulto = ObtenerColumnaDeTabla(HojaInventario, 1, "Precio por bulto ($)")
        ColumnaPrecioUnidad = ObtenerColumnaDeTabla(HojaInventario, 1, "Precio")
        ColumnaImporteCosto = ObtenerColumnaDeTabla(HojaInventario, 1, "Importe Costo")
        ColumnaImportePrecio = ObtenerColumnaDeTabla(HojaInventario, 1, "Importe Precio")
            'Ultima columna y fila del inventario
                UltimaFilaInventario = ObtenerUltimaFila(HojaInventario, ColumnaCodigo)
                UltimaColumnaInventario = ObtenerLetraDeColumna(ObtenerUltimaColumna(HojaInventario, 1))
        'Columnas de la hoja de clientes
        ColumnaIDCliente = ObtenerColumnaDeTabla(HojaClientes, 1, "ID")
        ColumnaNombreCliente = ObtenerColumnaDeTabla(HojaClientes, 1, "Nombre")
        ColumnaDireccionCliente = ObtenerColumnaDeTabla(HojaClientes, 1, "Direccion")
        ColumnaTelefonoCliente = ObtenerColumnaDeTabla(HojaClientes, 1, "Telefono")
        ColumnaCreditoCliente = ObtenerColumnaDeTabla(HojaClientes, 1, "Credito")
        ColumnaConsignacionCliente = ObtenerColumnaDeTabla(HojaClientes, 1, "Consignacion")
        ColumnaLimiteCreditoCliente = ObtenerColumnaDeTabla(HojaClientes, 1, "Limite Credito")
        ColumnaSaldoCreditoCliente = ObtenerColumnaDeTabla(HojaClientes, 1, "Saldo Credito")
        ColumnaSaldoConsignacionCliente = ObtenerColumnaDeTabla(HojaClientes, 1, "Saldo Consignacion")
        ColumnaPrestamoUSDCliente = ObtenerColumnaDeTabla(HojaClientes, 1, "Prestamo $")
        ColumnaPrestamoBRLCliente = ObtenerColumnaDeTabla(HojaClientes, 1, "Prestamo R$")
        ColumnaPrestamoVESCliente = ObtenerColumnaDeTabla(HojaClientes, 1, "Prestamo Bs")
        ColumnaDeudaTotalCliente = ObtenerColumnaDeTabla(HojaClientes, 1, "Deuda Total")
            'Ultima fila del listado de clientes
            UltimaFilaClientes = ObtenerUltimaFila(HojaClientes, ColumnaIDCliente)
        'Columnas de la hoja de cajas
        ColumnaIDResponsableCaja = ObtenerColumnaDeTabla(HojaCajas, 1, "ID Responsable Caja")
        ColumnaIDCaja = ObtenerColumnaDeTabla(HojaCajas, 1, "ID Caja")
        ColumnaSaldoCaja = ObtenerColumnaDeTabla(HojaCajas, 1, "Saldo")
            'Ultima fila de la hoja de cajas
            UltimaFilaCajas = ObtenerUltimaFila(HojaCajas, ColumnaIDCaja)



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Clientes'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


    'Hojas
    Set LibroClientes = Workbooks("Clientes.xlsm")
    Set HojaBaseClientes = Workbooks("Clientes.xlsm").Sheets("Base")
        'Columnas de la hoja base del libro de clientes
        ColumnaCodigoCliente = ObtenerColumnaDeTabla(HojaBaseClientes, 1, "Codigo")
        ColumnaProductoCliente = ObtenerColumnaDeTabla(HojaBaseClientes, 1, "Producto")
        ColumnaUnidadesPorBultoCliente = ObtenerColumnaDeTabla(HojaBaseClientes, 1, "Cantidad de unidades por bulto")
        ColumnaPrecioBultoCliente = ObtenerColumnaDeTabla(HojaBaseClientes, 1, "Precio por bulto ($)")
        ColumnaPrecioUnitarioCliente = ObtenerColumnaDeTabla(HojaBaseClientes, 1, "Precio")
        ColumnaExistenciaCliente = ObtenerColumnaDeTabla(HojaBaseClientes, 1, "Existencia")
        ColumnaImporteCliente = ObtenerColumnaDeTabla(HojaBaseClientes, 1, "Importe")
        ColumnaImporteTotalCliente = ObtenerColumnaDeTabla(HojaBaseClientes, 1, "Importe Total:") + 1



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Historial'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
     'Hojas
    Set HojaHistorial = Workbooks("Historial.xlsm").Sheets("Hoja1")
    Set HojaHistorialTemporal = Workbooks("Historial.xlsm").Sheets("Historial Temporal")
        'Columnas de la hoja de historial
        ColumnaFechaHistorial = ObtenerColumnaDeTabla(HojaHistorial, 1, "Fecha")
        ColumnaHoraHistorial = ObtenerColumnaDeTabla(HojaHistorial, 1, "Hora")
        ColumnaDevueltoHistorial = ObtenerColumnaDeTabla(HojaHistorial, 1, "Devuelto")
        ColumnaDEVAUXHistorial = ObtenerColumnaDeTabla(HojaHistorial, 1, "DEV-AUX")
        ColumnaTipoDeTransaccionHistorial = ObtenerColumnaDeTabla(HojaHistorial, 1, "Tipo de Transaccion")
        ColumnaID1Historial = ObtenerColumnaDeTabla(HojaHistorial, 1, "ID1 Correlativo")
        ColumnaID2Historial = ObtenerColumnaDeTabla(HojaHistorial, 1, "ID2 Correlativo")
        ColumnaIDCajaHistorial = ObtenerColumnaDeTabla(HojaHistorial, 1, "ID Caja")
        ColumnaSaldoAnteriorCajaHistorial = ObtenerColumnaDeTabla(HojaHistorial, 1, "Anterior Saldo Caja")
        ColumnaMontoCajaHistorial = ObtenerColumnaDeTabla(HojaHistorial, 1, "Monto Abonado/Descontado")
        ColumnaSaldoNuevoCajaHistorial = ObtenerColumnaDeTabla(HojaHistorial, 1, "Nuevo Saldo Caja")
        ColumnaDescripcionHistorial = ObtenerColumnaDeTabla(HojaHistorial, 1, "Descripcion")
        ColumnaIDClienteHistorial = ObtenerColumnaDeTabla(HojaHistorial, 1, "ID Cliente")
        ColumnaIDResponsableHistorial = ObtenerColumnaDeTabla(HojaHistorial, 1, "ID Responsable")
        ColumnaCodigoHistorial = ObtenerColumnaDeTabla(HojaHistorial, 1, "Codigo")
        ColumnaProductoHistorial = ObtenerColumnaDeTabla(HojaHistorial, 1, "Producto")
        ColumnaCantidadHistorial = ObtenerColumnaDeTabla(HojaHistorial, 1, "Cantidad")
        ColumnaNuevaExistenciaHistorial = ObtenerColumnaDeTabla(HojaHistorial, 1, "Nueva Existencia")
        ColumnaCostoHistorial = ObtenerColumnaDeTabla(HojaHistorial, 1, "Costo")
        ColumnaPrecioHistorial = ObtenerColumnaDeTabla(HojaHistorial, 1, "Precio")
        ColumnaImporteHistorial = ObtenerColumnaDeTabla(HojaHistorial, 1, "Importe")
            'Ultima fila de la hoja de historial temporal
            UltimaFilaHistorial = ObtenerUltimaFila(HojaHistorial, ColumnaFechaHistorial)
            'Ultima fila de la hoja de historial temporal
            UltimaFilaHistorialTemporal = ObtenerUltimaFila(HojaHistorialTemporal, ColumnaFechaHistorial)
    

End Sub



