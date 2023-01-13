VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form_Menu 
   Caption         =   "Menu"
   ClientHeight    =   11220
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   11940
   OleObjectBlob   =   "form_Menu.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "form_Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton_CerrarSesion_Click()
    BloquearAcceso
    GuardarDependencias
    Unload Me
    form_Login.Show
End Sub

Private Sub CommandButton_Compras_Click()
    form_Compras.Show
End Sub

Private Sub CommandButton_Consignaciones_Click()
    form_InventarioConsignaciones.Show
End Sub

Private Sub CommandButton_DescargarInventario_Click()
    form_DescargarInventario.Show
End Sub

Private Sub CommandButton_Devolucion_Click()
    form_Devolucion.Show
End Sub

Private Sub CommandButton_Extras_Click()
    form_Extras.Show
End Sub

Private Sub CommandButton_Inventario_Click()
    form_Inventario.Show
End Sub

Private Sub CommandButton_Pagos_Click()
    form_Pagos.Show
End Sub

Private Sub CommandButton_Prestamos_Click()
    form_Prestamos.Show
End Sub

Private Sub CommandButton_RecargarInventario_Click()
    form_RecargarInventario.Show
End Sub

Private Sub CommandButton_Facturar_Click()
    form_Facturar.Show
End Sub


Private Sub CommandButton_Historial_Click()
    form_Historial.Show
End Sub

Private Sub CommandButton_ModificarCliente_Click()
    form_ModificarCliente.Show
End Sub

Private Sub CommandButton_ModificarProducto_Click()
    form_ModificarProducto.Show
End Sub

Private Sub CommandButton_TransferenciaEntreCajas_Click()
    form_TransferenciaEntreCajas.Show
End Sub

Private Sub CommandButton_MovimientoDeMercancias_Click()
    form_MovimientoDeMercancias.Show
End Sub

Private Sub CommandButton_PagoConsignacion_Click()
    form_Pagos.Show
End Sub

Private Sub CommandButton_PagoCredito_Click()
    form_PagoCredito.Show
End Sub

Private Sub CommandButton_RegistrarCliente_Click()
    form_RegistrarCliente.Show
End Sub

Private Sub CommandButton_RegistrarProducto_Click()
    form_RegistrarProducto.Show
End Sub

Private Sub CommandButton_VisibilidadDependencias_Click()
    form_VisibilidadDependencias.Show
End Sub


Private Sub CommandButton_VisibilidadHojasDeGestion_Click()
    form_VisibilidadHojasDeGestion.Show
End Sub


Private Sub UserForm_Initialize()

    Inicializar
    
    Me.Caption = "Menu - Sesion Activa: " & HojaGestion.Range("B2")
    
    If Not HojaGestion.Range("B3") = "V-25017131" Then Frame_Programacion.Visible = False
    
    CommandButton_RegistrarProducto.Picture = LoadPicture(ThisWorkbook.Path & RutaImages & "añadir_producto.jpg")
    CommandButton_RegistrarCliente.Picture = LoadPicture(ThisWorkbook.Path & RutaImages & "añadir_cliente.jpg")
    CommandButton_ModificarProducto.Picture = LoadPicture(ThisWorkbook.Path & RutaImages & "modificar_producto.jpg")
    CommandButton_ModificarCliente.Picture = LoadPicture(ThisWorkbook.Path & RutaImages & "modificar_cliente.jpg")
    CommandButton_Historial.Picture = LoadPicture(ThisWorkbook.Path & RutaImages & "historial.jpg")
    CommandButton_Compras.Picture = LoadPicture(ThisWorkbook.Path & RutaImages & "cesta.jpg")
    CommandButton_PagoCredito.Picture = LoadPicture(ThisWorkbook.Path & RutaImages & "pagocredito.jpg")
    CommandButton_PagoConsignacion.Picture = LoadPicture(ThisWorkbook.Path & RutaImages & "pagoconsignacion.jpg")
    CommandButton_Facturar.Picture = LoadPicture(ThisWorkbook.Path & RutaImages & "facturar.jpg")
    CommandButton_VisibilidadDependencias.Picture = LoadPicture(ThisWorkbook.Path & RutaImages & "visibilidaddependencias.jpg")
    CommandButton_TransferenciaEntreCajas.Picture = LoadPicture(ThisWorkbook.Path & RutaImages & "movimientodecajas.jpg")
    CommandButton_CerrarSesion.Picture = LoadPicture(ThisWorkbook.Path & RutaImages & "cerrarsesion.jpg")
    CommandButton_RecargarInventario.Picture = LoadPicture(ThisWorkbook.Path & RutaImages & "recargarinventario.jpg")
    CommandButton_DescargarInventario.Picture = LoadPicture(ThisWorkbook.Path & RutaImages & "descargarinventario.jpg")
    
    CommandButton_Pagos.Picture = LoadPicture(ThisWorkbook.Path & RutaImages & "pagos.jpg")
    CommandButton_Prestamos.Picture = LoadPicture(ThisWorkbook.Path & RutaImages & "prestamo.jpg")
    CommandButton_Extras.Picture = LoadPicture(ThisWorkbook.Path & RutaImages & "extras.jpg")
    
    
    CommandButton_RegistrarProducto.Caption = Empty
    CommandButton_RegistrarCliente.Caption = Empty
    CommandButton_ModificarProducto.Caption = Empty
    CommandButton_ModificarCliente.Caption = Empty
    CommandButton_Historial.Caption = Empty
    CommandButton_Compras.Caption = Empty
    CommandButton_PagoCredito.Caption = Empty
    CommandButton_PagoConsignacion.Caption = Empty
    CommandButton_Facturar.Caption = Empty
    CommandButton_VisibilidadDependencias.Caption = Empty
    CommandButton_TransferenciaEntreCajas.Caption = Empty
    CommandButton_CerrarSesion.Caption = Empty
    CommandButton_RecargarInventario.Caption = Empty
    CommandButton_DescargarInventario.Caption = Empty
    
    CommandButton_Pagos.Caption = Empty
    CommandButton_Prestamos.Caption = Empty
    CommandButton_Extras.Caption = Empty
    
End Sub
