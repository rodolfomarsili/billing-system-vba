Attribute VB_Name = "CintaOpciones"
Option Explicit

'control As IRibbonControl

Sub CerrarSesion(Control As IRibbonControl)
    BloquearAcceso
    form_Login.Show
End Sub

Sub Facturar(Control As IRibbonControl)
    form_Menu.Show
End Sub

Sub Historial(Control As IRibbonControl)
    form_Historial.Show
End Sub

Sub ModificarProducto(Control As IRibbonControl)
    form_ModificarProducto.Show
End Sub

Sub ModificarCliente(Control As IRibbonControl)
    form_ModificarCliente.Show
End Sub

Sub MovimientoDeCajas(Control As IRibbonControl)
    form_MovimientoDeCajas.Show
End Sub

Sub MovimientoDeMercancias(Control As IRibbonControl)
'    form_MovimientoDeMercancias.Show
End Sub

Sub RegistrarCliente(Control As IRibbonControl)
    form_RegistrarCliente.Show
End Sub

Sub RegistrarProducto(Control As IRibbonControl)
    form_RegistrarProducto.Show
End Sub

Sub VisibilidadDependencias(Control As IRibbonControl)
    Inicializar
    If HojaGestion.Range("B3") = "V-25017131" Then form_VisibilidadDependencias.Show
End Sub

