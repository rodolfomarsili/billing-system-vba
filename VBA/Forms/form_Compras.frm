VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form_Compras 
   Caption         =   "Compras"
   ClientHeight    =   9120.001
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   12132
   OleObjectBlob   =   "form_Compras.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "form_Compras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ComboBox_Caja_Change()
    
    Label_AsteriscoCaja.Visible = False
    
    ActualizarSaldoCajaEnPantalla Label_SaldoCaja, ComboBox_Caja.Text
    
End Sub

Private Sub CommandButton_Facturar_Click()

Dim Cod As Variant
Dim NuevaExistencia As Long
Dim j As Integer
Dim i As Integer
Dim a As Integer
Dim ProcesarFactura As Byte
Dim FilaCaja As Byte
Dim Cantidad As Long
Dim Producto As String
Dim Codigo As String
Dim Costo As Single
Dim IDResponsable As String
Dim Fecha As Date
  
    Fecha = TextBox_Dia & "/" & TextBox_Mes & "/" & TextBox_Ano

    Inicializar
    
    'Ocultar todos los asteriscos
    Label_AsteriscoCaja.Visible = False
    Label_AsteriscoFormaDePago.Visible = False
    
    FilaCaja = ObtenerFila(HojaCajas, ComboBox_Caja.Text, ColumnaIDCaja)
    
    a = ListBox_Listado.ListCount
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Verificaciones'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    'Verificacion de que existan productos ingresados en la factura
    If a = 0 Then
        MsgBox "No hay productos agregados a la factura", , "Compras"
        Exit Sub
    End If
    
    'Verificacion de caja seleccionada
    If FilaCaja = 0 Then
        Label_AsteriscoCaja.Visible = True
        MsgBox "Selecciona una Caja valida", , "Compras"
        Exit Sub
    End If
    
    IDResponsable = HojaCajas.Cells(ObtenerFila(HojaCajas, ComboBox_Caja.Text, ColumnaIDCaja), ColumnaIDResponsableCaja)
    
    'Verificacion de forma de pago seleccionada
    If Not (ComboBox_FormaDePago.Text = "Contado" Or ComboBox_FormaDePago.Text = "Credito" Or ComboBox_FormaDePago.Text = "Consignacion") Then
        Label_AsteriscoFormaDePago.Visible = True
        MsgBox "Selecciona una Forma de Pago valida", , "Compras"
        Exit Sub
    End If
    
    If (HojaCajas.Cells(FilaCaja, ColumnaSaldoCaja) - Val(TextBox_Total)) < 0 Then
        Label_AsteriscoCaja.Visible = True
        MsgBox "Fondos insuficientes para realizar esta operacion", , "Compras"
        Exit Sub
    End If
    
    'Ultima verificacion antes de procesar la factura
    ProcesarFactura = MsgBox("¿Seguro que deseas procesar esta factura?", vbYesNo + vbExclamation, "Compras")
    If ProcesarFactura = vbNo Then Exit Sub
    
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Factura de Contado'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    'Procedimiento a ejecutar cuando la compra es de contado.
    If ComboBox_FormaDePago = "Contado" Then
    
    'Añadir existencias en el inventario para cada producto de la factura
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
        Costo = ListBox_Listado.List(i, 4)
        NuevaExistencia = Val(ListBox_Listado.List(i, 0))
        
        IncluirEnHistorial Frame_Correlativo, Label_CorrelativoPrefijo, Fecha, Codigo, Producto, ComboBox_Caja.Text, Cantidad, TextBox_Comentario.Text, , IDResponsable, Costo, False, NuevaExistencia
        
    Next i
    
    End If
    
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Factura de Credito'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    'Procedimiento a ejecutar cuando la compra es a credito.
    If ComboBox_FormaDePago = "Credito" Then
    
        MsgBox "La posibilidad de comprar a credito aun no se ha implementado", , "Compras"
        Exit Sub
    
    End If
    
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Factura de Consignacion'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    
    'Procedimiento a ejecutar cuando la compra es a consignacion.
    If ComboBox_FormaDePago = "Consignacion" Then
        
        MsgBox "La posibilidad de comprar a consignacion aun no se ha implementado", , "Compras"
        Exit Sub
    
    End If
       
   
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Limpieza de Formulario'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
    'Actualizar correlativo
    ActualizarCorrelativo Frame_Correlativo, Label_CorrelativoPrefijo
    
    'Actualizar correlativo en pantalla
    ActualizarCorrelativoEnPantalla Frame_Correlativo, Label_CorrelativoPrefijo
    
    'Actualizar monto de caja en pantalla
    ActualizarSaldoCajaEnPantalla Label_SaldoCaja, ComboBox_Caja.Text
    
    'Limpieza del formulario de factura
    ListBox_Listado.Clear
    TextBox_Comentario.Text = Empty
    TextBox_SubTotal.Text = Empty
    TextBox_Descuento.Text = 0
    TextBox_Total.Text = Empty

    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    ActualizarDashboard
    
    MsgBox "Compra ingresada exitosamente", , "Compras"
    
    
End Sub

Private Sub ComboBox_FormaDePago_Change()
 
    Label_AsteriscoFormaDePago.Visible = False
    
    'Actualizar correlativo en pantalla
    Label_CorrelativoPrefijo.Caption = ComboBox_FormaDePago.Text
    ActualizarCorrelativoEnPantalla Frame_Correlativo, Label_CorrelativoPrefijo
    
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


Private Sub CommandButton_EliminarItem_Click()

    'Si el listado de productos no esta vacio, se elmina el item elegido, de no elegirse ninguno se van eliminando uno a uno
    If (ListBox_Listado.ListIndex >= 0) Then
        ListBox_Listado.RemoveItem (ListBox_Listado.ListIndex)
        ActualizarSubTotal
    End If

End Sub

Private Sub CommandButton_IngresarItem_Click()
    sec_IngresarProductoEnCompra.Show
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
    
    Label_CorrelativoPrefijo.Caption = "Contado"
    
    CommandButton_IngresarItem.Picture = LoadPicture(ThisWorkbook.Path & "\Resources\images\mas.jpg")
    CommandButton_EliminarItem.Picture = LoadPicture(ThisWorkbook.Path & "\Resources\images\menos.jpg")
    
    ' Formato del listado
    ListBox_Listado.ColumnCount = 6
    ListBox_Listado.ColumnWidths = "60 pt; 100 pt; 298 pt; 50 pt; 60 pt; 70 pt"
    
    ' Se establece el valor del descuento en 0%
    TextBox_Descuento.Value = 0
    
    'Llenado del ComboBox de cajas
    For i = 2 To UltimaFilaCajas
         If (Mid(HojaCajas.Cells(i, ColumnaIDCaja), 1, 3) = "BRL") Then ComboBox_Caja.AddItem (HojaCajas.Cells(i, ColumnaIDCaja))
    Next i
    
    ComboBox_Caja.Text = "BRL-RODOLFO"

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

