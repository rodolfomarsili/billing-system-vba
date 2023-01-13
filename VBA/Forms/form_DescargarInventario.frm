VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form_DescargarInventario 
   Caption         =   "Descargar Inventario"
   ClientHeight    =   7080
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   11856
   OleObjectBlob   =   "form_DescargarInventario.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "form_DescargarInventario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton_Aceptar_Click()

Dim Cod As Variant
Dim NuevaExistencia As Long
Dim i As Integer
Dim a As Integer
Dim Cantidad As Long
Dim Codigo As String
Dim Producto As String
Dim Precio As Single
Dim IDResponsable As String
Dim Fecha As Date
Dim ProcesarFactura As Byte
Dim FilaDeCorrelativo As Byte

  
    Fecha = TextBox_Dia & "/" & TextBox_Mes & "/" & TextBox_Ano

    Inicializar

    a = ListBox_Listado.ListCount
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Verificaciones'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    IDResponsable = HojaGestion.Range("B3")
    
    'Verificacion de que existan productos ingresados en la factura
    If a = 0 Then
        MsgBox "No hay productos agregados a la factura", , "Descargo de inventario"
        Exit Sub
    End If
    
    'Verificacion de comentario obligatorio
    If (TextBox_Comentario.Text = Empty) Then
        MsgBox "Agrega un comentario a esta transaccion para tener una referencia futura", , "Descargo de inventario"
        Exit Sub
    End If
    
        'Ultima verificacion antes de procesar la factura
    ProcesarFactura = MsgBox("¿Seguro que deseas procesar esta transaccion?", vbYesNo + vbExclamation, "Descargo de inventario")
    If ProcesarFactura = vbNo Then Exit Sub
    
    

    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    
   
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Procesar Descuento'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

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
        
        IncluirEnHistorial Frame_Correlativo, Label_CorrelativoPrefijo, Fecha, Codigo, Producto, "USD", Cantidad, TextBox_Comentario.Text, , IDResponsable, Precio, , NuevaExistencia
    
    Next i
    

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Limpieza de Formulario'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
    'Actualizar correlativo
    ActualizarCorrelativo Frame_Correlativo, Label_CorrelativoPrefijo
    
    'Actualizar correlativo en pantalla
    ActualizarCorrelativoEnPantalla Frame_Correlativo, Label_CorrelativoPrefijo
    
    'Limpieza del formulario de factura
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
    
    MsgBox "Transaccion realizada exitosamente", , "Descargo de inventario"
        
End Sub

Private Sub CommandButton_Cancelar_Click()
    Unload Me
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
    sec_IngresarProductoEnFactura.Show
End Sub


Private Sub UserForm_Initialize()
        
    Inicializar
    
    TextBox_Dia = Day(Date)
    TextBox_Mes = Month(Date)
    TextBox_Ano = Year(Date)
    
    Label_CorrelativoPrefijo.Caption = "Descargo"
    
    CommandButton_IngresarItem.Picture = LoadPicture(ThisWorkbook.Path & "\Resources\images\mas.jpg")
    CommandButton_EliminarItem.Picture = LoadPicture(ThisWorkbook.Path & "\Resources\images\menos.jpg")
    
    ' Formato del listado
    ListBox_Listado.ColumnCount = 6
    ListBox_Listado.ColumnWidths = "60 pt; 100 pt; 298 pt; 50 pt; 60 pt; 70 pt"

    Set FormularioAnterior = Me
    
    'Actualizar correlativo en pantalla
    ActualizarCorrelativoEnPantalla Frame_Correlativo, Label_CorrelativoPrefijo

End Sub

Private Sub UserForm_Terminate()
    
    Set FormularioAnterior = Nothing
     
End Sub

