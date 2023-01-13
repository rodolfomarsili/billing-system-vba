VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form_ModificarProducto 
   Caption         =   "Modificar Producto"
   ClientHeight    =   7830
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   5652
   OleObjectBlob   =   "form_ModificarProducto.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "form_ModificarProducto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Private Sub CommandButton_Cancelar_Click()
Unload Me
End Sub

Private Sub CommandButton_Modificar_Click()

Dim ModificarRegistro As Integer
Dim FilaAModificarCliente As Integer
Dim FilaAModificar As Integer
Dim ComentarioOculto As String
Dim IDResponsable As String
Dim HojaCliente As Worksheet
Dim Fecha As Date
  
    Fecha = TextBox_Dia & "/" & TextBox_Mes & "/" & TextBox_Ano
 
 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Verificaciones'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    'Verificar que los campos del formulario esten todos llenos
    If (TextBox_Producto = Empty Or TextBox_Codigo = Empty Or TextBox_PresentacionPorUnidad = Empty Or TextBox_UnidadesPorBulto = Empty Or TextBox_CostoPorBulto = Empty Or TextBox_PrecioPorBulto = Empty) Then
        MsgBox "Debes rellenar todos los campos para continuar", , "Modificar Registro"
        Exit Sub
    End If


    Inicializar
    
    IDResponsable = HojaGestion.Range("B3")
    
    'Obtener numero de fila del producto que se va a modificar
    FilaAModificar = ObtenerFila(HojaInventario, TextBox_Codigo.Text, ColumnaCodigo)
    FilaAModificarCliente = ObtenerFila(HojaBaseClientes, TextBox_Codigo.Text, ColumnaCodigoCliente)
    
    If FilaAModificar = 0 Then
        MsgBox "Codigo de producto no encontrado", , "Modificar Registro"
        Exit Sub
    End If
    
    'Confirmacion de modificacion de registro
    ModificarRegistro = MsgBox("Seguro que deseas modificar este registro?", vbYesNo + vbExclamation, "Modificar Registro")
    If Not ModificarRegistro = 6 Then
       Exit Sub
    End If
            

    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
  
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Modificacion del Producto'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    ComentarioOculto = ""
    ComentarioOculto = "[Codigo de Producto: " & TextBox_Codigo & "] " + Chr(13)
    
    'Todos estos IF verifican si se ha hecho algun cambio en alguno de los campos del formulario, y de haberlos hecho
    'añaden automaticamente un comentario al historial señalando el cambio respectivo
    If (TextBox_Producto <> Label_RespaldoProducto) Then
        ComentarioOculto = ComentarioOculto & "[Modificacion de nombre " & Label_RespaldoProducto & " -> " & TextBox_Producto & "] " + Chr(13)
    End If
    If (TextBox_PresentacionPorUnidad <> Label_RespaldoPresentacionPorUnidad) Then
        ComentarioOculto = ComentarioOculto & "[Modificacion de presentacion " & Label_RespaldoPresentacionPorUnidad & " -> " & TextBox_PresentacionPorUnidad & "] " + Chr(13)
    End If
    If (TextBox_UnidadesPorBulto <> Label_RespaldoUnidadesPorBulto) Then
        ComentarioOculto = ComentarioOculto & "[Modificacion de unidades por bulto " & Label_RespaldoUnidadesPorBulto & " -> " & TextBox_UnidadesPorBulto & "] " + Chr(13)
    End If
    If (TextBox_CostoPorBulto <> Label_RespaldoCostoPorBulto) Then
        ComentarioOculto = ComentarioOculto & "[Modificacion de costo por bulto " & Label_RespaldoCostoPorBulto & " -> " & TextBox_CostoPorBulto & "] " + Chr(13)
    End If
    If (TextBox_PrecioPorBulto <> Label_RespaldoPrecioPorBulto) Then
        ComentarioOculto = ComentarioOculto & "[Modificacion de precio por bulto " & Label_RespaldoPrecioPorBulto & " -> " & TextBox_PrecioPorBulto & "] " + Chr(13)
    End If
       
    ' Llenado de la tabla de inventario con los datos ingresados
    HojaInventario.Cells(FilaAModificar, ColumnaProducto) = TextBox_Producto
    HojaInventario.Cells(FilaAModificar, ColumnaPresentacion) = TextBox_PresentacionPorUnidad
    HojaInventario.Cells(FilaAModificar, ColumnaUnidadesPorBulto) = TextBox_UnidadesPorBulto.Value
    HojaInventario.Cells(FilaAModificar, ColumnaCostoBulto) = TextBox_CostoPorBulto.Value
    HojaInventario.Cells(FilaAModificar, ColumnaPrecioBulto) = TextBox_PrecioPorBulto.Value
    
        
'''''''''''''''''''''''''''''''''''''''Modificar Producto en Tabla de Inventario de cada uno de los clientes''''''''''''''''''''''''''''''''''''''''''''
    If CheckBox_ModificarConsignaciones.Value = True Then
    
        For Each HojaCliente In LibroClientes.Sheets
            If Not HojaCliente.Name = "Inicio" Then
                ' Llenado de la tabla de inventario con los datos ingresados
                HojaCliente.Cells(FilaAModificarCliente, ColumnaProductoCliente) = TextBox_Producto
                HojaCliente.Cells(FilaAModificarCliente, ColumnaUnidadesPorBultoCliente) = TextBox_UnidadesPorBulto.Value
                HojaCliente.Cells(FilaAModificarCliente, ColumnaPrecioBultoCliente) = TextBox_PrecioPorBulto.Value
            End If
        Next HojaCliente
        
     End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Limpieza del Formulario'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    ' Reordenar alfabeticamente el inventario
    ReordenarInventario
    
    'Incluir en el historial
    IncluirEnHistorial Frame_Correlativo, Label_CorrelativoPrefijo, Fecha, , , , , ComentarioOculto, , IDResponsable

    'Actualizar correlativo
    ActualizarCorrelativo Frame_Correlativo, Label_CorrelativoPrefijo
    
    'Actualizar correlativo en pantalla
    ActualizarCorrelativoEnPantalla Frame_Correlativo, Label_CorrelativoPrefijo
    
    ' Limpieza de los campos de texto
    TextBox_Producto = Empty
    TextBox_Codigo = Empty
    TextBox_PresentacionPorUnidad = Empty
    TextBox_UnidadesPorBulto = Empty
    TextBox_CostoPorBulto = Empty
    TextBox_PrecioPorBulto = Empty
    TextBox_Comentario = Empty
    
    ' Establecer foco en el campo de codigo
    TextBox_Codigo.SetFocus
    
      
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    ActualizarDashboard
    
    MsgBox "Producto modificado exitosamente", , "Modificar Producto"
    

End Sub

Private Sub CommandButton_SeleccionarProducto_Click()
    sec_ListadoDeProductos.Show
End Sub

Private Sub TextBox_Codigo_AfterUpdate()
    
    On Error Resume Next
    
Dim FilaDeCodigo As String

    Inicializar
    
    FilaDeCodigo = ObtenerFila(HojaInventario, TextBox_Codigo.Text, ColumnaCodigo)
    
    TextBox_Producto = HojaInventario.Cells(FilaDeCodigo, ColumnaProducto)
    TextBox_UnidadesPorBulto = HojaInventario.Cells(FilaDeCodigo, ColumnaUnidadesPorBulto)
    TextBox_PresentacionPorUnidad = HojaInventario.Cells(FilaDeCodigo, ColumnaPresentacion)
    TextBox_CostoPorBulto = HojaInventario.Cells(FilaDeCodigo, ColumnaCostoBulto)
    TextBox_PrecioPorBulto = HojaInventario.Cells(FilaDeCodigo, ColumnaPrecioBulto)

End Sub

Private Sub TextBox_Codigo_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If Not (KeyAscii >= 48 And KeyAscii <= 57) Then
        KeyAscii = 0
    End If
End Sub

Private Sub TextBox_CostoPorBulto_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Dim UbicacionPunto As Integer

    UbicacionPunto = InStr(TextBox_CostoPorBulto.Text, ".")
    
    If (KeyAscii = 46 And UbicacionPunto > 0) Then
        KeyAscii = 0
    End If
    
    If Not (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 46) Then
        KeyAscii = 0
    End If
End Sub


Private Sub TextBox_PrecioPorBulto_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

Dim UbicacionPunto As Integer
    
        UbicacionPunto = InStr(TextBox_PrecioPorBulto.Text, ".")
    
    If (KeyAscii = 46 And UbicacionPunto > 0) Then
        KeyAscii = 0
    End If
    
    If Not (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 46) Then
        KeyAscii = 0
    End If
End Sub

Private Sub TextBox_UnidadesPorBulto_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If Not (KeyAscii >= 48 And KeyAscii <= 57) Then
        KeyAscii = 0
    End If
End Sub

Private Sub UserForm_Initialize()
    
    TextBox_Dia = Day(Date)
    TextBox_Mes = Month(Date)
    TextBox_Ano = Year(Date)
        
    Label_CorrelativoPrefijo.Caption = "Modificacion"
    
    Set FormularioAnterior = Me
        
    'Actualizar correlativo en pantalla
    ActualizarCorrelativoEnPantalla Frame_Correlativo, Label_CorrelativoPrefijo
    
End Sub

Private Sub UserForm_Terminate()

    Set FormularioAnterior = Nothing
    
End Sub
