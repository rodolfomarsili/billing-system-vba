VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form_RegistrarProducto 
   Caption         =   "Registrar Producto"
   ClientHeight    =   6690
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   5652
   OleObjectBlob   =   "form_RegistrarProducto.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "form_RegistrarProducto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton_Cancelar_Click()
    Unload Me
End Sub

Private Sub CommandButton_Registrar_Click()

Dim RegistroRepetido As Boolean
Dim IngresarRegistro As Byte
Dim FilaCodigo As Integer
Dim IDResponsable As String
Dim ComentarioOculto As String
Dim HojaCliente As Worksheet
Const NuevaFila As Byte = 2
Dim Fecha As Date
  
    Fecha = TextBox_Dia & "/" & TextBox_Mes & "/" & TextBox_Ano


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Verificaciones'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    If (TextBox_Producto = Empty Or TextBox_Codigo = Empty Or TextBox_PresentacionPorUnidad = Empty Or TextBox_UnidadesPorBulto = Empty Or TextBox_CostoPorBulto = Empty Or TextBox_PrecioPorBulto = Empty) Then
        MsgBox "Debes rellenar todos los campos para continuar"
        Exit Sub
    End If
    
    Inicializar
    
    IDResponsable = HojaGestion.Range("B3")
    
    ' Verificacion de registro repetido
    FilaCodigo = ObtenerFila(HojaInventario, TextBox_Codigo, ColumnaCodigo)
    
    If FilaCodigo <> 0 Then
        MsgBox "Registro repetido"
        Exit Sub
    End If
    
    IngresarRegistro = MsgBox("¿Seguro que deseas ingresar este registro?", vbYesNo + vbExclamation, "Ingresar Producto")
    If IngresarRegistro = vbNo Then Exit Sub
    
     
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    

'''''''''''''''''''''''''''''''''''''''''''''''''''Insertar Producto en Tabla de Inventario''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        ' Se inserta una fila entera en en principio de la tabla de inventario
        HojaInventario.Range(NuevaFila & ":" & NuevaFila).EntireRow.Insert
        ' Llenado de la tabla de inventario con los datos ingresados
        HojaInventario.Cells(NuevaFila, ColumnaProducto) = TextBox_Producto
        HojaInventario.Cells(NuevaFila, ColumnaCodigo) = TextBox_Codigo.Value
        HojaInventario.Cells(NuevaFila, ColumnaExistencia) = 0
        HojaInventario.Cells(NuevaFila, ColumnaPresentacion) = TextBox_PresentacionPorUnidad
        HojaInventario.Cells(NuevaFila, ColumnaUnidadesPorBulto) = TextBox_UnidadesPorBulto.Value
        HojaInventario.Cells(NuevaFila, ColumnaCostoBulto) = TextBox_CostoPorBulto.Value
        HojaInventario.Cells(NuevaFila, ColumnaPrecioBulto) = TextBox_PrecioPorBulto.Value
        'Se da el formato correspondiente a cada columna
        HojaInventario.Cells(NuevaFila, ColumnaProducto).NumberFormat = "General" 'General
        HojaInventario.Cells(NuevaFila, ColumnaCodigo).NumberFormat = "0"
        HojaInventario.Cells(NuevaFila, ColumnaPresentacion).NumberFormat = "General" 'General
        HojaInventario.Cells(NuevaFila, ColumnaCostoBulto).NumberFormat = "_-[$R$] * #,##0.00_-;[$R$] * -#,##0.00_-;_-[$R$] * ""-""??_-;_-@_-" ''Formato R$
        HojaInventario.Cells(NuevaFila, ColumnaPrecioBulto).NumberFormat = "_-[$$] * #,##0.00_-;[$$] * -#,##0.00_-;_-[$$] * ""-""??_-;_-@_-" 'Formato $
        ' Reordenar alfabeticamente el inventario
        ReordenarInventario
        
'''''''''''''''''''''''''''''''''''''''Insertar Producto en Tabla de Inventario de cada uno de los clientes''''''''''''''''''''''''''''''''''''''''''''
        
        For Each HojaCliente In LibroClientes.Sheets
        
            If Not HojaCliente.Name = "Inicio" Then
                ' Se inserta una fila entera en en principio de la tabla de inventario del cliente
                HojaCliente.Range(NuevaFila & ":" & NuevaFila).EntireRow.Insert

                'Llenado de la tabla de inventario de cada cliente con los datos ingresados
                HojaCliente.Cells(NuevaFila, ObtenerColumnaDeTabla(HojaCliente, 1, "Producto")) = TextBox_Producto
                HojaCliente.Cells(NuevaFila, ObtenerColumnaDeTabla(HojaCliente, 1, "Codigo")) = TextBox_Codigo.Value
                HojaCliente.Cells(NuevaFila, ObtenerColumnaDeTabla(HojaCliente, 1, "Existencia")) = 0
                HojaCliente.Cells(NuevaFila, ObtenerColumnaDeTabla(HojaCliente, 1, "Cantidad de unidades por bulto")) = TextBox_UnidadesPorBulto.Value
                HojaCliente.Cells(NuevaFila, ObtenerColumnaDeTabla(HojaCliente, 1, "Precio por bulto ($)")) = TextBox_PrecioPorBulto.Value
                
                'Se da el formato correspondiente a cada columna
                HojaCliente.Cells(NuevaFila, ColumnaProductoCliente).NumberFormat = "General" 'General
                HojaCliente.Cells(NuevaFila, ColumnaCodigoCliente).NumberFormat = "0" 'General
                HojaCliente.Cells(NuevaFila, ColumnaPrecioBultoCliente).NumberFormat = "_-[$$] * #,##0.00_-;[$$] * -#,##0.00_-;_-[$$] * ""-""??_-;_-@_-" 'Formato $
            
                ReordenarInventarioClientes HojaCliente
                
            End If
            
        Next HojaCliente
        
        
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Limpieza de Formulario''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        ComentarioOculto = ""
        ComentarioOculto = "[Codigo: " & TextBox_Codigo.Text & "]"
        ComentarioOculto = ComentarioOculto + Chr(13) + "[Producto: " & TextBox_Producto.Text & "]"
        ComentarioOculto = ComentarioOculto + Chr(13) + "[Unidades por bulto: " & TextBox_UnidadesPorBulto.Text & "]"
        ComentarioOculto = ComentarioOculto + Chr(13) + "[Presentacion por unidad: " & TextBox_PresentacionPorUnidad.Text & "]"
        ComentarioOculto = ComentarioOculto + Chr(13) + "[Costo por bulto: " & TextBox_CostoPorBulto.Text & "]"
        ComentarioOculto = ComentarioOculto + Chr(13) + "[Precio por bulto: " & TextBox_PrecioPorBulto.Text & "]"
        
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

        ' Establecer foco en el campo de codigo
        TextBox_Codigo.SetFocus
        
        
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    MsgBox "Producto registrado existosamente", , "Registrar Producto"
    

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
    
    Label_CorrelativoPrefijo.Caption = "Registro"

    Set FormularioAnterior = Me
    
    'Actualizar correlativo en pantalla
    ActualizarCorrelativoEnPantalla Frame_Correlativo, Label_CorrelativoPrefijo
    
End Sub

Private Sub UserForm_Terminate()
    
    Set FormularioAnterior = Nothing
     
End Sub
