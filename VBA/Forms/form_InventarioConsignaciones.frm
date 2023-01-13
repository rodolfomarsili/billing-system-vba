VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form_InventarioConsignaciones 
   Caption         =   "Inventario de Consignaciones"
   ClientHeight    =   7965
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   13560
   OleObjectBlob   =   "form_InventarioConsignaciones.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "form_InventarioConsignaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub ListBox_Consignaciones_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

        If ListBox_Consignaciones.ListIndex > -1 Then
                
                Load sec_ModificarPrecioConsignacion
                
                sec_ModificarPrecioConsignacion.TextBox_PrecioPorBulto = Format(ListBox_Consignaciones.List(ListBox_Consignaciones.ListIndex, 2), "0.00")
                sec_ModificarPrecioConsignacion.Label_PrecioPorBulto = sec_ModificarPrecioConsignacion.TextBox_PrecioPorBulto
                
                sec_ModificarPrecioConsignacion.Show
                
                ActualizarListado
        End If
End Sub

Private Sub TextBox_IDCliente_Change()
    
    ActualizarListado
    
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

Private Sub TextBox_IDCliente_AfterUpdate()

Dim FilaDeCliente As Integer

    Inicializar
    
    On Error Resume Next
    
    
    FilaDeCliente = ObtenerFila(HojaClientes, TextBox_IDCliente.Text, ColumnaIDCliente)
    
    TextBox_NombreCliente = HojaClientes.Cells(FilaDeCliente, ColumnaNombreCliente)
    TextBox_DireccionCliente = HojaClientes.Cells(FilaDeCliente, ColumnaDireccionCliente)
    TextBox_TelefonoCliente = HojaClientes.Cells(FilaDeCliente, ColumnaTelefonoCliente)
    
End Sub

Private Sub CommandButton_Cancelar_Click()
    Unload Me
End Sub

Private Sub CommandButton_SeleccionarCliente_Click()
    sec_ListadoDeClientes.Show
End Sub


Private Sub UserForm_Click()
    ListBox_Consignaciones.ListIndex = -1
End Sub

Private Sub UserForm_Initialize()

    
    CommandButton_SeleccionarCliente.Picture = LoadPicture(ThisWorkbook.Path & "\Resources\images\cliente.jpg")
    
    
    'Formato del listado de consignaciones
    ListBox_Consignaciones.ColumnCount = 7
    ListBox_Consignaciones.ColumnWidths = "100 pt; 298 pt; 60 pt; 30 pt; 50 pt; 60 pt; 60 pt"
    
    Set FormularioAnterior = Me
    
End Sub

Private Sub UserForm_Terminate()
    Set FormularioAnterior = Nothing
End Sub

Sub ActualizarListado()

Dim FilaDeCliente As Integer
Dim i As Integer
Dim a As Integer

    Inicializar
    
        ListBox_Consignaciones.Clear
    a = 0
    'Ingreso de las existencias en el inventario de consignacion del cliente seleccionado
    For i = 2 To UltimaFilaInventario
        If LibroClientes.Sheets(TextBox_IDCliente.Text).Cells(i, ColumnaExistenciaCliente) <> 0 Then
            
            ListBox_Consignaciones.AddItem
            ListBox_Consignaciones.List(a, 0) = LibroClientes.Sheets(TextBox_IDCliente.Text).Cells(i, ColumnaCodigoCliente)
            ListBox_Consignaciones.List(a, 1) = LibroClientes.Sheets(TextBox_IDCliente.Text).Cells(i, ColumnaProductoCliente)
            ListBox_Consignaciones.List(a, 2) = Format(LibroClientes.Sheets(TextBox_IDCliente.Text).Cells(i, ColumnaPrecioBultoCliente), "0.00")
            
            ListBox_Consignaciones.List(a, 4) = LibroClientes.Sheets(TextBox_IDCliente.Text).Cells(i, ColumnaExistenciaCliente)
            ListBox_Consignaciones.List(a, 5) = Format(LibroClientes.Sheets(TextBox_IDCliente.Text).Cells(i, ColumnaPrecioUnitarioCliente), "0.0000")
            ListBox_Consignaciones.List(a, 6) = Format(LibroClientes.Sheets(TextBox_IDCliente.Text).Cells(i, ColumnaImporteCliente), "0.0000")
            
            a = a + 1
        End If
    Next i
    
    ActualizarImporte
    
End Sub
Sub ActualizarImporte()
Dim a As Integer
Dim i As Integer
Dim Importe As Single

    a = ListBox_Consignaciones.ListCount
    Importe = 0
    If a > 0 Then
        For i = 0 To a - 1
            Importe = Importe + Val(ListBox_Consignaciones.List(i, 6))
        Next i
    End If
    
    Label_Importe.Caption = Format(Importe, "0,0.000")
    
End Sub
