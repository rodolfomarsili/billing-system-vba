VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} sec_ListadoDeProductos 
   Caption         =   "Listado de Productos"
   ClientHeight    =   7200
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   12516
   OleObjectBlob   =   "sec_ListadoDeProductos.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "sec_ListadoDeProductos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton_Cancelar_Click()
    Unload Me
End Sub


Private Sub CommandButton_IngresarProducto_Click()
    Anadir
End Sub


Private Sub ListBox_ListadoProductos_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    Anadir
    
End Sub

Private Sub ListBox_ListadoProductos_Click()
    
Dim FilaDeCodigo As Integer

    Inicializar
    
    On Error Resume Next
    
    FilaDeCodigo = ObtenerFila(HojaInventario, ListBox_ListadoProductos.List(ListBox_ListadoProductos.ListIndex, 0), ColumnaCodigo)

    Label_Producto.Caption = HojaInventario.Cells(FilaDeCodigo, ColumnaProducto)
    Label_UnidadesPorBulto.Caption = HojaInventario.Cells(FilaDeCodigo, ColumnaUnidadesPorBulto)
    Label_Presentacion.Caption = HojaInventario.Cells(FilaDeCodigo, ColumnaPresentacion)
    Label_Precio.Caption = Format(HojaInventario.Cells(FilaDeCodigo, ColumnaPrecioBulto), "0.00")
    Label_Existencia.Caption = HojaInventario.Cells(FilaDeCodigo, ColumnaExistencia)

    
    If Label_Existencia = 0 Then
        Label_Existencia.ForeColor = &HFF&
    Else
        Label_Existencia.ForeColor = &H80000012
    End If
    
End Sub

Private Sub ListBox_ListadoProductos_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then Unload Me
End Sub


Private Sub TextBox_Buscar_Change()
    FiltrarProductoEnListBox
End Sub

Private Sub UserForm_Initialize()

Dim i As Integer
Dim a As Integer
        
    Inicializar

    ' Formato del listado
    ListBox_ListadoProductos.ColumnCount = 3
    ListBox_ListadoProductos.ColumnWidths = "100 pt; 298 pt; 60 pt"
    
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    a = 0
    For i = 2 To UltimaFilaInventario
    
        ListBox_ListadoProductos.AddItem
        ListBox_ListadoProductos.List(a, 0) = HojaInventario.Cells(i, ColumnaCodigo)
        ListBox_ListadoProductos.List(a, 1) = HojaInventario.Cells(i, ColumnaProducto)
        ListBox_ListadoProductos.List(a, 2) = Format(HojaInventario.Cells(i, ColumnaCostoBulto), "0.00")
        
        
        a = a + 1
    Next i

    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    ''''''''''''''''''''''''''''''''''''''''''''''''''''

    
End Sub

Private Sub UserForm_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then Unload Me
End Sub


Sub Anadir()

On Error Resume Next
    
    FormularioAnterior.TextBox_Codigo = ListBox_ListadoProductos.List(ListBox_ListadoProductos.ListIndex, 0)
    FormularioAnterior.TextBox_Producto = ListBox_ListadoProductos.List(ListBox_ListadoProductos.ListIndex, 1)
    
    FormularioAnterior.TextBox_CostoPorBulto = Format(ListBox_ListadoProductos.List(ListBox_ListadoProductos.ListIndex, 2), "0.00")
    FormularioAnterior.TextBox_UnidadesPorBulto = Label_UnidadesPorBulto.Caption
    FormularioAnterior.TextBox_PresentacionPorUnidad = Label_Presentacion.Caption
    FormularioAnterior.Label_ExistenciaCantidad = Label_Existencia.Caption
    FormularioAnterior.TextBox_PrecioPorBulto = Label_Precio.Caption
    
    FormularioAnterior.Label_RespaldoProducto = ListBox_ListadoProductos.List(ListBox_ListadoProductos.ListIndex, 1)
    FormularioAnterior.Label_RespaldoCostoPorBulto = Format(ListBox_ListadoProductos.List(ListBox_ListadoProductos.ListIndex, 2), "0.00")
    FormularioAnterior.Label_RespaldoUnidadesPorBulto = Label_UnidadesPorBulto.Caption
    FormularioAnterior.Label_RespaldoPresentacionPorUnidad = Label_Presentacion.Caption
    FormularioAnterior.Label_ExistenciaCantidad = Label_Existencia.Caption
    FormularioAnterior.Label_RespaldoPrecioPorBulto = Label_Precio.Caption
    
    Unload Me

    
End Sub

Sub FiltrarProductoEnListBox()

Dim i As Integer
Dim a As Integer
Dim Producto As String
Dim Codigo As String

    Inicializar
    
    ListBox_ListadoProductos.Clear
    
    If TextBox_Buscar = "" Then
        
        a = 0
        For i = 2 To UltimaFilaInventario
        
            ListBox_ListadoProductos.AddItem
            ListBox_ListadoProductos.List(a, 0) = HojaInventario.Cells(i, ColumnaCodigo)
            ListBox_ListadoProductos.List(a, 1) = HojaInventario.Cells(i, ColumnaProducto)
            ListBox_ListadoProductos.List(a, 2) = Format(HojaInventario.Cells(i, ColumnaCostoBulto), "0.00")
            
            a = a + 1
            
        Next i
        
        Exit Sub
        
    End If
    
    For i = 2 To UltimaFilaInventario
        Producto = HojaInventario.Cells(i, ColumnaProducto)
        Codigo = HojaInventario.Cells(i, ColumnaCodigo)
        
        If UCase(Producto) Like "*" & UCase(TextBox_Buscar.Value) & "*" Then
        
            ListBox_ListadoProductos.AddItem
            ListBox_ListadoProductos.List(a, 0) = HojaInventario.Cells(i, ColumnaCodigo)
            ListBox_ListadoProductos.List(a, 1) = HojaInventario.Cells(i, ColumnaProducto)
            ListBox_ListadoProductos.List(a, 2) = Format(HojaInventario.Cells(i, ColumnaCostoBulto), "0.00")
            
            a = a + 1
            
        'He añadido todo este fragmento para que me busque al mismo tiempo por codigo.
        ElseIf Codigo Like "*" & UCase(TextBox_Buscar.Value) & "*" Then
        
            ListBox_ListadoProductos.AddItem
            ListBox_ListadoProductos.List(a, 0) = HojaInventario.Cells(i, ColumnaCodigo)
            ListBox_ListadoProductos.List(a, 1) = HojaInventario.Cells(i, ColumnaProducto)
            ListBox_ListadoProductos.List(a, 2) = Format(HojaInventario.Cells(i, ColumnaCostoBulto), "0.00")
            
            a = a + 1
            
        End If
        
    Next i


End Sub
