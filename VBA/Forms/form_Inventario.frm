VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form_Inventario 
   Caption         =   "Inventario"
   ClientHeight    =   10575
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   14292
   OleObjectBlob   =   "form_Inventario.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "form_Inventario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton_Salir_Click()
    Unload Me
End Sub

Private Sub ListBox_Inventario_Click()

Dim FilaDeCodigo As Integer

    Inicializar
    
    On Error Resume Next
    
    FilaDeCodigo = ObtenerFila(HojaInventario, ListBox_Inventario.List(ListBox_Inventario.ListIndex, 0), ColumnaCodigo)

    Label_UnidadesPorBulto.Caption = HojaInventario.Cells(FilaDeCodigo, ColumnaUnidadesPorBulto)
    Label_Presentacion.Caption = HojaInventario.Cells(FilaDeCodigo, ColumnaPresentacion)
        
End Sub

Private Sub ListBox_Inventario_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
        
        Load form_ModificarProducto
        
        Modificar
        
        form_ModificarProducto.Show
        
        ActualizarListado
        
End Sub

Private Sub UserForm_Initialize()
        
    'Formato del listado de consignaciones
    ListBox_Inventario.ColumnCount = 8
    ListBox_Inventario.ColumnWidths = "100 pt; 276 pt; 60 pt; 60 pt; 30 pt; 50 pt; 60 pt; 60 pt"
    
'    ''''''''''''''''''''''''''''''''''''''''''''''''''''
'    Application.ScreenUpdating = False
'    Application.Calculation = xlCalculationManual
'    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    
   ActualizarListado
    
'    ''''''''''''''''''''''''''''''''''''''''''''''''''''
'    Application.ScreenUpdating = True
'    Application.Calculation = xlCalculationAutomatic
'    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    
End Sub

Sub Modificar()

On Error Resume Next
    
    FormularioAnterior.TextBox_Codigo = ListBox_Inventario.List(ListBox_Inventario.ListIndex, 0)
    FormularioAnterior.TextBox_Producto = ListBox_Inventario.List(ListBox_Inventario.ListIndex, 1)
    
    FormularioAnterior.TextBox_CostoPorBulto = Format(ListBox_Inventario.List(ListBox_Inventario.ListIndex, 2), "0.00")
    FormularioAnterior.TextBox_PrecioPorBulto = Format(ListBox_Inventario.List(ListBox_Inventario.ListIndex, 3), "0.00")
    FormularioAnterior.TextBox_UnidadesPorBulto = Label_UnidadesPorBulto.Caption
    FormularioAnterior.TextBox_PresentacionPorUnidad = Label_Presentacion.Caption
    FormularioAnterior.Label_ExistenciaCantidad = ListBox_Inventario.List(ListBox_Inventario.ListIndex, 5)
    
    FormularioAnterior.Label_RespaldoProducto = ListBox_Inventario.List(ListBox_Inventario.ListIndex, 1)
    FormularioAnterior.Label_RespaldoCostoPorBulto = Format(ListBox_Inventario.List(ListBox_Inventario.ListIndex, 2), "0.00")
    FormularioAnterior.Label_RespaldoPrecioPorBulto = Format(ListBox_Inventario.List(ListBox_Inventario.ListIndex, 3), "0.00")
    FormularioAnterior.Label_RespaldoUnidadesPorBulto = Label_UnidadesPorBulto.Caption
    FormularioAnterior.Label_RespaldoPresentacionPorUnidad = Label_Presentacion.Caption
    FormularioAnterior.Label_ExistenciaCantidad = ListBox_Inventario.List(ListBox_Inventario.ListIndex, 5)

    
End Sub

Sub ActualizarListado()
     
Dim a As Integer
Dim i As Integer

             Inicializar
        
        ListBox_Inventario.Clear
        
        a = 0
    For i = 2 To UltimaFilaInventario
    
        ListBox_Inventario.AddItem
        ListBox_Inventario.List(a, 0) = HojaInventario.Cells(i, ColumnaCodigo)
        ListBox_Inventario.List(a, 1) = HojaInventario.Cells(i, ColumnaProducto)
        ListBox_Inventario.List(a, 2) = Format(HojaInventario.Cells(i, ColumnaCostoBulto), "0.00")
        ListBox_Inventario.List(a, 3) = Format(HojaInventario.Cells(i, ColumnaPrecioBulto), "0.00")
        ListBox_Inventario.List(a, 5) = HojaInventario.Cells(i, ColumnaExistencia)
        ListBox_Inventario.List(a, 6) = Format(HojaInventario.Cells(i, ColumnaPrecioUnidad), "0.0000")
        ListBox_Inventario.List(a, 7) = Format(HojaInventario.Cells(i, ColumnaImportePrecio), "0.00")
        
        a = a + 1
    Next i
    
    ActualizarImporte
    
End Sub

Sub ActualizarImporte()
Dim a As Integer
Dim i As Integer
Dim Importe As Single

    a = ListBox_Inventario.ListCount
    Importe = 0
    If a > 0 Then
        For i = 0 To a - 1
            Importe = Importe + Val(ListBox_Inventario.List(i, 7))
        Next i
    End If
    
    Label_Importe.Caption = Format(Importe, "0,0.000")
    
End Sub
