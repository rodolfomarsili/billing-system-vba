VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} test_CambiarIDProducto 
   Caption         =   "Cambiar ID Producto"
   ClientHeight    =   3315
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   4812
   OleObjectBlob   =   "test_CambiarIDProducto.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "test_CambiarIDProducto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub btn_Click()

Dim HojaCliente As Worksheet
Dim FilaAModificarCliente As Integer
Dim Fila As Integer
Dim i As Long

    Inicializar
    
    FilaAModificarCliente = ObtenerFila(HojaBaseClientes, ComboBox1.Text, ColumnaProductoCliente)
    Fila = ObtenerFila(HojaInventario, ComboBox1.Text, ColumnaProducto)
    
    For Each HojaCliente In LibroClientes.Sheets
        If Not HojaCliente.Name = "Inicio" Then
            ' Llenado de la tabla de inventario con los datos ingresados
            HojaCliente.Cells(FilaAModificarCliente, ColumnaCodigoCliente) = TextBox1.Value
        End If
    Next HojaCliente

    HojaInventario.Cells(Fila, ColumnaCodigo) = TextBox1.Value
    
    For i = 1 To UltimaFilaHistorial
        If HojaHistorial.Cells(i, ColumnaCodigoHistorial) = ID.Caption Then
            HojaHistorial.Cells(i, ColumnaCodigoHistorial) = TextBox1.Value
        End If
    Next i
    
    MsgBox "Listo"
End Sub

Private Sub ComboBox1_Change()
Dim FilaProducto As Integer

    FilaProducto = ObtenerFila(HojaInventario, ComboBox1.Text, ColumnaProducto)
    
    If FilaProducto > 0 Then
        TextBox1 = HojaInventario.Cells(FilaProducto, ColumnaCodigo)
        ID.Caption = HojaInventario.Cells(FilaProducto, ColumnaCodigo)
    End If
End Sub

Private Sub CommandButton1_Click()
Unload Me
End Sub

Private Sub UserForm_Initialize()
Dim i As Integer
Dim Producto As String

    Inicializar
    
    For i = 2 To UltimaFilaInventario
        Producto = HojaInventario.Cells(i, ColumnaProducto)
        ComboBox1.AddItem (Producto)
    Next i
    
End Sub
