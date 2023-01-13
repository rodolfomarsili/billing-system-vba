VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} test_CambiarIDCliente 
   Caption         =   "Cambiar ID Cliente"
   ClientHeight    =   3315
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   4812
   OleObjectBlob   =   "test_CambiarIDCliente.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "test_CambiarIDCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Private Sub btn_Click()

Dim FilaCliente As Integer
Dim IDCliente As String
Dim i As Long

    FilaCliente = ObtenerFila(HojaClientes, ComboBox1.Text, ColumnaNombreCliente)
    
    If FilaCliente > 0 Then
        
        
        IDCliente = ID.Caption
        
        'Ingreso de nueva hoja en libro de clientes
        Application.Windows("Clientes.xlsm").Visible = True
        LibroClientes.Activate
        LibroClientes.Sheets(IDCliente).Name = TextBox1.Text
        LibroClientes.Sheets("Inicio").Select
        Application.Windows("Clientes.xlsm").Visible = False
        
        HojaClientes.Cells(FilaCliente, ColumnaIDCliente) = TextBox1.Text
        
        For i = 1 To UltimaFilaHistorial
            If HojaHistorial.Cells(i, ColumnaIDClienteHistorial) = ID.Caption Then
                HojaHistorial.Cells(i, ColumnaIDClienteHistorial) = TextBox1.Value
            End If
        Next i
        
        MsgBox "Listo"
        
    End If
    
End Sub

Private Sub ComboBox1_Change()
Dim FilaCliente As Integer

    FilaCliente = ObtenerFila(HojaClientes, ComboBox1.Text, ColumnaNombreCliente)
    
    If FilaCliente > 0 Then
        TextBox1 = HojaClientes.Cells(FilaCliente, ColumnaIDCliente)
        ID.Caption = HojaClientes.Cells(FilaCliente, ColumnaIDCliente)
    End If
End Sub

Private Sub CommandButton1_Click()
Unload Me
End Sub

Private Sub UserForm_Initialize()
Dim i As Integer
Dim Cliente As String

    Inicializar
    
    For i = 2 To UltimaFilaClientes
        Cliente = HojaClientes.Cells(i, ColumnaNombreCliente)
        ComboBox1.AddItem (Cliente)
    Next i
    
End Sub
