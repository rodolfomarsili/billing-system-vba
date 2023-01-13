VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} sec_ListadoDeClientes 
   Caption         =   "Seleccionar Cliente"
   ClientHeight    =   7245
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   10548
   OleObjectBlob   =   "sec_ListadoDeClientes.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "sec_ListadoDeClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton_Cancelar_Click()
    Unload Me
End Sub

Private Sub CommandButton_Seleccionar_Click()

    Anadir

End Sub


Private Sub ListBox_ListadoClientes_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

   Anadir

End Sub

Private Sub ListBox_ListadoClientes_Click()
    
Dim FilaIDCliente As Integer

    Inicializar
    
    On Error Resume Next
    
    
    
    FilaIDCliente = ObtenerFila(HojaClientes, ListBox_ListadoClientes.List(ListBox_ListadoClientes.ListIndex, 0), ColumnaIDCliente)
    
    Label_IDCliente.Caption = HojaClientes.Cells(FilaIDCliente, ColumnaIDCliente)
    Label_NombreCliente.Caption = HojaClientes.Cells(FilaIDCliente, ColumnaNombreCliente)
    Label_DireccionCliente.Caption = HojaClientes.Cells(FilaIDCliente, ColumnaDireccionCliente)
    Label_TelefonoCliente.Caption = HojaClientes.Cells(FilaIDCliente, ColumnaTelefonoCliente)
    CheckBox_CreditoCliente.Value = HojaClientes.Cells(FilaIDCliente, ColumnaCreditoCliente)
    CheckBox_ConsignacionCliente.Value = HojaClientes.Cells(FilaIDCliente, ColumnaConsignacionCliente)
    
    'Mostrar saldo de consignacion
    Label_SaldoConsignacion.Caption = 0
    If HojaClientes.Cells(FilaIDCliente, ColumnaSaldoConsignacionCliente).Value <> 0 Then
        Label_SaldoConsignacion.Caption = Format(HojaClientes.Cells(FilaIDCliente, ColumnaSaldoConsignacionCliente), "0.0000")
    End If
    
    'Mostrar los datos de credito del cliente
    If Label_IDCliente.Caption <> "V-00000000" Then
        Label_LimiteCredito.Caption = HojaClientes.Cells(FilaIDCliente, ColumnaLimiteCreditoCliente)
        Label_SaldoCredito.Caption = HojaClientes.Cells(FilaIDCliente, ColumnaSaldoCreditoCliente)
    Else
        Label_LimiteCredito.Caption = 0
        Label_SaldoCredito.Caption = 0
    End If
    
    'Limite de credito excedido
    If Val(Label_SaldoCredito) > Val(Label_LimiteCredito) Then
        Label_SaldoCredito.ForeColor = &HFF&
        Label_CreditoExcedido.Visible = True
    Else
        Label_SaldoCredito.ForeColor = &H80000012
        Label_CreditoExcedido.Visible = False
    End If
    
End Sub

Private Sub ListBox_ListadoClientes_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then Unload Me
End Sub

Private Sub TextBox_Buscar_Change()
    FiltrarClienteEnListBox
End Sub

Private Sub UserForm_Initialize()

Dim i As Integer
Dim a As Integer
        
    Inicializar
    
    Label_CreditoExcedido.Visible = False

    ' Formato del listado
    ListBox_ListadoClientes.ColumnCount = 2
    ListBox_ListadoClientes.ColumnWidths = "80 pt; 250 pt"
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    a = 0
    For i = 2 To UltimaFilaClientes
    
        ListBox_ListadoClientes.AddItem
        ListBox_ListadoClientes.List(a, 0) = HojaClientes.Cells(i, ColumnaIDCliente)
        ListBox_ListadoClientes.List(a, 1) = HojaClientes.Cells(i, ColumnaNombreCliente)

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

    FormularioAnterior.TextBox_IDCliente = ListBox_ListadoClientes.List(ListBox_ListadoClientes.ListIndex, 0)
    FormularioAnterior.TextBox_NombreCliente = ListBox_ListadoClientes.List(ListBox_ListadoClientes.ListIndex, 1)
    FormularioAnterior.TextBox_DireccionCliente = Label_DireccionCliente.Caption
    FormularioAnterior.TextBox_TelefonoCliente = Label_TelefonoCliente.Caption
    FormularioAnterior.TextBox_LimiteCreditoCliente = Label_LimiteCredito.Caption
    FormularioAnterior.Label_SaldoCreditoCliente = Format(Label_SaldoCredito, "0.00")
    FormularioAnterior.CheckBox_CreditoCliente = CheckBox_CreditoCliente
    FormularioAnterior.CheckBox_ConsignacionCliente = CheckBox_CreditoCliente
    
    
    FormularioAnterior.Label_RespaldoNombre.Caption = ListBox_ListadoClientes.List(ListBox_ListadoClientes.ListIndex, 1)
    FormularioAnterior.Label_RespaldoDireccion.Caption = Label_DireccionCliente.Caption
    FormularioAnterior.Label_RespaldoTelefono.Caption = Label_TelefonoCliente.Caption
    FormularioAnterior.Label_RespaldoLimiteCredito = Label_LimiteCredito.Caption
    FormularioAnterior.CheckBox_RespaldoCredito = CheckBox_CreditoCliente
    FormularioAnterior.CheckBox_RespaldoConsignacion = CheckBox_CreditoCliente
    
    
    Unload Me
    
End Sub

Sub FiltrarClienteEnListBox()

Dim i As Integer
Dim a As Integer
Dim Nombre As String
Dim ID As String

    Inicializar
    
    ListBox_ListadoClientes.Clear
    
    If TextBox_Buscar = "" Then
        
        a = 0
        For i = 2 To UltimaFilaClientes
        
            ListBox_ListadoClientes.AddItem
            ListBox_ListadoClientes.List(a, 0) = HojaClientes.Cells(i, ColumnaIDCliente)
            ListBox_ListadoClientes.List(a, 1) = HojaClientes.Cells(i, ColumnaNombreCliente)
    
            a = a + 1
                
        Next i
        
        Exit Sub
        
    End If
    
    For i = 2 To UltimaFilaClientes
        Nombre = HojaClientes.Cells(i, ColumnaNombreCliente)
        ID = HojaClientes.Cells(i, ColumnaIDCliente)
        
        If UCase(Nombre) Like "*" & UCase(TextBox_Buscar.Value) & "*" Then
        
            ListBox_ListadoClientes.AddItem
            ListBox_ListadoClientes.List(a, 0) = HojaClientes.Cells(i, ColumnaIDCliente)
            ListBox_ListadoClientes.List(a, 1) = HojaClientes.Cells(i, ColumnaNombreCliente)

        a = a + 1
            
        'He añadido todo este fragmento para que me busque al mismo tiempo por codigo.
        ElseIf UCase(ID) Like "*" & UCase(TextBox_Buscar.Value) & "*" Then
        
            ListBox_ListadoClientes.AddItem
            ListBox_ListadoClientes.List(a, 0) = HojaClientes.Cells(i, ColumnaIDCliente)
            ListBox_ListadoClientes.List(a, 1) = HojaClientes.Cells(i, ColumnaNombreCliente)
    
            a = a + 1
            
        End If
        
    Next i


End Sub
