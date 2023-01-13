VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form_PagoCredito 
   Caption         =   "Pago de Credito"
   ClientHeight    =   4050
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8895.001
   OleObjectBlob   =   "form_PagoCredito.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "form_PagoCredito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ComboBox_Caja_Change()

Dim FilaCaja As Byte
    
    On Error Resume Next
    
    Label_AsteriscoCaja.Visible = False
    
    FilaCaja = ObtenerFila(HojaCajas, ComboBox_Caja.Text, ColumnaIDCaja)
    Label_SaldoCaja.Caption = "Saldo Actual: " & Format(HojaCajas.Cells(FilaCaja, ColumnaSaldoCaja), "0.00") & " $"
    
End Sub

Private Sub CommandButton_Pagar_Click()

Dim FilaCliente As Integer
Dim FilaCaja As Byte
Dim ProcesarAbono As Byte
Dim Fecha As String
Dim IDResponsable As String
  
    Fecha = TextBox_Mes & "/" & TextBox_Dia & "/" & TextBox_Ano

    Inicializar
    
    Label_AsteriscoMontoAbonado.Visible = False
    Label_AsteriscoCaja.Visible = False
    
    FilaCliente = ObtenerFila(HojaClientes, TextBox_IDCliente.Text, ColumnaIDCliente)
    FilaCaja = ObtenerFila(HojaCajas, ComboBox_Caja.Text, ColumnaIDCaja)
    
    IDResponsable = HojaCajas.Cells(FilaCaja, ColumnaIDResponsableCaja)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Verificaciones'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    If FilaCliente = 0 Then
        MsgBox "Debes seleccionar un cliente valido para realizar esta operacion", , "Pago de credito"
        Exit Sub
    End If
    
    If FilaCaja = 0 Then
        Label_AsteriscoCaja.Visible = True
        MsgBox "Selecciona una caja valida", , "Pago de credito"
        Exit Sub
    End If
    
    If Val(TextBox_MontoAbonado) <= 0 Then
        Label_AsteriscoMontoAbonado.Visible = True
        MsgBox "Ingresa el monto a abonar", , "Pago de credito"
        Exit Sub
    End If
    
    'Ultima verificacion antes de procesar el abono de credito
    ProcesarAbono = MsgBox("¿Seguro que deseas procesar esta operacion?", vbYesNo + vbExclamation, "Pago de credito")
    If ProcesarAbono = vbNo Then Exit Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Modificar Credito'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Establecer el nuevo saldo de credito para el cliente
    HojaClientes.Cells(FilaCliente, ColumnaSaldoCreditoCliente) = Val(Label_SaldoCreditoRestanteCliente)
    
    'Añadir el dinero abonado a la caja correspondiente
    AbonarSaldoACaja ComboBox_Caja, Val(TextBox_MontoAbonado)
    
    'Actualizar el saldo de la caja en la pantalla
    Label_SaldoCaja.Caption = "Saldo Actual: " & Format(HojaCajas.Cells(FilaCaja, ColumnaSaldoCaja), "0.00") & " $"

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Limpieza del Formulario'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Limpieza del formulario de factura
    TextBox_NombreCliente = Empty
    TextBox_IDCliente = Empty
    TextBox_DireccionCliente = Empty
    TextBox_TelefonoCliente = Empty
    TextBox_MontoAbonado = Empty
    Label_SaldoCreditoCliente = Empty
    Label_SaldoCreditoRestanteCliente = Empty
    
    'Incluir en el hitorial
    IncluirEnHistorial Fecha, , , Label_CorrelativoPrefijo.Caption, ComboBox_Caja.Text, , "Abono: " & TextBox_MontoAbonado.Text & " $", TextBox_IDCliente.Text, IDResponsable
        
    'Actualizar correlativo
    ActualizarCorrelativo Label_CorrelativoPrefijo.Caption
    
    ActualizarCorrelativoEnPantalla Label_CorrelativoPrefijo.Caption
    
    MsgBox "Pago realizado exitosamente", , "Pago de credito"
    
End Sub

Private Sub CommandButton_Cancelar_Click()
    Unload Me
End Sub

Private Sub CommandButton_SeleccionarCliente_Click()
    sec_ListadoDeClientes.Show
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

Private Sub TextBox_MontoAbonado_Change()
    
    Label_AsteriscoMontoAbonado.Visible = False
    If Label_SaldoCreditoCliente <> Empty Then Label_SaldoCreditoRestanteCliente = Format(Val(Label_SaldoCreditoCliente) - Val(TextBox_MontoAbonado), "0.00")
    If TextBox_MontoAbonado = Empty Then Label_SaldoCreditoRestanteCliente = Empty
    
End Sub

Private Sub TextBox_MontoAbonado_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

Dim UbicacionPunto As Integer

    UbicacionPunto = InStr(TextBox_MontoAbonado.Text, ".")
    
    If (KeyAscii = 46 And UbicacionPunto > 0) Then
        KeyAscii = 0
    End If
    
    If Not (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 46) Then
        KeyAscii = 0
    End If
    
End Sub

Private Sub UserForm_Initialize()

Dim i As Integer
Dim FilaCaja As Byte

    Inicializar
    
    TextBox_Dia = Day(Date)
    TextBox_Mes = Month(Date)
    TextBox_Ano = Year(Date)
    
    Label_AsteriscoMontoAbonado.Visible = False
    Label_AsteriscoCaja.Visible = False
    
    Label_CorrelativoPrefijo.Caption = "P-Credito"
    
    CommandButton_SeleccionarCliente.Picture = LoadPicture(ThisWorkbook.Path & "\Resources\images\cliente.jpg")
    
    'Llenado del ComboBox de cajas
    For i = 3 To UltimaFilaCajas
         If (Mid(HojaCajas.Cells(i, ColumnaIDCaja), 1, 3) = "USD") Then ComboBox_Caja.AddItem (HojaCajas.Cells(i, ColumnaIDCaja))
    Next i
    ComboBox_Caja.Text = "USD-DEIBYS"
    FilaCaja = ObtenerFila(HojaCajas, ComboBox_Caja.Text, ColumnaIDCaja)
    Label_SaldoCaja.Caption = "Saldo Actual: " & Format(HojaCajas.Cells(FilaCaja, ColumnaSaldoCaja), "0.00") & " $"
    
    Set FormularioAnterior = Me
    
    ActualizarCorrelativoEnPantalla Label_CorrelativoPrefijo.Caption
    
End Sub

Private Sub UserForm_Terminate()
    Set FormularioAnterior = Nothing
End Sub
