VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form_Prestamos 
   Caption         =   "Prestamos"
   ClientHeight    =   5010
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   8868.001
   OleObjectBlob   =   "form_Prestamos.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "form_Prestamos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ComboBox_Caja_Change()

    Label_AsteriscoCaja.Visible = False
    
    ActualizarSaldoCajaEnPantalla Label_SaldoCaja, ComboBox_Caja.Text
    
    ActualizarSaldoPrestamo
    
End Sub


Private Sub CommandButton_Aceptar_Click()

Dim FilaCliente As Integer
Dim FilaCaja As Byte
Dim ProcesarPrestamo As Byte
Dim Comentario As String
Dim Divisa As String
Dim IDResponsable As String
Dim Fecha As Date
  
    Fecha = TextBox_Dia & "/" & TextBox_Mes & "/" & TextBox_Ano

    Inicializar
    
    Label_AsteriscoMonto.Visible = False
    Label_AsteriscoCaja.Visible = False
    
    FilaCliente = ObtenerFila(HojaClientes, TextBox_IDCliente.Text, ColumnaIDCliente)
    FilaCaja = ObtenerFila(HojaCajas, ComboBox_Caja.Text, ColumnaIDCaja)
    
    IDResponsable = HojaCajas.Cells(FilaCaja, ColumnaIDResponsableCaja)
    
''''''''''''''''''''''Verificaciones'''''''''''''''''''''''''''''

    If FilaCliente = 0 Then
        MsgBox "Debes seleccionar un cliente valido para realizar esta operacion", , "Prestamo"
        Exit Sub
    End If
    
    If FilaCaja = 0 Then
        Label_AsteriscoCaja.Visible = True
        MsgBox "Selecciona una caja valida", , "Prestamo"
        Exit Sub
    End If
    
    If Val(TextBox_Monto) = 0 Then
        Label_AsteriscoMonto.Visible = True
        MsgBox "Ingresa el monto a abonar", , "Prestamo"
        Exit Sub
    End If
    
    If (HojaCajas.Cells(FilaCaja, ColumnaSaldoCaja) - Val(TextBox_Monto)) < 0 Then
        Label_AsteriscoCaja.Visible = True
        MsgBox "Fondos insuficientes para realizar esta operacion", , "Prestamo"
        Exit Sub
    End If
    
    'Ultima verificacion antes de procesar el prestamo
    ProcesarPrestamo = MsgBox("¿Seguro que deseas procesar esta operacion?", vbYesNo + vbExclamation, "Prestamo")
    If ProcesarPrestamo = vbNo Then Exit Sub

''''''''''''''''''''''''''Modificar Prestamo'''''''''''''''''''''''''''''''''''''''
    'Establecer el nuevo saldo de credito para el cliente
    Select Case Mid(ComboBox_Caja.Text, 1, 3)
    
        Case "USD":
                    HojaClientes.Cells(FilaCliente, ColumnaPrestamoUSDCliente) = Val(Label_PrestamoRestante)
                    Divisa = " $"
        Case "BRL":
                    HojaClientes.Cells(FilaCliente, ColumnaPrestamoBRLCliente) = Val(Label_PrestamoRestante)
                    Divisa = " R$"
        Case "VES":
                    HojaClientes.Cells(FilaCliente, ColumnaPrestamoVESCliente) = Val(Label_PrestamoRestante)
                    Divisa = " Bs"
    
    End Select

''''''''''''''''''''''''Limpieza del Formulario'''''''''''''''''''''''''''''''''''''
    
        If Val(TextBox_Monto) < 0 Then
        Comentario = "Monto abonado por el cliente: " & Abs(Val(TextBox_Monto))
    ElseIf Val(TextBox_Monto) > 0 Then
        Comentario = "Monto prestado al cliente: " & Abs(Val(TextBox_Monto))
    End If
    
    'Incluir en el hitorial
    IncluirEnHistorial Frame_Correlativo, Label_CorrelativoPrefijo, Fecha, , , ComboBox_Caja.Text, , Comentario & Divisa, TextBox_IDCliente.Text, IDResponsable, Val(TextBox_Monto), False
        
    'Actualizar saldo de caja en pantalla
    ActualizarSaldoCajaEnPantalla Label_SaldoCaja, ComboBox_Caja.Text
    
    'Actualizar correlativo
    ActualizarCorrelativo Frame_Correlativo, Label_CorrelativoPrefijo
    
    'Actualizar correlativo
    ActualizarCorrelativoEnPantalla Frame_Correlativo, Label_CorrelativoPrefijo
    
    'Limpieza del formulario
    Select Case Mid(ComboBox_Caja.Text, 1, 3)
    
        Case "USD": Label_PrestamoCliente.Caption = HojaClientes.Cells(FilaCliente, ColumnaPrestamoUSDCliente)
        Case "BRL": Label_PrestamoCliente.Caption = HojaClientes.Cells(FilaCliente, ColumnaPrestamoBRLCliente)
        Case "VES": Label_PrestamoCliente.Caption = HojaClientes.Cells(FilaCliente, ColumnaPrestamoVESCliente)
    
    End Select
    TextBox_Monto = Empty
    
    ActualizarDashboard
    
    MsgBox "Prestamo ingresado exitosamente", , "Prestamo"

    
End Sub

Private Sub CommandButton_Cancelar_Click()
    Unload Me
End Sub

Private Sub CommandButton_SeleccionarCliente_Click()
    sec_ListadoDeClientes.Show
End Sub


Private Sub TextBox_IDCliente_Change()
    ActualizarSaldoPrestamo
End Sub

Private Sub TextBox_Monto_Change()
    
    Label_AsteriscoMonto.Visible = False
    If Label_PrestamoCliente <> Empty Then Label_PrestamoRestante = Format(Val(Label_PrestamoCliente) + Val(TextBox_Monto), "0.00")
    If TextBox_Monto = Empty Then Label_PrestamoRestante = Empty
    
End Sub

Private Sub TextBox_Monto_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

Dim UbicacionPunto As Integer
Dim UbicacionMenos As Integer

    UbicacionPunto = InStr(TextBox_Monto.Text, ".")
    UbicacionMenos = InStr(TextBox_Monto.Text, "-")
    
    If (KeyAscii = 46 And UbicacionPunto > 0) Then
        KeyAscii = 0
    End If
    
    If (KeyAscii = 45 And UbicacionMenos > 0) Then
        KeyAscii = 0
    End If
    
    If Not (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 46 Or KeyAscii = 45) Then
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
    
    Label_CorrelativoPrefijo.Caption = "Prestamo"
    
    Label_AsteriscoCaja.Visible = False
    
    CommandButton_SeleccionarCliente.Picture = LoadPicture(ThisWorkbook.Path & "\Resources\images\cliente.jpg")
    
    'Llenado del ComboBox de cajas
    For i = 2 To UltimaFilaCajas
        ComboBox_Caja.AddItem (HojaCajas.Cells(i, ColumnaIDCaja))
    Next i
    
    ComboBox_Caja.Text = "BRL-RODOLFO"
    ActualizarSaldoCajaEnPantalla Label_SaldoCaja, ComboBox_Caja.Text
    
    Set FormularioAnterior = Me
    
    ActualizarCorrelativoEnPantalla Frame_Correlativo, Label_CorrelativoPrefijo
    
End Sub

Private Sub UserForm_Terminate()
    Set FormularioAnterior = Nothing
End Sub

Sub ActualizarSaldoPrestamo()

Dim FilaCliente As Integer
    
    Inicializar

    FilaCliente = ObtenerFila(HojaClientes, TextBox_IDCliente.Text, ColumnaIDCliente)
    
    Select Case Mid(ComboBox_Caja.Text, 1, 3)
        Case "USD":
                        If FilaCliente > 0 Then Label_PrestamoCliente.Caption = Format(HojaClientes.Cells(FilaCliente, ColumnaPrestamoUSDCliente), "0,0.00")
                    Label_SignoSaldo.Caption = "$"
                    Label_SignoMonto.Caption = "$"
                    Label_SignoRestante.Caption = "$"
        Case "BRL":
                        If FilaCliente > 0 Then Label_PrestamoCliente.Caption = Format(HojaClientes.Cells(FilaCliente, ColumnaPrestamoBRLCliente), "0,0.00")
                    Label_SignoSaldo.Caption = "R$"
                    Label_SignoMonto.Caption = "R$"
                    Label_SignoRestante.Caption = "R$"
        Case "VES":
                        If FilaCliente > 0 Then Label_PrestamoCliente.Caption = Format(HojaClientes.Cells(FilaCliente, ColumnaPrestamoVESCliente), "0,0.00")
                    Label_SignoSaldo.Caption = "Bs"
                    Label_SignoMonto.Caption = "Bs"
                    Label_SignoRestante.Caption = "Bs"
    End Select
    
End Sub

