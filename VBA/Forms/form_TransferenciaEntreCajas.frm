VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form_TransferenciaEntreCajas 
   Caption         =   "Transferencia Entre Cajas"
   ClientHeight    =   6135
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   8160
   OleObjectBlob   =   "form_TransferenciaEntreCajas.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "form_TransferenciaEntreCajas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ComboBox_CajaOrigen_Change()
   
Dim FilaOrigen As Byte
    
    On Error Resume Next
  
    
    If Not (TextBox_MontoEnviado = Empty) Then
        TextBox_MontoEnviado = Empty
    End If
    
    'Se obtienen la fila correspondiente a las caja origen
    FilaOrigen = ObtenerFila(HojaCajas, ComboBox_CajaOrigen.Text, ColumnaIDCaja)
    
    'Se obtiene el saldo actual que tiene la caja de origen
    Label_SaldoActualOrigenMonto.Caption = Format(HojaCajas.Cells(FilaOrigen, ColumnaSaldoCaja), "0.00")
    
    'Si no esta seleccionada alguna de las cajas, no se ejecuta la actualizacion
    If (ComboBox_CajaOrigen = Empty Or ComboBox_CajaDestino = Empty) Then
        Exit Sub
    End If
    
    If (ComboBox_CajaOrigen = ComboBox_CajaDestino) Then
        MsgBox "La caja de destino, no puede ser la misma que la de origen", , "Movimiento de cajas"
        Exit Sub
    End If

    
    
    If (Mid(ComboBox_CajaOrigen.Text, 1, 3) = Mid(ComboBox_CajaDestino.Text, 1, 3)) Then

        'Misma divisa
        TextBox_MontoRecibido.Locked = True
        TextBox_MontoRecibido.BackColor = &HE0E0E0
        TextBox_MontoRecibido = TextBox_MontoEnviado

    Else

        'Divisas diferentes
        TextBox_MontoRecibido.Locked = False
        TextBox_MontoRecibido.BackColor = &HFFFFFF

    End If
    
    
End Sub

Private Sub ComboBox_CajaDestino_Change()

Dim FilaDestino As Integer
    
    On Error Resume Next
    
    If Not (TextBox_MontoRecibido = Empty) Then
        TextBox_MontoRecibido = Empty
    End If
    
    'Se obtienen la fila correspondiente a las caja destino
    FilaDestino = ObtenerFila(HojaCajas, ComboBox_CajaDestino.Text, ColumnaIDCaja)
    
    'Se obtiene el saldo actual que tiene la caja de destino
    Label_SaldoActualDestinoMonto.Caption = Format(HojaCajas.Cells(FilaDestino, ColumnaSaldoCaja), "0.00")
    
    'Si no esta seleccionada alguna de las cajas, no se ejecuta la actualizacion
    If (ComboBox_CajaOrigen = Empty Or ComboBox_CajaDestino = Empty) Then
        Exit Sub
    End If
    
    If (ComboBox_CajaOrigen = ComboBox_CajaDestino) Then
        MsgBox "La caja de destino, no puede ser la misma que la de origen", , "Movimiento de cajas"
        Exit Sub
    End If
    
    If (Mid(ComboBox_CajaOrigen.Text, 1, 3) = Mid(ComboBox_CajaDestino.Text, 1, 3)) Then
        'Misma divisa
        
        TextBox_MontoRecibido.Locked = True
        TextBox_MontoRecibido.BackColor = &HE0E0E0
        TextBox_MontoRecibido = TextBox_MontoEnviado
    Else
        'Divisas diferentes
        TextBox_MontoRecibido.Locked = False
        TextBox_MontoRecibido.BackColor = &HFFFFFF
    End If
    
End Sub

Private Sub CommandButton_Cancelar_Click()
    Unload Me
End Sub


Private Sub CommandButton_Enviar_Click()

Dim FilaOrigen As Integer
Dim FilaDestino As Integer
Dim IDResponsable As String
Dim ComentarioOculto As String
Dim Transferir As Byte
Dim Fecha As Date
  
    Fecha = TextBox_Dia & "/" & TextBox_Mes & "/" & TextBox_Ano
    
    Inicializar

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Verificaciones'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    
    IDResponsable = HojaGestion.Range("B3")
    
    'Se obtienen la fila correspondiente a las caja origen
    FilaOrigen = ObtenerFila(HojaCajas, ComboBox_CajaOrigen.Text, ColumnaIDCaja)
    FilaDestino = ObtenerFila(HojaCajas, ComboBox_CajaDestino.Text, ColumnaIDCaja)
    
    If (ComboBox_CajaOrigen = Empty Or ComboBox_CajaDestino = Empty) Then
        MsgBox "Debes seleccionar ambas cajas para proseguir", , "Movimiento de cajas"
        Exit Sub
    End If

    If (ComboBox_CajaOrigen = ComboBox_CajaDestino) Then
        MsgBox "La caja de destino, no puede ser la misma que la de origen", , "Movimiento de cajas"
        Exit Sub
    End If
    
    If (ComboBox_CajaOrigen = Empty Or ComboBox_CajaDestino = Empty Or TextBox_MontoRecibido = Empty Or TextBox_MontoEnviado = Empty) Then
        MsgBox "Debes completar todos los campos", , "Movimiento de cajas"
        Exit Sub
    End If
    
    If (HojaCajas.Cells(FilaOrigen, ColumnaSaldoCaja) - Val(TextBox_MontoEnviado)) < 0 Then
        MsgBox "Fondos insuficientes para realizar esta operacion", , "Movimiento de cajas"
        Exit Sub
    End If
    
    Transferir = MsgBox("Seguro que deseas hacer esta transferencia?", vbYesNo + vbExclamation, "Movimineto de Cajas")
    If Transferir <> 6 Then Exit Sub

    

    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Movimiento de saldos'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    ComentarioOculto = "[Transferencia de caja " & ComboBox_CajaOrigen.Text & " -> " & ComboBox_CajaDestino.Text & "] " + Chr(13) _
                     & "[Monto Enviado " & TextBox_MontoEnviado & " -> " & "Monto Recibido " & TextBox_MontoRecibido & "] " + Chr(13) _
                     & "[Nuevo saldo Caja Origen " & Label_NuevoSaldoOrigenMonto & " -> " & "Nuevo saldo Caja Destino " & Label_NuevoSaldoDestinoMonto & "] " + Chr(13) _
                     & "[" & TextBox_Comentario & "] "
                     
        'Incluir en el historial
    IncluirEnHistorial Frame_Correlativo, Label_CorrelativoPrefijo, Fecha, , , ComboBox_CajaOrigen.Text, 1, ComentarioOculto, , IDResponsable, Val(TextBox_MontoEnviado), False
    IncluirEnHistorial Frame_Correlativo, Label_CorrelativoPrefijo, Fecha, , , ComboBox_CajaDestino.Text, 1, ComentarioOculto, , IDResponsable, Val(TextBox_MontoRecibido), True
    
                
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Limpieza del formulario'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    'Aumentar numero de correlativo
    ActualizarCorrelativo Frame_Correlativo, Label_CorrelativoPrefijo
    
    'Actualizar correlativo en pantalla
    ActualizarCorrelativoEnPantalla Frame_Correlativo, Label_CorrelativoPrefijo
    
    'Limpieza de los campos
    ComboBox_CajaOrigen = Empty
    ComboBox_CajaDestino = Empty
    TextBox_MontoEnviado = Empty
    TextBox_MontoRecibido = Empty
    Label_SaldoActualOrigenMonto.Caption = Empty
    Label_SaldoActualDestinoMonto.Caption = Empty
    Label_NuevoSaldoOrigenMonto.Caption = Empty
    Label_NuevoSaldoDestinoMonto.Caption = Empty
    
    TextBox_MontoRecibido.Locked = False
    TextBox_MontoRecibido.BackColor = &HFFFFFF
    
    ComboBox_CajaOrigen.SetFocus
    
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    ActualizarDashboard
    
    MsgBox "Transferencia realizada exitosamente", , "Movimiento de cajas"
    
    
End Sub

Private Sub TextBox_MontoEnviado_Change()
    
    'Si la casilla de monto recibido esta bloqueda significa que las divisas son iguales, por tanto éste se iguala al monto enviado.
    If TextBox_MontoRecibido.Locked Then TextBox_MontoRecibido = TextBox_MontoEnviado
    
    If Not (TextBox_MontoEnviado.Value = Empty) Then
        Label_NuevoSaldoOrigenMonto.Caption = Format(Val(Label_SaldoActualOrigenMonto.Caption) - Val(TextBox_MontoEnviado.Value), "0.00")
    Else
        Label_NuevoSaldoOrigenMonto.Caption = ""
    End If
    
End Sub

Private Sub TextBox_MontoRecibido_Change()
    
    If Not (TextBox_MontoRecibido.Value = Empty) Then
        Label_NuevoSaldoDestinoMonto.Caption = Format(Val(Label_SaldoActualDestinoMonto.Caption) + Val(TextBox_MontoRecibido.Value), "0.00")
    Else
        Label_NuevoSaldoDestinoMonto.Caption = ""
    End If
    
End Sub

Private Sub TextBox_MontoEnviado_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

Dim UbicacionPunto As Integer

    UbicacionPunto = InStr(TextBox_MontoEnviado.Text, ".")
    
    If (KeyAscii = 46 And UbicacionPunto > 0) Then
        KeyAscii = 0
    End If
    
    If Not (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 46) Then
        KeyAscii = 0
    End If
    
End Sub

Private Sub TextBox_MontoRecibido_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

Dim UbicacionPunto As Integer

    UbicacionPunto = InStr(TextBox_MontoRecibido.Text, ".")
    
    If (KeyAscii = 46 And UbicacionPunto > 0) Then
        KeyAscii = 0
    End If
    
    If Not (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 46) Then
        KeyAscii = 0
    End If
    
End Sub

Private Sub UserForm_Initialize()

Dim i As Integer
    
    Inicializar
    
    TextBox_Dia = Day(Date)
    TextBox_Mes = Month(Date)
    TextBox_Ano = Year(Date)
    
    For i = 2 To UltimaFilaCajas
        ComboBox_CajaOrigen.AddItem (HojaCajas.Cells(i, ColumnaIDCaja))
        ComboBox_CajaDestino.AddItem (HojaCajas.Cells(i, ColumnaIDCaja))
    Next i
    
    Label_CorrelativoPrefijo.Caption = "Movimiento"
    
    Set FormularioAnterior = Me
        
    'Actualizar correlativo en pantalla
    ActualizarCorrelativoEnPantalla Frame_Correlativo, Label_CorrelativoPrefijo
    
End Sub

Private Sub UserForm_Terminate()

    Set FormularioAnterior = Nothing
    
End Sub



