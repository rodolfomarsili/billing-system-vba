VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form_Extras 
   Caption         =   "Ingresos Extras"
   ClientHeight    =   6585
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   5460
   OleObjectBlob   =   "form_Extras.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "form_Extras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton_Aceptar_Click()

Dim FilaCaja As Byte
Dim Procesar As Byte
Dim IDResponsable As String
Dim Comentario As String
Dim Fecha As Date
  
    Fecha = TextBox_Dia & "/" & TextBox_Mes & "/" & TextBox_Ano

    Inicializar
    
    Label_AsteriscoMonto.Visible = False
    Label_AsteriscoCaja.Visible = False
    
    FilaCaja = ObtenerFila(HojaCajas, ComboBox_Caja.Text, ColumnaIDCaja)
    
    IDResponsable = HojaCajas.Cells(FilaCaja, ColumnaIDResponsableCaja)
    
''''''''''''''''''''''Verificaciones'''''''''''''''''''''''''''''

    
    If FilaCaja = 0 Then
        Label_AsteriscoCaja.Visible = True
        MsgBox "Selecciona una caja valida", , "Extras"
        Exit Sub
    End If
    
    If Val(TextBox_Monto) = 0 Then
        Label_AsteriscoMonto.Visible = True
        MsgBox "Ingresa el monto a abonar", , "Extras"
        Exit Sub
    End If
    
    If (HojaCajas.Cells(FilaCaja, ColumnaSaldoCaja) + Val(TextBox_Monto)) < 0 Then
        Label_AsteriscoCaja.Visible = True
        MsgBox "Fondos insuficientes para realizar esta operacion", , "Extras"
        Exit Sub
    End If
    
    'Verificacion de comentario obligatorio
    If (TextBox_Comentario.Text = Empty) Then
        MsgBox "Agrega un comentario a esta transaccion para tener una referencia futura", , "Extras"
        Exit Sub
    End If
    
    'Ultima verificacion antes de procesar el prestamo
    Procesar = MsgBox("¿Seguro que deseas procesar esta operacion?", vbYesNo + vbExclamation, "Extras")
    If Procesar = vbNo Then Exit Sub

''''''''''''''''''''''Modificar saldo de caja'''''''''''''''''''''''''''''
    Comentario = ""
    Comentario = "[" & "Monto: " & TextBox_Monto.Text & " $" & "] " + Chr(13)
    Comentario = Comentario & "[" & "Comentario: " & TextBox_Comentario.Text & "] "
        
        'Incluir en el hitorial
    IncluirEnHistorial Frame_Correlativo, Label_CorrelativoPrefijo, Fecha, , , ComboBox_Caja.Text, , Comentario, , IDResponsable, Val(TextBox_Monto), True

''''''''''''''''''''''Limpieza del formulario'''''''''''''''''''''''''''''
    
    
    'Actualizar saldo de caja en pantalla
    ActualizarSaldoCajaEnPantalla Label_SaldoCaja, ComboBox_Caja.Text
        
    'Actualizar Correlativo
    ActualizarCorrelativo Frame_Correlativo, Label_CorrelativoPrefijo
    
    'Actualizar Correlativo
    ActualizarCorrelativoEnPantalla Frame_Correlativo, Label_CorrelativoPrefijo
    
    'Limpieza del formulario
    TextBox_Comentario = Empty
    TextBox_Monto = Empty
    
    
    ActualizarDashboard
    
    MsgBox "Operacion ejecutada exitosamente", , "Extras"
    
End Sub

Private Sub CommandButton_Cancelar_Click()
    Unload Me
End Sub


Private Sub ComboBox_Caja_Change()

    Label_AsteriscoCaja.Visible = False
    
    ActualizarSaldoCajaEnPantalla Label_SaldoCaja, ComboBox_Caja.Text
    
    ActualizarSignoDivisa

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

Dim i As Byte

    Inicializar
    
    TextBox_Dia = Day(Date)
    TextBox_Mes = Month(Date)
    TextBox_Ano = Year(Date)
    
    Label_CorrelativoPrefijo.Caption = "Extras"
    
    Label_AsteriscoCaja.Visible = False
    
    'Llenado del ComboBox de cajas
    For i = 2 To UltimaFilaCajas
         ComboBox_Caja.AddItem (HojaCajas.Cells(i, ColumnaIDCaja))
    Next i
    
    ComboBox_Caja.Text = "USD-DEIBYS"
    ActualizarSaldoCajaEnPantalla Label_SaldoCaja, ComboBox_Caja.Text
    
    Set FormularioAnterior = Me
    
'    'Actualizar correlativo en pantalla
    ActualizarCorrelativoEnPantalla Frame_Correlativo, Label_CorrelativoPrefijo

End Sub

Private Sub UserForm_Terminate()
    
    Set FormularioAnterior = Nothing
     
End Sub

Sub ActualizarSignoDivisa()

    Select Case Mid(ComboBox_Caja.Text, 1, 3)
        Case "USD":
                    Label_SignoMonto.Caption = "$"
        Case "BRL":
                    Label_SignoMonto.Caption = "R$"
        Case "VES":
                    Label_SignoMonto.Caption = "Bs"
    End Select
    
End Sub
