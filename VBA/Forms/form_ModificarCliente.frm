VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form_ModificarCliente 
   Caption         =   "Modificar Cliente"
   ClientHeight    =   5115
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   10128
   OleObjectBlob   =   "form_ModificarCliente.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "form_ModificarCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton_Cancelar_Click()
Unload Me
End Sub

Private Sub CommandButton_Modificar_Click()

Dim ComentarioOculto As String
Dim IDResponsable As String
Dim LargoTelefono As Integer
Dim FilaCliente As Integer
Dim ModificarCliente As Integer
Dim Fecha As Date
  
    Fecha = TextBox_Dia & "/" & TextBox_Mes & "/" & TextBox_Ano



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Verificaciones'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    'Verificacion de las demas casillas vacias
    If (TextBox_NombreCliente = Empty Or TextBox_DireccionCliente = Empty Or TextBox_TelefonoCliente = Empty) Then
        MsgBox "Debes rellenar todos los campos", , "Modificar Cliente"
        Exit Sub
    End If
    
    'Verificacion de largo del numero de telefono
    LargoTelefono = Len(TextBox_TelefonoCliente)
    
    If (LargoTelefono <> 12) Then
        MsgBox "Ingresa un numero de telefono valido", , "Modificar Cliente"
        Exit Sub
    End If
    
    ModificarCliente = MsgBox("Seguro que deseas modificar este registro?", vbYesNo + vbExclamation, "Modificar Cliente")
    If Not ModificarCliente = 6 Then
       Exit Sub
    End If
    
      
        
            

    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    
    Inicializar

    IDResponsable = HojaGestion.Range("B3")
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Modificacion del cliente'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        ComentarioOculto = ""
        ComentarioOculto = "[ID Cliente: " & TextBox_IDCliente & "] " + Chr(13)
    
    'Todos estos IF verifican si se ha hecho algun cambio en alguno de los campos del formulario, y de haberlos hecho
    'añaden automaticamente un comentario al historial señalando el cambio respectivo
    If (TextBox_NombreCliente <> Label_RespaldoNombre) Then
        ComentarioOculto = ComentarioOculto & "[Modificacion de nombre " & Label_RespaldoNombre & " -> " & TextBox_NombreCliente & "] " + Chr(13)
    End If
    If (TextBox_TelefonoCliente <> Label_RespaldoTelefono) Then
        ComentarioOculto = ComentarioOculto & "[Modificacion de telefono " & Label_RespaldoTelefono & " -> " & TextBox_TelefonoCliente & "] " + Chr(13)
    End If
    If (TextBox_DireccionCliente <> Label_RespaldoDireccion) Then
        ComentarioOculto = ComentarioOculto & "[Modificacion de direccion " & Label_RespaldoDireccion & " -> " & TextBox_DireccionCliente & "] " + Chr(13)
    End If
    If (TextBox_LimiteCreditoCliente <> Label_RespaldoLimiteCredito) Then
        ComentarioOculto = ComentarioOculto & "[Modificacion de limite de credito " & Label_RespaldoLimiteCredito & " -> " & TextBox_LimiteCreditoCliente & "] " + Chr(13)
    End If
    If (CheckBox_CreditoCliente <> CheckBox_RespaldoCredito) Then
        ComentarioOculto = ComentarioOculto & "[Modificacion de permiso de credito " & CheckBox_RespaldoCredito & " -> " & CheckBox_CreditoCliente & "] " + Chr(13)
    End If
    If (CheckBox_ConsignacionCliente <> CheckBox_RespaldoConsignacion) Then
        ComentarioOculto = ComentarioOculto & "[Modificacion de permiso de consignacion " & CheckBox_RespaldoConsignacion & " -> " & CheckBox_ConsignacionCliente & "] " + Chr(13)
    End If
    
        'Fila del cliente a modificar
        FilaCliente = ObtenerFila(HojaClientes, TextBox_IDCliente.Text, ColumnaIDCliente)
        'LLenado del registro del nuevo cliente
        HojaClientes.Cells(FilaCliente, ColumnaNombreCliente) = TextBox_NombreCliente
        HojaClientes.Cells(FilaCliente, ColumnaDireccionCliente) = TextBox_DireccionCliente
        HojaClientes.Cells(FilaCliente, ColumnaTelefonoCliente) = TextBox_TelefonoCliente
        HojaClientes.Cells(FilaCliente, ColumnaLimiteCreditoCliente) = Val(TextBox_LimiteCreditoCliente)
        HojaClientes.Cells(FilaCliente, ColumnaCreditoCliente) = CheckBox_CreditoCliente
        HojaClientes.Cells(FilaCliente, ColumnaConsignacionCliente) = CheckBox_ConsignacionCliente
        

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Limpieza del Formulario'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
    'Incluir en el historial
    IncluirEnHistorial Frame_Correlativo, Label_CorrelativoPrefijo, Fecha, , , , , ComentarioOculto, , IDResponsable

    'Actualizar correlativo
    ActualizarCorrelativo Frame_Correlativo, Label_CorrelativoPrefijo
    
    'Actualizar correlativo en pantalla
    ActualizarCorrelativoEnPantalla Frame_Correlativo, Label_CorrelativoPrefijo
    
    'Reordenamiento de lista de clientes
    ReordenarClientes
    
    'Limpieza de los campos del formulario
    TextBox_IDCliente = Empty
    TextBox_NombreCliente = Empty
    TextBox_DireccionCliente = Empty
    TextBox_TelefonoCliente = Empty
    TextBox_LimiteCreditoCliente = Empty
    CheckBox_CreditoCliente.Value = False
    CheckBox_ConsignacionCliente.Value = False
    
    
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    
    MsgBox "Cliente modificado existosamente", , "Modificar Cliente"
    
    
End Sub


Private Sub TextBox_TelefonoCliente_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    'Cuando el numero de caracteres es igual a  en la casilla del numero de telefono, éste ya no recibe mas caracteres.
    If Len(TextBox_TelefonoCliente) = 12 Then KeyAscii = 0
    'Agregado automatico de un guion despues de agregar el codigo de area del telefono
    If Len(TextBox_TelefonoCliente) = 4 Then TextBox_TelefonoCliente = TextBox_TelefonoCliente & "-"
    
    If Not (KeyAscii >= 48 And KeyAscii <= 57) Then
        KeyAscii = 0
    End If
    
End Sub


Private Sub CommandButton_SeleccionarCliente_Click()
    sec_ListadoDeClientes.Show
End Sub


Private Sub UserForm_Initialize()
    
    TextBox_Dia = Day(Date)
    TextBox_Mes = Month(Date)
    TextBox_Ano = Year(Date)
    
    Label_CorrelativoPrefijo.Caption = "Modificacion"
    
    Set FormularioAnterior = Me
        
    'Actualizar correlativo en pantalla
    ActualizarCorrelativoEnPantalla Frame_Correlativo, Label_CorrelativoPrefijo
    
End Sub

Private Sub UserForm_Terminate()

    Set FormularioAnterior = Nothing
    
End Sub
