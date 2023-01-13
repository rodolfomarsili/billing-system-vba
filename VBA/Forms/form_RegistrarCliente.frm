VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form_RegistrarCliente 
   Caption         =   "Registrar Cliente"
   ClientHeight    =   4395
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   10212
   OleObjectBlob   =   "form_RegistrarCliente.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "form_RegistrarCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ComboBox_TipoID_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If Not (KeyAscii = 69 Or KeyAscii = 71 Or KeyAscii = 74 Or KeyAscii = 80 Or KeyAscii = 86 Or KeyAscii = 101 Or KeyAscii = 103 Or KeyAscii = 106 Or KeyAscii = 112 Or KeyAscii = 118) Then KeyAscii = 0
End Sub

Private Sub CommandButton_Cancelar_Click()
Unload Me
End Sub

Private Sub CommandButton_Registrar_Click()

Dim ID As String
Dim IDResponsable As String
Dim ComentarioOculto As String
Dim ColumnaIDClientes As Integer
Dim ComprobarExistencia As Integer
Dim LargoID As Integer
Dim LargoTelefono As Integer
Dim IngresarCliente As Integer
Const NuevaFila As Byte = 2
Dim Fecha As Date
  
    Fecha = TextBox_Dia & "/" & TextBox_Mes & "/" & TextBox_Ano


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Verificaciones'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    'Verificacion de casilla de letra de identificacion vacia
    If (ComboBox_TipoID = Empty) Then
        MsgBox "Selecciona el tipo de Identificacion", , "Registrar Cliente"
        Exit Sub
    End If
    'Verificacion de casilla de identificacion vacia
    If (TextBox_ID = Empty) Then
        MsgBox "Ingresa el numero de Identificacion", , "Registrar Cliente"
        Exit Sub
    End If
    'Verificacion de las demas casillas vacias
    If (TextBox_Nombre = Empty Or TextBox_Direccion = Empty Or TextBox_Telefono = Empty) Then
        MsgBox "Debes rellenar todos los campos", , "Registrar Cliente"
        Exit Sub
    End If
    
    'Verificacion de largo de los numeros de identificacion
    LargoID = Len(TextBox_ID.Text)

    If ((ComboBox_TipoID = "V" Or ComboBox_TipoID = "E") And LargoID <> 8) Then
        MsgBox "Numero de identificacion incorrecto" + Chr(13) + Chr(13) + "Si el numero de identificacion es menor a 8" + Chr(13) + "rellena con 0 (Ceros) delante del mismo", , "Ingresar Cliente"
        Exit Sub
    End If

    If ((ComboBox_TipoID = "J" Or ComboBox_TipoID = "G") And LargoID <> 9) Then
        MsgBox "Numero de identificacion incorrecto", , "Registrar Cliente"
        Exit Sub
    End If
    
    'Verificacion de largo del numero de telefono
    LargoTelefono = Len(TextBox_Telefono)
    
    If (LargoTelefono <> 12) Then
        MsgBox "Ingresa un numero de telefono valido", , "Registrar Cliente"
        Exit Sub
    End If
        
    
    Inicializar
    
    IDResponsable = HojaGestion.Range("B3")
    
    'Concatenacion de la letra del ID con el numero
    ID = ComboBox_TipoID.Text & "-" & TextBox_ID.Text
    
    'Verificacion de existencia de cliente
    ComprobarExistencia = ObtenerFila(HojaClientes, ID, ColumnaIDCliente)
    If ComprobarExistencia <> 0 Then
        MsgBox "El cliente ya esta registrado", , "Registrar Cliente"
        Exit Sub
    End If
    
    'Ultima comprobacion antes de registrar el cliente
    IngresarCliente = MsgBox("¿Seguro que deseas ingresar este registro?", vbYesNo + vbExclamation, "Ingresar Cliente")
    If IngresarCliente = vbNo Then Exit Sub
    
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Registro del cliente'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        'Ingreso de nueva hoja en libro de clientes
        Application.Windows("Clientes.xlsm").Visible = True
        LibroClientes.Activate
        HojaBaseClientes.Copy After:=Sheets("Base")
        LibroClientes.Sheets("Base (2)").Name = ID
        LibroClientes.Sheets("Inicio").Select
        Application.Windows("Clientes.xlsm").Visible = False


        'Se inserta una fila entera en en principio de la tabla de clientes
        HojaClientes.Range(NuevaFila & ":" & NuevaFila).EntireRow.Insert
        'LLenado del registro del nuevo cliente
        HojaClientes.Cells(NuevaFila, ColumnaIDCliente) = ID
        HojaClientes.Cells(NuevaFila, ColumnaNombreCliente) = TextBox_Nombre
        HojaClientes.Cells(NuevaFila, ColumnaDireccionCliente) = TextBox_Direccion
        HojaClientes.Cells(NuevaFila, ColumnaTelefonoCliente) = TextBox_Telefono
        HojaClientes.Cells(NuevaFila, ColumnaLimiteCreditoCliente) = Val(TextBox_LimiteCreditoCliente)
        
        HojaClientes.Cells(NuevaFila, ColumnaSaldoCreditoCliente) = 0
        HojaClientes.Cells(NuevaFila, ColumnaSaldoConsignacionCliente) = "='[Clientes.xlsm]" & ID & "'!$J$1" 'Aqui se coloca autamaticamente el valor de las consignaciones en la hoja de clientes de la base de datos
        HojaClientes.Cells(NuevaFila, ColumnaPrestamoUSDCliente) = 0
        HojaClientes.Cells(NuevaFila, ColumnaPrestamoBRLCliente) = 0
        HojaClientes.Cells(NuevaFila, ColumnaPrestamoVESCliente) = 0
        
        HojaClientes.Cells(NuevaFila, ColumnaCreditoCliente) = CheckBox_Credito
        HojaClientes.Cells(NuevaFila, ColumnaConsignacionCliente) = CheckBox_Consignacion
        
        'Formato de cada columna del registro del nuevo cliente
        HojaClientes.Cells(NuevaFila, ColumnaIDCliente).NumberFormat = "General" 'General
        HojaClientes.Cells(NuevaFila, ColumnaNombreCliente).NumberFormat = "General" 'General
        HojaClientes.Cells(NuevaFila, ColumnaDireccionCliente).NumberFormat = "General" 'General
        HojaClientes.Cells(NuevaFila, ColumnaTelefonoCliente).NumberFormat = "General" 'General
        HojaClientes.Cells(NuevaFila, ColumnaPrestamoUSDCliente).NumberFormat = "_-[$$] * #,##0.00_-;[$$] * -#,##0.00_-;_-[$$] * ""-""??_-;_-@_-" ''Formato $
        HojaClientes.Cells(NuevaFila, ColumnaPrestamoBRLCliente).NumberFormat = "_-[$R$] * #,##0.00_-;[$R$] * -#,##0.00_-;_-[$R$] * ""-""??_-;_-@_-" ''Formato R$
        HojaClientes.Cells(NuevaFila, ColumnaPrestamoVESCliente).NumberFormat = "_-[$Bs] * #,##0.00_-;[$Bs] * -#,##0.00_-;_-[$Bs] * ""-""??_-;_-@_-" ''Formato Bs
        HojaClientes.Cells(NuevaFila, ColumnaDeudaTotalCliente).NumberFormat = "_-[$$] * #,##0.00_-;[$$] * -#,##0.00_-;_-[$$] * ""-""??_-;_-@_-" ''Formato $
        

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Limpieza del Formulario'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ComentarioOculto = ""
        ComentarioOculto = "[ID: " & ID & "]"
        ComentarioOculto = ComentarioOculto + Chr(13) + "[Nombre: " & TextBox_Nombre.Text & "]"
        ComentarioOculto = ComentarioOculto + Chr(13) + "[Telefono: " & TextBox_Telefono.Text & "]"
        ComentarioOculto = ComentarioOculto + Chr(13) + "[Direccion: " & TextBox_Direccion.Text & "]"
        ComentarioOculto = ComentarioOculto + Chr(13) + "[Limite de credito: " & TextBox_LimiteCreditoCliente.Text & "]"
        ComentarioOculto = ComentarioOculto + Chr(13) + "[Creditos permitidos: " & CheckBox_Credito & "]"
        ComentarioOculto = ComentarioOculto + Chr(13) + "[Consignaciones permitidas: " & CheckBox_Consignacion & "]"
        
        'Incluir en el historial
        IncluirEnHistorial Frame_Correlativo, Label_CorrelativoPrefijo, Fecha, , , , , ComentarioOculto, , IDResponsable
        
        'Actualizar correlativo
        ActualizarCorrelativo Frame_Correlativo, Label_CorrelativoPrefijo
        
        'Actualizar correlativo en pantalla
        ActualizarCorrelativoEnPantalla Frame_Correlativo, Label_CorrelativoPrefijo
        
        'Reordenamiento de lista de clientes
        ReordenarClientes
        
        'Limpieza de los campos del formulario
        ComboBox_TipoID = Empty
        TextBox_ID = Empty
        TextBox_Nombre = Empty
        TextBox_Direccion = Empty
        TextBox_Telefono = Empty
        TextBox_LimiteCreditoCliente = 0
        CheckBox_Credito.Value = False
        CheckBox_Consignacion.Value = False
        
        ComboBox_TipoID.SetFocus
    
        
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
   
    MsgBox "Cliente registrado existosamente", , "Registrar Cliente"
    
    
End Sub


Private Sub TextBox_ID_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    'No se puede escribir en la casilla del numero de identificacion, si no se selecciona una letra.
    If (ComboBox_TipoID = Empty) Then
        MsgBox "Selecciona el tipo de Identificacion", , "Registrar Cliente"
        KeyAscii = 0
    End If
    'Cuando el numero de caracteres es igual a 8 en la casilla del numero de identificacion, éste ya no recibe mas caracteres.
    If (ComboBox_TipoID = "V" Or ComboBox_TipoID = "E") Then
        If Not (KeyAscii >= 48 And KeyAscii <= 57) Then KeyAscii = 0
        If Len(TextBox_ID) = 8 Then KeyAscii = 0
    End If
    
    If (ComboBox_TipoID = "J" Or ComboBox_TipoID = "G") Then
        If Not (KeyAscii >= 48 And KeyAscii <= 57) Then KeyAscii = 0
        If Len(TextBox_ID) = 9 Then KeyAscii = 0
    End If
    
End Sub

Private Sub TextBox_Telefono_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    'Cuando el numero de caracteres es igual a  en la casilla del numero de telefono, éste ya no recibe mas caracteres.
    If Len(TextBox_Telefono) = 12 Then KeyAscii = 0
    'Agregado automatico de un guion despues de agregar el codigo de area del telefono
    If Len(TextBox_Telefono) = 4 Then TextBox_Telefono = TextBox_Telefono & "-"
    
    If Not (KeyAscii >= 48 And KeyAscii <= 57) Then
        KeyAscii = 0
    End If
    
End Sub


Private Sub UserForm_Initialize()
    
    TextBox_Dia = Day(Date)
    TextBox_Mes = Month(Date)
    TextBox_Ano = Year(Date)
    
    Label_CorrelativoPrefijo.Caption = "Registro"
    
    ComboBox_TipoID.AddItem ("V")
    ComboBox_TipoID.AddItem ("E")
    ComboBox_TipoID.AddItem ("J")
    ComboBox_TipoID.AddItem ("P")
    ComboBox_TipoID.AddItem ("G")
    
    TextBox_LimiteCreditoCliente = 0

    Set FormularioAnterior = Me
    
    'Actualizar correlativo en pantalla
    ActualizarCorrelativoEnPantalla Frame_Correlativo, Label_CorrelativoPrefijo
    
End Sub

Private Sub UserForm_Terminate()
    
    Set FormularioAnterior = Nothing
     
End Sub
