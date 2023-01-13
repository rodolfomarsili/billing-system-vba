VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form_Login 
   Caption         =   "Login"
   ClientHeight    =   3510
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   7116
   OleObjectBlob   =   "form_Login.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "form_Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton_Cancelar_Click()
Unload Me
End Sub

Private Sub CommandButton_Ingresar_Click()

Dim FilaUsuario As Byte
Dim ColumnaUsuario As Byte
Dim ColumnaContraseña As Byte
Dim ColumnaNombre As Byte
Dim ColumnaID As Byte

    Inicializar
    
    ColumnaUsuario = ObtenerColumnaDeTabla(HojaUsuarios, 1, "Usuario")
    ColumnaContraseña = ObtenerColumnaDeTabla(HojaUsuarios, 1, "Contraseña")
    ColumnaNombre = ObtenerColumnaDeTabla(HojaUsuarios, 1, "Nombre")
    ColumnaID = ObtenerColumnaDeTabla(HojaUsuarios, 1, "ID")
    
    FilaUsuario = ObtenerFila(HojaUsuarios, TextBox_Usuario, ColumnaUsuario)
    
    If TextBox_Usuario = Empty Then
        MsgBox "Ingresa tu usuario y contraseña", , "Login"
        Exit Sub
    End If
    
    If FilaUsuario = 0 Then
        MsgBox "Usuario no existe", , "Login"
        Exit Sub
    End If
    
    If TextBox_Contraseña = Empty Then
        MsgBox "Ingresa tu contraseña", , "Login"
        Exit Sub
    End If
    
    If Not TextBox_Contraseña.Text = HojaUsuarios.Cells(FilaUsuario, ColumnaContraseña) Then
        MsgBox "Contraseña Incorrecta", , "Login"
        TextBox_Contraseña.SetFocus
    Else
        HojaGestion.Range("B2") = HojaUsuarios.Cells(FilaUsuario, ColumnaNombre)
        HojaGestion.Range("B3") = HojaUsuarios.Cells(FilaUsuario, ColumnaID)
        HojaGestion.Range("B4") = HojaUsuarios.Cells(FilaUsuario, ColumnaUsuario)
        HojaGestion.Range("B5") = "Desbloqueado"
        
        'Mostrar Dashboard
        ThisWorkbook.Sheets("Dashboard").Visible = xlSheetVisible
        'Ocultar Inicio
        ThisWorkbook.Sheets("Inicio").Visible = xlSheetHidden
        
        ActualizarDashboard
        
        Unload Me
'        Load form_Menu
'        form_Menu.Show
        
    End If
        
End Sub

Private Sub UserForm_Initialize()
    Image_Login.Picture = LoadPicture(ThisWorkbook.Path & "\Resources\images\login.jpg")
End Sub

Private Sub UserForm_Terminate()
    
    Inicializar
    
    If HojaGestion.Range("B5") = "Bloqueado" Then
        Application.DisplayFormulaBar = True
        Application.DisplayStatusBar = True
        ActiveWindow.DisplayWorkbookTabs = True
        CerrarDependencias (False)
        Workbooks("Gestion.xlsm").Close (False)
    End If
    
End Sub
