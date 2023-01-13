VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form_VisibilidadDependencias 
   Caption         =   "Visibilidad de las Dependencias"
   ClientHeight    =   2865
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   3756
   OleObjectBlob   =   "form_VisibilidadDependencias.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "form_VisibilidadDependencias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton_Aceptar_Click()
    
    If (ComboBox_Dependencia.Text = "Todas" And ComboBox_Visibilidad.Text = "Mostrar") Then
        MostrarDependencias
        Unload Me
        Exit Sub
    End If
    
    If (ComboBox_Dependencia.Text = "Todas" And ComboBox_Visibilidad.Text = "Ocultar") Then
        OcultarDependencias
        Unload Me
        Exit Sub
    End If
    
    If ComboBox_Visibilidad.Text = "Mostrar" Then
        Application.Windows(ComboBox_Dependencia.Text & ".xlsm").Visible = True
    Else
        Application.Windows(ComboBox_Dependencia.Text & ".xlsm").Visible = False
    End If
    
    Unload Me
    
End Sub

Private Sub CommandButton_Cancelar_Click()
    Unload Me
End Sub

Private Sub CommandButton_MostrarInterfaz_Click()
    Application.DisplayFormulaBar = True
    Application.DisplayStatusBar = True
    ActiveWindow.DisplayWorkbookTabs = True
    
    MostrarHojas
End Sub

Private Sub CommandButton_OcultarInterfaz_Click()
    Application.DisplayFormulaBar = False
    Application.DisplayStatusBar = False
    ActiveWindow.DisplayWorkbookTabs = False
    
    OcultarHojas
End Sub

Private Sub UserForm_Initialize()

    ComboBox_Dependencia.AddItem ("Todas")
    ComboBox_Dependencia.AddItem ("Base de datos")
    ComboBox_Dependencia.AddItem ("Clientes")
    ComboBox_Dependencia.AddItem ("Historial")
    
    ComboBox_Visibilidad.AddItem ("Mostrar")
    ComboBox_Visibilidad.AddItem ("Ocultar")
    
    ComboBox_Dependencia.Text = "Todas"
    ComboBox_Visibilidad.Text = "Mostrar"
End Sub
