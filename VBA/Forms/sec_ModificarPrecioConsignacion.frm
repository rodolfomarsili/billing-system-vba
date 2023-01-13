VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} sec_ModificarPrecioConsignacion 
   Caption         =   "Modificar Precio"
   ClientHeight    =   1875
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   3540
   OleObjectBlob   =   "sec_ModificarPrecioConsignacion.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "sec_ModificarPrecioConsignacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton_Aceptar_Click()

Dim FilaDeCodigo As Integer
Dim Codigo As String
Dim HojaCliente As Worksheet
        
        If Not Val(TextBox_PrecioPorBulto.Text) = Val(Label_PrecioPorBulto.Caption) Then

                Inicializar
                
                On Error Resume Next
                
                Codigo = form_InventarioConsignaciones.ListBox_Consignaciones.List(form_InventarioConsignaciones.ListBox_Consignaciones.ListIndex, 0)
                Set HojaCliente = LibroClientes.Sheets(form_InventarioConsignaciones.TextBox_IDCliente.Text)
                FilaDeCodigo = ObtenerFila(HojaCliente, Codigo, ColumnaCodigoCliente)

        
                HojaCliente.Cells(FilaDeCodigo, ColumnaPrecioBultoCliente) = Val(TextBox_PrecioPorBulto.Text)
                
                Unload Me
        
        End If
        
End Sub

Private Sub CommandButton_Cancelar_Click()
        Unload Me
End Sub

Private Sub TextBox_PrecioPorBulto_Enter()
        With TextBox_PrecioPorBulto
                .SetFocus
                .SelStart = 0
                .SelLength = Len(.Text)
         End With
End Sub

Private Sub TextBox_PrecioPorBulto_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

Dim UbicacionPunto As Integer
    
        UbicacionPunto = InStr(TextBox_PrecioPorBulto.Text, ".")
    
    If (KeyAscii = 46 And UbicacionPunto > 0) Then
        KeyAscii = 0
    End If
    
    If Not (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 46) Then
        KeyAscii = 0
    End If
End Sub

