VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} sec_Cantidad 
   Caption         =   "Cantidad "
   ClientHeight    =   1965
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   2460
   OleObjectBlob   =   "sec_Cantidad.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "sec_Cantidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton_Aceptar_Click()
    FormularioAnterior.Label_Cantidad_Auxiliar.Caption = TextBox_Cantidad
    Unload Me
End Sub

Private Sub TextBox_Cantidad_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    If KeyAscii = 27 Then Unload Me
    
    If Not (KeyAscii >= 48 And KeyAscii <= 57) Then
        KeyAscii = 0
    End If

End Sub

Private Sub UserForm_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then Unload Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If Not (FormularioAnterior.Label_Cantidad_Auxiliar.Caption = 0 Or FormularioAnterior.Label_Cantidad_Auxiliar.Caption = Empty) Then
    Else
        FormularioAnterior.Label_Cantidad_Auxiliar.Caption = 0
    End If
    
End Sub
