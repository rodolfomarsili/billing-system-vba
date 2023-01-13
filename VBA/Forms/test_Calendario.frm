VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} test_Calendario 
   Caption         =   "Credito"
   ClientHeight    =   3780
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   2868
   OleObjectBlob   =   "test_Calendario.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "test_Calendario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Click()
    Call Calendario.InicializarCalendario
End Sub
