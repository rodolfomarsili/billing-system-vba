Attribute VB_Name = "Dashboard"
Option Explicit

Sub ActualizarDashboard()

    'Actualiza los datos de las tablas dinamicas del dashboard
    ThisWorkbook.RefreshAll
    
    'Ajusta el ancho de las columnas para que los datos sean visibles
    Columns("C:Z").EntireColumn.AutoFit
    
End Sub

