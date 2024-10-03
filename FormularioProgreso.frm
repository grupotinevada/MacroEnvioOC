VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormularioProgreso 
   Caption         =   "UserForm1"
   ClientHeight    =   1230
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12540
   OleObjectBlob   =   "FormularioProgreso.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "FormularioProgreso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    Me.EtiquetaProgreso.Width = 0 ' Inicializa la barra en 0%
    Me.MarcoProgreso.Caption = "0% Completado" ' Muestra 0% al inicio
    Me.Caption = "Progreso de Envío de Correos" ' Fijar un título que no cambie
End Sub

Public Sub Inicializar(total As Integer)
    Me.EtiquetaProgreso.Width = 0 ' Inicializa la barra en 0%
    Me.Tag = total ' Guarda el total para usar en el cálculo del progreso
    Me.MarcoProgreso.Caption = "0% Completado" ' Inicializa el marco con 0%
End Sub

Public Sub ActualizarProgreso(completado As Integer)
    Dim porcentaje As Double
    porcentaje = (completado / Me.Tag) * 100
    Me.EtiquetaProgreso.Width = Me.MarcoProgreso.Width * (completado / Me.Tag) ' Ajusta el tamaño del Label
    Me.MarcoProgreso.Caption = Round(porcentaje, 0) & "% Completado" ' Muestra el porcentaje en el marco
    DoEvents ' Permite que el formulario se actualice
End Sub

Public Sub Cerrar()
    Unload Me ' Cierra el formulario
End Sub

