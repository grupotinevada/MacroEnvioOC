Attribute VB_Name = "Módulo1"
Sub EnviarCorreo(Email As String, sArchivo As String, sAsunto As String, sCuerpo As String, Optional CC As String = "", Optional BCC As String = "")
    Dim OutApp As Object
    Dim OutMail As Object
    
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0) ' 0 = olMailItem

    With OutMail
        .To = Email ' Dirección de correo del destinatario
        .CC = CC ' Dirección de correo en copia (opcional)
        .BCC = BCC ' Dirección de correo en copia oculta (opcional)
        .Subject = sAsunto ' Asunto del correo tomado del Excel
        .Body = sCuerpo ' Cuerpo del mensaje tomado del Excel

        ' Adjuntar solo un archivo
        If sArchivo <> "" Then
            If Dir(sArchivo) <> "" Then ' Verifica si el archivo existe
                .Attachments.Add sArchivo
            Else
                MsgBox "No se encontró el archivo: " & sArchivo, vbExclamation, "Error de archivo"
            End If
        End If

        .Send ' Enviar el correo
    End With

    Set OutMail = Nothing
    Set OutApp = Nothing
End Sub

Sub EnvioOC()
    Dim RutaArchivo As String
    Dim Libro As Workbook
    Dim Hoja As Worksheet
    Dim UltimaFila As Long
    Dim i As Long
    Dim sNomArchivo As String, sEmail As String
    Dim sAsunto As String, sCuerpoPersonalizado As String, sCuerpoGenerico As String, sCuerpoFinal As String
    Dim sCC As String, sBCC As String
    Dim DelaySeconds As Integer
    Dim frm As FormularioProgreso ' Referencia al UserForm de progreso

    ' Definir el retraso en segundos entre el envío de cada correo
    DelaySeconds = 5

    ' Solicitar al usuario la ruta del archivo Excel
    RutaArchivo = Application.GetOpenFilename(FileFilter:="Archivos de Excel (*.xls*),*.xls*", Title:="Seleccionar archivo")

    If RutaArchivo <> "" Then
        ' Abrir el archivo Excel
        Set Libro = Workbooks.Open(RutaArchivo)

        ' Seleccionar la primera hoja del libro
        Set Hoja = Libro.Worksheets(1)

        ' Encontrar la última fila con datos en la primera columna (A)
        UltimaFila = Hoja.Cells(Hoja.Rows.Count, 1).End(xlUp).Row

        ' Inicializar el UserForm de progreso
        Set frm = New FormularioProgreso
        frm.Inicializar UltimaFila - 1 ' -1 para restar la fila de encabezado
        frm.Show vbModeless ' Mostrar el formulario

        ' Recorrer las filas y leer los datos de las columnas
        For i = 2 To UltimaFila
            sAsunto = Hoja.Cells(i, 1).Value ' Asunto (columna A)
            sCuerpoPersonalizado = Hoja.Cells(i, 2).Value ' Parte Personalizada (columna B)
            sCuerpoGenerico = Hoja.Cells(i, 3).Value ' Parte Genérica (columna C)
            sCuerpoFinal = sCuerpoPersonalizado & vbCrLf & vbCrLf & sCuerpoGenerico ' Concatenar las partes del cuerpo
            sNomArchivo = Hoja.Cells(i, 4).Value ' Documento(s) (columna D)
            sEmail = Hoja.Cells(i, 5).Value ' Correo (columna E)
            sCC = Hoja.Cells(i, 6).Value ' Con copia (columna F)
            sBCC = Hoja.Cells(i, 7).Value ' Con copia oculta (columna G)

            If sEmail <> "" Then
                ' Llamar al subproceso EnviarCorreo con los valores del Excel
                Call EnviarCorreo(sEmail, sNomArchivo, sAsunto, sCuerpoFinal, sCC, sBCC)
                
                ' Actualizar la barra de progreso
                frm.ActualizarProgreso i - 1 ' -1 para ajustar el conteo
                
                ' Esperar el tiempo definido antes de enviar el siguiente correo
                Application.Wait (Now + TimeValue("00:00:" & DelaySeconds))
            End If
        Next i
        
        ' Cerrar el UserForm de progreso
        frm.Cerrar
        
        MsgBox "Correos enviados exitosamente", vbOKOnly + vbInformation, "Proceso terminado"

        ' Cerrar el libro sin guardar cambios
        ' Libro.Close SaveChanges:=False
        
    End If
End Sub

