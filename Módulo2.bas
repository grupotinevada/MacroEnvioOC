Attribute VB_Name = "Módulo2"
Sub GenerarMatriz()
    Dim wsSAP As Worksheet
    Dim wsBaseDatos As Worksheet
    Dim wsMatriz As Worksheet
    Dim wsRevision As Worksheet
    Dim ultimaFilaSAP As Long
    Dim ultimaFilaBaseDatos As Long
    Dim ultimaFilaMatriz As Long
    Dim ultimaFilaRevision As Long
    Dim i As Long, j As Long
    Dim proveedorSAP As String
    Dim proveedorBaseDatos As String
    Dim documentosSinCoincidencia As String
    Dim carpetaDocumentos As String
    Dim fDialog As FileDialog
    Dim ordenCompraSAP As String
    Dim archivoOC As String
    Dim filaDestino As Long
    Dim diagnostico As String
    Dim correoConCopia As String

    ' Inicializar las hojas
    Set wsSAP = ThisWorkbook.Sheets("oc SAP")
    Set wsBaseDatos = ThisWorkbook.Sheets("BASE DATOS")
    Set wsMatriz = ThisWorkbook.Sheets("MATRIZ")
    Set wsRevision = ThisWorkbook.Sheets("Revisión")
    
    ' Limpiar las hojas "MATRIZ" y "Revisión", pero mantener los encabezados
    wsMatriz.Rows("2:" & wsMatriz.Rows.Count).ClearContents
    wsRevision.Rows("2:" & wsRevision.Rows.Count).ClearContents

    ' Obtener las últimas filas con datos
    ultimaFilaSAP = wsSAP.Cells(wsSAP.Rows.Count, 1).End(xlUp).Row
    ultimaFilaBaseDatos = wsBaseDatos.Cells(wsBaseDatos.Rows.Count, 1).End(xlUp).Row

    ' Inicializar la variable de documentos sin coincidencia
    documentosSinCoincidencia = ""

    ' Mostrar el diálogo para seleccionar la carpeta donde están los documentos
    Set fDialog = Application.FileDialog(msoFileDialogFolderPicker)
    With fDialog
        .Title = "Seleccionar Carpeta de Documentos de Órdenes de Compra"
        If .Show = -1 Then ' Si el usuario selecciona una carpeta
            carpetaDocumentos = .SelectedItems(1)
        Else ' Si el usuario cancela la selección
            MsgBox "No se seleccionó ninguna carpeta. La macro será cancelada.", vbExclamation
            Exit Sub
        End If
    End With

    ' Recorrer la hoja "oc SAP" y buscar coincidencias en "BASE DATOS"
    For i = 2 To ultimaFilaSAP
        proveedorSAP = Trim(wsSAP.Cells(i, 3).Value)
        ordenCompraSAP = Trim(wsSAP.Cells(i, 1).Value) ' Suponiendo que la OC está en la columna 1

        ' Inicializar la variable diagnóstico
        diagnostico = ""

        ' Buscar coincidencia en la hoja "BASE DATOS"
        For j = 2 To ultimaFilaBaseDatos
            proveedorBaseDatos = Trim(wsBaseDatos.Cells(j, 3).Value)
            correoConCopia = wsBaseDatos.Cells(j, 6).Value ' Columna F, "Con Copia"

            If InStr(1, proveedorSAP, proveedorBaseDatos, vbTextCompare) > 0 Then
                ' Buscar el archivo correspondiente en la carpeta seleccionada
                archivoOC = ""

                ' Extraer los últimos 5 dígitos de la orden de compra
                If Len(ordenCompraSAP) >= 5 Then
                    ordenCompraSAP = Right(ordenCompraSAP, 5)
                End If

                ' Buscar cualquier archivo que comience con "OC " y el número de la OC
                archivoOC = Dir(carpetaDocumentos & "\OC " & ordenCompraSAP & "*.pdf")

                ' Determinar si la fila se copia a "MATRIZ" o a "Revisión"
                If archivoOC <> "" And wsBaseDatos.Cells(j, 4).Value <> "" Then
                    ' Coincidencia encontrada y archivo disponible
                    filaDestino = wsMatriz.Cells(wsMatriz.Rows.Count, 1).End(xlUp).Row + 1
                    wsMatriz.Cells(filaDestino, 1).Value = "Orden de Compra " & wsSAP.Cells(i, 1).Value
                    wsMatriz.Cells(filaDestino, 2).Value = "Estimado " & wsBaseDatos.Cells(j, 5).Value
                    wsMatriz.Cells(filaDestino, 3).Value = "Cuerpo genérico..."
                    wsMatriz.Cells(filaDestino, 4).Value = carpetaDocumentos & "\" & archivoOC
                    wsMatriz.Cells(filaDestino, 5).Value = wsBaseDatos.Cells(j, 4).Value
                    wsMatriz.Cells(filaDestino, 6).Value = correoConCopia ' Añadir "Con Copia" a la columna F de MATRIZ
                Else
                    ' Problema encontrado: archivo o correo faltante
                    filaDestino = wsRevision.Cells(wsRevision.Rows.Count, 1).End(xlUp).Row + 1
                    wsRevision.Cells(filaDestino, 1).Value = wsSAP.Cells(i, 1).Value ' OC
                    wsRevision.Cells(filaDestino, 2).Value = proveedorSAP ' Proveedor en SAP
                    wsRevision.Cells(filaDestino, 3).Value = wsBaseDatos.Cells(j, 4).Value ' Correo
                    wsRevision.Cells(filaDestino, 4).Value = archivoOC ' Archivo PDF
                    wsRevision.Cells(filaDestino, 5).Value = correoConCopia ' Añadir "Con Copia" a la columna F de Revisión
                    If archivoOC = "" Then
                        diagnostico = "No se encuentra el documento PDF"
                    End If
                    If wsBaseDatos.Cells(j, 4).Value = "" Then
                        diagnostico = diagnostico & IIf(diagnostico <> "", " y ", "") & "No se encuentra el correo del proveedor"
                    End If
                    wsRevision.Cells(filaDestino, 6).Value = diagnostico
                End If
                Exit For
            End If
        Next j

        ' Si no se encontró coincidencia, agregar a la lista de documentos sin coincidencia
        If j > ultimaFilaBaseDatos Then
            documentosSinCoincidencia = documentosSinCoincidencia & wsSAP.Cells(i, 1).Value & vbNewLine
            ' También agregar la OC sin coincidencia a "Revisión"
            ultimaFilaRevision = wsRevision.Cells(wsRevision.Rows.Count, 1).End(xlUp).Row + 1
            wsRevision.Cells(ultimaFilaRevision, 1).Value = wsSAP.Cells(i, 1).Value
            wsRevision.Cells(ultimaFilaRevision, 2).Value = proveedorSAP
            wsRevision.Cells(ultimaFilaRevision, 5).Value = "No encontrado"
            wsRevision.Cells(ultimaFilaRevision, 6).Value = "No se encontró coincidencia en la base de datos"
        End If
    Next i

    ' Mostrar mensaje con los documentos sin coincidencia
    If documentosSinCoincidencia <> "" Then
        MsgBox "Las siguientes órdenes de compra no tienen coincidencias en la base de datos:" & vbNewLine & documentosSinCoincidencia, vbExclamation
    Else
        MsgBox "Todas las órdenes de compra se han procesado correctamente.", vbInformation
    End If
End Sub


Sub ProcesarMatriz()
    Dim HojaMatriz As Worksheet
    Dim UltimaFila As Long
    Dim i As Long, j As Long
    Dim OCActual As String
    Dim Documentos As String
    Dim rng As Range
    Dim Fila As Variant ' Cambiado a Variant

    ' Definir la hoja MATRIZ
    Set HojaMatriz = ThisWorkbook.Sheets("MATRIZ")
    
    ' Encontrar la última fila con datos en la hoja MATRIZ
    UltimaFila = HojaMatriz.Cells(HojaMatriz.Rows.Count, 1).End(xlUp).Row
    
    ' Inicializar el rango de datos (excluyendo la fila de encabezado)
    Set rng = HojaMatriz.Range("A2:G" & UltimaFila)
    
    ' Recorrer cada fila de la matriz
    For Each Fila In rng.Rows
        OCActual = Fila.Cells(1, 1).Value ' Columna A, Doc. Compr.
        Documentos = Fila.Cells(1, 4).Value ' Columna D, Documento(s)
        
        ' Buscar filas con la misma OC
        For j = Fila.Row + 1 To UltimaFila
            If HojaMatriz.Cells(j, 1).Value = OCActual Then
                ' Concatenar los documentos
                Documentos = Documentos & ";" & HojaMatriz.Cells(j, 4).Value
                ' Eliminar la fila repetida
                HojaMatriz.Rows(j).Delete
                j = j - 1
                UltimaFila = UltimaFila - 1
            End If
        Next j
        
        ' Actualizar la celda de documentos con todos los documentos concatenados
        Fila.Cells(1, 4).Value = Documentos
    Next Fila
    
    MsgBox "Proceso completado. Filas duplicadas combinadas y eliminadas.", vbInformation
End Sub

