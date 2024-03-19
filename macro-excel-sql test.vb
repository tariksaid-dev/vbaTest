Sub GenerarInformeConID()
    Dim wbOriginalData As Workbook, wbNuevo As Workbook
    Dim wsOriginalData As Worksheet, wsTemplate As Worksheet, wsPreguntasRespuestas As Worksheet, wsResumen As Worksheet
    Dim rutaOriginalData As String
    Dim i As Long, destinoFila As Long, filaID As Long
    Dim idUsuario As Variant
    Dim encontrado As Range

    ' Propiedades para la hoja de Preguntas y Respuestas
    Dim id As String
    Dim level As String
    Dim points As Integer
    Dim killer As Integer
    
    ' Propiedades para la hoja de Resumen
    
    ' Solicitar al usuario que introduzca el ID
    idUsuario = Application.InputBox("Introduce el ID para buscar:", "Buscar por ID", Type:=1)

    ' Verificar si el usuario presionó Cancelar en el InputBox
    If idUsuario = False Then
        MsgBox "Operación cancelada por el usuario.", vbExclamation
        Exit Sub
    End If

    ' Ruta completa del archivo de producción
    ' rutaOriginalData = "C:\Users\tarik.said\Desktop\vbaTest\Cuestionario SQL (Producción).xlsx"
    rutaOriginalData = "C:\Users\miriam.romero\Documents\macro-excel\vbaTest\Cuestionario SQL (Producción).xlsx"

    ' Intentar abrir el libro de producción
    Set wbOriginalData = Workbooks.Open(rutaOriginalData)
    Set wsOriginalData = wbOriginalData.Sheets(1)

    ' Buscar el ID en la primera columna
    Set encontrado = wsOriginalData.Columns(1).Find(What:=idUsuario, LookIn:=xlValues, LookAt:=xlWhole)

    ' Verificar si se encontró el ID
    If Not encontrado Is Nothing Then
        filaID = encontrado.Row
        Debug.Print "El valor de filaID es: " & filaID
    Else
        MsgBox "El ID que has introducido no existe.", vbExclamation
        wbOriginalData.Close SaveChanges:=False
        Exit Sub
    End If

    Set wsTemplate = ThisWorkbook.Sheets("answer template")

    ' Crear un nuevo libro para los resultados
    Set wbNuevo = Workbooks.Add
    
    Set wsPreguntasRespuestas = wbNuevo.Sheets(1)
    wbNuevo.Sheets(1).Name = "Preguntas y respuestas"
    
    Set wsResumen = wbNuevo.Worksheets.Add(After:=wbNuevo.Sheets(1))
    wsResumen.Name = "Resumen"

    ' ##### Acciones en la hoja de Preguntas y respuestas #####

    ' Cabeceras
    wsPreguntasRespuestas.Cells(1, 1).Value = "question"
    wsPreguntasRespuestas.Cells(1, 2).Value = "user answer"
    wsPreguntasRespuestas.Cells(1, 3).Value = "id"
    wsPreguntasRespuestas.Cells(1, 4).Value = "Level"
    wsPreguntasRespuestas.Cells(1, 5).Value = "Points"
    wsPreguntasRespuestas.Cells(1, 6).Value = "Killer answer"
    
    wsPreguntasRespuestas.Columns("C").NumberFormat = "@"

    destinoFila = 2 ' Iniciar en la segunda fila para los datos copiados tras las cabeceras

    ' Determinar la última columna con datos en la fila 1
    Dim ultimaColumna As Long
    ultimaColumna = wsOriginalData.Cells(1, Columns.Count).End(xlToLeft).Column

    ' Iterar sobre cada celda en la fila 1 hasta la última columna con datos
    For i = 1 To ultimaColumna
        ' Verificar si el valor de la celda comienza con "["
        If Left(wsOriginalData.Cells(1, i).Value, 1) = "[" Then
            ' Transponer los valores de la fila 1 y la fila encontrada a las columnas A y B en el nuevo libro
            wsPreguntasRespuestas.Cells(destinoFila, 1).Value = Trim(wsOriginalData.Cells(1, i).Value)
            wsPreguntasRespuestas.Cells(destinoFila, 2).Value = Trim(wsOriginalData.Cells(filaID, i).Value)

            ' Obtener ID
            id = Mid(wsOriginalData.Cells(1, i).Value, 2, 4) & "." & Mid(wsOriginalData.Cells(filaID, i).Value, 2, 2)
            wsPreguntasRespuestas.Cells(destinoFila, 3).Value = id
            
            ' Obtener Level
            level = WorksheetFunction.VLookup(id, wsTemplate.Range("C:F"), 2, False)
            wsPreguntasRespuestas.Cells(destinoFila, 4).Value = level
            
            ' Obtener Points
            points = WorksheetFunction.VLookup(id, wsTemplate.Range("C:F"), 3, False)
            wsPreguntasRespuestas.Cells(destinoFila, 5).Value = points
            
            ' Obtener Killer answer
            killer = WorksheetFunction.VLookup(id, wsTemplate.Range("C:F"), 4, False)
            wsPreguntasRespuestas.Cells(destinoFila, 6).Value = killer

            destinoFila = destinoFila + 1
        End If
    Next i
    
    ' ##### Acciones en la hoja de Resumen #####
    
    CrearGraficoApilado wsResumen

    filaActualResumen = filaActualResumen + 20 ' simulamos que se coloca el gráfico

    ' ####### Agregar el cuadro resumen #######
    filaActualResumen = filaActualResumen + 2 ' Fila en blanco de espacio entre el gráfico y el cuadro resumen

    wsResumen.Cells(filaActualResumen, 5).Value = "Puntuación final"
    wsResumen.Cells(filaActualResumen, 6).Value = "100%" ' ! HARDCODED

    filaActualResumen = filaActualResumen + 1
    wsResumen.Cells(filaActualResumen, 5).Value = "Nivel"
    wsResumen.Cells(filaActualResumen, 6).Value = "3 - Avanzado" ' ! HARDCODED

    ' ####### Rellenar tabla #######
    filaActualResumen = filaActualResumen + 2 ' Fila en blanco de espacio entre el cuadro resumen y la tabla

    Dim cabeceraResumen As Variant
    Dim niveles As Variant

    cabeceraResumen = Array("Dificultad", "Total preguntas", "Aciertos", "Errores", "% aciertos", "Killer answers", "% Killer answers")
    niveles = Array("Basic", "Intermediate", "Advanced")

    ' Insertar títulos de la cabecera
    
    For i = LBound(cabeceraResumen) To UBound(cabeceraResumen)
        wsResumen.Cells(filaActualResumen, i + 2).Value = cabeceraResumen(i)
    Next i
    
    filaActualResumen = filaActualResumen + 1
    
    ' Insertar niveles
    For i = LBound(niveles) To UBound(niveles)
        wsResumen.Cells(i + filaActualResumen, 2).Value = niveles(i)
        
        ' Insertar fórmulas
        With wsResumen
            ' Total preguntas por dificultad
            .Cells(i + filaActualResumen, 3).Formula = "=COUNTIFS('Preguntas y respuestas'!$D:$D, Resumen!B" & i + filaActualResumen & ")"
            ' Aciertos
            .Cells(i + filaActualResumen, 4).Formula = "=SUMIFS('Preguntas y respuestas'!$E:$E, 'Preguntas y respuestas'!$D:$D, Resumen!B" & i + filaActualResumen & ", 'Preguntas y respuestas'!$D:$D, Resumen!B" & i + filaActualResumen & ")"
            ' Errores (total preguntas - aciertos)
            .Cells(i + filaActualResumen, 5).Formula = "=C" & i + filaActualResumen & "-D" & i + filaActualResumen
            ' % aciertos
            .Cells(i + filaActualResumen, 6).Formula = "=IFERROR(C" & i + filaActualResumen & "/(C" & i + filaActualResumen & "+D" & i + filaActualResumen & "),0)"
            .Cells(i + filaActualResumen, 6).NumberFormat = "0%"
            ' Killer answers
            .Cells(i + filaActualResumen, 7).Formula = "=SUMIFS('Preguntas y respuestas'!$F:$F, 'Preguntas y respuestas'!$D:$D, Resumen!B" & i + filaActualResumen & ", 'Preguntas y respuestas'!$D:$D, Resumen!B" & i + filaActualResumen & ")"
            ' % Killer answers
            .Cells(i + filaActualResumen, 8).Formula = "=IFERROR(G" & i + filaActualResumen & "/C" & i + filaActualResumen & ",0)"
            .Cells(i + filaActualResumen, 8).NumberFormat = "0%"
        End With
    Next i

    ' Cálculo de totales
    filaActualResumen = filaActualResumen + (UBound(niveles) - LBound(niveles) + 1)
    
    wsResumen.Cells(filaActualResumen, 2).Value = "Total"
    
    With wsResumen
        ' Total preguntas
        .Cells(filaActualResumen, 3).Formula = "=SUM(C26:C28)"
        ' Aciertos
        .Cells(filaActualResumen, 4).Formula = "=SUM(D26:D28)"
        ' Errores (total preguntas - aciertos)
        .Cells(filaActualResumen, 5).Formula = "=C29-D29"
        ' % aciertos
        .Cells(filaActualResumen, 6).Formula = "=IF(C29<>0,D29/C29,0)"
        .Cells(filaActualResumen, 6).NumberFormat = "0%"
        ' Killer answers
        .Cells(filaActualResumen, 7).Formula = "=SUM(G26:G28)"
        ' % Killer answers
        .Cells(filaActualResumen, 8).Formula = "=IF(C29<>0,G29/C29,0)"
        .Cells(filaActualResumen, 8).NumberFormat = "0%"
    End With

    wsResumen.Columns.AutoFit

    ' ##### Guardar archivo resultante #####
    
    Dim nombreUsuario As String

    Dim celdaNombreApellidos As Range
    Set celdaNombreApellidos = wsOriginalData.Rows(1).Find("Nombre y Apellidos")

    If celdaNombreApellidos Is Nothing Then
        ' Si no se encuentra la columna, asigna el ID del usuario como nombre
        nombreUsuario = "ID-" & CStr(idUsuario)
    Else
        nombreUsuario = wsOriginalData.Cells(filaID, celdaNombreApellidos.Column).Value
    End If
    
    Dim nombreArchivo As String
    nombreArchivo = "Resultados SQL - " & nombreUsuario & ".xlsx"

    ' Comprobamos si hay que cerrar el Workbook en caso de que estuviese abierto antes para evitar errores
    For Each wb In Workbooks
        If wb.Name = nombreArchivo Then
            ' Cerrar el libro si está abierto
            wb.Close SaveChanges:=False
            Exit For ' Salir del bucle una vez que se cierra el libro
        End If
    Next wb

    ' Guardar el nuevo libro con los resultados
    wbNuevo.SaveAs ".\" & nombreArchivo

    ' Cerrar el libro de producción sin guardar cambios
    wbOriginalData.Close SaveChanges:=False

    MsgBox "Informe generado con éxito como: " & nombreArchivo, vbInformation
End Sub

Sub CrearGraficoApilado(wsResumen As Worksheet)
    ' Asumiendo que wsResumen es la hoja de trabajo pasada como parámetro
    
    ' Eliminar cualquier gráfico existente en wsResumen
    Dim chtObj As ChartObject
    For Each chtObj In wsResumen.ChartObjects
        chtObj.Delete
    Next chtObj

    ' Añadir el gráfico y configurarlo
    Set chtObj = wsResumen.ChartObjects.Add(Left:=10, Width:=375, Top:=10, Height:=300)
    With chtObj.Chart
        .ChartType = xlColumnStacked

        ' Establecer rango de datos y aplicar a las series del gráfico
        .SeriesCollection.NewSeries
        .SeriesCollection(1).Name = "Aciertos"
        .SeriesCollection(1).Values = wsResumen.Range("D26:D28")
        .SeriesCollection(1).XValues = wsResumen.Range("B26:B28")
        
        .SeriesCollection.NewSeries
        .SeriesCollection(2).Name = "Errores"
        .SeriesCollection(2).Values = wsResumen.Range("E26:E28")

        ' Añadir las etiquetas de datos
        .SeriesCollection(1).ApplyDataLabels
        .SeriesCollection(1).DataLabels.ShowValue = True
        .SeriesCollection(2).ApplyDataLabels
        .SeriesCollection(2).DataLabels.ShowValue = True

        ' Añadir título y personalizar el gráfico
        .HasTitle = True
        .ChartTitle.Text = "RESUMEN TEST SQL"
        .Axes(xlCategory, xlPrimary).HasTitle = False ' Ocultar título del eje X
        .Axes(xlValue, xlPrimary).HasTitle = False ' Ocultar título del eje Y
        .Legend.Position = xlLegendPositionTop ' Mover leyenda arriba

        ' Opciones de formato adicionales
        .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(146, 208, 80) ' Verde para aciertos
        .SeriesCollection(2).Format.Fill.ForeColor.RGB = RGB(192, 0, 0) ' Rojo para errores
    End With

    ' Ajustar tamaño de la fuente del título del gráfico
    With chtObj.Chart.ChartTitle.Font
        .Size = 14
        .Bold = True
    End With
End Sub

