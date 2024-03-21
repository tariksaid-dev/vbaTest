Sub GenerarInformeConID()
    Dim wbOriginalData As Workbook, wbNuevo As Workbook
    Dim wsOriginalData As Worksheet, wsTemplate As Worksheet, wsPreguntasRespuestas As Worksheet, wsResumen As Worksheet
    Dim rutaBase As String
    Dim rutaOriginalData As String
    Dim i As Long, destinoFila As Long, filaID As Long
    Dim idUsuario As Variant
    Dim encontrado As Range
    
    ' Solicitar al usuario que introduzca el ID
    idUsuario = Application.InputBox("Introduce el ID para buscar:", "Buscar por ID", Type:=1)

    ' Verificar si el usuario presionó Cancelar en el InputBox
    If idUsuario = False Then
        MsgBox "Operación cancelada por el usuario.", vbExclamation
        Exit Sub
    End If

    rutaBase = ThisWorkbook.Path

    ' Ruta completa del archivo de producción
    rutaOriginalData = rutaBase & "\Cuestionario SQL (Producción).xlsx"
    
    ' Intentar abrir el libro de producción
    Set wbOriginalData = Workbooks.Open(rutaOriginalData)
    Set wsOriginalData = wbOriginalData.Sheets(1)

    ' Buscar el ID en la primera columna
    Set encontrado = wsOriginalData.Columns(1).Find(What:=idUsuario, LookIn:=xlValues, LookAt:=xlWhole)

    ' Verificar si se encontró el ID
    If Not encontrado Is Nothing Then
        filaID = encontrado.Row
        ' Debug.Print "El valor de filaID es: " & filaID
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

    RellenarHojaPreguntasRespuestas wsOriginalData, wsTemplate, wsPreguntasRespuestas, filaID

    ' ##### Acciones en la hoja de Resumen #####
    
    CrearGraficoApilado wsResumen
    CrearGraficoBarrasKiller wsResumen

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

    cabeceraResumen = Array("Dificultad", "Total preguntas", "Aciertos", "Errores", "% aciertos", "Total killer", "Killer answers", "% Killer answers")
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
            .Cells(i + filaActualResumen, 6).Formula = "=IF(C" & i + filaActualResumen & "<>0, D" & i + filaActualResumen & "/C" & i + filaActualResumen & ",0)"

            .Cells(i + filaActualResumen, 6).NumberFormat = "0%"

            ' Total killer answers
            .Cells(i + filaActualResumen, 7).Value = CalcularTotalKiller(wsTemplate, niveles(i))

            ' Killer answers
            .Cells(i + filaActualResumen, 8).Formula = "=SUMIFS('Preguntas y respuestas'!$F:$F, 'Preguntas y respuestas'!$D:$D, Resumen!B" & i + filaActualResumen & ", 'Preguntas y respuestas'!$D:$D, Resumen!B" & i + filaActualResumen & ")"

            ' % Killer answers
            .Cells(i + filaActualResumen, 9).Formula = "=IF(G" & i + filaActualResumen & "<>0, H" & i + filaActualResumen & "/G" & i + filaActualResumen & ",0)"

            .Cells(i + filaActualResumen, 9).NumberFormat = "0%"
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
        .Cells(filaActualResumen, 8).Formula = "=SUM(H26:H28)"
        ' % Killer answers
        .Cells(filaActualResumen, 9).Formula = "=IF(G29<>0,H29/G29,0)"
        .Cells(filaActualResumen, 9).NumberFormat = "0%"
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

    Dim rutaArchivoFinal As String

    rutaArchivoFinal = GuardarArchivoResultado(wbNuevo, rutaBase, nombreUsuario)

    ' Cerrar el libro de producción sin guardar cambios
    wbOriginalData.Close SaveChanges:=False

    MsgBox "Informe generado con éxito en: " & rutaArchivoFinal, vbInformation
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

Function GuardarArchivoResultado(wbNuevo As Workbook, rutaBase As String, nombreUsuario As String) As String

    Dim nombreArchivo As String
    nombreArchivo = "Resultados SQL - " & nombreUsuario & ".xlsx"

    ' Comprobamos si hay que cerrar el Workbook en caso de que estuviese abierto antes para evitar errores
    For Each wb In Workbooks
        If wb.Name = nombreArchivo Then
            wb.Close SaveChanges:=False
            Exit For
        End If
    Next wb
    
    Dim rutaArchivoFinal As String
    rutaArchivoFinal = rutaBase & "\" & nombreArchivo

    wbNuevo.SaveAs rutaArchivoFinal
    
    GuardarArchivoResultado = rutaArchivoFinal

End Function

Function RellenarHojaPreguntasRespuestas(wsOriginalData As Worksheet, wsTemplate As Worksheet, wsPreguntasRespuestas As Worksheet, filaID As Long)

    ' Cabeceras
    wsPreguntasRespuestas.Cells(1, 1).Value = "question"
    wsPreguntasRespuestas.Cells(1, 2).Value = "user answer"
    wsPreguntasRespuestas.Cells(1, 3).Value = "id"
    wsPreguntasRespuestas.Cells(1, 4).Value = "Level"
    wsPreguntasRespuestas.Cells(1, 5).Value = "Points"
    wsPreguntasRespuestas.Cells(1, 6).Value = "Killer answer"
    
    wsPreguntasRespuestas.Columns("C").NumberFormat = "@"

    Dim destinoFila As Long
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

End Function

Sub CrearGraficoBarrasKiller(wsResumen As Worksheet)
    ' Asumiendo que wsResumen es la hoja de trabajo pasada como parámetro
    
    ' Eliminar cualquier gráfico existente que se llame "GraficoKiller"
    Dim graficoExistente As ChartObject
    For Each graficoExistente In wsResumen.ChartObjects
        If graficoExistente.Name = "GraficoKiller" Then
            graficoExistente.Delete
        End If
    Next graficoExistente

    ' Añadir el gráfico de columnas y configurarlo
    Dim chtObj As ChartObject
    Set chtObj = wsResumen.ChartObjects.Add(Left:=400, Width:=375, Top:=10, Height:=300)
    chtObj.Name = "GraficoKiller" ' Asignar un nombre al gráfico para poder referenciarlo luego
    With chtObj.Chart
        .ChartType = xlColumnClustered ' Gráfico de columnas agrupadas
        
        ' Establecer rango de datos para las "Killer Answers"
        .SeriesCollection.NewSeries
        .SeriesCollection(1).Name = "Killer Answers"
        .SeriesCollection(1).Values = wsResumen.Range("H26:H28")
        .SeriesCollection(1).XValues = wsResumen.Range("B26:B28")

        ' Añadir las etiquetas de datos
        .SeriesCollection(1).ApplyDataLabels
        .SeriesCollection(1).DataLabels.ShowValue = True

        ' Añadir título y personalizar el gráfico
        .HasTitle = True
        .ChartTitle.Text = "Número de Killer Answers"
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.Text = "Dificultad"
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Text = "Cantidad"
        
        ' Personalizar colores de las series
        .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(79, 129, 189) ' Azul para "Killer Answers"
        
        .Legend.Position = xlLegendPositionTop ' Mover leyenda arriba
    End With

    ' Ajustar tamaño de la fuente del título del gráfico
    With chtObj.Chart.ChartTitle.Font
        .Size = 14
        .Bold = True
    End With
End Sub

Function CalcularTotalKiller(wsTemplate As Worksheet, nivel As Variant) As Integer
    ' Elimina On Error GoTo ErrorHandler para poder depurar
    Dim totalKiller As Integer
    totalKiller = 0

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim lastRow As Long
    lastRow = wsTemplate.Cells(wsTemplate.Rows.Count, "D").End(xlUp).Row

    Dim questionCode As String
    Dim questionLevel As String
    Dim isKiller As Integer

    For i = 2 To lastRow
        questionCode = wsTemplate.Cells(i, "C").Value
        Debug.Print "Question Code: "; questionCode ' Imprime el código de la pregunta
        If InStr(questionCode, ".") > 0 Then
            questionCode = Split(questionCode, ".")(0)
        End If

        questionLevel = wsTemplate.Cells(i, "D").Value
        Debug.Print "Question Level: "; questionLevel ' Imprime el nivel de la pregunta
        isKiller = Val(wsTemplate.Cells(i, "F").Value)

        Debug.Print "Is Killer: "; isKiller ' Imprime si es killer

        If IsError(isKiller) Then
            MsgBox "La celda F" & i & " no contiene un número válido."
            Exit Function
        End If

        If questionLevel = nivel And isKiller = 1 And Not dict.Exists(questionCode) Then
            dict.Add questionCode, True
            totalKiller = totalKiller + 1
        End If
    Next i

    CalcularTotalKiller = totalKiller
    Exit Function
End Function
