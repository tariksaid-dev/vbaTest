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
    rutaOriginalData = rutaBase & "\Cuestionario SQL (Producción).xlsx"
    
    ' Intentar abrir el libro de producción
    Set wbOriginalData = Workbooks.Open(rutaOriginalData)
    Set wsOriginalData = wbOriginalData.Sheets(1)

    ' Buscar el ID en la primera columna
    Set encontrado = wsOriginalData.Columns(1).Find(What:=idUsuario, LookIn:=xlValues, LookAt:=xlWhole)

    ' Verificar si se encontró el ID
    If Not encontrado Is Nothing Then
        filaID = encontrado.Row
    Else
        MsgBox "El ID que has introducido no existe.", vbExclamation
        wbOriginalData.Close SaveChanges:=False
        Exit Sub
    End If

    Dim nombreUsuario As String
    Dim fecha As String
    Dim celdaFecha as Range
    Dim celdaNombreApellidos As Range
    Set celdaNombreApellidos = wsOriginalData.Rows(1).Find("Nombre y Apellidos")
    Set celdaFecha = wsOriginalData.Rows(1).Find("Hora de finalización")

    If celdaNombreApellidos Is Nothing Then
        ' Si no se encuentra la columna, asigna el ID del usuario como nombre
        nombreUsuario = "ID-" & CStr(idUsuario)
    Else
        nombreUsuario = wsOriginalData.Cells(filaID, celdaNombreApellidos.Column).Value
    End If

    If celdaFecha Is Nothing Then
        fecha = "Sin fecha"
    Else
        ' Primero obtenemos el valor de la celda
        Dim valorCelda As Variant
        valorCelda = wsOriginalData.Cells(filaID, celdaFecha.Column).Value

        ' Luego verificamos si el valorCelda no es Nothing y si es una cadena no vacía
        If Not IsError(valorCelda) And Not IsEmpty(valorCelda) Then
            ' Si es una cadena, aplicamos Split y obtenemos la primera parte
            Dim partes() As String
            partes = Split(CStr(valorCelda), " ")
            If UBound(partes) >= 0 Then fecha = partes(0) Else fecha = "Formato inesperado"
        Else
            fecha = "Sin fecha"
        End If

        Debug.Print fecha
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
    
    CrearGraficoApilado wsResumen, nombreUsuario, fecha
    CrearGraficoBarrasKiller wsResumen

    filaActualResumen = filaActualResumen + 20 ' simulamos que se coloca el gráfico

    ' ####### Agregar el cuadro resumen #######
    filaActualResumen = filaActualResumen + 2 ' Fila en blanco de espacio entre el gráfico y el cuadro resumen

    With wsResumen.Cells(filaActualResumen, 5)
      .Value = "Puntuación final"
      .Interior.Color = RGB(203, 213, 225)
    End With

    With wsResumen.Cells(filaActualResumen, 6)
      .Value = "-"
      .Interior.Color = RGB(15, 23, 42)
      .Font.Color = RGB(255, 255, 255)
      .Font.Bold = True
    End With

    filaActualResumen = filaActualResumen + 1

    With wsResumen.Cells(filaActualResumen, 5)
      .Value = "Nivel"
      .Interior.Color = RGB(203, 213, 225)
    End With

    With wsResumen.Cells(filaActualResumen, 6)
      .Value = "-"
      .Interior.Color = RGB(15, 23, 42)
      .Font.Color = RGB(255, 255, 255)
      .Font.Bold = True
    End With

    ' ####### Rellenar tabla #######
    filaActualResumen = filaActualResumen + 2 ' Fila en blanco de espacio entre el cuadro resumen y la tabla

    Dim cabeceraResumen As Variant
    Dim niveles As Variant

    cabeceraResumen = Array("Dificultad", "Total preguntas", "Aciertos", "Errores", "% aciertos", "Total killer", "Killer answers", "% Killer answers")
    niveles = Array("Basic", "Intermediate", "Advanced")

    ' Insertar títulos de la cabecera
    
    For i = LBound(cabeceraResumen) To UBound(cabeceraResumen)
      With wsResumen.Cells(filaActualResumen, i + 2)
        .Value = cabeceraResumen(i)
        .Interior.Color = RGB(0, 112, 192)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
      End With
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
    


    Dim rutaArchivoFinal As String

    rutaArchivoFinal = GuardarArchivoResultado(wbNuevo, rutaBase, nombreUsuario)

    ' Cerrar el libro de producción sin guardar cambios
    wbOriginalData.Close SaveChanges:=False

    MsgBox "Informe generado con éxito en: " & rutaArchivoFinal, vbInformation
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

Sub CrearGraficoApilado(wsResumen As Worksheet, nombreUsuario As String, fecha As String)
    ' Eliminar cualquier gráfico existente en wsResumen
    Dim chtObj As ChartObject
    For Each chtObj In wsResumen.ChartObjects
        chtObj.Delete
    Next chtObj

    ' Añadir el gráfico y configurarlo
    Set chtObj = wsResumen.ChartObjects.Add(Left:=10, Width:=300, Top:=10, Height:=250)
    With chtObj.Chart
        .ChartType = xlColumnStacked
        .HasTitle = True
        .ChartTitle.Text = "Test SQL - " & nombreUsuario & " (" & fecha & ")"
        .Axes(xlCategory, xlPrimary).HasTitle = False ' Ocultar título del eje X
        .Axes(xlValue, xlPrimary).HasTitle = False ' Ocultar título del eje Y
        .Legend.Position = xlLegendPositionTop ' Mover leyenda arriba
        .HasAxis(xlCategory, xlPrimary) = True
        .HasAxis(xlValue, xlPrimary) = False ' No mostrar el eje Y y sus medidas

        
        ' Quitar líneas de cuadrícula del eje X (aunque generalmente no hay, por si acaso)
        .Axes(xlCategory).MajorGridlines.Format.Line.Visible = msoFalse
        ' Quitar líneas de cuadrícula del eje Y
        .Axes(xlValue).MajorGridlines.Format.Line.Visible = msoFalse

        ' Configuración común para todas las series
        Dim serie As Series
        Set serie = .SeriesCollection.NewSeries
        With serie
            .Name = "Aciertos"
            .Values = wsResumen.Range("D26:D28")
            .XValues = wsResumen.Range("B26:B28")
            .ApplyDataLabels
            With .DataLabels
                .ShowValue = True
                .Font.Color = RGB(255, 255, 255) ' Texto blanco
                .Font.Bold = True
            End With
            .Format.Fill.ForeColor.RGB = RGB(112, 173, 71) ' Verde para aciertos
        End With

        Set serie = .SeriesCollection.NewSeries
        With serie
            .Name = "Errores"
            .Values = wsResumen.Range("E26:E28")
            .ApplyDataLabels
            With .DataLabels
                .ShowValue = True
                .Font.Color = RGB(255, 255, 255) ' Texto blanco
                .Font.Bold = True
            End With
            .Format.Fill.ForeColor.RGB = RGB(192, 0, 0) ' Rojo para errores
        End With

        ' Opciones de formato adicionales para el gráfico
        With .ChartTitle.Font
            .Size = 18
            .Bold = True
        End With
    End With
End Sub

Sub CrearGraficoBarrasKiller(wsResumen As Worksheet)
    ' Eliminar cualquier gráfico existente que se llame "GraficoKiller"
    Dim graficoExistente As ChartObject
    For Each graficoExistente In wsResumen.ChartObjects
        If graficoExistente.Name = "GraficoKiller" Then
            graficoExistente.Delete
        End If
    Next graficoExistente

    ' Añadir el gráfico de columnas y configurarlo
    Dim chtObj As ChartObject
    Set chtObj = wsResumen.ChartObjects.Add(Left:=400, Width:=300, Top:=10, Height:=250)
    chtObj.Name = "GraficoKiller"
    With chtObj.Chart
        .ChartType = xlColumnClustered ' Gráfico de columnas agrupadas
        .HasTitle = True
        .ChartTitle.Text = "Número de Killer Answers"
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Text = "Cantidad"
        .Legend.Position = xlLegendPositionTop ' Mover leyenda arriba
        ' Configuración para eliminar eje Y y líneas de cuadrícula
        .HasAxis(xlCategory, xlPrimary) = True
        .HasAxis(xlValue, xlPrimary) = True

        With .Axes(xlValue)
          .HasMajorGridlines = True
          .MajorUnit = 1
          .MinimumScale = 0
        End With

        ' .Axes(xlCategory).MajorGridlines.Format.Line.Visible = msoFalse
        ' Quitar líneas de cuadrícula del eje Y
        ' .Axes(xlValue).MajorGridlines.Format.Line.Visible = msoFalse

        
        ' Configurar serie de datos en una operación con With
        Dim serie As Series
        Set serie = .SeriesCollection.NewSeries

        With serie
            .Name = "Killer Answers"
            .Values = wsResumen.Range("H26:H28")
            .XValues = wsResumen.Range("B26:B28")
            .ApplyDataLabels
            With .DataLabels
                .ShowValue = True
                .Font.Color = RGB(255, 255, 255) ' Texto blanco
                .Font.Bold = True 'texto blanco para mejor contraste
            End With
            .Format.Fill.ForeColor.RGB = RGB(79, 129, 189) ' Azul para "Killer Answers"
        End With

        ' Ajustes adicionales del gráfico
        With .ChartTitle.Font
            .Size = 18
            .Bold = True
        End With
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
        ' Debug.Print "Question Code: "; questionCode ' Imprime el código de la pregunta
        If InStr(questionCode, ".") > 0 Then
            questionCode = Split(questionCode, ".")(0)
        End If

        questionLevel = wsTemplate.Cells(i, "D").Value
        ' Debug.Print "Question Level: "; questionLevel ' Imprime el nivel de la pregunta
        isKiller = Val(wsTemplate.Cells(i, "F").Value)

        ' Debug.Print "Is Killer: "; isKiller ' Imprime si es killer

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
