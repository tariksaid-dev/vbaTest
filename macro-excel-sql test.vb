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
    
    ' Dibujar gráfico
    
    filaActualResumen = filaActualResumen + 10 ' simulamos que se coloca el gráfico

    ' ####### Agregar el cuadro resumen #######
    filaActualResumen = filaActualResumen + 2 ' Fila en blanco de espacio entre el gráfico y el cuadro resumen

    wsResumen.Cells(filaActualResumen, 5).Value = "Puntuación final"
    wsResumen.Cells(filaActualResumen, 6).Value = "100%" ' ! HARDCODED

    filaActualResumen = filaActualResumen + 1
    wsResumen.Cells(filaActualResumen, 5).Value = "Nivel"
    wsResumen.Cells(filaActualResumen, 6).Value = "3 - Avanzado" ' ! HARDCODED

    ' ####### Rellenar tabla #######
    filaActualResumen = filaActualResumen + 2 ' Fila en blanco de espacio entre el cuadro resumen y la tabla

    Dim niveles As Variant
    ' Definir los niveles
    niveles = Array("Basic", "Intermediate", "Advanced", "Total")

    ' Insertar títulos
    wsResumen.Cells(filaActualResumen, 2).Value = "Dificultad"
    wsResumen.Cells(filaActualResumen, 3).Value = "Total preguntas"
    wsResumen.Cells(filaActualResumen, 4).Value = "Aciertos"
    wsResumen.Cells(filaActualResumen, 5).Value = "Errores"
    wsResumen.Cells(filaActualResumen, 6).Value = "% aciertos"
    wsResumen.Cells(filaActualResumen, 7).Value = "Killer answers"
    wsResumen.Cells(filaActualResumen, 8).Value = "% Killer answers"
    
    filaActualResumen = filaActualResumen + 1
    
    ' Insertar niveles
    For i = LBound(niveles) To UBound(niveles)
        wsResumen.Cells(i + filaActualResumen, 2).Value = niveles(i)
    Next i

    wsResumen.Columns.AutoFit

    ' ##### Guardar archivo resultante #####
    
    ' wbNuevo.SaveAs "C:\Users\tarik.said\Desktop\vbaTest\Resultados.xlsx"
    wbNuevo.SaveAs "C:\Users\miriam.romero\Documents\macro-excel\vbaTest\Resultados SQL - Usuario Prueba.xlsx"

    ' Cerrar el libro de producción sin guardar cambios
    wbOriginalData.Close SaveChanges:=False

    MsgBox "Informe generado con éxito en: C:\Users\miriam.romero\Documents\macro-excel\vbaTest\Resultados.xlsx", vbInformation
End Sub


