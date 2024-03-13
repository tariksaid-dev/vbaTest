
Sub GenerarInforme()
    ' id As Long
    Dim wbOriginalData As Workbook, wbNuevo As Workbook
    Dim wsOriginalData As Worksheet, wsTemplate As Worksheet, wsPreguntasRespuestas As Worksheet, wsEstadisticas As Worksheet
    
    Dim rutaOriginalData As String
    Dim idUsuario As Variant
    Dim encontrado As Range
    
      ' Solicitar al usuario que introduzca el ID
    idUsuario = Application.InputBox("Introduce el ID para buscar:", "Buscar por ID", Type:=1)

    ' Verificar si el usuario presionó Cancelar en el InputBox
    If idUsuario = False Then
        MsgBox "Operación cancelada por el usuario.", vbExclamation
        Exit Sub
    End If

    ' 1. Cargar archivos excel
    rutaOriginalData = "C:\Users\tarik.said\Desktop\vbaTest\Cuestionario SQL (Producción).xlsx"
    
    Set wsTemplate = ThisWorkbook.Sheets("answer template")
    Set wbOriginalData = Workbooks.Open(rutaOriginalData)
    Set wsOriginalData = wbOriginalData.Sheets(1) 

    Set encontrado = wsOriginalData.Columns(1).Find(What:=idUsuario, LookIn:=xlValues, LookAt:=xlWhole)

      ' Verificar si se encontró el ID
    If Not encontrado Is Nothing Then
        filaID = encontrado.Row
    Else
        MsgBox "El ID que has introducido no existe.", vbExclamation
        wbOriginalData.Close SaveChanges:=False
        Exit Sub
    End If
    
    ' 2. Formatear los datos originales y obtener una tabla con las preguntas y respuestas
    Set wsPreguntasRespuestas = FormatearTablaPreguntasRespuestas(wsOriginalData)
    
    ' 3. Obtener la hoja de estadísticas
    Set wsEstadisticas = CalcularEstadisticas(wsPreguntasRespuestas)
    
    ' 4. Crear un nuevo libro de Excel
    Set wbNuevo = Application.Workbooks.Add
    
    ' 5. Añadir hoja de preguntas y respuestas al nuevo archivo
    wsPreguntasRespuestas.Copy Before:=wbNuevo.Sheets(1)
    wbNuevo.Sheets(1).Name = "Preguntas y respuestas"
    
    ' 6. Añadir hoja de estadísticas al nuevo archivo
    wsEstadisticas.Copy After:=wbNuevo.Sheets(1)
    wbNuevo.Sheets(2).Name = "Estadísticas"
    
    ' 7. Guardar el nuevo archivo en la misma ruta
    wbNuevo.SaveAs "C:\Users\tarik.said\Desktop\vbaTest\Resultados.xlsx"
    
    wbNuevo.Close
    wbOriginalData.Close SaveChanges:=False

    MsgBox "Archivo creado correctamente.", vbInformation

End Sub
Function FormatearTablaPreguntasRespuestas(wsOriginalData As Worksheet) As Worksheet
    Dim wbNuevo As Workbook
    Dim wsOriginal As Worksheet
    Dim wsReturn As Worksheet
    Dim wsTemplate As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim filaDestino As Integer
    Dim id As String
    Dim level As String
    Dim points As Integer
    Dim killer As Integer
    
    'Set wsOriginal = ThisWorkbook.Sheets("respuesta original")
    Set wsOriginal = wsOriginalData
    
    Set wsReturn = Worksheets.Add
    
    Set wsTemplate = ThisWorkbook.Sheets("answer template")
    
    ' Encabezados
    wsReturn.Cells(1, 1).Value = "question"
    wsReturn.Cells(1, 2).Value = "user answer"
    wsReturn.Cells(1, 3).Value = "id"
    wsReturn.Cells(1, 4).Value = "Level"
    wsReturn.Cells(1, 5).Value = "Points"
    wsReturn.Cells(1, 6).Value = "Killer answer"
    
    wsReturn.Columns("C").NumberFormat = "@"
    
    ' Buscar las preguntas y respuestas y copiarlas en la nueva hoja
    filaDestino = 2
    For i = 1 To wsOriginal.Cells(1, wsOriginal.Columns.Count).End(xlToLeft).Column
        If Left(wsOriginal.Cells(1, i).Value, 1) = "[" Then
            ' Copiar pregunta y respuesta
            wsReturn.Cells(filaDestino, 1).Value = wsOriginal.Cells(1, i).Value
            wsReturn.Cells(filaDestino, 2).Value = wsOriginal.Cells(2, i).Value
            
            ' Obtener ID
            id = Mid(wsOriginal.Cells(1, i).Value, 2, 4) & "." & Mid(wsOriginal.Cells(2, i).Value, 2, 2)
            wsReturn.Cells(filaDestino, 3).Value = CStr(id)
            
            ' Obtener Level
            level = WorksheetFunction.VLookup(id, wsTemplate.Range("C:F"), 2, False)
            wsReturn.Cells(filaDestino, 4).Value = level
            
            ' Obtener Points
            points = WorksheetFunction.VLookup(id, wsTemplate.Range("C:F"), 3, False)
            wsReturn.Cells(filaDestino, 5).Value = points
            
            ' Obtener Killer answer
            killer = WorksheetFunction.VLookup(id, wsTemplate.Range("C:F"), 4, False)
            wsReturn.Cells(filaDestino, 6).Value = killer
            
            filaDestino = filaDestino + 1
        End If
    Next i
    
    ' Ajustar anchos de las columnas
    'wsReturn.Columns("A:F").AutoFit
    
    Set FormatearTablaPreguntasRespuestas = wsReturn
        
End Function

Function CalcularEstadisticas(wsDatosFormateados As Worksheet) As Worksheet
    Dim wbNuevo As Workbook
    Dim wsOriginal As Worksheet
    Dim wsEstadisticas As Worksheet
    Dim niveles As Variant
    
    Dim lastRow As Long
    Dim i As Long
    
    ' Definir los niveles
    niveles = Array("Basic", "Intermediate", "Advanced", "Total")
    
    ' Definir la hoja original y la hoja de estadísticas
    Set wsOriginal = wsDatosFormateados
    
    Set wsEstadisticas = Worksheets.Add
    
    ' Insertar títulos
    wsEstadisticas.Cells(1, 2).Value = "Nivel"
    wsEstadisticas.Cells(1, 3).Value = "Total"
    wsEstadisticas.Cells(1, 4).Value = "Ok"
    wsEstadisticas.Cells(1, 5).Value = "% Ok"
    wsEstadisticas.Cells(1, 6).Value = "Killer answers"
    wsEstadisticas.Cells(1, 7).Value = "% Killer answers"
    
    ' Insertar niveles
    For i = LBound(niveles) To UBound(niveles)
        wsEstadisticas.Cells(i + 2, 1).Value = niveles(i)
    Next i
    
    ' Insertar fórmulas
    wsEstadisticas.Cells(2, 2).Formula = "" ' total en el nivel "Basic"
    wsEstadisticas.Cells(3, 2).Formula = "" ' total en el nivel "Intermediate"
    wsEstadisticas.Cells(4, 2).Formula = "" ' total en el nivel "Advanced"
    wsEstadisticas.Cells(5, 2).Formula = "" ' total general
    
    wsEstadisticas.Cells(2, 3).Formula = "" ' total de "Ok" en el nivel "Basic"
    wsEstadisticas.Cells(3, 3).Formula = "" ' total de "Ok" en el nivel "Intermediate"
    wsEstadisticas.Cells(4, 3).Formula = "" ' total de "Ok" en el nivel "Advanced"
    wsEstadisticas.Cells(5, 3).Formula = "" ' total de "Ok" general
    
    ' Insertar las demás fórmulas
    
    Set CalcularEstadisticas = wsEstadisticas
End Function

