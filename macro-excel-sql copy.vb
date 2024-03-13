' Sub GenerarInforme()
'   Dim wbOriginalData As Workbook, wbTemplate As Workbook, wbNuevo As Workbook
'   Dim wsOriginalData As Worksheet, wsTemplate As Worksheet
'   Dim rutaOriginalData As String, rutaTemplate As String

'   ' Rutas de los archivos
'   rutaOriginalData = ".\Cuestionario SQL (Producción).xlsx"
'   rutaTemplate = ".\answer template.xlsx"

'   Abrir el libro de producción y el de plantilla
'   Set wbOriginalData = Workbooks.Open(rutaOriginalData)
'   Set wbTemplate = Workbooks.Open(rutaTemplate)

'   ' Suponiendo que los datos y las fórmulas están en la primera hoja de cada libro
'   Set wsOriginalData = wbOriginalData.Sheets(1)
'   Set wsTemplate = wbTemplate.Sheets(1)
  
'   ' Crear un nuevo libro para los resultados
'   Set wbNuevo = Workbooks.Add

'   ' Copiar preguntas y respuestas al libro nuevo
'   wsOriginalData.Range("A:B").Copy Destination:=wbNuevo.Sheets(1).Range("A1")

'   ' Aquí necesitarás adaptar cómo aplicar las fórmulas del libro de plantilla.
'   ' Esto puede involucrar copiar las fórmulas directamente y ajustar referencias de celda, o reconstruir las fórmulas en VBA y aplicarlas.
  
'   ' Guardar el libro de resultados
'   wbNuevo.SaveAs ".\Resultados.xlsx"

'   ' Cerrar libros sin guardar cambios
'   wbOriginalData.Close SaveChanges:=False
'   wbTemplate.Close SaveChanges:=False

'   MsgBox "Informe generado con éxito.", vbInformation
' End Sub
Sub GenerarInforme()
    Dim wbOriginalData As Workbook, wbNuevo As Workbook
    Dim wsOriginalData As Worksheet, wsNuevo As Worksheet
    Dim rutaOriginalData As String
    Dim i As Long
    Dim destinoFila As Long
    Dim filaSeleccionada As Long

    ' Solicitar al usuario que introduzca el número de fila
    filaSeleccionada = Application.InputBox("Introduce el número de la fila adicional para copiar:", "Seleccionar Fila", Type:=1)

    ' Verificar si el usuario presionó Cancelar en el InputBox
    If filaSeleccionada = False Then
        MsgBox "Operación cancelada por el usuario.", vbExclamation
        Exit Sub
    End If

    ' Ruta completa del archivo de producción
    rutaOriginalData = "C:\Users\tarik.said\Desktop\vbaTest\Cuestionario SQL (Producción).xlsx"

    ' Intentar abrir el libro de producción
    Set wbOriginalData = Workbooks.Open(rutaOriginalData)
    Set wsOriginalData = wbOriginalData.Sheets(1)
    
    ' Crear un nuevo libro para los resultados
    Set wbNuevo = Workbooks.Add
    Set wsNuevo = wbNuevo.Sheets(1)
    
    destinoFila = 1 ' Iniciar en la primera fila para los datos copiados
    
    ' Determinar la última columna con datos en la fila 1
    Dim ultimaColumna As Long
    ultimaColumna = wsOriginalData.Cells(1, Columns.Count).End(xlToLeft).Column
    
    ' Iterar sobre cada celda en la fila 1 hasta la última columna con datos
    For i = 1 To ultimaColumna
        ' Verificar si el valor de la celda comienza con "["
        If Left(wsOriginalData.Cells(1, i).Value, 1) = "[" Then
            ' Transponer los valores de la fila 1 y la fila seleccionada a las columnas A y B en el nuevo libro
            wsNuevo.Cells(destinoFila, 1).Value = Trim(wsOriginalData.Cells(1, i).Value)
            wsNuevo.Cells(destinoFila, 2).Value = Trim(wsOriginalData.Cells(filaSeleccionada, i).Value)
            destinoFila = destinoFila + 1
        End If
    Next i
    
    ' Guardar el nuevo libro con los resultados
    wbNuevo.SaveAs "C:\Users\tarik.said\Desktop\vbaTest\Resultados.xlsx"

    ' Cerrar el libro de producción sin guardar cambios
    wbOriginalData.Close SaveChanges:=False

    MsgBox "Informe generado con éxito en: C:\Users\tarik.said\Desktop\vbaTest\Resultados.xlsx", vbInformation
End Sub





' Sub CopiarYFiltrarVerticalSinHuecos()
'     Dim wsSource As Worksheet, wsDest As Worksheet
'     Dim ultimaColumna As Long
'     Dim celda As Range
'     Dim filaInicioDestino As Long, columnaInicioDestino As Long
'     Dim destinoActual As Long

'     ' Hoja de inicio y destino '
'     Set wsSource = ThisWorkbook.Sheets("respuesta original")
'     Set wsDest = ThisWorkbook.Sheets("test")
'     ultimaColumna = wsSource.Cells(2, wsSource.Columns.Count).End(xlToLeft).Column
    
'     filaInicioDestino = 8 ' Fila de inicio en la hoja "test" para pegar los datos
'     columnaInicioDestino = 2 ' Columna de inicio en la hoja "test"
'     destinoActual = filaInicioDestino ' Inicializar la fila destino para pegado vertical
    
'     Application.CutCopyMode = False

'     For i = 1 To ultimaColumna
'         If Left(wsSource.Cells(2, i).Value, 1) = "[" Then
'             ' Copiar celda por celda verticalmente solo si comienza con "["
'             wsDest.Cells(destinoActual, columnaInicioDestino).Value = wsSource.Cells(2, i).Value
'             destinoActual = destinoActual + 1 ' Moverse a la siguiente posición vertical
'         End If
'     Next i
' End Sub