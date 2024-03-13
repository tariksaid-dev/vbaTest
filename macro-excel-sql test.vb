Sub GenerarInformeConID()
    Dim wbOriginalData As Workbook, wbNuevo As Workbook
    Dim wsOriginalData As Worksheet, wsNuevo As Worksheet
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

    ' Ruta completa del archivo de producción
    rutaOriginalData = "C:\Users\tarik.said\Desktop\vbaTest\Cuestionario SQL (Producción).xlsx"

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
            ' Transponer los valores de la fila 1 y la fila encontrada a las columnas A y B en el nuevo libro
            wsNuevo.Cells(destinoFila, 1).Value = Trim(wsOriginalData.Cells(1, i).Value)
            wsNuevo.Cells(destinoFila, 2).Value = Trim(wsOriginalData.Cells(filaID, i).Value)
            destinoFila = destinoFila + 1
        End If
    Next i

    ' Guardar el nuevo libro con los resultados
    wbNuevo.SaveAs "C:\Users\tarik.said\Desktop\vbaTest\Resultados.xlsx"

    ' Cerrar el libro de producción sin guardar cambios
    wbOriginalData.Close SaveChanges:=False

    MsgBox "Informe generado con éxito en: C:\Users\tarik.said\Desktop\vbaTest\Resultados.xlsx", vbInformation
End Sub
