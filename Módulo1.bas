Attribute VB_Name = "Módulo1"
Sub OptimizarTicketEventos()
    ' Declaración de hojas
    Dim wsOrigen As Worksheet
    Dim wsMuestra As Worksheet
    Dim wsDetalle As Worksheet
    Dim wsLog As Worksheet
    Dim wsGrafico As Worksheet
    
    Dim hojaDatos As String
    Dim incidente As String
    Dim fechaInicio As String
    Dim fechaFin As String
    
    'Actualización de la página detenida
    Application.ScreenUpdating = False
    
    ' Pedir el nombre de la hoja
    hojaDatos = InputBox("Ingrese el nombre de la hoja de donde desea tomar los datos", "Seleccionar Hoja")
    incidente = InputBox("Ingrese el nombre del INCIDENTE", "INCIDENTE")
    fechaInicio = InputBox("Ingrese la fecha de inicio para buscar las transacciones", "FECHA INICIO")
    fechaFin = InputBox("Ingrese la fecha de limite o fin para buscar las transacciones", "FECHA FIN")

    ' Validar y asignar la hoja de origen
    Set wsOrigen = ObtenerHoja(hojaDatos)
    If wsOrigen Is Nothing Then Exit Sub

    ' Crear la hoja de Muestra
    Set wsMuestra = CrearHoja("Muestra")

    ' Copiar la cabecera de la hoja de origen a la hoja de muestra
    CopiarCabecera wsOrigen, wsMuestra
    
    ' Añadir columna "ERROR" al final de la cabecera
    AgregarColumnaError wsMuestra
    
    ' Seleccionar aleatoriamente 15 casos y pegarlos en la hoja Muestra
    SeleccionarCasosAleatorios wsOrigen, wsMuestra, 15
    
    ' Crear hoja de Detalle
    Set wsDestalle = CrearHoja("Detalle")
    
    ' Crear hoja de Log
    Set wsLog = CrearHoja("Log")
    
    ' Crear hoja de Gráfico
    Set wsGrafico = CrearHoja("Gráfico")
    
     ' Intentar conectar a SQL Server
    If ConectarSQL() Then
        ' Si la conexión fue exitosa, ejecutar la consulta
        ' EjecutarConsultaSQL incidente, fechaInicio, fechaFin, "2210910586", "405110077323", "800.00"
        ' Llamar función para ejecutar múltiples consultas desde "Muestra"
        EjecutarConsultasDesdeMuestra wsMuestra, wsDetalle, incidente, fechaInicio, fechaFin
        
        ' Llamar al modulo para consultar los logs
        ObtenerLogsDesdeDetalle incidente
    End If
    
    ' Cerrar conexión después de la consulta
    CerrarConexionSQL
    
    Application.ScreenUpdating = True
    
End Sub
Private Sub EjecutarConsultasDesdeMuestra(wsMuestra As Worksheet, wsDetalle As Worksheet, incidente As String, fechaInicio As String, fechaFin As String)
    Dim lastRow As Long
    Dim fila As Integer
    Dim ordenante As String
    Dim beneficiario As String
    Dim monto As String
    Dim colOrdenante As Integer
    Dim colBeneficiario As Integer
    Dim colMonto As Integer
 
    ' Encontrar la última fila con datos en la hoja Muestra
    lastRow = wsMuestra.Cells(wsMuestra.Rows.Count, 1).End(xlUp).row
 
    ' Encontrar las columnas necesarias
    colOrdenante = BuscarColumna(wsMuestra, "N° DE CUENTA ORDENANTE")
    colBeneficiario = BuscarColumna(wsMuestra, "N° DE CUENTA BENEFICIARIA")
    colMonto = BuscarColumna(wsMuestra, "VALOR ORIGEN TRX")
 
    ' Validar si se encontraron las columnas
    If colOrdenante = 0 Or colBeneficiario = 0 Or colMonto = 0 Then
        MsgBox "No se encontraron las columnas requeridas en la hoja 'Muestra'.", vbCritical, "Error"
        Exit Sub
    End If
 
    ' Recorrer cada fila desde la segunda (evitar cabecera)
    For fila = 2 To lastRow
        ordenante = Trim(wsMuestra.Cells(fila, colOrdenante).Value)
        beneficiario = Trim(wsMuestra.Cells(fila, colBeneficiario).Value)
        monto = Trim(wsMuestra.Cells(fila, colMonto).Value)
 
        ' Validar que los datos no estén vacíos
        If ordenante <> "" And beneficiario <> "" And monto <> "" Then
            ' Ejecutar la consulta con los valores de la fila actual
            EjecutarConsultaSQL incidente, fechaInicio, fechaFin, ordenante, beneficiario, monto
        Else
            ' Registrar en la consola si hay datos vacíos
            Debug.Print "Fila " & fila & ": Datos incompletos, consulta no ejecutada."
        End If
    Next fila
End Sub
Private Function BuscarColumna(ws As Worksheet, nombreColumna As String) As Integer
    Dim lastColumn As Integer
    Dim col As Integer
 
    lastColumn = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
 
    For col = 1 To lastColumn
        If Trim(ws.Cells(1, col).Value) = nombreColumna Then
            BuscarColumna = col
            Exit Function
        End If
    Next col
 
    BuscarColumna = 0 ' Retorna 0 si no encuentra la columna
End Function
Private Function ObtenerHoja(nombreHoja As String) As Worksheet
    On Error Resume Next
    Set ObtenerHoja = ActiveWorkbook.Sheets(nombreHoja)
    On Error GoTo 0
    
    If ObtenerHoja Is Nothing Then
        MsgBox "La hoja '" & nombreHoja & "' no existe en este libro. Verifique el nombre e intente nuevamente.", vbCritical
    End If
End Function

Private Function CrearHoja(nombreHoja As String) As Worksheet
    On Error Resume Next
    Set CrearHoja = ActiveWorkbook.Sheets(nombreHoja)
    On Error GoTo 0

    If CrearHoja Is Nothing Then
        Set CrearHoja = ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count))
        CrearHoja.Name = nombreHoja
    End If
End Function

Private Sub CopiarCabecera(wsOrigen As Worksheet, wsMuestra As Worksheet)
    Dim lastColumn As Long
    lastColumn = wsOrigen.Cells(1, 1).End(xlToRight).Column
    
    ' Copiar la cabecera
    wsOrigen.Range(wsOrigen.Cells(1, 1), wsOrigen.Cells(1, lastColumn)).Copy
    wsMuestra.Cells(1, 1).PasteSpecial Paste:=xlPasteAll
    
    'MsgBox "Cabecera copiada exitosamente de '" & wsOrigen.Name & "' a 'Muestra'.", vbInformation
End Sub
Private Sub AgregarColumnaError(wsMuestra As Worksheet)
    Dim lastColumn As Long
    Dim nuevaCelda As Range
    
    lastColumn = wsMuestra.Cells(1, 1).End(xlToRight).Column
    
    ' Copiar el valor de la última celda de la cabecera
    Set nuevaCelda = wsMuestra.Cells(1, lastColumn + 1)
    wsMuestra.Cells(1, lastColumn).Copy
    nuevaCelda.PasteSpecial Paste:=xlPasteAll
    nuevaCelda.Value = "ERROR"
    
    'Deselecionar
    Application.CutCopyMode = False
    'MsgBox "Se agrego la columna 'ERROR'.", vbInformation
End Sub
Private Sub SeleccionarCasosAleatorios(wsOrigen As Worksheet, wsMuestra As Worksheet, numCasos As Long)
    Dim lastRow As Long
    Dim filasSeleccionadas As Collection
    Dim i As Long
    Dim filaRandom As Long
    Dim existe As Boolean

    ' Encontrar la última fila con datos en la hoja de origen
    lastRow = wsOrigen.Cells(wsOrigen.Rows.Count, 1).End(xlUp).row

    ' Verificar que haya suficientes filas
    If lastRow < numCasos + 1 Then
        MsgBox "No hay suficientes filas en la hoja de origen para seleccionar " & numCasos & " casos.", vbCritical
        Exit Sub
    End If

    ' Crear una colección para rastrear filas seleccionadas
    Set filasSeleccionadas = New Collection

    ' Limpiar filas previas en wsMuestra, si es necesario
    Dim nextRow As Long
    nextRow = wsMuestra.Cells(wsMuestra.Rows.Count, 1).End(xlUp).row + 1   ' Encuentra la siguiente fila vacía en wsMuestra
    
    ' Seleccionar casos aleatorios y copiarlos a wsMuestra
    For i = 1 To numCasos
        Do
            existe = False
            ' Generar un número aleatorio entre 2 y lastRow
            filaRandom = Application.WorksheetFunction.RandBetween(2, lastRow)

            ' Verificar si ya ha sido seleccionada
            On Error Resume Next
            filasSeleccionadas.Add filaRandom, CStr(filaRandom) ' Intentar agregar el número
            If Err.Number = 0 Then
                ' Si no hay error, esto significa que la fila no ha sido seleccionada
                existe = False
            Else
                ' Si hubo un error, la fila ya ha sido seleccionada
                existe = True
            End If
            On Error GoTo 0

        Loop While existe ' Repetir si la fila ya ha sido seleccionada

        ' Copiar la fila aleatoria al wsMuestra
        wsOrigen.Rows(filaRandom).Copy
        wsMuestra.Rows(nextRow).PasteSpecial Paste:=xlPasteAll

        nextRow = nextRow + 1  ' Incrementar la fila de destino
    Next i

    MsgBox numCasos & " casos aleatorios copiados exitosamente a la hoja 'Muestra'.", vbInformation
    
End Sub


