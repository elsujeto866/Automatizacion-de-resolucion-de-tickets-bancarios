Attribute VB_Name = "Módulo4"
Option Explicit
 
' Función para extraer GUIDs desde la pestaña "Detalle" y obtener logs
Public Sub ObtenerLogsDesdeDetalle(inc As String)
    Dim wsDetalle As Worksheet, wsLog As Worksheet
    Dim lastRow As Long, rowNum As Long
    Dim incidente As String, fechaIngreso As String
    Dim fechaInicio As String, fechaFin As String
    Dim estado As String, guid As String
    Dim guids As Object  ' Diccionario para evitar duplicados
    Dim mensajeError As String
    Dim logLastRow As Long
 
    ' Definir las hojas
    Set wsDetalle = ActiveWorkbook.Sheets("Detalle")
    Set wsLog = ActiveWorkbook.Sheets("Log")
 
    ' Borrar datos previos en la hoja "Log"
    wsLog.Cells.Clear
 
    ' Encontrar la última fila en la hoja "Detalle"
    lastRow = wsDetalle.Cells(Rows.Count, 1).End(xlUp).row
    rowNum = 1  ' Comenzar en la primera fila
 
    ' Verificar conexión antes de ejecutar las consultas
    If Not ConectarSQL() Then
        MsgBox "No se pudo conectar a SQL Server.", vbCritical
        Exit Sub
    End If
 
    ' Recorrer la hoja "Detalle"
    Do While rowNum <= lastRow
        ' Verificar si la primera columna tiene "ODT" (indica cabecera de nuevo bloque)
        If Trim(wsDetalle.Cells(rowNum, 1).Value) = "ODT" Then
            ' Inicializar nuevo caso y diccionario de GUIDs
            Set guids = CreateObject("Scripting.Dictionary")
            incidente = wsDetalle.Cells(rowNum + 1, 1).Value ' Obtener incidente de la siguiente fila
            
        Else
            ' Leer los datos de la fila actual
            mensajeError = wsDetalle.Cells(rowNum, 1).Value
            guid = Trim(wsDetalle.Cells(rowNum, 4).Value) ' TSH_GUID en columna 4 (D)
            fechaIngreso = wsDetalle.Cells(rowNum, 6).Value ' TSH_FECHA en columna 6 (F)
            estado = wsDetalle.Cells(rowNum, 2).Value ' TSH_ESTADO en columna 2 (B)
            
 
            ' Verificar si hay mensaje de error
            If mensajeError = "No se encontraron datos con los parámetros proporcionados." Then
                logLastRow = wsLog.Cells(Rows.Count, 1).End(xlUp).row + 1
                 ' Si la hoja está vacía, la primera fila de datos será la fila 1
                If logLastRow = 2 Then logLastRow = 1
                Call ImprimirCabeceraEnHoja(logLastRow, wsLog)
                wsLog.Cells(logLastRow + 1, 1).Value = mensajeError
                rowNum = rowNum + 1
                GoTo SiguienteBloque
            End If
 
            ' Agregar GUIDs si el estado es diferente de 10
            If estado <> "10" And guid <> "" Then
                If Not guids.Exists(guid) Then
                    guids.Add guid, True
                End If
            End If
        End If
 
        ' Si encontramos otra cabecera "ODT" o llegamos al final, ejecutamos la consulta
        If (rowNum = lastRow) Or (Trim(wsDetalle.Cells(rowNum + 1, 1).Value) = "ODT") Then
            If guids.Count > 0 Then
                ' Calcular el rango de ±1 hora para la búsqueda
                fechaInicio = Format(CDate(fechaIngreso) - TimeValue("00:05:00"), "yyyy-mm-dd HH:MM:SS")
                fechaFin = Format(CDate(fechaIngreso) + TimeValue("00:05:00"), "yyyy-mm-dd HH:MM:SS")
 
                ' Ejecutar la consulta SQL
                Call EjecutarConsultaLogs(inc, fechaInicio, fechaFin, guids)
                        
            Else
                ' Si todos los GUIDs tienen estado 10, registrar "SIN NOVEDAD"
                logLastRow = wsLog.Cells(Rows.Count, 1).End(xlUp).row + 1
                 ' Si la hoja está vacía, la primera fila de datos será la fila 1
                If logLastRow = 2 Then logLastRow = 1
                Call ImprimirCabeceraEnHoja(logLastRow, wsLog)
                wsLog.Cells(logLastRow + 1, 1).Value = "SIN NOVEDAD"
            End If
        End If
        
SiguienteBloque:
        rowNum = rowNum + 1
    Loop
 
    ' Cerrar conexión SQL
    conn.Close
    Set conn = Nothing
    
    MsgBox "Consultas de logs completadas.", vbInformation
End Sub
Public Sub EjecutarConsultaLogs(incidente As String, fechaInicio As String, fechaFin As String, guids As Object)
    On Error GoTo ErrHandler
    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")
    
    ' Verificar si la conexión está abierta antes de ejecutar la consulta
    If conn Is Nothing Then
        MsgBox "La conexión a SQL Server no está establecida.", vbExclamation, "Error"
        Exit Sub
    End If
 
    ' Construcción de la consulta SQL con WITH GUID_LIST AS (...)
    Dim consultaSQL As String
    consultaSQL = "WITH GUID_LIST AS ("
    
    Dim guid As Variant
    Dim first As Boolean: first = True
    For Each guid In guids.Keys
        If Not first Then consultaSQL = consultaSQL & " UNION ALL "
        consultaSQL = consultaSQL & "SELECT '" & guid & "' AS LTR_GUID"
        first = False
    Next guid
    
    consultaSQL = consultaSQL & ") " & _
                  "SELECT TOP 100 '" & incidente & "' AS ODT, l.* " & _
                  "FROM [BDD_OMNI_TRANS_HIST].[LOGS].[LOG_TRANSACCIONAL] l WITH(NOLOCK) " & _
                  "INNER JOIN GUID_LIST g ON l.LTR_GUID = g.LTR_GUID " & _
                  "WHERE LTR_FECHA_HORA >= '" & fechaInicio & "' " & _
                  "AND LTR_FECHA_HORA <= '" & fechaFin & "'"
 
    ' Ejecutar la consulta SQL
    rs.Open consultaSQL, conn
 
    ' Verificar si hay datos
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Sheets("Log")
 
    Dim nextRow As Long
    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row + 1
    
    ' Si la hoja está vacía, la primera fila de datos será la fila 1
    If nextRow = 2 Then nextRow = 1
    
    ' Insertar las cabeceras en cada consulta
    Dim i As Integer
    For i = 0 To rs.Fields.Count - 1
        ws.Cells(nextRow, i + 1).Value = rs.Fields(i).Name
    Next i
    
    
    ' Si la consulta no devuelve datos, agregar mensaje en la siguiente fila vacía
    If rs.EOF Then
        ws.Cells(nextRow + 1, 1).Value = "No se encontraron datos con los parámetros proporcionados."
        Exit Sub
    End If
 
    ' Copiar los datos en la siguiente fila vacía
    ws.Cells(nextRow + 1, 1).CopyFromRecordset rs
 
    ' Cerrar el recordset
    rs.Close
    Set rs = Nothing
    Exit Sub
 
ErrHandler:
    MsgBox "Error en la consulta SQL: " & Err.Description, vbCritical, "Error"
    Debug.Print "Error en la consulta SQL: " & Err.Description
End Sub
Sub ImprimirCabeceraEnHoja(filaInicio As Long, wsLog As Worksheet)
    Dim cabeceras As Variant
    Dim i As Integer

    ' Define el rango de cabeceras
    cabeceras = Array("ODT", "LTR_CODIGO", "LTR_EMPRESA", "LTR_CANAL", "LTR_MEDIO", "LTR_APLICACION", _
                      "LTR_AGENCIA", "LTR_IDIOMA", "LTR_USUARIO", "LTR_IP_SRV_SOLICITANTE", "LTR_SESION", _
                      "LTR_UNICIDAD", "LTR_GUID", "LTR_FECHA_HORA", "LTR_FECHA_EJECUCION", "LTR_IP_ISP", _
                      "LTR_DISPOSITIVO", "LTR_GEOLOCALIZACION", "LTR_ID_CLIENTE", "LTR_TIPO_ID_CLIENTE", _
                      "LTR_NOMBRE_ORQ", "LTR_METODO_ORQ", "LTR_TIPO_TRANSACCION_ORQ", "LTR_NOMBRE_MS", _
                      "LTR_METODO_MS", "LTR_TIPO_TRANSACCION_MS", "LTR_FECHA_HORA_INICIO", "LTR_FECHA_HORA_FIN", _
                      "LTR_CODIGO_RESPUESTA", "LTR_MENSAJE_RESPUESTA", "LTR_MENSAJE_NEGO_RESPUESTA", _
                      "LTR_SEVERIDAD_RESPUESTA", "LTR_DIA", "LTR_MES", "LTR_ELASTICO")

    ' Imprimir las cabeceras en la fila especificada de la hoja proporcionada
    For i = LBound(cabeceras) To UBound(cabeceras)
        wsLog.Cells(filaInicio, i + 1).Value = cabeceras(i)
    Next i
    
   
End Sub


