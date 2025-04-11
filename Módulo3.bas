Attribute VB_Name = "M�dulo3"
Option Explicit
 
' Funci�n para ejecutar una consulta SQL con par�metros y copiar los datos en Excel
Public Sub EjecutarConsultaSQL(incidente As String, fechaInicio As String, fechaFin As String, ordenante As String, beneficiario As String, monto As String)
    On Error GoTo ErrHandler
 
    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")
 
    ' Verificar si la conexi�n est� abierta antes de ejecutar la consulta
    If conn Is Nothing Then
        MsgBox "La conexi�n a SQL Server no est� establecida.", vbExclamation, "Error"
        Exit Sub
    End If
    
    ' Convertir el monto a valor absoluto

    
    ' Convertir el monto a valor absoluto
    Dim montoAbsoluto As Double
    montoAbsoluto = Abs(CDbl(monto))
    
    Dim montoFormateado As String
    montoFormateado = Replace(CStr(montoAbsoluto), ",", ".")

    
    
    
    ' Construcci�n de la consulta SQL con par�metros din�micos
    Dim consultaSQL As String
    consultaSQL = "DECLARE @fechaInicio DATETIME, @fechaFin DATETIME, @ordenante VARCHAR(50), @beneficiario VARCHAR(50), @tsh_monto DECIMAL(18,4); " & _
                  "SET @fechaInicio = '" & fechaInicio & "'; " & _
                  "SET @fechaFin = '" & fechaFin & "'; " & _
                  "SET @ordenante = '" & ordenante & "'; " & _
                  "SET @beneficiario = '" & beneficiario & "'; " & _
                  "SET @tsh_monto = " & montoFormateado & "; " & _
                  "SELECT '" & incidente & "' AS ODT, " & _
                  "T0.TSH_ESTADO_TRANSACCION, T0.TSH_CODIGO, T0.TSH_GUID, T0.TSH_MONTO, " & _
                  "T0.TSH_FECHA_INGRESO, T0.TSH_ID_ORDENANTE, T0.TSH_ID_BENEFICIARIO, " & _
                  "T0.TSH_PRODUCTO_BENEFICIARIO, T0.TSH_PRODUCTO_ORDENANTE, T0.TSH_TIPO_TRANSACCION, " & _
                  "T0.TSH_JSON_ELASTICO, T1.REV_ID, T1.REV_ESTADO, " & _
                  "T1.REV_FECHA_INGRESO, T1.REV_FECHA_EJECUCION " & _
                  "FROM [TRANSACCION].[TRN_TRANSACCION_SWITCH] T0 WITH(NOLOCK) " & _
                  "LEFT JOIN [TRANSACCION].[TRN_REVERSO] T1 WITH(NOLOCK) ON T0.TSH_GUID = T1.REV_GUID " & _
                  "WHERE T0.TSH_FECHA_EJECUCION BETWEEN @fechaInicio AND @fechaFin " & _
                  "AND T0.TSH_PRODUCTO_ORDENANTE = @ordenante " & _
                  "AND T0.TSH_PRODUCTO_BENEFICIARIO = @beneficiario " & _
                  "AND T0.TSH_MONTO = @tsh_monto;"
 
    ' "AND T0.TSH_MONTO = " & Replace(monto, ",", ".") & " " & _
    ' Ejecutar la consulta
    rs.Open consultaSQL, conn
 
    ' Verificar si hay datos
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Sheets("Detalle") ' Asegurar que escriba en la hoja correcta
 
    Dim nextRow As Long
    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row + 1 ' Buscar la siguiente fila vac�a
 
    ' Si la hoja est� vac�a, la primera fila de datos ser� la fila 1
      If nextRow = 2 Then nextRow = 1
 
    ' Insertar las cabeceras en cada consulta
    Dim i As Integer
    For i = 0 To rs.Fields.Count - 1
        ws.Cells(nextRow, i + 1).Value = rs.Fields(i).Name
    Next i
 
    ' Mover el cursor a la siguiente fila disponible despu�s de la cabecera
    nextRow = nextRow + 1
 
    ' Si la consulta no devuelve datos, agregar mensaje en la siguiente fila vac�a
    If rs.EOF Then
        ws.Cells(nextRow, 1).Value = "No se encontraron datos con los par�metros proporcionados."
        Exit Sub
    End If
 
    ' Copiar los datos en la siguiente fila vac�a
    ws.Cells(nextRow, 1).CopyFromRecordset rs
    
    ' Formatear columnas de fecha despu�s de copiar los datos
    Call FormatearFechas(ws, nextRow, rs.Fields.Count)
 
    ' Cerrar el recordset
    rs.Close
    Set rs = Nothing
 
    Exit Sub
 
ErrHandler:
    MsgBox "Error en la consulta SQL: " & Err.Description, vbCritical, "Error"
    Debug.Print "Error en la consulta SQL: " & Err.Description
End Sub

Private Sub FormatearFechas(ws As Worksheet, startRow As Long, totalColumns As Integer)
    Dim col As Integer
    Dim row As Long
    Dim lastRow As Long
 
    ' Encontrar la �ltima fila de datos
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
 
    ' Recorrer todas las columnas para identificar las que contienen fechas
    For col = 1 To totalColumns
        ' Verificar si la cabecera indica que es una columna de fecha
        If InStr(1, ws.Cells(1, col).Value, "TSH_FECHA_INGRESO", vbTextCompare) > 0 Then
            ' Aplicar formato de fecha en la columna desde startRow hasta la �ltima fila
            For row = startRow To lastRow
                If IsNumeric(ws.Cells(row, col).Value) Then
                    ws.Cells(row, col).NumberFormat = "yyyy-mm-dd hh:mm:ss"
                End If
            Next row
        End If
    Next col
End Sub


