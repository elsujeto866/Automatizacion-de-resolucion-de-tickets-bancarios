Attribute VB_Name = "Módulo2"
Option Explicit
 
Public conn As Object ' Variable global para la conexión
 
' Función para conectar a SQL Server
Public Function ConectarSQL() As Boolean
    On Error GoTo ErrHandler
 
    ' Crear el objeto de conexión si aún no existe
    If conn Is Nothing Then
        Set conn = CreateObject("ADODB.Connection")
    ElseIf conn.State = 1 Then
        ' MsgBox "Ya hay una conexión abierta.", vbInformation, "Aviso"
        ConectarSQL = True
        Exit Function
    End If
 
    ' Definir la cadena de conexión con autenticación de Windows (SSPI)
    Dim connectionString As String
    connectionString = "Provider=SQLOLEDB; Data Source=DRECBPPRQ140\Q140DCA,11440; Initial Catalog=BDD_OMNI_TRANS; Integrated Security=SSPI;"
    
    ' Abrir la conexión
    conn.Open connectionString
 
    ' Verificar si la conexión fue exitosa
    If conn.State = 1 Then
        ConectarSQL = True
        MsgBox "Conexión exitosa a SQL Server.", vbInformation, "Éxito"
        Debug.Print "Conexión exitosa a SQL Server." ' Verificar en la Ventana Inmediata (CTRL+G)
    Else
        ConectarSQL = False
    End If
    Exit Function
 
ErrHandler:
    MsgBox "Error al conectar con SQL Server: " & Err.Description, vbCritical, "Error"
    Debug.Print "Error al conectar con SQL Server: " & Err.Description
    ConectarSQL = False
End Function
 
' Función para cerrar la conexión
Public Sub CerrarConexionSQL()
    If Not conn Is Nothing Then
        conn.Close
        Set conn = Nothing
        MsgBox "Conexión cerrada correctamente.", vbInformation, "Cierre de Conexión"
        Debug.Print "Conexión cerrada correctamente."
    End If
End Sub
