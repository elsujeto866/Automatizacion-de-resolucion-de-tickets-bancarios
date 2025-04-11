Attribute VB_Name = "M�dulo2"
Option Explicit
 
Public conn As Object ' Variable global para la conexi�n
 
' Funci�n para conectar a SQL Server
Public Function ConectarSQL() As Boolean
    On Error GoTo ErrHandler
 
    ' Crear el objeto de conexi�n si a�n no existe
    If conn Is Nothing Then
        Set conn = CreateObject("ADODB.Connection")
    ElseIf conn.State = 1 Then
        ' MsgBox "Ya hay una conexi�n abierta.", vbInformation, "Aviso"
        ConectarSQL = True
        Exit Function
    End If
 
    ' Definir la cadena de conexi�n con autenticaci�n de Windows (SSPI)
    Dim connectionString As String
    connectionString = "Provider=SQLOLEDB; Data Source=DRECBPPRQ140\Q140DCA,11440; Initial Catalog=BDD_OMNI_TRANS; Integrated Security=SSPI;"
    
    ' Abrir la conexi�n
    conn.Open connectionString
 
    ' Verificar si la conexi�n fue exitosa
    If conn.State = 1 Then
        ConectarSQL = True
        MsgBox "Conexi�n exitosa a SQL Server.", vbInformation, "�xito"
        Debug.Print "Conexi�n exitosa a SQL Server." ' Verificar en la Ventana Inmediata (CTRL+G)
    Else
        ConectarSQL = False
    End If
    Exit Function
 
ErrHandler:
    MsgBox "Error al conectar con SQL Server: " & Err.Description, vbCritical, "Error"
    Debug.Print "Error al conectar con SQL Server: " & Err.Description
    ConectarSQL = False
End Function
 
' Funci�n para cerrar la conexi�n
Public Sub CerrarConexionSQL()
    If Not conn Is Nothing Then
        conn.Close
        Set conn = Nothing
        MsgBox "Conexi�n cerrada correctamente.", vbInformation, "Cierre de Conexi�n"
        Debug.Print "Conexi�n cerrada correctamente."
    End If
End Sub
