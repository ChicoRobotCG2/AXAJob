Public conn As ADODB.Connection
Public recordset As ADODB.recordset

Sub conexionGama()
    On Error GoTo Errores
    Dim host As String, database As String
    host = "10.133.42.2"
    database = "BAC_2021"
    Set conn = New ADODB.Connection
    conn.Open "Driver={SQL Server};Server=" & host & ";Database=" & database & ";"
    Debug.Print "Conexión Exitosa a la Base de Datos"
    Exit Sub
Errores:
    MsgBox Err.Description, vcCritical
End Sub



Sub cerrarConexion()
    If conn Is Nothing Then Exit Sub
    conn.Close
    Set conn = Nothing
    Debug.Print "Se ha cerrado la Conexion"
End Sub
