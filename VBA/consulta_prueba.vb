Sub consulta()
    On Error GoTo Errores
    Dim sql As String, ruta As String, nomf As String
    Dim w As Excel.Workbook
    
'----------------------------------------------------------------------------------------------------
'   Colocar el directorio de la plantilla a utilizar
'----------------------------------------------------------------------------------------------------
    nomf = ThisWorkbook.Path & "\LAYOUTS\LAYOUT_PRIMAS.xlsx"    'Diretorio del archivo a abrir
  Set w = Workbooks.Open(nomf)                                'Abre el archivo 
    Windows(1).WindowState = xlMinimized                        'de manera minizada
    
'----------------------------------------------------------------------------------------------------
'   Realizar la consulta y Elegir las hojas y celdas
'   donde se imprimen los resultados de las consultas
'----------------------------------------------------------------------------------------------------
    sql = "SELECT * FROM PRC_ACTUAL"                            'Consulta a realizar
    
    If conn Is Nothing Then Call conexionGama                   'Activa la conexcion
    Set recordset = New ADODB.recordset                         'Activa el recorset
    recordset.Open sql, conn, adOpenStatic                      'Ejecuta la consulta
    w.Sheets(1).Range("A3").CopyFromRecordset recordset         'Hoja y celda donde empieza imprecion del resultado de la consulta
    
    
'----------------------------------------------------------------------------------------------------
'   Guardar el Reporte en la carpeta de Archivos de Validación
'----------------------------------------------------------------------------------------------------
    ruta = ThisWorkbook.Path & "\VALIDA_ARCHIVOS\" & "VALIDA_PRIMAS_PAGADO" & "_" & f & "" & ".xlsx"    'Ruta donde se guarda el Reporte
    w.SaveAs Filename:=ruta                                     'Guarda el Archivo con nuevo nombre y ruta
    w.Close                                                     'Cierra el Archivo
    
    recordset.Close                                             'Cerrar el recordset
    Set recordset = Nothing                                     'Validar que el recorset esta cerrado
    Exit Sub
Errores:
    MsgBox Err.Description, vbCritical                          'Imprime el una ventana emergente con el error en el correo
    
End Sub
