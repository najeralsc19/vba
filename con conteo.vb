Sub BuscarYCopiarFilas()
    Dim folderPath As String
    Dim fileName As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim cell As Range
    Dim searchValue As String
    Dim newWb As Workbook
    Dim newRow As Long
    Dim lastRow As Long
    Dim totalRows As Long
    
    ' Ruta de la carpeta con los archivos de Excel
    folderPath = "E:\huejutla\FORMATO IMSS BIENESTAR JUNIO\01 totales\"
    searchValue = "26.06.2024" ' Cadena a buscar
    
    ' Crear un nuevo libro de trabajo
    Set newWb = Workbooks.Add
    
    ' Recorrer todos los archivos en la carpeta
    fileName = Dir(folderPath & "*.xls*")
    Do While fileName <> ""
        ' Abrir el archivo
        Set wb = Workbooks.Open(folderPath & fileName)
        
        For Each ws In wb.Worksheets
            
            Set cell = ws.Cells.Find(searchValue, LookIn:=xlValues, LookAt:=xlWhole)
            If Not cell Is Nothing Then
                
                newRow = newWb.Sheets(1).Cells(Rows.Count, 1).End(xlUp).Row + 1
                lastRow = ws.Cells(ws.Rows.Count, cell.Column).End(xlUp).Row
                ws.Rows(cell.Row & ":" & lastRow).Copy newWb.Sheets(1).Rows(newRow)
                totalRows = totalRows + (lastRow - cell.Row + 1)
            End If
        Next ws
        
        ' Cerrar el archivo sin guardar cambios
        wb.Close SaveChanges:=False
        
        ' Siguiente archivo
        fileName = Dir
    Loop
    
    ' Limpiar el portapapeles
    Application.CutCopyMode = False
    
    ' Guardar el nuevo libro de trabajo
    newWb.SaveAs folderPath & "ConcentradoFinal.xlsx"
    newWb.Close SaveChanges:=False
    
    MsgBox "Se han copiado " & totalRows & " filas correctamente y se ha limpiado el portapapeles.", vbInformation
End Sub
