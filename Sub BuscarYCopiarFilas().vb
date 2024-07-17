Sub CopiarRegistros()
    Dim newWb As Workbook
    Dim newRow As Long
    Dim lastRow As Long
    Dim numRegistros As Long ' Variable para contar registros
    
    ' Ruta de la carpeta con los archivos de Excel
    folderPath = "E:\huejutla\CLUES JUR. HUEJUTLA\concentrados\"
    searchValue = "10.07.2024" ' Cadena a buscar
    
    ' Crear un nuevo libro de trabajo
    Set newWb = Workbooks.Add
    
    ' Recorrer todos los archivos en la carpeta
    fileName = Dir(folderPath & "*.xls*")
    Do While fileName <> ""
        ' Abrir el archivo
        Set wb = Workbooks.Open(folderPath & fileName)
        
        ' Recorrer todas las hojas en el archivo
        For Each ws In wb.Worksheets
            ' Buscar la cadena en todas las filas
            Set cell = ws.Cells.Find(searchValue, LookIn:=xlValues, LookAt:=xlWhole)
            If Not cell Is Nothing Then
                ' Copiar desde la primera fila hasta la última fila
                newRow = newWb.Sheets(1).Cells(Rows.Count, 1).End(xlUp).Row + 1
                lastRow = ws.Cells(ws.Rows.Count, cell.Column).End(xlUp).Row
                ws.Rows(cell.Row & ":" & lastRow).Copy newWb.Sheets(1).Rows(newRow)
                
                ' Incrementar el contador de registros
                numRegistros = numRegistros + (lastRow - cell.Row + 1)
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
    newWb.SaveAs folderPath & "Resultados.xlsx"
    newWb.Close SaveChanges:=False
    
    ' Mostrar el número de registros copiados
    MsgBox "Se han copiado " & numRegistros & " registros y se ha limpiado el portapapeles.", vbInformation
End Sub


