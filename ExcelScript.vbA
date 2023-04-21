Sub ObtenerExcelValue()
    
    Dim objfso As Object
    Dim objshell As Object
    Dim valor As Variant
    Dim LocalUser As String
    Dim filePath As String
    Dim tmp As String
    
    
    Set objfso = CreateObject("Scripting.FileSystemObject")
    Set objshell = CreateObject("WScript.Shell")
    
    LocalUser = Environ("USERNAME")
    
    valor = ThisWorkbook.Sheets("Hoja3").Range("C13").Value
    
    ' Crear archivo ExcelValue.txt y escribir el valor en él
    Set archivo = objfso.CreateTextFile("C:\Users\" & LocalUser & "\Documents\ExcelValue.txt", True)
    archivo.WriteLine CStr(valor)
    archivo.Close
    
    ' Escribir el valor en el archivo excel.txt
    filePath = "C:\Users\" & LocalUser & "\Documents\excel.txt"
    Dim fileNum As Integer
    fileNum = FreeFile
    Open filePath For Output As #fileNum
    Write #fileNum, valor
    Close #fileNum
    
    ' Leer la ruta de la carpeta desde el archivo Folder.txt
    rutaArchivo = "C:\Users\" & LocalUser & "\Documents\Folder.txt"
    If objfso.FileExists(rutaArchivo) Then
        Set archivo = objfso.OpenTextFile(rutaArchivo, 1)
        contenido = archivo.ReadAll
        archivo.Close
        Path = Replace(contenido, vbCr, "")
        contenido = Replace(Path, vbLf, "")
        tmp = "D:\EMPLEADOS\JONATHAN\CORTES\" & CStr(contenido) + "\"
        
        ' Cambiar el directorio activo a la carpeta especificada
            MsgBox tmp
            
            ChDir tmp
            
        
            ThisWorkbook.SaveAs Filename:=tmp + "Plantilla-Diaria-Por-Usuario.xls", FileFormat:=xlExcel8, Password:="", WriteResPassword:="", ReadOnlyRecommended:=False, CreateBackup:=False
    Else
        MsgBox "No se encontró el archivo Folder.txt."
    End If
    
    'objshell.Popup "DONE", 1, "Exportar valor", 0 + 64
    
End Sub


