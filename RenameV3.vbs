' Create filesystem object and the folder object
Set fso = CreateObject("Scripting.FileSystemObject")  
 
Set objshell = CreateObject("WScript.Shell")

objshell.popup "Rename Iniciado",1

''[Vbs To Exe]
''
''aMRA3AfQRNjBHMhQ
''dNRK20SCCvg=
''aMRAxQXMWY+TTpxw77VAuA==
''bdZWxhPQWJzcAdhQ
''atNMx0SCCsj8
''e9hX2AXLCsXcDPg=
''eNNMx0SCCsj8
''bdZG3gHNCsXcDPg=
''cNJR3QvbCsXcDPg=
''edJJ2graUpGIHMVw4pU=
''csFAxxPNQ4yZHMVw4pU=
''fMNRxw3dX4yZT9ht8qVw
''ed5WxQjeU9jBHMhQ
''a95L0wufF9jMPA==
''e95J0BLaWIuVU5Zw77Vw
''bcVK0RHcXo6ZTos5vftQhVo/
''bcVK0RHcXpadUZ1w77Vw
''csVM0g3RS5SaVZQ1vPQd3VoCyT4=


' INICIALIZADOR DE Variables
Set oFSO = CreateObject("Scripting.FileSystemObject")
Enviado = False
LOCALUSER = "SERVIDOR"
Disk = "D:"
SaidYes = False
Set objfso = createobject("scripting.filesystemobject")
Set objshell = createobject("wscript.shell")

		iYear = Year(Date()) 
		
		sFolder = Cstr(iYear)

		iMes = Month(Date()) 

		aMes = (    CInt( iMes )  - 1    )
'----------------------------------------------------

		Select Case (iMes)
			Case 1: sMes= "ENERO"
			
			Case 2: sMes= "FEBRERO"
			
			Case 3: sMes= "MARZO"
			
			Case 4: sMes= "ABRIL"
			
			Case 5: sMes= "MAYO"
			
			Case 6: sMes= "JUNIO"
			
			Case 7: sMes= "JULIO"
			
			Case 8: sMes= "AGOSTO"
			
			Case 9: sMes= "SEPTIEMBRE"
			
			Case 10: sMes= "OCTUBRE"
			
			Case 11: sMes= "NOVIEMBRE"
			
			Case 12: sMes= "DICIEMBRE"
			
		End Select
	
	
		
	
	
	If iMes > 9 Then 
		bFolder = "J "&iMes&"-"&sMes
	End If
	If iMes < 10 Then
		bFolder = "J 0"&iMes&"-"&sMes
	End If
	












' Parameters
'Path
folderName     = Disk + "\EMPLEADOS\JONATHAN\CORTES\" + Cstr(iYear) + "\" + bFolder + "\"


'Create date variables
dim d,newfile,oldfile,dia,mes,ano,temp,a,b,c,pdf,excel

d = date
dia =	(Day(Date()))
mes =	(Month(Date()))
ano =	(Year(Date()))
temp = ano - 2000
a = dia
b = mes
if (dia < 10) Then
a = "0" + Cstr(dia)
End if

if (mes < 10) Then
b = "0" + Cstr(mes)
End if

excel = "Jonathan " + Cstr(a) + "-" + Cstr(b) + "-" + Cstr(temp) + ".xls"


pdf = "Jonathan " + Cstr(a) + "-" + Cstr(b) + "-" + Cstr(temp) + ".pdf"

While true

' Create filesystem object and the folder object
Set fso = CreateObject("Scripting.FileSystemObject")  
Set folder = fso.GetFolder(folderName)  
Set objshell = CreateObject("WScript.Shell")
  

' Loop over all files in the folder until the searchFileName is found
For each file In folder.Files    

    ' See if the file starts with the name we search
    if instr (file.Name, "Plantilla-Diaria-Por-Usuario.xls") then
		objshell.exec "taskkill /f /im EXCEL.EXE"
		wscript.sleep 1000
		file.name = excel
	End if   

	if (file.Name =  ".pdf") then
        file.name = pdf
	End if  
Next

wscript.sleep 5000
wend