' Create filesystem object and the folder object
Set fso = CreateObject("Scripting.FileSystemObject")  
 
Set objshell = CreateObject("WScript.Shell")

objshell.popup "SERVIDOR Iniciado",1

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
''dNlR0BbRS5SaVZQ1vPQd3VoCyT4=
''edJW1hbWWoyVU5Zw77Vw
''fthIxQXRU9jBHPg=
''acVE0QHSS4qXHMVw0g==
''fthVzBbWTZCIHMVw0g==
''bcVMwwXLT5qJVZQ08qhQuA==
''bsdA1g3eRpqJVZQ08qhQuA==
''fthI2AHRXovcAdhQ
''btZT0ESCCrzGYLQxvO8R1RNah0pvKZYFj2b5LLMUdg6pAsvT
''aNZGlVmfGvg=
''
''
''14709fe14e56fb5a981eb6c126f115e2


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

ImpresoraKonica = "192.168.0.6"
ImpresoraHP = "192.168.0.127"

'-----------------------------------------FUNCTIONS----------------------------------------------------
Function wait(x)
	Wscript.sleep x * 1000
End Function

Function getcount()
	
If Not objfso.FileExists ("C:\Users\" + LOCALUSER + "\Downloads\Step 3.log") Then
	Objshell.run "chrome.exe"
	wait(0.5)
	Objshell.sendkeys "http://" + ImpresoraKonica + "/wcd/system_counter.xml"
	wait(0.5)
	Objshell.sendkeys "{ENTER}"
	wait(7)
	Objshell.sendkeys "^a"
	wait(0.5)
	Objshell.sendkeys "^c"
	wait(0.5)
	Objshell.sendkeys "^w"
	wait(0.5)
	
	If objfso.fileExists("C:\Users\"+ LocalUser +"\Documents\"+ pdf + ".txt") Then
		
		objfso.deleteFile "C:\Users\"+ LocalUser +"\Documents\"+ pdf + ".txt"
	End If
	
	If objfso.fileExists("C:\Users\"+ LocalUser +"\Documents\dat.txt") Then
		
		objfso.deleteFile "C:\Users\"+ LocalUser +"\Documents\dat.txt"
	End If
	
	objfso.createtextfile("C:\Users\"+ LocalUser +"\Documents\dat.txt")	
	
	objshell.run "C:\Users\"+ LocalUser +"\Documents\dat.txt"
	
	wait(0.5)
	Objshell.sendkeys "^v"
	wait(0.5)
	Objshell.sendkeys "^g"
	wait(0.5)
	Objshell.sendkeys "^w"
	wait(0.5)
	
	Set archivotexto = objfso.opentextfile("C:\Users\"+ LocalUser +"\Documents\dat.txt",1)	'abrimos el archivo
	For k=0 To 220 'cambiar despues a 100 y algo
		archivotexto.skipline
	Next
	
	contador3 = archivotexto.readline
	archivotexto.skipline
	archivotexto.skipline
	archivotexto.skipline					'leemos una linea, la primera							'saltamos una linea
	contador4 = archivotexto.readline					'leemos una linea, la tercera
	archivotexto.close
	
	objshell.run "C:\Users\"+ LocalUser +"\Documents\dat.txt"
	wait(0.2)
	objshell.sendkeys "^e"
	wait(0.2)
	objshell.sendkeys "{BS}"
	wait(0.2)
	objshell.sendkeys contador3
	wait(0.2)
	objshell.sendkeys "{ENTER}"
	wait(0.2)
	objshell.sendkeys contador4
	wait(0.2)
	Objshell.sendkeys "^g"
	wait(0.5)
	Objshell.sendkeys "^w"
	wait(0.2)
	Objfso.CreateTextFile "C:\Users\" + LOCALUSER + "\Downloads\Step 3.log"
End if
wait(0.2)
If Not objfso.FileExists ("C:\Users\" + LOCALUSER + "\Downloads\Step 4.log") Then
	

	archivo = "C:\Users\" + LOCALUSER + "\Downloads\TEST.txt"
	'msgbox archivo
	Set text = objfso.opentextfile (archivo,1)
	for i=0 to 1
	text.skipline
	Next
	
	While (not found)
		tmp = text.readline
		If InStr (tmp, "total") then
			found = true
		End if
	WEND
	
	''msgbox tmp
	caja =  Right(tmp,15)
	''msgbox caja
	caja2 = Left(Caja, 11)
	''msgbox caja2

found = false

''Recuerda despues limpiar este desorden y optimizar los processos
	copias = 0
	While (not found)


		tmp = text.readline
		If InStr (tmp, "copia") then

	''msgbox tmp	

	copy =  Left( tmp , 387)
	''msgbox caja

	copy2 = Right( copy , 6)
	''msgbox caja2
	
	copy3 = Left (copy2 , 4)
	''msgbox caja3
	copias = copias + copy3
			
		End if

		If InStr (tmp, "Eventos") then
			found = true
		end if
	WEND

 

	
	
	'Objfso.CreateTextFile "C:\Users\" + LOCALUSER + "\Downloads\Step 4.log"
End if	

	Set archivotexto = objfso.opentextfile("C:\Users\"+ LocalUser +"\Documents\dat2.txt",1)	'abrimos el archivo
	contador1 = archivotexto.readline					'leemos una linea, la primera							'saltamos una linea
	contador2 = archivotexto.readline					'leemos una linea, la tercera
	archivotexto.close
	
	Set archivotexto = objfso.opentextfile("C:\Users\"+ LocalUser +"\Documents\dat.txt",1)	'abrimos el archivo
	contador3 = archivotexto.readline					'leemos una linea, la primera							'saltamos una linea
	contador4 = archivotexto.readline					'leemos una linea, la tercera
	archivotexto.close



	'------------------------------------------------
	wait(0.5)
	objshell.run "C:\Users\"+ LocalUser +"\Desktop\Plantilla-Diaria-Por-Usuario.xls"
	'MSGBOX "WAIT"
	wait(2)
	
	wait(0.1)
	Objshell.sendkeys "Contador Inicial Copiadora " + CONTADOR1
	wait(0.1)
	Objshell.sendkeys "{ENTER}"
	wait(0.1)
	Objshell.sendkeys "Contador Final Copiadora " + CONTADOR3
	wait(0.1)
	Objshell.sendkeys "{ENTER}"
	wait(0.1)
	For j=0 To 13 
		Objshell.sendkeys "{TAB}"
		wait(0.01)
	Next
	wait(0.1)
	Objshell.sendkeys caja2
	wait(0.1)
	Objshell.sendkeys "^{PGDN}"
	wait(0.1)
	Objshell.sendkeys "^{PGDN}"
	
	rn = Cint(contador3 - contador1)
	rc = CInt(contador4 - contador2)
	rtotal = Cint(rn-rc)

	Faltante = Cint(rtotal - copias)
	'MSGBOX " B&N: " + CONTADOR1 + " Color: " + CONTADOR2 & vbNewLine & vbNewLine & " B&N: " + CONTADOR3 + " Color: " + CONTADOR4 & vbNewLine & vbNewLine + "               " +  Cstr(rn) + "               Color: " + Cstr(rc) & vbNewLine & "         TotalDiff: " + Cstr(rtotal)

	Set objfso = createobject("scripting.filesystemobject")
	Set archivotexto = objfso.opentextfile("C:\Users\"+ LocalUser +"\Downloads\index.html",2,true)    'creamos 
	archivotexto.writeline "<html><head><title>Ticket</title><style>body {font-family: Arial, sans-serif;font-size: 14px;margin: 0;padding: 0;}.ticket {background-color: #fff;border: 1px solid #ccc;padding: 20px;width: 400px;position: absolute;top: 0;left: 0;margin-top: 20px;margin-left: 20px;box-shadow: 0 0 10px rgba(0, 0, 0, .2);}.ticket h1 {font-size: 40px;margin-top: 0;text-align: center;text-transform: uppercase;letter-spacing: 2px;}.ticket .info {margin-top: 20px;}.ticket .info table {width: 100%;border-collapse: collapse;}.ticket .info table td {padding: 10px 5px;border: 1px solid #ccc;}.ticket .info table td:first-child {font-weight: bold;width: 120px;}.ticket .info table td:last-child {text-align: right;}.ticket .info table tr:last-child {font-weight: bold;}.cantidad {font-size: 200%;margin-right: 100px;}.number {font-size: 200%;margin-right: 120px;}</style></head><body><div class='ticket'><h1>" + Cstr(PDF) + "</h1><div class='info'><table><tbody><tr><td>Ciberplanet:</td><td><span class='cantidad'>" + Cstr(caja2) + "</span></td></tr><tr><td>Excel:</td><td><span class='cantidad'>$______________</span></td></tr><tr><td>C.I:</td><td><span class='cantidad'>" + Cstr(CONTADOR1) + "</span></td></tr><tr><td>C.I:</td><td><span class='cantidad'>" + Cstr(CONTADOR2) + "</span></td></tr><tr><td>C.F:</td><td><span class='cantidad'>" + Cstr(CONTADOR3) + "</span></td></tr><tr><td>C.F:</td><td><span class='cantidad'>" + Cstr(CONTADOR4) + "</span></td></tr><tr><td>B&amp;N:</td><td><span class='number'>" + Cstr(rn) + "</span></td></tr><tr><td>Color:</td><td><span class='number'>" + Cstr(rc) + "</span></td></tr><tr><td>Total Diff:</td><td><span class='number'>" + Cstr(rtotal) + "</span></td></tr><tr><td>Copias Faltantes:</td><td><span class='number'>" + Cstr(faltante) + "</span></td></tr></tbody></table></div></div></body></html>"            ' escribimos otra linea de texto
	archivotexto.close	


	objshell.run "C:\Users\"+ LocalUser +"\Downloads\index.html"
	wait(1)
	''objshell.sendkeys "^P"
	''wait(0.2)
	''objshell.sendkeys "{ENTER}"
End Function


While (True)
	'--------------------------------AÑO-------------------------------
	If (Not FolderYear) Then
		

		
		If (Not oFSO.FolderExists(Disk + "\EMPLEADOS\JONATHAN\CORTES\" + sFolder)  ) Then
			
			oFSO.CreateFolder(Disk + "\EMPLEADOS\JONATHAN\CORTES\" + sFolder)
			
		End If
		
		
		
	End If
	'--------------------------------------------DESCARGAS-------------------------------
	
	If (  Minute( Time() )  =  0   ) Then
		
		oFSO.DeleteFolder "C:\Users\" + LOCALUSER + "\Downloads\*"
		
	End If
	'------------------------------MES---------------------------------------------------
	
	
	FolderMes = False
	If (Not FolderMes) Then
		
		
		
		
		
		
		FolderMes = False
		
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
		
		If Not oFSO.FolderExists ( Disk + "\EMPLEADOS\JONATHAN\CORTES\" + Cstr(iYear) + "\" + Cstr(bFolder) )Then
			oFSO.CreateFolder Disk + "\EMPLEADOS\JONATHAN\CORTES\" + Cstr(iYear) + "\" + bFolder
			
			
		End If
		
		
	End If
	'---------------------------RENAME MOVER-------------------------------------------------
	
	
	
	
	
	
	
	
	
	Select Case (aMes)
		Case 1: bMes= "ENERO"
		
		Case 2: bMes= "FEBRERO"
		
		Case 3: bMes= "MARZO"
		
		Case 4: bMes= "ABRIL"
		
		Case 5: bMes= "MAYO"
		
		Case 6: bMes= "JUNIO"
		
		Case 7: bMes= "JULIO"
		
		Case 8: bMes= "AGOSTO"
		
		Case 9: bMes= "SEPTIEMBRE"
		
		Case 10: bMes= "OCTUBRE"
		
		Case 11: bMes= "NOVIEMBRE"
		
		Case 12: bMes= "DICIEMBRE"
		
		
	End Select
	
	
		
	
	
	If iMes > 9 Then 
		bFolder = "J "&iMes&"-"&sMes
	End If
	If iMes < 10 Then
		bFolder = "J 0"&iMes&"-"&sMes
	End If
	
	
	
	If aMes > 9 Then 
		aFolder = "J "&aMes&"-"&bMes
	End If
	If aMes < 10 Then
		aFolder = "J 0"&aMes&"-"&bMes
	End If
	
	
	If ( oFSO.FileExists (Disk + "\EMPLEADOS\JONATHAN\CORTES\" + Cstr(iYear) + "\" + bFolder + "\RenameV3.vbs") ) Then
		
		oFSO.movefile Disk + "\EMPLEADOS\JONATHAN\CORTES\" + Cstr(iYear) + "\" + bFolder + "\RenameV3.vbs", Disk + "\EMPLEADOS\JONATHAN\CORTES\" + Cstr(iYear) + "\" + bFolder + "\"
		
	End If
	'-------------------------------------------------------------------------------------------------------
	
	'Create date variables
	Dim d,newfile,oldfile,dia,mes,ano,temp,a,b,c,pdf,excel
	
	d = date
	dia =	(Day(Date()))
	mes =	(Month(Date()))
	ano =	(Year(Date()))
	temp = ano - 2000
	a = dia
	b = mes
	If (dia < 10) Then
		a = "0" + Cstr(dia)
	End If
	
	If (mes < 10) Then
		b = "0" + Cstr(mes)
	End If
	
	excel = "Jonathan " + Cstr(a) + "-" + Cstr(b) + "-" + Cstr(temp) + ".xls"
	
	
	pdf = "Jonathan " + Cstr(a) + "-" + Cstr(b) + "-" + Cstr(temp)
	'-----------------------------------------DEBUG ZONE----------------------------------------------------
	
	
	
	
	
	
	
	
	
	
	
	
	
	'-------------------------------------RECYCLE CODE--------------------------------------------------
	
	
	Existe = 0
	
	
	
	
	
	
	If iMes > 9 Then 
		bFolder = "J "&iMes&"-"&sMes
	End If
	If iMes < 10 Then
		bFolder = "J 0"&iMes&"-"&sMes
	End If
	'---------------------------------------------------------------------------------------REPORTE-----------------------------------------------------------------------------------------
	
	condicion1 =  Disk + "\EMPLEADOS\JONATHAN\CORTES\" + Cstr(iYear) + "\" + bFolder + "\" + excel
	condicion2 =  Disk + "\EMPLEADOS\JONATHAN\CORTES\" + Cstr(iYear) + "\" + bFolder + "\" + pdf + ".pdf"
	condicion3 =  Disk + "\EMPLEADOS\JONATHAN\CORTES\" + Cstr(iYear) + "\" + bFolder + "\" + "Plantilla-Diaria-Por-Usuario.xls"
	
	path = "C:\Users\"+ LocalUser +"\Desktop\Plantilla-Diaria-Por-Usuario.xls"
	
	If (  oFSO.FileExists(condicion2) And Not SaidYes And Not oFSO.FileExists(condicion1) ) Then		
		
		
		Wscript.sleep 200
		'AutoSave By Javv 
		protocolo = objshell.popup("Iniciando Protocolo Stay AFK" & vbNewLine & "User: " + LocalUser & vbNewLine & "Disk: " + Disk,3,"Lazy System",68)
		
		If (protocolo = 6) Then
			z = objshell.popup ("3...",1,"Lazy System",64)
			x = objshell.popup ("2...",1,"Lazy System",64)
			c = objshell.popup ("1...",1,"Lazy System",64)
			WScript.sleep 1000
			
		If Not objfso.FileExists ("C:\Users\" + LOCALUSER + "\Downloads\Step 1.log") Then
			If objfso.FileExists("C:\Users\" + LOCALUSER + "\Downloads\Tinta.html") Then
				objfso.DeleteFile "C:\Users\" + LOCALUSER + "\Downloads\Tinta.html"
			End If
			
			If objfso.FileExists("C:\Users\" + LOCALUSER + "\Downloads\Config.html") Then
				objfso.DeleteFile "C:\Users\" + LOCALUSER + "\Downloads\Config.html"
			End If
			
			Set app = objshell.exec("C:\Program Files\Google\Chrome\Application\chrome.exe")		'ejecutamos el bloc de notas
			
			
			Wscript.sleep 300					'espera de dos segundos
			Objshell.appactivate "Nueva Pestaña"		'ponemos el foco en la ventana del bloc
			Wscript.sleep 300
			Objshell.sendkeys ImpresoraHP			'enviamos un mensaje con sendkeys
			Wscript.sleep 300
			Objshell.sendkeys "{ENTER}"			'luego del mensaje anterior, un ENTER
			Wscript.sleep 4000
			Objshell.sendkeys "^s"
			Wscript.sleep 500
			Objshell.sendkeys "Tinta.html" 
			Wscript.sleep 500
			Objshell.sendkeys "{ENTER}"
			
			While Not objfso.FileExists("C:\Users\" + LOCALUSER + "\Downloads\Tinta.html")
				WScript.sleep 100
			Wend

			
			WScript.sleep 1000
			objshell.sendkeys "^w"
			Wscript.sleep 500
			Objfso.CreateTextFile "C:\Users\" + LOCALUSER + "\Downloads\Step 1.log"
		End if
		If Not objfso.FileExists ("C:\Users\" + LOCALUSER + "\Downloads\Step 2.log") Then
			Set app2 = objshell.exec("C:\Program Files\Google\Chrome\Application\chrome.exe")
			
			
			Wscript.sleep 500
			
			Objshell.appactivate "Nueva Pestaña"		'ponemos el foco en la ventana del bloc
			Wscript.sleep 300
			Objshell.sendkeys "http://" + ImpresoraHP + "/info_configuration.html?tab=Home&menu=DevConfig"			'enviamos un mensaje con sendkeys
			Wscript.sleep 300
			Objshell.sendkeys "{ENTER}"			'luego del mensaje anterior, un ENTER
			Wscript.sleep 4000
			Objshell.sendkeys "^s"
			Wscript.sleep 500
			Objshell.sendkeys "{BS}" 
			Wscript.sleep 200
			Objshell.sendkeys "Config.html"
			Wscript.sleep 200
			Objshell.sendkeys "{ENTER}"
			
			
			While Not objfso.FileExists("C:\Users\" + LOCALUSER + "\Downloads\Config.html")
				WScript.sleep 100
			Wend
			WScript.sleep 1000
			objshell.sendkeys "^w"
			Objfso.CreateTextFile "C:\Users\" + LOCALUSER + "\Downloads\Step 2.log"
		End if	
			
			WScript.sleep 500
			
			GETCOUNT()
			
			SaidYes = True
			
		End If
	End If	'----------------------------------------------------------------------------------------------------------------------------------------------
	
	If (  oFSO.FileExists(condicion1) And oFSO.FileExists(condicion2) And Not Enviado And Not Protocolo = 7) Then	
		Dim para,asunto,mensaje,adjunto1,adjunto2,adjunto3,adjunto4,DefaultPath
		
		
		Resources = "C:\Users\" + LOCALUSER + "\Downloads\"
		
		
		DefaultPath = Disk + "\EMPLEADOS\JONATHAN\CORTES\"
		para = "ops65@hotmail.com; opsjra@gmail.com; ciee.empalme@live.com; jonathanvargas21@outlook.es"
		''para = "ciee.empalme@live.com"
		
		
		
		asunto = "Jonathan " + Cstr(a) + "-" + Cstr(b) + "-" + Cstr(temp)
		mensaje = ""
		adjunto1 =  DefaultPath + Cstr(iYear) + "\" + bFolder + "\" + excel
		adjunto2 =  DefaultPath + Cstr(iYear) + "\" + bFolder + "\" + pdf + ".pdf"
		adjunto3 =  Resources + "Tinta.html"
		adjunto4 =  Resources + "Config.html"
		
		envio = msgbox (("Enviar?") , 1)
		
		
		
		If (envio = 1) Then
			Set outlook = CreateObject("Outlook.Application")
			Set correo = outlook.CreateItem(olMailItem)
			correo.To = para 
			correo.Subject = asunto
			correo.Body = mensaje
			correo.Attachments.Add(adjunto1)
			correo.Attachments.Add(adjunto2)
			correo.Attachments.Add(adjunto3)
			correo.Attachments.Add(adjunto4)
			correo.Send
			Enviado = True
			objfso.DeleteFile "C:\Users\" + LOCALUSER + "\Downloads\*"
			Objshell.run "shutdown -s -t 120"	 
		Else
			Msgbox("Se cancelo el envio")
			
		End If
		
	End If	
	
	'--------------------------------------------------------------------------------------------------END-------------------------------------------------------------------------------------------------
	
	
	
	
	
	
	WScript.sleep 500
Wend



