' Create filesystem object and the folder object
Set fso = CreateObject("Scripting.FileSystemObject")  
 
Set objshell = CreateObject("WScript.Shell")

objshell.popup "SERVIDOR Iniciado Mod Kass",1

' INICIALIZADOR DE Variables
Set oFSO = CreateObject("Scripting.FileSystemObject")
Enviado = False
LOCALUSER = objShell.ExpandEnvironmentStrings("%UserProfile%")



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
	
If Not objfso.FileExists (LocalUser + "\Downloads\Step 3.log") Then
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
	
	If objfso.fileExists(LocalUser +"\Documents\"+ pdf + ".txt") Then
		
		objfso.deleteFile LocalUser +"\Documents\"+ pdf + ".txt"
	End If
	
	If objfso.fileExists(LocalUser +"\Documents\dat.txt") Then
		
		objfso.deleteFile LocalUser +"\Documents\predat.txt"
	End If
	
	objfso.createtextfile(LocalUser +"\Documents\predat.txt")	
	
	objshell.run LocalUser +"\Documents\predat.txt"
	
	wait(0.5)
	Objshell.sendkeys "^v"
	wait(0.5)
	Objshell.sendkeys "^g"
	wait(0.5)
	Objshell.sendkeys "^w"
	wait(0.5)
	
	Set archivotexto = objfso.opentextfile(LocalUser +"\Documents\predat.txt",1)	'abrimos el archivo
	For k=0 To 220 'cambiar despues a 100 y algo
		archivotexto.skipline
	Next
	
	contador3 = archivotexto.readline
	archivotexto.skipline
	archivotexto.skipline
	archivotexto.skipline					'leemos una linea, la primera							'saltamos una linea
	contador4 = archivotexto.readline					'leemos una linea, la tercera
	archivotexto.close
	
	objshell.run LocalUser +"\Documents\predat.txt"
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
	Objfso.CreateTextFile LocalUser + "\Downloads\Step 3.log"
End if
wait(0.2)
If Not objfso.FileExists (LocalUser + "\Downloads\Step 4.log") Then
	

	archivo = LocalUser + "\Downloads\test.txt"
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
	text.close
	'msgbox tmp
	caja =  Right(tmp,15)
	'msgbox caja
	caja2 = Left(Caja, 11)
	'msgbox caja2








	
	
	'Objfso.CreateTextFile LocalUser + "\Downloads\Step 4.log"
End if	

	Set archivotexto = objfso.opentextfile(LocalUser +"\Documents\dat.txt",1)	'abrimos el archivo
	contador1 = archivotexto.readline					'leemos una linea, la primera							'saltamos una linea
	contador2 = archivotexto.readline					'leemos una linea, la tercera
	archivotexto.close
	
	Set archivotexto = objfso.opentextfile(LocalUser +"\Documents\predat.txt",1)	'abrimos el archivo
	contador3 = archivotexto.readline					'leemos una linea, la primera							'saltamos una linea
	contador4 = archivotexto.readline					'leemos una linea, la tercera
	archivotexto.close



	'------------------------------------------------
	wait(0.5)
	objshell.run LocalUser +"\Desktop\Plantilla-Diaria-Por-Usuario.xls"
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
	
	rn = Cint(contador3 - contador1) 'Remember Change This variables
	rc = CInt(contador4 - contador2)
	rtotal = Cint(rn-rc)
	''MSGBOX " B&N: " + CONTADOR1 + " Color: " + CONTADOR2 & vbNewLine & vbNewLine & " B&N: " + CONTADOR3 + " Color: " + CONTADOR4 & vbNewLine & vbNewLine + "               " +  Cstr(rn) + "               Color: " + Cstr(rc) & vbNewLine & "         TotalDiff: " + Cstr(rtotal)

	Set objfso = createobject("scripting.filesystemobject")
	Set archivotexto = objfso.opentextfile(LocalUser +"\Downloads\index.html",2,true)    'creamos 
	archivotexto.writeline "<html><head><title>Ticket</title><style>body{font-family:Arial,sans-serif;font-size:14px;margin:0;padding:0}.ticket{background-color:#fff;border:1px solid #ccc;padding:20px;width:400px;height:500px;position:absolute;top:0;left:0;margin-top:20px;margin-left:20px;box-shadow:0 0 10px rgba(0,0,0,.2)}.ticket h1 {font-size: 24px;margin-top: 0;text-align: center;text-transform: uppercase;letter-spacing: 2px;animation: color-change 2s linear infinite;}@keyframes color-change{0%{color:#ff0000}50%{color:#00ff00}100%{color:#0000ff}} .info{margin-top:20px}.ticket .info table{width:100%;height:90%;border-collapse:collapse}.ticket .info table td{padding:10px 5px;border:1px solid #ccc}.ticket .info table td:first-child{font-weight:bold;width:120px}.ticket .info table td:last-child{text-align:right}.ticket .info table tr:last-child{font-weight:bold}.cantidad{font-size:110%; margin-right:100px}.number{font-size:110%; margin-right:120px}.excelcantidad{font-size:110%; margin-right:90px}</style></head><body><div class='ticket'><h1>" + pdf + "</h1><div class='info'><table><tbody><tr><td>Ciberplanet:</td><td><span class='cantidad'>" + Cstr(caja2) + "</span></td></tr><tr><td>Excel:</td><td><span class='excelcantidad'>$________</span></td></tr><tr><td>C.I:</td><td><span class='cantidad'>" + Cstr(CONTADOR1) + "</span></td></tr><tr><td>C.I:</td><td><span class='cantidad'>" + Cstr(CONTADOR2) + "</span></td></tr><tr><td>C.F:</td><td><span class='cantidad'>" + Cstr(CONTADOR3) + "</span></td></tr><tr><td>C.F:</td><td><span class='cantidad'>" + Cstr(CONTADOR4) + "</span></td></tr><tr><td>B&amp;N:</td><td><span class='number'>" + Cstr(rn) + "</span></td></tr><tr><td>Color:</td><td><span class='number'>" + Cstr(rc) + "</span></td></tr><tr><td>Total Diff:</td><td><span class='number'>" + Cstr(rtotal) + "</span></td></tr></tbody></table></div></div></body></html>"
	archivotexto.close	


	objshell.run LocalUser +"\Downloads\index.html"
	wait(1)
	''objshell.sendkeys "^P"
	''wait(0.2)
	''objshell.sendkeys "{ENTER}"
End Function


While (True)
	'--------------------------------AÑO-------------------------------
	If (Not oFSO.FolderExists(Disk + "\EMPLEADOS\Kassandra\CORTES\" + sFolder) ) Then
		
	
		oFSO.CreateFolder (Disk + "\EMPLEADOS\Kassandra\CORTES\" + sFolder)
			
		
	End If
	'--------------------------------------------DESCARGAS-------------------------------
	

	'Deleted


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
			bFolder = "K "&iMes&"-"&sMes
		End If
		If iMes < 10 Then
			bFolder = "K 0"&iMes&"-"&sMes
			
		End If
		
		If Not oFSO.FolderExists ( Disk + "\EMPLEADOS\Kassandra\CORTES\" + Cstr(iYear) + "\" + Cstr(bFolder) )Then
			oFSO.CreateFolder Disk + "\EMPLEADOS\Kassandra\CORTES\" + Cstr(iYear) + "\" + bFolder
			
			
		End If
		
		
	End If
	'---------------------------RENAME MOVER-------------------------------------------------
	
		'Deleted

	'------------------------------------------------------------------------------------------------
	
	'Create date variables
	Dim d,newfile,oldfile,dia,mes,ano,temp,a,b,c,pdf,excel 'obsolete
	
	
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
	
	excel = "Kassandra " + Cstr(a) + "-" + Cstr(b) + "-" + Cstr(temp) + ".xls"
	
	
	pdf = "Kassandra " + Cstr(a) + "-" + Cstr(b) + "-" + Cstr(temp)
	'-----------------------------------------DEBUG ZONE----------------------------------------------------
	
	
	

	
	
	
	
	
	
	
	
	
	'-------------------------------------RECYCLE CODE--------------------------------------------------
	
	
	Existe = 0
	
	
	
	
	
	
	If iMes > 9 Then 
		bFolder = "K "&iMes&"-"&sMes
	End If
	If iMes < 10 Then
		bFolder = "K 0"&iMes&"-"&sMes
	End If
	'---------------------------------------------------------------------------------------REPORTE-----------------------------------------------------------------------------------------
	
	condicion1 =  Disk + "\EMPLEADOS\Kassandra\CORTES\" + Cstr(iYear) + "\" + bFolder + "\" + excel
	condicion2 =  Disk + "\EMPLEADOS\Kassandra\CORTES\" + Cstr(iYear) + "\" + bFolder + "\" + pdf + ".pdf"
	condicion3 =  Disk + "\EMPLEADOS\Kassandra\CORTES\" + Cstr(iYear) + "\" + bFolder + "\" + "Plantilla-Diaria-Por-Usuario.xls"
	
	path = LocalUser +"\Desktop\Plantilla-Diaria-Por-Usuario.xls"
	
	If (  oFSO.FileExists(condicion2) And Not SaidYes And Not oFSO.FileExists(condicion1) ) Then		
		
		
		Wait (0.2)
		'AutoSave By Javv 
		protocolo = objshell.popup("Iniciando Protocolo Stay AFK" & vbNewLine & "User: " + LocalUser & vbNewLine & "Disk: " + Disk,3,"Lazy System",68)
		
		If (protocolo = 6) Then
			
			z = objshell.popup ("3...2...1...",1,"Lazy System",64)
			Wait(1)
			
		If Not objfso.FileExists (LocalUser + "\Downloads\Step 1.log") Then
			If objfso.FileExists(LocalUser + "\Downloads\Tinta.html") Then
				objfso.DeleteFile LocalUser + "\Downloads\Tinta.html"
			End If
			
			If objfso.FileExists(LocalUser + "\Downloads\Config.html") Then
				objfso.DeleteFile LocalUser + "\Downloads\Config.html"
			End If
			
			Set app = objshell.exec("C:\Program Files\Google\Chrome\Application\chrome.exe")		'ejecutamos el bloc de notas
			
			
			Wait(0.3) 					'espera de dos segundos
			Objshell.appactivate "Nueva Pestaña"		'ponemos el foco en la ventana del bloc
			Wait(0.3) 
			Objshell.sendkeys ImpresoraHP			'enviamos un mensaje con sendkeys
			Wait(0.3)
			Objshell.sendkeys "{ENTER}"			'luego del mensaje anterior, un ENTER
			Wait(4) 
			Objshell.sendkeys "^s"
			Wait(0.5) 
			Objshell.sendkeys "Tinta.html" 
			Wait(0.5)
			Objshell.sendkeys "{ENTER}"
			
			While Not objfso.FileExists(LocalUser + "\Downloads\Tinta.html")
				Wait(0.1) 
			Wend

			
			Wait(1) 
			objshell.sendkeys "^w"
			Wait(0.5) 
			Objfso.CreateTextFile LocalUser + "\Downloads\Step 1.log"
		End if
		If Not objfso.FileExists (LocalUser + "\Downloads\Step 2.log") Then
			Set app2 = objshell.exec("C:\Program Files\Google\Chrome\Application\chrome.exe")
			
			
			Wait(0.5) 
			
			Objshell.appactivate "Nueva Pestaña"		'ponemos el foco en la ventana del bloc
			Wait(0.3) 
			Objshell.sendkeys "http://" + ImpresoraHP + "/info_configuration.html?tab=Home&menu=DevConfig"			'enviamos un mensaje con sendkeys
			Wait(0.3) 
			Objshell.sendkeys "{ENTER}"			'luego del mensaje anterior, un ENTER
			Wait(4) 
			Objshell.sendkeys "^s"
			Wait(0.5) 
			Objshell.sendkeys "{BS}" 
			Wait(0.2) 
			Objshell.sendkeys "Config.html"
			Wait(0.2) 
			Objshell.sendkeys "{ENTER}"
			
			
			While Not objfso.FileExists(LocalUser + "\Downloads\Config.html")
				Wait(0.1)
			Wend
			Wait(1) 
			objshell.sendkeys "^w"
			Objfso.CreateTextFile LocalUser + "\Downloads\Step 2.log"
		End if	
			
			Wait(0.5) 
			
			GETCOUNT()
			
			SaidYes = True
			
		End If
	End If	'----------------------------------------------------------------------------------------------------------------------------------------------
	
	If (  oFSO.FileExists(condicion1) And oFSO.FileExists(condicion2) And Not Enviado And Not Protocolo = 7) Then	
		Dim para,asunto,mensaje,adjunto1,adjunto2,adjunto3,adjunto4,DefaultPath
		
		
		Resources = LocalUser + "\Downloads\"
		
		
		DefaultPath = Disk + "\EMPLEADOS\Kassandra\CORTES\"
		para = "ops65@hotmail.com; opsjra@gmail.com; jonathanvargas21@outlook.es"
		''para = "ciee.empalme@live.com"
		
		
		
		asunto = "Kassandra " + Cstr(a) + "-" + Cstr(b) + "-" + Cstr(temp)
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
			objfso.DeleteFile LocalUser + "\Downloads\*"
			Objshell.run "shutdown -s -t 120"	 
		Else
			Msgbox("Se cancelo el envio")
			
		End If
		
	End If	
	
	'--------------------------------------------------------------------------------------------------END-------------------------------------------------------------------------------------------------
	
	
	
	
	
	
	Wait (0.5)
Wend



