
Set objfso = createobject("scripting.filesystemobject")
Set objshell = createobject("wscript.shell")
Enviado = False
LOCALUSER = objShell.ExpandEnvironmentStrings("%USERPROFILE%")
Disk = "D:"
SaidYes = False

FolderYear = Cstr(Year(Date()))

MesActual = Month(Date())

MesAnterior = (    CInt( MesActual )  - 1    )

ImpresoraKonica = "192.168.0.6"
ImpresoraHP = "192.168.0.127"

UserPath = "\EMPLEADOS\JONATHAN\CORTES\"

objshell.popup "SERVIDOR Iniciado",1

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

            objfso.deleteFile LocalUser +"\Documents\dat.txt"
        End If

        objfso.createtextfile(LocalUser +"\Documents\dat.txt")

        objshell.run LocalUser +"\Documents\dat.txt"

        wait(0.5)
        Objshell.sendkeys "^v"
        wait(0.5)
        Objshell.sendkeys "^g"
        wait(0.5)
        Objshell.sendkeys "^w"
        wait(0.5)

        Set archivotexto = objfso.opentextfile(LocalUser +"\Documents\dat.txt",1)	'abrimos el archivo
        For k=0 To 220
            archivotexto.skipline
            Next

            contador3 = archivotexto.readline
            archivotexto.skipline
            archivotexto.skipline
            archivotexto.skipline
            contador4 = archivotexto.readline
            archivotexto.close

            objshell.run LocalUser +"\Documents\dat.txt"
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


            ' archivo = LocalUser + "\Downloads\TEST.txt"

            Set text = objfso.opentextfile (LocalUser + "\Downloads\TEST.txt",1)
            for i=0 to 1
                text.skipline
                Next

                While (not found)
                    LineaLeida = text.readline
                    If InStr (LineaLeida, "total") then
                        found = true
                    End if
                WEND


                caja =  Right(LineaLeida,15)

                caja2 = Left(Caja, 11)


                found = false

                
                CopiasVendidas = 0
                While (not found)


                    LineaLeida = text.readline
                    If InStr (LineaLeida, "copia") then



                        copy =  Left( LineaLeida , 387)

                        copy2 = Right( copy , 6)

                        copy3 = Left (copy2 , 4)

                        CopiasVendidas = CopiasVendidas + copy3

                    End if

                    If InStr (LineaLeida, "Eventos") then
                        found = true
                    end if
                WEND





                Objfso.CreateTextFile LocalUser + "\Downloads\Step 4.log"
            End if

            Set archivotexto = objfso.opentextfile(LocalUser +"\Documents\dat2.txt",1)
            contador1 = archivotexto.readline
            contador2 = archivotexto.readline
            archivotexto.close

            Set archivotexto = objfso.opentextfile(LocalUser +"\Documents\dat.txt",1)
            contador3 = archivotexto.readline
            contador4 = archivotexto.readline
            archivotexto.close



            '------------------------------------------------
            wait(0.5)
            objshell.run LocalUser +"\Desktop\Plantilla-Diaria-Por-Usuario.xls"

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

                TotalCount = Cint(contador3 - contador1)
            	ColorCount = CInt(contador4 - contador2)
                TotalDiff = Cint(TotalCount - ColorCount)

                Faltante = Cint(TotalDiff - CopiasVendidas)

                Set archivotexto = objfso.opentextfile(LocalUser +"\Downloads\index.html",2,true)
				archivotexto.writeline "<html><head><title>Ticket</title><style>body{font-family:Arial,sans-serif;font-size:14px;margin:0;padding:0}.ticket{background-color:#fff;border:1px solid #ccc;padding:20px;width:400px;position:absolute;top:0;left:0;margin-top:20px;margin-left:20px;box-shadow:0 0 10px rgba(0,0,0,.2)}.ticket h1{font-size:40px;margin-top:0;text-align:center;text-transform:uppercase;letter-spacing:2px}.ticket .info{margin-top:20px}.ticket .info table{width:100%;border-collapse:collapse}.ticket .info table td{padding:10px 5px;border:1px solid #ccc}.ticket .info table td:first-child{font-weight:bold;width:120px}.ticket .info table td:last-child{text-align:right}.ticket .info table tr:last-child{font-weight:bold}.cantidad{font-size:200%;margin-right:100px}.number{font-size:200%;margin-right:120px}.remove-button{background-color:#ff6347;border:none;color:#fff;padding:10px;font-size:16px;cursor:pointer;margin-top:20px}.remove-button:hover{background-color:#ff2e00}</style></head><body><div class='ticket'><h1>" + Cstr(PDF) + "</h1><div class='info'><table><tbody><tr><td>Ciberplanet:</td><td><span class='cantidad'>" + Cstr(caja2) + "</span></td></tr><tr><td>Excel:</td><td><span class='cantidad'>$______________</span></td></tr><tr><td>C.I:</td><td><span class='cantidad'>" + Cstr(CONTADOR1) + "</span></td></tr><tr><td>C.I:</td><td><span class='cantidad'>" + Cstr(CONTADOR2) + "</span></td></tr><tr><td>C.F:</td><td><span class='cantidad'>" + Cstr(CONTADOR3) + "</span></td></tr><tr><td>C.F:</td><td><span class='cantidad'>" + Cstr(CONTADOR4) + "</span></td></tr><tr><td>B&amp;N:</td><td><span class='number'>" + Cstr(BNCopy) + "</span></td></tr><tr><td>Color:</td><td><span class='number'>" + Cstr(rc) + "</span></td></tr><tr><td>Total Diff:</td><td><span class='number'>" + Cstr(rtotal) + "</span></td></tr><tr><td>CopiasVendidas Faltantes:</td><td><span class='number'>" + Cstr(faltante) + "</span></td></tr></tbody></table></div></div><script>function removeCopiasVendidasFaltantes(){var table=document.getElementsByTagName('table')[0];table.deleteRow(table.rows.length-1);}document.addEventListener('keydown',function(event){if(event.keyCode===74){removeCopiasVendidasFaltantes();}});</script></body></html>"
    			archivotexto.close


    objshell.run LocalUser +"\Downloads\index.html"
    wait(1)
    ''objshell.sendkeys "^P"
    ''wait(0.2)
    ''objshell.sendkeys "{ENTER}"
End Function


While (True)
    '--------------------------------AÑO-------------------------------
    If (Not FolderYear) Then



        If (Not ObjFso.FolderExists(Disk + UserPath + FolderYear)  ) Then

        ObjFso.CreateFolder(Disk + UserPath + FolderYear)

        End If



    End If
    '------------------------------MES---------------------------------------------------


    FolderMes = False
    If (Not FolderMes) Then






        FolderMes = False

        Select Case (MesActual)
         Case 1: StrMonth= "ENERO"

         Case 2: StrMonth= "FEBRERO"

         Case 3: StrMonth= "MARZO"

         Case 4: StrMonth= "ABRIL"

         Case 5: StrMonth= "MAYO"

         Case 6: StrMonth= "JUNIO"

         Case 7: StrMonth= "JULIO"

         Case 8: StrMonth= "AGOSTO"

         Case 9: StrMonth= "SEPTIEMBRE"

         Case 10: StrMonth= "OCTUBRE"

         Case 11: StrMonth= "NOVIEMBRE"

         Case 12: StrMonth= "DICIEMBRE"

        End Select





        If MesActual > 9 Then
            bFolder = "J "&MesActual&"-"&StrMonth
        End If
        If MesActual < 10 Then
            bFolder = "J 0"&MesActual&"-"&StrMonth

        End If

        If Not ObjFso.FolderExists ( Disk + UserPath + Cstr(Year(Date())) + "\" + Cstr(bFolder) )Then
            ObjFso.CreateFolder Disk + UserPath + Cstr(Year(Date())) + "\" + bFolder


        End If


    End If
    '---------------------------RENAME MOVER-------------------------------------------------









    Select Case (MesAnterior)
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





    If MesActual > 9 Then
        bFolder = "J "&MesActual&"-"&StrMonth
    End If
    If MesActual < 10 Then
        bFolder = "J 0"&MesActual&"-"&StrMonth
    End If



    If MesAnterior > 9 Then
        aFolder = "J "&MesAnterior&"-"&bMes
    End If
    If MesAnterior < 10 Then
        aFolder = "J 0"&MesAnterior&"-"&bMes
    End If


    If ( ObjFso.FileExists (Disk + UserPath + Cstr(Year(Date())) + "\" + bFolder + "\RenameV3.vbs") ) Then

        ObjFso.movefile Disk + UserPath + Cstr(Year(Date())) + "\" + bFolder + "\RenameV3.vbs", Disk + UserPath + Cstr(Year(Date())) + "\" + bFolder + "\"

    End If
    '-------------------------------------------------------------------------------------------------------

    dia =	(Day(Date()))
    mes =	(Month(Date()))
    shortyear = (Year(Date())) - 2000

    If (dia < 10) Then
        a = "0" + Cstr(dia)
	
	Else
		a = Cstr(dia)
    End If

    If (mes < 10) Then
        b = "0" + Cstr(mes)
	Else
		b = Cstr(mes)
    End If

    excel = "Jonathan " + Cstr(a) + "-" + Cstr(b) + "-" + Cstr(shortyear) + ".xls"


    pdf = "Jonathan " + Cstr(a) + "-" + Cstr(b) + "-" + Cstr(shortyear)
    '-----------------------------------------DEBUG ZONE----------------------------------------------------













    '-------------------------------------RECYCLE CODE--------------------------------------------------


    Existe = 0






    If MesActual > 9 Then
        bFolder = "J "&MesActual&"-"&StrMonth
    End If
    If MesActual < 10 Then
        bFolder = "J 0"&MesActual&"-"&StrMonth
    End If
    '---------------------------------------------------------------------------------------REPORTE-----------------------------------------------------------------------------------------

    condicion1 =  Disk + UserPath + Cstr(Year(Date())) + "\" + bFolder + "\" + excel
    condicion2 =  Disk + UserPath + Cstr(Year(Date())) + "\" + bFolder + "\" + pdf + ".pdf"
    condicion3 =  Disk + UserPath + Cstr(Year(Date())) + "\" + bFolder + "\" + "Plantilla-Diaria-Por-Usuario.xls"

    path = LocalUser +"\Desktop\Plantilla-Diaria-Por-Usuario.xls"

    If (  ObjFso.FileExists(condicion2) And Not SaidYes And Not ObjFso.FileExists(condicion1) ) Then


        Wscript.sleep 200

        'AutoSave By Javv

        protocolo = objshell.popup("Iniciando Protocolo Stay AFK" & vbNewLine & "User: " + Right(LocalUser,5) & vbNewLine & "Disk: " + Disk,3,"Lazy System",68)

        If (protocolo = 6) Then
            z = objshell.popup ("3...2...1...",1,"Lazy System",64)

            WScript.sleep 1000

            If Not objfso.FileExists (LocalUser + "\Downloads\Step 1.log") Then
                If objfso.FileExists(LocalUser + "\Downloads\Tinta.html") Then
                    objfso.DeleteFile LocalUser + "\Downloads\Tinta.html"
                End If

                If objfso.FileExists(LocalUser + "\Downloads\Config.html") Then
                    objfso.DeleteFile LocalUser + "\Downloads\Config.html"
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

                While Not objfso.FileExists(LocalUser + "\Downloads\Tinta.html")
                    WScript.sleep 100
                Wend


                WScript.sleep 1000
                objshell.sendkeys "^w"
                Wscript.sleep 500
                Objfso.CreateTextFile LocalUser + "\Downloads\Step 1.log"
            End if
            If Not objfso.FileExists (LocalUser + "\Downloads\Step 2.log") Then
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


                While Not objfso.FileExists(LocalUser + "\Downloads\Config.html")
                    WScript.sleep 100
                Wend
                WScript.sleep 1000
                objshell.sendkeys "^w"
                Objfso.CreateTextFile LocalUser + "\Downloads\Step 2.log"
            End if

            WScript.sleep 500

            GETCOUNT()

            SaidYes = True

        End If
    End If	'----------------------------------------------------------------------------------------------------------------------------------------------

    If (  ObjFso.FileExists(condicion1) And ObjFso.FileExists(condicion2) And Not Enviado And Not Protocolo = 7) Then
        


        Resources = LocalUser + "\Downloads\"


        DefaultPath = Disk + UserPath
        para = "ops65@hotmail.com; opsjra@gmail.com; ciee.empalme@live.com; jonathanvargas21@outlook.es"
        ''para = "ciee.empalme@live.com"



        asunto = "Jonathan " + Cstr(a) + "-" + Cstr(b) + "-" + Cstr(shortyear)
        mensaje = ""
        adjunto1 =  DefaultPath + Cstr(Year(Date())) + "\" + bFolder + "\" + excel
        adjunto2 =  DefaultPath + Cstr(Year(Date())) + "\" + bFolder + "\" + pdf + ".pdf"
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






    WScript.sleep 500
Wend



