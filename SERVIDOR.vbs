Set objfso = createobject("scripting.filesystemobject")
Set objshell = createobject("wscript.shell")
Enviado = False
LOCALUSER = objShell.ExpandEnvironmentStrings("%USERPROFILE%")
Disk = "D:"
SaidYes = False
Error = false
Found = false
FolderMes = False

FolderYear = Cstr(Year(Date()))

MesActual = Month(Date())

MesAnterior = (    CInt( MesActual )  - 1    )

ImpresoraKonica = "192.168.0.6"
ImpresoraHP = "192.168.0.127"

UserPath = "\EMPLEADOS\JONATHAN\CORTES\"

'objshell.popup "SERVIDOR Iniciado",1

CopiasVendidas = 0

Dim PDF
Dim Caja2
Dim ExcelValor 
Dim Contador1
Dim Contador2
Dim Contador3
Dim Contador4
Dim TotalCount
Dim ColorCount
Dim TotalDiff
Dim Faltante


'-----------------------------------------FUNCTIONS----------------------------------------------------
Function wait(x)
    Wscript.sleep x * 1000
End Function

 Function SetupPrinter()  'Algoritmo para la impresora 
    Objshell.sendkeys "^p"


    wait(3)
    for i=0 to 5
    Objshell.sendkeys "{TAB}"
    next
    wait(0.1)
    Objshell.sendkeys "{ENTER}"
    wait(0.4)
    Objshell.sendkeys "{DOWN}"
    wait(0.1)
    Objshell.sendkeys "{ENTER}"
    wait(1.5)
    for i=0 to 3
    Objshell.sendkeys "{TAB}"
    wait(0.1)
    next
    wait(0.2)
    Objshell.sendkeys "{ENTER}"
    For i=0 To 3
        wait(0.1)
    Objshell.sendkeys "{TAB}"
    Next 
    wait(0.1)
    Objshell.sendkeys "{ENTER}"
    wait(0.1)
    Objshell.sendkeys "{DOWN}"
    wait(0.1)
    Objshell.sendkeys "{ENTER}"
    wait(0.7)
    Objshell.sendkeys "^a"
    wait(0.2)
    Objshell.sendkeys "{BS}"
    wait(0.2)
    Objshell.sendkeys "40"
    for i=1 to 16
        wait(0.05)
        Objshell.sendkeys "{TAB}"

    next
        Objshell.sendkeys "{ENTER}"
    wait (21)

    Objshell.sendkeys "^p"



    wait(3)
    for i=0 to 5
    Objshell.sendkeys "{TAB}"
    next
    wait(0.1)
    Objshell.sendkeys "{ENTER}"
    wait(0.4)
    Objshell.sendkeys "{UP}"
    wait(0.1)
    Objshell.sendkeys "{ENTER}"
    wait(1.5)
    for i=0 to 3
    Objshell.sendkeys "{TAB}"
    wait(0.1)
    next
    wait(0.2)
    Objshell.sendkeys "{ENTER}"
    For i=0 To 3
        wait(0.1)
    Objshell.sendkeys "{TAB}"
    Next 
    wait(0.1)
    Objshell.sendkeys "{ENTER}"
    wait(0.1)
    Objshell.sendkeys "{UP}"
    wait(0.1)
    Objshell.sendkeys "{ENTER}"
    wait(0.1)
    For i = 0 To 4
        wait(0.1)
    Objshell.sendkeys "{TAB}"
    Next 
    wait(0.3)
    Objshell.sendkeys "{ENTER}"

End Function

Function getcount() 

    If Not objfso.FileExists (LocalUser + "\Downloads\Step 3.log") And not error Then
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
        
        On error resume Next
        For k=0 To 220
            archivotexto.skipline 'change this
        Next

        CatchError("Vuelve a correr el script, La pagina de los contadores no cargo correctamente")

        On error goto 0

            contador3 = archivotexto.readline
            archivotexto.skipline
            archivotexto.skipline
            archivotexto.skipline
            contador4 = archivotexto.readline
            archivotexto.close

            objshell.run LocalUser + "\Documents\dat.txt"
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
        


            ' archivo = LocalUser + "\Downloads\TEST.txt"

            Set text = objfso.opentextfile (LocalUser + "\Downloads\TEST.txt",1)

            k = 0
            While (not found)
                LineaLeida = text.readline
                If InStr (LineaLeida, "Recaudaci") then
                    k = k + 1
                    if (k = 3) then
                        found = true
                    end if

                End if
            WEND
		LineaLeida = Replace (LineaLeida," ","")
		
                                
                    for l = 1 to 100
                        texto = Left( LineaLeida , l)
                        lastcharacter = Right( texto , 1)
                        
                        
                        If lastcharacter = "$" Then
			tmp = l
                            

                        End if
			If lastcharacter = "." Then
			tmp2 = l
                    texto = (Left( LineaLeida , l + 2) )
                    caja2 = Right ( texto , 3 + (tmp2 - tmp) )
				     l = 100
             End if

                        Next
            If InStr(caja2, ".00") Then
            caja2 = Replace(caja2,".00","")
            End if
	

            found = false

            
            While (not found)


                LineaLeida = text.readline
                If InStr (LineaLeida, "copia") then


                    ' copy =  Left( LineaLeida , 500)

                    TextResume = Replace ( Left( LineaLeida , 500) ," ","")
                    
                    for p = 1 to 100
                        copy = Left( TextResume, p)
                        copy2 = Right( copy , 1)
                        
                        
                        If copy2 = "$" Then
                            copy = (Left( TextResume , p - 1) )
                            copy2 = Right ( copy , 3 )
                            p = 100
                            

                        End if

                        Next



                        

                        If VarType( Cint(copy2) )  = 2 Then
                        CopiasVendidas = CopiasVendidas + Cint(Copy2)
                            
                        End If

                    End if

                    If InStr (LineaLeida, "Eventos") then
                        found = true
                    end if
                    
                WEND
                
                On error resume Next
                set archivo = objfso.opentextfile(LocalUser + "\Documents\ExcelValue.txt")
                CatchError("No sabemos el total de excel, ve a excel y clickea el boton GetSheetValue")
                ExcelValor = archivo.readline
                archivo.close
                On error goto 0


            

            Set archivotexto = objfso.opentextfile(LocalUser +"\Documents\dat2.txt",1)
            contador1 = archivotexto.readline
            contador2 = archivotexto.readline
            archivotexto.close

            Set archivotexto = objfso.opentextfile(LocalUser +"\Documents\dat.txt",1)
            contador3 = archivotexto.readline
            contador4 = archivotexto.readline
            archivotexto.close

            Set archivotexto = objfso.opentextfile(LocalUser +"\Documents\Folder.txt",2,true)
            archivotexto.writeline Cstr(currentFolder)
            archivotexto.close

            '------------------------------------------------
            wait(0.5)
            objshell.run 

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


   
End Function

Function CatchError(Solution)
    
    If (err <> 0) Then
    set ErrorPage = objfso.opentextfile(LocalUser +"\Downloads\ERROR.html",2,true)
    ErrorPage.writeline "<html><head><title>Error</title><style>body{background-color:#252525;font-family:Arial,sans-serif;}.error-container{background-color:transparent;border-radius:10px;box-shadow:0 2px 4px rgba(0,0,0,0.1);margin:50px auto;max-width:400px;padding:30px;text-align:center;}h1{font-size:4rem;margin-bottom:20px;text-transform:uppercase;color:#fff;text-shadow:0 0 10px #fff,0 0 20px #fff,0 0 30px #fff,0 0 40px #0ff,0 0 70px #0ff,0 0 80px #0ff,0 0 100px #0ff,0 0 150px #0ff;animation:neon 1.5s ease-in-out infinite alternate;-webkit-animation:neon 1.5s ease-in-out infinite alternate;-moz-animation:neon 1.5s ease-in-out infinite alternate;-o-animation:neon 1.5s ease-in-out infinite alternate;}p{color:#999;font-size:21px;margin-bottom:30px;}.back-btn{background-color:#2196f3;border:none;border-radius:5px;color:white;font-size:16px;padding:10px 20px;text-decoration:none;transition:background-color 0.2s ease-in-out;}.back-btn:hover{background-color:#1976d2;}@keyframes neon{from{text-shadow:0 0 10px #fff,0 0 20px #fff,0 0 30px #fff,0 0 40px #0ff,0 0 70px #0ff,0 0 80px #0ff,0 0 100px #0ff,0 0 150px #0ff;}to{text-shadow:0 0 10px #fff,0 0 20px #fff,0 0 30px #fff,0 0 40px #f0f,0 0 70px #f0f,0 0 80px #f0f,0 0 100px #f0f,0 0 150px #f0f;}}</style></head><body><div class='error-container'><h1>Error " + Cstr(err.Number) + "</h1><p>" + Cstr(err.DescriptioN) + "<br>" + "How to fix: " + Solution + "</p><a href='#' class='back-btn'>Volver</a></div></body></html>"
    objshell.run LocalUser +"\Downloads\ERROR.html"
    Error = true
    End if
    wait(1)
End Function

'----------------------- Error stages ----------------------------

On error Resume next 
ObjFso.CreateTextFile(Disk + UserPath + "asd.txt")

CatchError("Conecte el Disco y asegurese de que la ruta " + Disk + UserPath + " Existe ó en su defecto cambie el Path")

On error goto 0


'----------------------- Error stages ----------------------------

On error Resume next

Set ffffff = objfso.opentextfile (LocalUser + "\Downloads\TEST.txt",1)

CatchError("Asegurese de que el archivo TXT que contiene el reporte del dia este en la carpeta de Descargas")

On error goto 0

'----------------------- Error stages ----------------------------

On error Resume next 

Set ffffff = objfso.opentextfile(LocalUser +"\Documents\dat2.txt",1)

CatchError("Asegurese de que el archivo TXT que contiene los contadores iniciales este en la carpeta de Documentos")

On error goto 0
'----------------------- Error stages ----------------------------




If Not error then
ObjFso.DeleteFile(Disk + UserPath + "asd.txt")
End if



'Inicio del flujo del codigo


While (Not error)
    '--------------------------------AÑO-------------------------------
    If (Not FolderYear) Then



        If (Not ObjFso.FolderExists(Disk + UserPath + FolderYear)  ) Then

            ObjFso.CreateFolder(Disk + UserPath + FolderYear)

        End If



    End If
    '------------------------------MES---------------------------------------------------


    
    If (Not FolderMes) Then






        

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
            FolderMonth = "J "&MesActual&"-"&StrMonth
        Else
            FolderMonth = "J 0"&MesActual&"-"&StrMonth
        End If

        If Not ObjFso.FolderExists ( Disk + UserPath + FolderYear + "\" + Cstr(FolderMonth) )Then
            ObjFso.CreateFolder Disk + UserPath + FolderYear + "\" + FolderMonth

            FolderMes = True
        End If


    End If
    '-------------------------------------------------------------------------------------------------------

    dia =	(Day(Date()))
    mes =	(Month(Date()))
    shortyear = (Year(Date())) - 2000

    If (dia < 10) Then
        DayNumber = "0" + Cstr(dia)

    Else
        DayNumber = Cstr(dia)
    End If

    If (mes < 10) Then
        MonthNumber = "0" + Cstr(mes)
    Else
        MonthNumber = Cstr(mes)
    End If

    excel = "Jonathan " + Cstr(DayNumber) + "-" + Cstr(MonthNumber) + "-" + Cstr(shortyear) + ".xls"

    currentFolder = folderYear + "\" + FolderMonth

    pdf = "Jonathan " + Cstr(DayNumber) + "-" + Cstr(MonthNumber) + "-" + Cstr(shortyear)


    

    '-------------------------------------RECYCLE CODE--------------------------------------------------

    If MesActual > 9 Then
        FolderMonth = "J "&MesActual&"-"&StrMonth
    Else
        FolderMonth = "J 0"&MesActual&"-"&StrMonth

    End If
    '---------------------------------------------------------------------------------------REPORTE-----------------------------------------------------------------------------------------

    ExisteElExcel =  Disk + UserPath + FolderYear + "\" + FolderMonth + "\" + excel
    ExisteElPDF =  Disk + UserPath + FolderYear + "\" + FolderMonth + "\" + pdf + ".pdf"
    
    PlantillaDeUsuario = LocalUser +  "\Desktop\Plantilla-Diaria-Por-Usuario.xls"

    If (  ObjFso.FileExists(ExisteElPDF) And Not SaidYes And Not ObjFso.FileExists(ExisteElExcel) ) Then


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


                Wscript.sleep 300					
                Objshell.appactivate "Nueva Pestaña"		
                Wscript.sleep 300
                Objshell.sendkeys ImpresoraHP			
                Wscript.sleep 300
                Objshell.sendkeys "{ENTER}"			
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

            On error resume next

            Set fff = objfso.opentextfile (LocalUser + "\Downloads\Tinta.html")

            CatchError("Asegurate de que el archivo Tinta.html existe de lo contrario borra el step1.log")

            On error goto 0
            
            
            If Not objfso.FileExists (LocalUser + "\Downloads\Step 2.log") And not error Then
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

            On error resume next

            Set fff = objfso.opentextfile (LocalUser + "\Downloads\Config.html")

            CatchError("Asegurate de que el archivo Config.html existe de lo contrario borra el step1.log")

            On error goto 0


            WScript.sleep 500

            GETCOUNT()

            SaidYes = True

        End If
    End If	'----------------------------------------------------------------------------------------------------------------------------------------------

    If (  ObjFso.FileExists(ExisteElExcel) And ObjFso.FileExists(ExisteElPDF) And Not Enviado And Not Protocolo = 7) Then

        Set archivotexto = objfso.opentextfile(LocalUser +"\Downloads\index.html",2,true)
        archivotexto.writeline "<html><head><title>Ticket</title><style>body{font-family:Arial,sans-serif;font-size:14px;margin:0;padding:0}.ticket{background-color:#fff;border:1px solid #ccc;padding:20px;width:400px;position:absolute;top:0;left:0;margin-top:20px;margin-left:20px;box-shadow:0 0 10px rgba(0,0,0,.2)}.ticket h1{font-size:40px;margin-top:0;text-align:center;text-transform:uppercase;letter-spacing:2px}.ticket .info{margin-top:20px}.ticket .info table{width:100%;border-collapse:collapse}.ticket .info table td{padding:10px 5px;border:1px solid #ccc}.ticket .info table td:first-child{font-weight:bold;width:120px}.ticket .info table td:last-child{text-align:right}.ticket .info table tr:last-child{font-weight:bold}.cantidad{font-size:200%;margin-right:100px}.number{font-size:200%;margin-right:120px}.remove-button{background-color:#ff6347;border:none;color:#fff;padding:10px;font-size:16px;cursor:pointer;margin-top:20px}.remove-button:hover{background-color:#ff2e00}</style></head><body><div class='ticket'><h1>" + Cstr(PDF) + "</h1><div class='info'><table><tbody><tr><td>Ciberplanet:</td><td><span class='cantidad'>" + Cstr(caja2) + "</span></td></tr><tr><td>Excel:</td><td><span class='cantidad'>" + "$" + ExcelValor + "</span></td></tr><tr><td>C.I:</td><td><span class='cantidad'>" + Cstr(CONTADOR1) + "</span></td></tr><tr><td>C.I:</td><td><span class='cantidad'>" + Cstr(CONTADOR2) + "</span></td></tr><tr><td>C.F:</td><td><span class='cantidad'>" + Cstr(CONTADOR3) + "</span></td></tr><tr><td>C.F:</td><td><span class='cantidad'>" + Cstr(CONTADOR4) + "</span></td></tr><tr><td>B&amp;N:</td><td><span class='number'>" + Cstr(TotalCount) + "</span></td></tr><tr><td>Color:</td><td><span class='number'>" + Cstr(ColorCount) + "</span></td></tr><tr><td>Total Diff:</td><td><span class='number'>" + Cstr(TotalDiff) + "</span></td></tr><tr><td>CopiasVendidas Faltantes:</td><td><span class='number'>" + Cstr(Faltante) + "</span></td></tr></tbody></table></div></div><script>function removeCopiasVendidasFaltantes(){var table=document.getElementsByTagName('table')[0];table.deleteRow(table.rows.length-1);}document.addEventListener('keydown',function(event){if(event.keyCode===74){removeCopiasVendidasFaltantes();}});</script></body></html>"
            archivotexto.close
        
        
            objshell.run LocalUser +"\Downloads\index.html"

            wait(2)
            


        Resources = LocalUser + "\Downloads\"


        DefaultPath = Disk + UserPath
        para = "ops65@hotmail.com; opsjra@gmail.com; ciee.empalme@live.com; jonathanvargas21@outlook.es"
        ''para = "ciee.empalme@live.com"



        asunto = "Jonathan " + Cstr(DayNumber) + "-" + Cstr(MonthNumber) + "-" + Cstr(shortyear)
        mensaje = ""
        adjunto1 =  DefaultPath + FolderYear + "\" + FolderMonth + "\" + excel
        adjunto2 =  DefaultPath + FolderYear + "\" + FolderMonth + "\" + pdf + ".pdf"
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
            SetupPrinter()
            Wait(30)
            ObjFso.DeleteFile LocalUser +"\Documents\dat2.txt"
            Objshell.run "shutdown -s -t 120"
        Else
            Msgbox("Se cancelo el envio")
            protocolo = 7
        End If

    End If

    '--------------------------------------------------------------------------------------------------END-------------------------------------------------------------------------------------------------






    WScript.sleep 200
Wend


'-----------------------------------------DEBUG ZONE----------------------------------------------------

