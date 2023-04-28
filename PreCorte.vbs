Set objfso = createobject("scripting.filesystemobject")
Set objshell = createobject("wscript.shell")

LocalUser = "SERVIDOR"

function wait(x)
	Wscript.sleep x * 1000
End function


Objshell.run "chrome.exe"
wait(0.5)
Objshell.sendkeys "http://192.168.0.6/wcd/system_counter.xml"
wait(0.5)
Objshell.sendkeys "{ENTER}"
wait(6.5)
Objshell.sendkeys "^a"
wait(0.1)
Objshell.sendkeys "^c"
wait(0.1)
Objshell.sendkeys "^w"
wait(0.1)

if objfso.fileExists("C:\Users\"+ LocalUser +"\Documents\dat2.txt") then

	objfso.deleteFile "C:\Users\"+ LocalUser +"\Documents\dat2.txt"
End if


objfso.createtextfile("C:\Users\"+ LocalUser +"\Documents\dat2.txt")	
					
objshell.run "C:\Users\"+ LocalUser +"\Documents\dat2.txt"
	


wait(0.1)
Objshell.sendkeys "^v"
wait(0.1)
Objshell.sendkeys "^g"
wait(0.1)
Objshell.sendkeys "^w"
wait(0.1)
Set TxT = objfso.opentextfile("C:\Users\"+ LocalUser +"\Documents\dat2.txt",1)	'abrimos el archivo

        For k=0 To 90
            TxT.skipline 'change this
        Next

            contador5 = TxT.readline
            for i=0 to 130
            TxT.skipline
            next


wait(0.1)
contador1 = TxT.readline					'leemos una linea, la primera
TxT.skipline
TxT.skipline
TxT.skipline							'saltamos una linea
contador2 = TxT.readline					'leemos una linea, la tercera
TxT.close
Objshell.sendkeys "^w"
wait(0.1)
objshell.run "C:\Users\"+ LocalUser +"\Documents\dat2.txt"
wait(0.1)
Objshell.sendkeys "^e"
wait(0.1)
Objshell.sendkeys "{BS}"
wait(0.1)
objshell.sendkeys contador1
wait(0.1)
objshell.sendkeys "{ENTER}"
wait(0.1)
objshell.sendkeys contador2
wait(0.1)
objshell.sendkeys "{ENTER}"
wait(0.1)
objshell.sendkeys contador5
wait(0.1)
Objshell.sendkeys "^g"
wait(0.5)
Objshell.sendkeys "^w"



