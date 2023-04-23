Set objfso = createobject("scripting.filesystemobject")
Set objshell = createobject("wscript.shell")
Enviado = False
LOCALUSER = objShell.ExpandEnvironmentStrings("%USERPROFILE%")
Disk = "D:"
SaidYes = False

Found = false
FolderMes = False

FolderYear = Cstr(Year(Date()))

MesActual = Month(Date())

MesAnterior = (    CInt( MesActual )  - 1    )

ImpresoraKonica = "192.168.0.6"
ImpresoraHP = "192.168.0.127"

UserPath = "\EMPLEADOS\JONATHAN\CORTES\"

'objshell.popup "SERVIDOR Iniciado",1

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

Function wait(x)
    Wscript.sleep x * 1000
End Function


objshell.popup "Init" , 3,"xs",68

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


