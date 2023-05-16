Set objshell = createobject("wscript.shell")

Objshell.Exec("taskkill /f /im WScript.exe")

