Function PrvdMon(strProvider1, strProvider2)
	MsgBox("ipconfig find ")
	' 21/08/2009 - Nova fun��o PrvdMon em teste. Eliminando a depend�ncia do script PRVDMON.CMD.
	' 22/08/2009 - Testei a fun��o online e parece que tenho problemas ao executar comandos com pipeline! ainda n�o aposentamos o PRVDMON.CMD.
	' Obs : Ainda � necess�rio testar OnLine!!!!!!!!!!!!!!
	' A adapta��o ainda n�o est� completa. Eliminar depend�ncia do Script PRVDMON.CMD
	'--------------------------------Set objShell = CreateObject("Wscript.Shell")
	'--------------------------------PrvdMon = objShell.Run ("C:\Windows\System32\prvdmon.cmd " & strProvider1 & " " & strProvider2,0,true)
	' IPCONFIG | FIND "%1" > nul
	MsgBox("ipconfig | find ")
	Dim intResult
	Set objShell = CreateObject("Wscript.Shell")
	'intResult = objShell.Run ("%windir%\System32\ipconfig | %windir%\System32\find " & Chr(34) & strProvider & Chr(34) ,0,true)
	intResult = objShell.Run ("ipconfig | find " & Chr(34) & strProvider1 & Chr(34) ,0,true)
	MsgBox("ipconfig | find " & Chr(34) & strProvider1 & Chr(34)  & " " & intResult)
	If intResult <> 0 Then
		' Desconecta o strProvider1
		objShell.Run "%windir%\System32\rasphone -H " & strProvider1 ,0,true
		' Conecta ao strProvider2
		objShell.Run "%windir%\System32\rasphone -D " & strProvider2 ' ,0,true
	End If
End Function

MsgBox("ipconf")
PrvdMon "MS","iG2"