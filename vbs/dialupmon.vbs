' Dial-Up Monitor - Gerencia conex�es PPP.
'	Vers�o : 0.01_19082009
'	18/08/2009 - Cria��o. Ap�s acria��o do arquivo original em BAT, foram detectado alguns Bugs que seriam mais facilmente solucionados em VBS.
'	20/08/2009 - Adicionada fun��o WriteRlt para a gerar relat�rios e efetuar testes. Bugs corrigidos vers�o praticamente us�vel... Ainda h� necessidade de eliminar a depend�ncia dos scripts ONLINE.CMD e PRVDMON.CMD
'	21/08/2009 - Alguns BugFixes, e elimina��o das depend�ncias de scripts externos. (Falta testar!)
'	22/08/2009 - Alguns BugFixes, Problemas ao executar Obj Shell com pipeline. Por causa desses problemas continuamos usando o script PRVDMON.CMD. Apartir de agora os S�bados que forem feriado retornar�o todos os seus hor�rios como provedor de B�nus.
'	13/09/2009 - Criei agora Modos de conex�o All(A) conecta em todo hor�rio resuzido Night(N) Somente de madrugada. Mudei tambem PRVDMON.CMD para suportar multiplas conex�es
'	23/09/2009 - Bugfix: Mudei o PRVDMON.CMD, por�m n�o estava chamando com 2 argumentos, por isso o script ficou uma semana sem trocar a conex�o!!!!
'	Vers�o : 0.04_07102009
'	07/10/2009 - Apartir da vers�o 0.03 o script passou a se chamar dialupmon.vbs, pois existem alguns programas conhecidos com o nome de pppmon... adicionado o modo Night que s� conecta de madrugada;
'	Obs:
'	* Preciso Criar uma fun��o para otimizar, ou seja gastar a franquia restante de uma linha com o provedor de b�nus
'	* Melhorar resultado dos Relat�rios

' Variaveis
	Dim ParamOpt
	Dim ParamPrvdNorm
	Dim ParamPrvdBnus
	Dim ParamRlt
	Dim ParamRltYear
	Dim PrvdBnusTotalH
	Dim PrvdBnusTotalM(12)

' Constantes
private const ScrVer = "0.04"

'Exibe uma caixa de Dialogo
private const ModeDialog = "d"
'Conecta se estiver OffLine - Somente no Provedor Normal
private const ModeOnline = "o"
'Conecta se estiver OffLine - Somente no Provedor B�nus
private const ModeForceBonus = "b"
'Conecta em todos os Hor�rios de Tarifa Reduzida
private const ModeAll = "a"
'Conecta Somente de Madrugada
private const ModeNight = "n"
'Gera relat�rio com h�ras de navega��o
private const ModeRlt = "r"
'Controla o desligamento autom�tico segundo o modo de conex�o
private const ModeShutdown = "s"

' Dialogos
private const Dialog_Title = "PPP Monitor"
private const Dialog_PrvdNorm = "Hor�rio de conex�o correspondente ao Provedor Normal."
private const Dialog_PrvdBonus = "Hor�rio de conex�o correspondente ao Provedor de B�nus."
private const Dialog_NoParam = "Execu��o sem argumentos!"
private const Dialog_ErrParam = "Parametros Inv�lidos!"

' Ping Hosts
private const Host1 = "www.google.com"
private const Host2 = "www.globo.com"
private const Host3 = "a.ntp.br"

' Cores do Relat�rio
private const ClrPrvdNorm = "#CC0000"
private const ClrPrvdBonus = "#00CC00"
private const ClrPrvdBnusHD = "#00CCCC"
private const ClrCellHead = "#CCCCCC"

' Textos do Relat�rio
private const TxtPrvdNorm = "Normal"
private const TxtPrvdBonus = "B�nus"


Function Ping(strHost, intCount, intWait)
	' 21/08/2009 - Nova fun��o Ping em teste. Eliminando a depend�ncia do script ONLINE.CMD.
	' Obs : Ainda � necess�rio testar OnLine!!!!!!!!!!!!!!
	REM ' A adapta��o ainda n�o est� completa. Eliminar depend�ncia do Script ONLINE.CMD
	REM Set objShell = CreateObject("Wscript.Shell")
	REM Ping = objShell.Run ("C:\Windows\System32\ONLINE.CMD",0,true)
	Set objShell = CreateObject("Wscript.Shell")
	Ping = objShell.Run ("%windir%\System32\ping " & strHost & " -n " & intCount & " -w " & intWait ,0,true)
End Function

Function OnLine()
	OnLine = True
	If Ping(Host1,5,35000) <> 0 Then
		If Ping(Host2,5,35000) <> 0 Then
			If Ping(Host3,5,35000) <> 0 Then
				OnLine = False
			End If
		End If
	End If
End Function

Function PrvdMon(strProvider1, strProvider2)
	' 21/08/2009 - Nova fun��o PrvdMon em teste. Eliminando a depend�ncia do script PRVDMON.CMD.
	' 22/08/2009 - Testei a fun��o online e parece que tenho problemas ao executar comandos com pipeline! ainda n�o aposentamos o PRVDMON.CMD.
	' Obs : Ainda � necess�rio testar OnLine!!!!!!!!!!!!!!
	' A adapta��o ainda n�o est� completa. Eliminar depend�ncia do Script PRVDMON.CMD
	Set objShell = CreateObject("Wscript.Shell")
	PrvdMon = objShell.Run ("C:\Windows\System32\prvdmon.cmd " & strProvider1 & " " & strProvider2,0,true)
	' IPCONFIG | FIND "%1" > nul
	REM Dim intResult
	REM Set objShell = CreateObject("Wscript.Shell")
	'intResult = objShell.Run ("%windir%\System32\ipconfig | %windir%\System32\find " & Chr(34) & strProvider & Chr(34) ,0,true)
	REM intResult = objShell.Run ("ipconfig | find " & Chr(34) & strProvider & Chr(34) ,0,true)
	REM MsgBox("ipconfig | find " & Chr(34) & strProvider & Chr(34)  & " " & intResult)
	REM If intResult <> 0 Then
		REM objShell.Run "%windir%\System32\rasdial /disconect" ,0,true
		REM objShell.Run "%windir%\System32\rasphone -D " & strProvider ' ,0,true
	REM End If
End Function

'	Retorna a Data da Pascoa no ano recebido
'	22/08/2009 - Cria��o desta Fun��o
'	Obs: Corrigir o nome Pascoa em ingl�s na fun��o!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Function CalcPasch(intPYear)
	Dim X, Y, A, B, C, D, E, intPDay, intPMonth
	X = 24
	Y = 5
	A = intPYear mod 19
	B = intPYear mod 4
	C = intPYear mod 7
	D = (19 * A + X) mod 30
	E = (2 * B + 4 * C + 6 * D + Y) mod 7
	If ( D + E ) > 9 Then
		intPDay = ( D + E - 9 )
		intPMonth = 4
	Else
		intPDay = ( D + E + 22)
		intPMonth = 3
	End If
	CalcPasch = DateValue(intPDay & "/" & intPMonth  & "/" & intPYear)
End Function

Function Hollyday(HDate)
	'HDate = DateValue(HDate)
	Dim HMonth
	Dim HDay
	Dim HPasch
	Dim HCarnaval
	Dim HCopusChristi
	Dim HPaixao
	HMonth = Month(HDate)
	HDay = Day(HDate)
	' Calcula P�scoa
	HPasch = CalcPasch(Year(HDate))
	' Calcula Carnaval
	HCarnaval = DateAdd("y",-47,HPasch)
	' Calcula Corpus Christi
	HCopusChristi = DateAdd("y",60,HPasch)
	' Calcula Paix�o de Cristo
	HPaixao = DateAdd("y",-2,HPasch)
	Select Case(HMonth)
		Case(1):
			' Confraterniza��o Universal 01/01
			If HDay = 1 Then
				Hollyday = true
			End If
		Case(4):
			' Tiradentes 21/04
			If HDay = 21 Then
				Hollyday = true
			End If
		Case(5):
			' Dia do Trabalho 01/05
			If HDay = 1 Then
				Hollyday = true
			End If
		Case(9):
			' Independ�ncia do Brasil 07/09
			If HDay = 7 Then
				Hollyday = true
			End If
		Case(10):
			' Aparecida 12/10
			If HDay = 12 Then
				Hollyday = true
			End If
		Case(11):
			' Finados 02/11
			If HDay = 02 Then
				Hollyday = true
			Else
			' Proclama��o da Rep�blica 15/11
				If HDay = 15 Then Hollyday = true
			End If
		Case(12):
			' Natal 25/12
			If HDay = 25 Then
				Hollyday = true
			End If
		Case Else:
			Hollyday = false
	End Select
	If (HDate = HPaixao) Then
		Hollyday = true
	ElseIf (HDate = HCarnaval) Then
		Hollyday = true
	ElseIf (HDate = HCopusChristi) Then
		Hollyday = true
	End If
End Function

Function WorkHour(WDate, WTime, WMode)
	Dim WWeekday
	Dim WHour
	WWeekday = Weekday(WDate)
	WHour = Hour(WTime)
	Select Case(WMode)
		' ModeAll - Conecta em todos os hor�rios de tarifa reduzida
		Case(ModeAll):
			Select Case(WWeekday)
				Case(1):
					' Domingo
					WorkHour = false
				Case(7):
					' S�bado
					' Se for feriado
					If Hollyday(WDate) Then
						WorkHour = false
					Else
						If ((WHour <= 13) and (WHour >= 6)) Then
							WorkHour = true
						Else
							WorkHour = false
						End If
					End If
				Case Else:
					' Dias de Semana
					If Hollyday(WDate) Then
						WorkHour = false
					Else
						If ((WHour <= 23) and (WHour >= 6)) Then
							WorkHour = true
						Else
							WorkHour = false
						End If
					End If
			End Select
		' ModeNight - Conecta apenas de 00hs as 06hs (Madrugada)
		Case(ModeNight):
			If ((WHour <= 23) and (WHour >= 6)) Then
				WorkHour = true
			Else
				WorkHour = false
			End If
	End Select
End Function

Sub MonthRlt(objFile, MYear, MMonth, MCDays, MMode )
	For Ind = 1 to MCDays
		RDate = DateValue(Ind & "/" & MMonth & "/" & MYear)
		objFile.WriteLine("<tr>")
		objFile.WriteLine("<td bgcolor=" & Chr(34) & ClrCellHead & Chr(34) & "><p class=" & Chr(34) & "datecell" & Chr(34) & ">" & WeekdayName(Weekday(RDate)) & " " & RDate & "</p></td>")
		For IndHour = 0 to 23
			RTime = IndHour & ":00"
			If (WorkHour(RDate, RTime,MMode)) Then
				objFile.WriteLine("<td bgcolor=" & Chr(34) & ClrPrvdNorm & Chr(34) & "><p class=" & Chr(34) & "cell" & Chr(34) & ">" & TxtPrvdNorm & "</p></td>")
			Else
				If Hollyday(RDate) Then
					objFile.WriteLine("<td bgcolor=" & Chr(34) & ClrPrvdBnusHD & Chr(34) & "><p class=" & Chr(34) & "cell" & Chr(34) & ">" & TxtPrvdBonus & "</p></td>")
				Else
					objFile.WriteLine("<td bgcolor=" & Chr(34) & ClrPrvdBonus & Chr(34) & "><p class=" & Chr(34) & "cell" & Chr(34) & ">" & TxtPrvdBonus & "</p></td>")
				End If
				PrvdBnusTotalH = PrvdBnusTotalH + 1
				PrvdBnusTotalM(MMonth) = PrvdBnusTotalM(MMonth) + 1
			End If
		Next
		objFile.WriteLine("</tr>")
	Next
End Sub

Sub WriteRlt(RYear, RMode)
	Dim RDate
	Dim RTime
	Set RltFS = CreateObject("Scripting.FileSystemObject")
	Set RltFile = RltFS.CreateTextFile("Rlt_pppmon-" & RYear & ".htm", True)
	RltFile.WriteLine("<html><head><title> PPPMon Relat�rio do ano " & RYear & "</title></head><style>" & CssRltBody & "</style><body>")
	RltFile.WriteLine("<center><h1>PPPMon Relat�rio do ano " & RYear & "</h1></center>")
	RltFile.WriteLine("<table>")
	RltFile.WriteLine("<tr>")
	RltFile.WriteLine("<td>Dia/Hora</td>")
		For IndHour = 0 to 23
				RltFile.WriteLine("<td bgcolor=" & Chr(34) & ClrCellHead & Chr(34) & ">" & IndHour & ":00</td>")
		Next
	RltFile.WriteLine("</tr>")
	' Janeiro
	MonthRlt RltFile, RYear, 1, 31, RMode
	' Fevereiro
	If ((RYear Mod 4) > 0) Then
		' Ano Comum
		MonthRlt RltFile, RYear, 2, 28, RMode
	Else
		' Ano Bisexto
		MonthRlt RltFile, RYear, 2, 29, RMode
	End If
	' Mar�o
	MonthRlt RltFile, RYear, 3, 31, RMode
	' Abril
	MonthRlt RltFile, RYear, 4, 30, RMode
	' Maio
	MonthRlt RltFile, RYear, 5, 31, RMode
	' Junho
	MonthRlt RltFile, RYear, 6, 30, RMode
	' Julho
	MonthRlt RltFile, RYear, 7, 31, RMode
	' Agosto
	MonthRlt RltFile, RYear, 8, 31, RMode
	' Setembro
	MonthRlt RltFile, RYear, 9, 30, RMode
	' Outubro
	MonthRlt RltFile, RYear, 10, 31, RMode
	' Novembro
	MonthRlt RltFile, RYear, 11, 30, RMode
	' Dezembro
	MonthRlt RltFile, RYear, 12, 31, RMode
	RltFile.WriteLine("</table>")
	RltFile.WriteLine("<b>Total de Horas:</b><i>" & PrvdBnusTotalH & "</i><br>")
	RltFile.WriteLine("<b>Janeiro:</b><i>" & PrvdBnusTotalM(1) & "</i><br>")
	RltFile.WriteLine("<b>Fevereiro:</b><i>" & PrvdBnusTotalM(2) & "</i><br>")
	RltFile.WriteLine("<b>Mar�o:</b><i>" & PrvdBnusTotalM(3) & "</i><br>")
	RltFile.WriteLine("<b>Abril:</b><i>" & PrvdBnusTotalM(4) & "</i><br>")
	RltFile.WriteLine("<b>Maio:</b><i>" & PrvdBnusTotalM(5) & "</i><br>")
	RltFile.WriteLine("<b>Junho:</b><i>" & PrvdBnusTotalM(6) & "</i><br>")
	RltFile.WriteLine("<b>Julho:</b><i>" & PrvdBnusTotalM(7) & "</i><br>")
	RltFile.WriteLine("<b>Agosto:</b><i>" & PrvdBnusTotalM(8) & "</i><br>")
	RltFile.WriteLine("<b>Setembro:</b><i>" & PrvdBnusTotalM(9) & "</i><br>")
	RltFile.WriteLine("<b>Outubro:</b><i>" & PrvdBnusTotalM(10) & "</i><br>")
	RltFile.WriteLine("<b>Novembro:</b><i>" & PrvdBnusTotalM(11) & "</i><br>")
	RltFile.WriteLine("<b>Dezembro:</b><i>" & PrvdBnusTotalM(12) & "</i><br>")
	RltFile.WriteLine("<small><i>PPPMon vers�o:</i>" & ScrVer & "</small>")
	RltFile.WriteLine("</body>")
End Sub

Function SetMode(Opt)
	SetMode = ModeAll
	If (InStr(Opt, ModeAll)) <> 0 Then
		SetMode = ModeAll
	ElseIf (InStr(Opt, ModeNight)) <> 0 Then
		SetMode = ModeNight
	End If
End Function

'Shutdown - controla o desligamento de acordo com o resultado de Workday e o Modo de conex�o
'30/09/2009 - Cria��o
Sub Shutdown(strOpt)
	Dim SHour
	SHour = Hour(Time)
	'S� executa quando for 6hs
	If (SHour = 6) Then
		'ModeAll - Sempre conectado
		If (InStr(strOpt, ModeAll) <> 0) Then
			If (WorkHour(Date, Time,ModeAll)) Then
				Set objShell = CreateObject("Wscript.Shell")
				objShell.Run "%windir%\System32\shutdown /s /t 0 /f",0,true
			End If
		Else
		'Outros Modes - Sempre desliga as 6hs
			Set objShell = CreateObject("Wscript.Shell")
			objShell.Run "%windir%\System32\shutdown /s /t 0 /f",0,true
		End If
	End If
End Sub

Sub PPPMon
		Dim ParamMode
		'Verifica a necessidade de desligar o PC
		If (InStr(ParamOpt, ModeShutdown) <> 0) Then Shutdown ParamOpt
		ParamMode = SetMode(ParamOpt)
		If (WorkHour(Date, Time,ParamMode)) Then
			'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
			' Provedor Normal
			'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
			' Modo Dialogo
			If (InStr(ParamOpt, ModeDialog)) <> 0 Then
				Set WshShell = CreateObject("WScript.Shell")
				Result = WshShell.Popup(Dialog_PrvdNorm, 10, Dialog_Title, 65)
			End If
			'Se N�o estiver conectado
			If (Not OnLine) Then
				If (InStr(ParamOpt, ModeOnline)) <> 0 Then
					PrvdMon ParamPrvdNorm, ParamPrvdBnus
				End If
			Else
				'Se estiver conectado verifica o Provedor
				PrvdMon ParamPrvdNorm, ParamPrvdBnus
			End If
	Else
			'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
			'Provedor de B�nus
			'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
			'Modo Dialogo
			If (InStr(ParamOpt, ModeDialog) <> 0) Then
				Set WshShell = CreateObject("WScript.Shell")
				Result = WshShell.Popup(Dialog_PrvdBonus, 10, Dialog_Title, 65)
			End If
			'Se N�o estiver conectado
			If (Not OnLine) Then
				If (InStr(ParamOpt, ModeForceBonus)) <> 0 Then
					PrvdMon ParamPrvdBnus, ParamPrvdNorm
				End If
			Else
				'Se estiver conectado verifica o Provedor
				PrvdMon ParamPrvdBnus, ParamPrvdNorm
			End If
	End If
End Sub
' ........................................................................................
' Main
' ........................................................................................

Select Case(WScript.Arguments.Count)
	Case(0):
		MsgBox(Dialog_NoParam)
		WScript.Quit 1
	Case(1):
		ParamOpt = LCase(WScript.Arguments.Item(0))
	Case(2):
		' Assume que argumentos sejam:
		' Para a gera��o de relat�rio
		ParamOpt = LCase(WScript.Arguments.Item(0))
		ParamPrvdNorm = WScript.Arguments.Item(1)
		If (InStr(ParamOpt, ModeRlt) <> 0) Then
		'If (ParamOpt) = ModeRlt Then
			WriteRlt ParamPrvdNorm,SetMode(ParamOpt)
		End If
		WScript.Quit 0
	Case(3):
		' Assume que argumentos sejam:
		' Op��es ProvedorNormal ProvedordeBonus
		ParamOpt = LCase(WScript.Arguments.Item(0))
		ParamPrvdNorm = WScript.Arguments.Item(1)
		ParamPrvdBnus = WScript.Arguments.Item(2)
	Case Else:
		MsgBox(Dialog_ErrParam)
		WScript.Quit 1
End Select
PPPMon