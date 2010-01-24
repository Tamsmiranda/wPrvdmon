@ECHO OFF
:: Verifica se o provedor é o mesmo que o parâmetro %1
:: 13/08/2009 - CRIAÇÃO

IF %1'==' GOTO SYNTAX
IPCONFIG | FIND "%1" > nul
IF %ERRORLEVEL%==1 GOTO PRVDERROR
GOTO PRVDOK

:PRVDOK
:: Provedor é o correto
IF "%3"=="/T" ECHO Provedor correto.
EXIT /B 0

:PRVDERROR
:: Provedor é diferente do parâmetro %1
IF "%3"=="/T" (
ECHO Provedor diferente.
) ELSE (
:: Com RASPHONE seria necessário especificar o provedor
REM RASDIAL /DISCONNECT
ECHO deconectando %1
RASPHONE -H %2
:: Com RASDIAL recebi menssagens de erro durante a conexão
ECHO deconectando %1
RASPHONE -D %1
)
EXIT /B 1

:SYNTAX
type prvdmon.txt