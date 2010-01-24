# Dial-Up Monitor - Gerencia conex�es PPP.
#	Vers�o em Perl
#!/usr/bin/perl

# Constantes
my $ScrVer = "0.04";

# Exibe uma caixa de Dialogo - Tempor�riamente desabilitado
my $ModeDialog = "d";
# Conecta se estiver OffLine - Somente no Provedor Normal
my $ModeOnline = "o";
# Conecta se estiver OffLine - Somente no Provedor B�nus
my $ModeForceBonus = "b";
# Conecta em todos os Hor�rios de Tarifa Reduzida
my $ModeAll = "a";
# Conecta Somente de Madrugada
my $ModeNight = "n";
# Gera relat�rio com h�ras de navega��o
my $ModeRlt = "r";
# Controla o desligamento autom�tico segundo o modo de conex�o
my $ModeShutdown = "s";

# Dialogos
my $Dialog_Title = "PPP Monitor";
my $Dialog_PrvdNorm = "Hor�rio de conex�o correspondente ao Provedor Normal.";
my $Dialog_PrvdBonus = "Hor�rio de conex�o correspondente ao Provedor de B�nus.";
my $Dialog_NoParam = "Execu��o sem argumentos!";
my $Dialog_ErrParam = "Parametros Inv�lidos!";

# Ping Hosts
my $Host1 = "www.google.com";
my $Host2 = "www.globo.com";
my $Host3 = "a.ntp.br";

# Cores do Relat�rio
my $ClrPrvdNorm = "#CC0000";
my $ClrPrvdBonus = "#00CC00";
my $ClrPrvdBnusHD = "#00CCCC";
my $ClrCellHead = "#CCCCCC";

# Textos do Relat�rio
my $TxtPrvdNorm = "Normal";
my $TxtPrvdBonus = "B�nus";

# Ping - Ping o Host retorna TRUE com sucesso
sub Ping#(strHost, intCount, intWait)
{
	0;
}

#
sub OnLine
{
	my $Online = 0;
	if (Ping($Host1,5,35000)) <> 0
	{
		if (Ping($Host2,5,35000)) <> 0
		{
			if (Ping($Host1,5,35000)) <> 0
			{
				$Online = 1;
			}
		}
	}
	$Online;
}

sub PrvdMon
{
	0;
}

sub CalcPasch($Year)
{
	my $X = 24;
	my $Y = 5;
	my $A = $Year mod 19;
	my $B = $Year mod 4;
	my $C = $Year mod 7;
	my $D = (19 * $A + $X) mod 30;
	my $E = (2 * $B + 4 * $C + 6 * $D + $Y) mod 7;
	if ( $D + $E ) > 9
	{
		my $Day = ( $D + $E  - 9 );
		my $Month  = 4;
	}
	else
	{
		my $Day = ( $D + $E + 22 );
		my $Month = 3;
	}
	#
	#	RETORNO $DAY / $ MONTH / $ YEAR
	#
	#
}

sub Hollyday
{
	#HDate = DateValue(HDate)
	#Dim HMonth
	#Dim HDay
	#Dim HPasch
	#Dim HCarnaval
	#Dim HCopusChristi
	#Dim HPaixao
	#HMonth = Month(HDate)
	#HDay = Day(HDate)
	#Calcula P�scoa
	my $Pasch = CalcPasch($Year); # calcula pascoa
	#Calcula Carnaval
	#my $Carnaval = DateAdd("y",-47,HPasch)
	#' Calcula Corpus Christi
	#HCopusChristi = DateAdd("y",60,HPasch)
	#' Calcula Paix�o de Cristo
	#HPaixao = DateAdd("y",-2,HPasch)
	#Janeiro
	if ( $Month == 1 )
	{
		if ($Day == 1 )
		{
			#Confraterniza��o Universal 01/01
			my $Hollyday = 0;
		}
	}
	else
	{
		# Abril
		if ( $Month == 4 )
		{
			if ( $Day == 21 )
			{
				#Tiradentes 21/04
				my $Hollyday = 0;
			}
		}
		# Maio
		else
		{
			if ( $Month == 5 )
			{
				if ( $Day == 1 )
				{
					#Dia do Trabalho 01/05
					my $Hollyday = 0;
				}
			}
			else
			# Setembro
			{
				if ( $Month == 9 )
				{
					if ( $Day == 7 )
					{
						#Independ�ncia do Brasil 07/09
						my $Hollyday = 0;
					}
				}
				else
				# Outubro
				{
					if ( $Month == 10 )
					{
						if ( $Day == 12 )
						{
							# Aparecida 12/10
							my $Hollyday = 0;
						}
					}
					else
					# Novembro
					{
						if ( $Month == 11 )
						{
							if ( $Day == 2 )
							{
								# Finados 02/11
								my $Hollyday = 0;
							}
							else
							if ( $Day == 15 )
							{
								# Proclama��o da Rep�blica 15/11
								my $Hollyday = 0;
							}
						}
						else
						# Dezembro
						{
							if ( $Month == 12 )
							{
								if ( $Day == 25 )
								{
									# Natal 25/12
									my Hollyday = 0;
								}
							}
							else
							{
								my $Hollyday = 1;
							}
						}
					}
				}
			}
		}
	}
	# FINAL COM DATAS MOVEIS MUDAR................
	#If (HDate = HPaixao) Then
	#	Hollyday = true
	#ElseIf (HDate = HCarnaval) Then
	#	Hollyday = true
	#ElseIf (HDate = HCopusChristi) Then
	#	Hollyday = true
	#End If
}

sub WorkHour
{
	my $DoW = Day_of_Week($year, $month, $day);
	my $Hour ; # HORA WHour = Hour(WTime)
	if ($Mode == $ModeAll)
	{
		# ModeAll - Conecta em todos os hor�rios de tarifa reduzida
		if ( $DoW == 1 )
		{
			# Domingo
			my $WorkHour = 1;
		}
		else
		{
			if ( $DoW == 7 )
			{
				# S�bado
				if (Hollyday($day, $month, $year))
				{
					my $WorkHour = 1;
				}
				else
				if ( $Hour <=13 )  && ( $Hour >= 6 ))
				{
					my $WorkHour = 0;
				}
				else
				{
					my $WorkHour = 1;
				}
			}
			else
			{
				if (Hollyday($day, $month, $year))
				{
					my $WorkHour = 1;
				}
				else
				{
					if ( $Hour <= 23 ) && ( $Hour >= 6 ))
					{
						my $WorkHour = 0;
					}
					else
					{
						my $WorkHour = 1;
					}
				}
			}
		}
	}
	else
	{
		if ( $Mode==$ModeNight )
		{
			if ( $Hour <= 23 ) && ( $Hour >=6 )
			{
				my $WorkHour = 0;
			}
			else
			{
				my $WorkHour = 1;
			}
		}
	}
	# RETORNA WORK HOUR
}
