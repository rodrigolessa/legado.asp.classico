<%
Function compDataHora(cdData1, cdData2)
	'VARIÁVEIS DA FUNÇÃO
	DIM strData1, dia1, mes1, ano1, hora1, minuto1, segundo1, dataValida1, aDataHora1, strHora1
	DIM strData2, dia2, mes2, ano2, hora2, minuto2, segundo2, dataValida2, aDataHora2, strHora2
	DIM retComp, compTipo


	'VERIFICA SE A PRIMEIRA DATA FOI INFORMADA
	if len(trim(cdData1)) > 6 then
	
		'VERIFICA SE A PRIMEIRA DATA É VÁLIDA
		if valDatHor(cdData1) then
		
			aDataHora1	= split(cdData1, " ")

			if ubound(aDataHora1) <> 1 then
			'VÁLIDA SOMENTE A DATA INFORMADA
				strData1	= split(aDataHora1(0), "/")
				dia1		= d2(strData1(0))
				mes1		= d2(strData1(1))
				ano1		= strData1(2)
				
				compTipo	= "D"
			else
			'VÁLIDA A DATA E A HORA INFORMADAS
				strData1	= split(aDataHora1(0), "/")
				dia1		= d2(strData1(0))
				mes1		= d2(strData1(1))
				ano1		= strData1(2)
				
				strHora1	= split(aDataHora1(1), ":")
				hora1		= d2(strHora1(0))
				minuto1		= d2(strHora1(1))
				segundo1	= d2(strHora1(2))
				
				compTipo	= "H"
			end if
			
		end if
		
	end if
			
			
	'VERIFICA SE A SEGUNDA DATA FOI INFORMADA
	if len(trim(cdData2)) > 6 then

		'VERIFICA SE A SEGUNDA DATA É VÁLIDA
		if valDatHor(cdData2) then

			aDataHora2	= split(cdData2, " ")

			if ubound(aDataHora2) <> 1 then
			'VÁLIDA SOMENTE A DATA INFORMADA
				strData2	= split(aDataHora2(0), "/")
				dia2		= d2(strData2(0))
				mes2		= d2(strData2(1))
				ano2		= strData2(2)

				compTipo	= "D"
			else
			'VÁLIDA A DATA E A HORA INFORMADAS
				strData2	= split(aDataHora2(0), "/")
				dia2		= d2(strData2(0))
				mes2		= d2(strData2(1))
				ano2		= strData2(2)

				strHora2	= split(aDataHora2(1), ":")
				hora2		= d2(strHora2(0))
				minuto2		= d2(strHora2(1))
				segundo2	= d2(strHora2(2))

				compTipo	= "H"
			end if

		end if
		
	else

		dia2	= d2(day(now()))
		mes2	= d2(month(now()))
		ano2	= year(now())
			
		if cStr(compTipo) = "H" then
			hora2		= d2(hour(now()))
			minuto2		= d2(minute(now()))
			segundo2	= d2(second(now()))
		end if

	end if
	
	
	'VERIFICA QUAL VAI SER O TIPO DA COMPARAÇÃO
	if cStr(compTipo) = "D" then
	
		dataValida1	= ano1 & mes1 & dia1
		if len(dataValida1) > 0 then
			dataValida1	= cDbl(dataValida1)
		end if
		
		dataValida2	= ano2 & mes2 & dia2
		if len(dataValida2) > 0 then
			dataValida2	= cDbl(dataValida2)
		end if
		
	elseif cStr(compTipo) = "H" then
	
		dataValida1	= ano1 & mes1 & dia1 & hora1 & minuto1 & segundo1
		if len(dataValida1) > 0 then
			dataValida1	= cDbl(dataValida1)
		end if
		
		dataValida2	= ano2 & mes2 & dia2 & hora2 & minuto2 & segundo2
		if len(dataValida2) > 0 then
			dataValida2	= cDbl(dataValida2)
		end if
		
	end if
			
			
	'VERIFICA SE AS DATAS TRABALHADAS SÃO VÁLIDAS
	if len(dataValida1) > 0 then
	
		if len(dataValida2) > 0 then
			'FAZ A COMPARAÇÃO ENTRE AS DATAS
			if dataValida1 > dataValida2 then
				retComp = 1
			elseif dataValida1 < dataValida2 then
				retComp = -1
			else
				retComp = 0
			end if	
			
		end if
		
	end if
	
	'RETORNA SE A PRIMEIRA DATA É MAIOR, MENOR OU IGUAL A SEGUNDA
	compDataHora = retComp
	
end function
%>
