<%

function valDatHor(strDatHor)

	DIM vDHRetorno
	DIM arrayDataHora
	DIM arrayData
	DIM arrayHora
	DIM strData
	DIM strHora
	DIM dia, mes, ano
	DIM hora, minuto, segundo

	if len(strDatHor) < 10 then
	
		vDHRetorno		= false
		
	else

		vDHRetorno		= true
		
		arrayDataHora	= split(strDatHor, " ")

		if ubound(arrayDataHora) <> 1 then
		
			'=========================================
			'VÁLIDA SOMENTE A DATA PASSADA
			'=========================================
			strData		= arrayDataHora(0)

			arrayData	= split(strData, "/")

			if ubound(arrayData) <> 2 then
			
				vDHRetorno = false
				
			else
		
				dia	= arrayData(0)
				mes	= arrayData(1)
				ano	= arrayData(2)
				
				if not isNumeric(dia) then
					vDHRetorno = false
				elseif not isNumeric(mes) then
					vDHRetorno = false
				elseif not isNumeric(ano) then
					vDHRetorno = false
				end if
				
				if vDHRetorno = true then

					dia	= cint(dia)
					mes	= cint(mes)
					ano	= cint(ano)

					if not (mes > 0 and mes < 13) then
						vDHRetorno = false
					elseif dia < 1 then
						vDHRetorno = false
					elseif mes = 2 and (dia > 28 or (dia > 29 and (ano mod 4) = 0)) then
						vDHRetorno = false
					elseif (mes = 4 or mes = 6 or mes = 9 or mes = 11) and dia > 30 then
						vDHRetorno = false
					elseif dia > 31 then
						vDHRetorno = false
					elseif ano > 2072 then
						vDHRetorno = false
					end if

				end if
				
			end if
			
		else
		
			'=========================================
			'VÁLIDA A DATA E A HORA PASSADA
			'=========================================
			strData		= arrayDataHora(0)
			strHora		= arrayDataHora(1)

			arrayData	= split(strData, "/")

			if ubound(arrayData) <> 2 then
			
				vDHRetorno = false
				
			else
			
				dia	= arrayData(0)
				mes	= arrayData(1)
				ano	= arrayData(2)

				if not isNumeric(dia) then
					vDHRetorno = false
				elseif not isNumeric(mes) then
					vDHRetorno = false
				elseif not isNumeric(ano) then
					vDHRetorno = false
				end if

				if vDHRetorno = true then

					dia	= cint(dia)
					mes	= cint(mes)
					ano	= cint(ano)

					if not (mes > 0 and mes < 13) then
						vDHRetorno = false
					elseif dia < 1 then
						vDHRetorno = false
					elseif mes = 2 and (dia > 28 or (dia > 29 and (ano mod 4) = 0)) then
						vDHRetorno = false
					elseif (mes = 4 or mes = 6 or mes = 9 or mes = 11) and dia > 30 then
						vDHRetorno = false
					elseif dia > 31 then
						vDHRetorno = false
					elseif ano > 2072 then
						vDHRetorno = false
					end if

				end if

				if vDHRetorno = true then

					arrayHora = split(strHora, ":")

					if ubound(arrayHora) <> 2 then
					
						vDHRetorno = false
						
					else
					
						hora	= arrayHora(0)
						minuto	= arrayHora(1)
						segundo	= arrayHora(2)

						if not isNumeric(hora) then
							vDHRetorno = false
						elseif not isNumeric(minuto) then
							vDHRetorno = false
						elseif not isNumeric(segundo) then
							vDHRetorno = false
						end if

						if vDHRetorno = true then

							hora	= cint(hora)
							minuto	= cint(minuto)
							segundo	= cint(segundo)

							if not (hora >= 0 and hora < 24) then
								vDHRetorno = false
							elseif not (minuto >= 0 and minuto < 60) then
								vDHRetorno = false
							elseif not (segundo >=0 and segundo < 60) then
								vDHRetorno = false
							end if

						end if
						
					end if
					
				end if
				
			end if
			
		end if

	end if

	valDatHor = vDHRetorno

end function

%>