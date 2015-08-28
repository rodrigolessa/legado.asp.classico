<%
function verSQL(valor, tipo, destino)

DIM temp, tData, tHora

select case uCase(destino)
case "B"
	'INICIO PARA FAZER INSERT OU UPDATE NO BD
	if len(valor) > 0 Then
		select case uCase(tipo)
			case "S" ' para char, varchar
				verSQL = "'" & valor & "'"
			case "D" ' para smalldatetime
				if isDate(valor) Then
					temp	= split(valor, "/")
					verSQL	= "'" & temp(2) & "-" & temp(1) & "-" & temp(0) & "'"
				else
					verSQL	= "Null"
				end if
			case "H" ' para datetime
				if isDate(valor) Then
					temp	= split(valor, " ")
					tData	= split(temp(0), "/")
					tHora	= split(temp(1), ":")
					verSQL =	"'" & _
									tData(2)	& "-" & _
									tData(1)	& "-" & _
									tData(0)	& " " & _
									tHora(0)	& ":" & _
									tHora(1)	& ":" & _
												  "99" & _
								"'"								
					if ubound(tHora) = 2 then
						verSQL = replace(VerSQL, "99", tHora(2))
					else
						verSQL = replace(VerSQL, "99", "00")
					end if
				else
					verSQL =	"Null"
				end if
			case "I" ' para int
				verSQL = "" & valor & ""
			case "N" ' para numeric
				if isNumeric(valor) Then
					verSQL = "" & replace(valor, ",", ".") & ""
				else
					verSQL = "Null"
				end if
			case "U" ' para string em maiuscula
				verSQL = "'" & uCase(valor) & "'"
			case "L" ' para string em minuscula
				verSQL = "'" & lCase(valor) & "'"
			case else
				verSQL = "" & valor & ""
		end select
	else
		verSQL = "Null"
	end if
	'FIM PARA FAZER INSERT OU UPDATE NO BD
case "F", "R"
	'INICIO RETORNAR OS CAMPOS DE UM SELECT NO BD
	if len(valor) > 0 Then
	
		'INICIO RETORNAR OS CAMPOS DE UM REQUEST
		if uCase(destino) = "R" then
			valor = replace(valor, "'", "`")
			valor = trim(valor)
		end if
		'FIM RETORNAR OS CAMPOS DE UM REQUEST

		select case uCase(tipo)
			case "S" ' para string
				verSQL = "" & valor & ""
			case "D" ' para datas
				if isDate(valor) Then
					verSQL = "" & d2(day(valor)) & "/" & d2(month(valor)) & "/" & year(valor) & ""
				else
					verSQL = valor
				end if
			case "H" ' para data e hora
				if isDate(valor) Then
					verSQL = "" & d2(day(valor)) & "/" & d2(month(valor)) & "/" & year(valor) & " " & d2(hour(valor)) & ":" & d2(minute(valor)) & ":" & d2(second(valor))
				else
					verSQL = valor
				end if
			case "I" ' para inteiros
				verSQL = valor
			case "N" ' para numericos
				if isNumeric(replace(valor, ".", ",")) Then
					'retSQL = formatNumber(replace(valor, ".", ","), 2)
					verSQL = formatNumber(valor, 2)
				else
					verSQL = valor
				end if
			case "U" ' para string em maiuscula
				verSQL = "" & uCase(valor) & ""
			case "L" ' para string em minucula
				verSQL = "" & lCase(valor) & ""
			case else
				verSQL = valor
		end select
	else
		verSQL = ""
	end if
	'FIM RETORNAR OS CAMPOS DE UM SELECT NO BD
end select

end function
%>