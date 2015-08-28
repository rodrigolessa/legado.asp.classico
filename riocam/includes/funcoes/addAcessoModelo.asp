<%
Function addAcessoModelo(aaModelo)

	DIM aAcessos, mes, ano, contador, dataAcesso
	
	'===== ACESSO DOS USUÁRIOS AS MODELOS =====
	sSQL =	"SELECT	ma.modelo, ma.mes, ma.ano, contador, ma.dataAcesso " & _
			"FROM			XCA_modelos_acessos ma " & _
			"WHERE	ma.modelo = " & aaModelo & " " & _
			"	AND	ma.mes = " & month(date()) & " " & _
			"	AND	ma.ano = " & year(date()) & " "
	aAcessos	= getArray(sSQL)
	
	if isArray(aAcessos) then
	
		mes				= aAcessos(1,0)
		ano				= aAcessos(2,0)
		contador		= aAcessos(3,0) + 1
		dataAcesso		= verSQL(dateAdd("h", 1, aAcessos(4,0)), "H", "F")

		if compDataHora(dataAcesso, "") < 1 then
			sSQL =	"UPDATE	XCA_modelos_acessos SET " & _
					"		dataAcesso	= getDate(), " & _
					"		contador	= contador + 1 " & _
					"WHERE	modelo = " & aaModelo & " " & _
					"	AND	mes = " & mes & " " & _
					"	AND	ano = " & ano & " "
			conn.Execute(sSQL)
		end if
		
	else
	
		sSQL =	"INSERT INTO XCA_modelos_acessos (modelo, mes, ano, contador, dataAcesso) " & _
				"VALUES ( " & aaModelo & ", " & month(date()) & ", " & year(date()) & ", 1, getDate()) "
		conn.Execute(sSQL)
		contador = 1
		
	end if
	
	addAcessoModelo = contador

End Function
%>