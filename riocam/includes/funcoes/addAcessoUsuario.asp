<%
Function addAcessoUsuario(aaUsuario)

	DIM aAcessos, mes, ano, contador, dataAcesso
	
	'===== CONTROLE ACESSO DOS USUARIOS =====
	sSQL =	"SELECT	ma.usuario, ma.mes, ma.ano, contador, ma.dataAcesso " & _
			"FROM			XCA_usuarios_acessos ma " & _
			"WHERE	ma.usuario = " & aausuario & " " & _
			"	AND	ma.mes = " & month(date()) & " " & _
			"	AND	ma.ano = " & year(date()) & " "
	aAcessos	= getArray(sSQL)
	
	if isArray(aAcessos) then
	
		mes				= aAcessos(1,0)
		ano				= aAcessos(2,0)
		contador		= aAcessos(3,0) + 1
		dataAcesso		= verSQL(dateAdd("h", 1, aAcessos(4,0)), "H", "F")

		if compDataHora(dataAcesso, "") < 1 then
			sSQL =	"UPDATE	XCA_usuarios_acessos SET " & _
					"		dataAcesso	= getDate(), " & _
					"		contador	= contador + 1 " & _
					"WHERE	usuario_acesso = " & aaUsuario & " " & _
					"	AND	mes = " & mes & " " & _
					"	AND	ano = " & ano & " "
			conn.Execute(sSQL)
		end if
		
	else
	
		sSQL =	"INSERT INTO XCA_usuarios_acessos (usuario, mes, ano, contador, dataAcesso) " & _
				"VALUES (" & aaUsuario & ", " & month(date()) & ", " & year(date()) & ", 1, getDate()) "
		conn.Execute(sSQL)
		contador = 1
		
	end if
	
	addAcessoUsuario = contador

End Function
%>