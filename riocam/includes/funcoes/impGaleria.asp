<%
Function impGaleria(origem)

	DIM aModelos, temp, ind, strInd, intTotal
	DIM strControle, qtModelo, strLead
	
	CONST	maxLeadMenu = 50
	
	if uCase(trim(origem)) = "H" then
		qtModelo	= 2
	else
		qtModelo	= 2
	end if
	
	'===== MODELOS E GALERIAS =====
	sSQL =	("SELECT	m.modelo, mg.modelo_galeria, gf.modelo_galeria_foto, mg.titulo, mg.descricao, gf.nomeArquivo2 " & _
			"FROM	(		XCA_modelos m " & _
			"	INNER JOIN	XCA_modelos_galerias mg			ON m.modelo = mg.modelo) " & _
			"	INNER JOIN	XCA_modelos_galerias_fotos gf	ON mg.modelo_galeria = gf.modelo_galeria " & _
			"WHERE	m.status = 'A' " & _
			"	AND	mg.status = 'A' " & _
			"	AND	gf.status = 'A' " & _
			"ORDER BY	mg.modelo_galeria DESC;")
			
	aModelos	= getArray(sSQL)
	if isArray(aModelos) then
	
		intTotal = cInt(uBound(aModelos, 2))
		if intTotal > qtModelo then
			intTotal = qtModelo
		end if
	
		for i = 0 to intTotal
		
			strLead = left(aModelos(4,i), maxLeadMenu)
			if len(aModelos(2,i)) > maxLeadMenu then
				while (right(strLead, 1) <> " ") and (len(strLead) > 0)
					strLead = left(strLead, len(strLead) - 1)
				wend
				strLead = strLead & "..."
			end if
			
			strControle = strControle & " (" & ind & ") "
	
			temp =	temp &	"<div id=""menu_item"">" & _
							"<div id=""menu_foto"">" & _
							"<a href=""modelos.asp?oper=G&modelo=" & aModelos(0,i) & "&galeria=" & aModelos(1,i) & "&foto=" & aModelos(2,i) & """>" & _
							"<img src=""modelos/" & aModelos(0,i) & "/galeria/" & aModelos(5,i) & """ width=62 height=72 border=0>" & _
							"</a>" & _
							"</div>" & _
							"<div id=""menu_texto"">" & _
							"<b>" & aModelos(3,i) & "</b><br>" & strLead & _
							"</div>" & _
							"</div>"
		next
	end if
	
	impGaleria = temp

End Function
%>