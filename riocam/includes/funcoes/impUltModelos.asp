<%
Function impUltModelos(origem)

	DIM aModelos, temp, ind, strInd, intTotal
	DIM strControle, qtModelo, strLead
	
	CONST	maxLeadMenu = 50
	
	if uCase(trim(origem)) = "H" then
		qtModelo	= 2
	else
		qtModelo	= 2
	end if
	
	'===== MODELOS =====
	sSQL = "SELECT	modelo AS id, nomeArquivo2, dataCadastro, nome, descricao AS lead " & _ 
			"FROM	XCA_modelos " & _ 
			"WHERE	status='A' " & _ 
			"ORDER BY dataCadastro DESC"
			
	aModelos	= getArray(sSQL)
	if isArray(aModelos) then
	
		intTotal = cInt(uBound(aModelos, 2))
		if intTotal > qtModelo then
			intTotal = qtModelo
		end if
	
		for i = 0 to intTotal
		
			strLead = left(aModelos(4,i), maxLeadMenu)
			if len(aModelos(4,i)) > maxLeadMenu then
				while (right(strLead, 1) <> " ") and (len(strLead) > 0)
					strLead = left(strLead, len(strLead) - 1)
				wend
				strLead = strLead & "..."
			end if
	
			temp =	temp &	"<div id=""ultModelos_item"">" & _
							"<div id=""ultModelos_foto"">" & _
							"<a href=""modelos.asp?modelo=" & aModelos(0,i) & """>" & _
							"<img src=""modelos/" & aModelos(0,i) & "/" & aModelos(1,i) & """ width=62 height=72 border=0>" & _
							"</a>" & _
							"</div>" & _
							"<div id=""ultModelos_texto"">" & _
							"<b>" & aModelos(3,i) & "</b><br>" & strLead & _
							"</div>" & _
							"</div>"
		next
	end if
	
	impUltModelos = temp

End Function
%>