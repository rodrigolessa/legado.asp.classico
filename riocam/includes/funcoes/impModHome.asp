<%
Function impModHome(origem)

	DIM aModelos, temp, ind, strInd, intTotal
	DIM strControle, qtModelo
	
	if uCase(trim(origem)) = "H" then
		qtModelo	= 0
	else
		qtModelo	= 0
	end if
	
	'===== FOTO GRANDE DE MODELOS PARA EXIBIR NA HOME DE FORMA ALEATÓRIA =====
	sSQL = "SELECT	modelo AS id, nomeArquivo, nome, dataCadastro " & _ 
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
		
			Randomize
			ind = cInt(RND * UBound(aModelos,2))
			
			while inStr(strControle, "("&ind&")")
			
				Randomize
				ind = cInt(RND * UBound(aModelos,2))
			
			wend
			
			strControle = strControle & " (" & ind & ") "
	
			temp =	temp &	"<div id=""img_foto_home"">" & _
							"<div id=""titulo_modelo_home""> :: " & aModelos(2,ind) & " :: </div>" & _
							"<a href=""modelos.asp?modelo=" & aModelos(0,ind) & """>" & _
							"<img src=""modelos/" & aModelos(0,ind) & "/" & aModelos(1,ind) & """ width=275 height=260 border=0>" & _
							"</a>" & _
							"</div>"
		next
	end if
	
	impModHome = temp

End Function
%>