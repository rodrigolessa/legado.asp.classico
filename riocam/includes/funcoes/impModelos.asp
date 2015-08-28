<%
Function impModelos(origem)

	DIM aModelos, temp, ind, strInd, intTotal
	DIM strControle, qtModelo
	
	if uCase(trim(origem)) = "H" then
		qtModelo	= 9
	else
		qtModelo	= 8
	end if
	
	'===== MODELOS =====
	sSQL = "SELECT	modelo, nomeArquivo2, dataCadastro " & _
			"FROM	XCA_modelos " & _
			"WHERE	status = 'A' " & _
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
	
			temp =	temp &	"<div id=""img_catalogo"">" & _
							"<a href=""modelos.asp?modelo=" & aModelos(0,ind) & """>" & _
							"<img src=""modelos/" & aModelos(0,ind) & "/" & aModelos(1,ind) & """ width=62 height=72 border=0>" & _
							"</a>" & _
							"</div>"
		next
	end if
	
	impModelos = temp

End Function
%>