<%
Function impUltNoticias(origem)

	DIM aNoticias, temp, strLead
	
	'===== NOTÍCIAS =====
	sSQL = "SELECT TOP 2 noticia AS id, titulo, lead, data " & _ 
			"FROM	XCA_noticias " & _ 
			"WHERE	data<=getDate() AND (dataExpiracao>=getDate() OR dataExpiracao IS NULL) AND status='A' " & _ 
			"ORDER BY data DESC"
			
	aNoticias	= getArray(sSQL)
	if isArray(aNoticias) then
	
		temp = "<div id=""ultNoticias_titulo"">" & id_getText("home_tit_04") & "</div>"
	
		for i = 0 to uBound(aNoticias, 2)
		
			strLead = cortaTxt(aNoticias(2,i), 80)
		
			temp =	temp &	"<table width=""100%"" cellpadding=0 cellspacing=1 border=0>" & _
							"<tr>" & _
							"	<td id=""ultNoticias_titulo2"">" & aNoticias(1,i) & "</td>" & _
							"</tr>" & _
							"<tr>" & _
							"	<td id=""ultNoticias_texto""><a href=""noticias.asp?oper=V&noticia=" & aNoticias(0,i) & """ class=""lNoticia"">" & strLead & "</a></td>" & _
							"</tr>" & _
							"</table>"
		next
	
	end if
	
	impUltNoticias = temp

End Function
%>