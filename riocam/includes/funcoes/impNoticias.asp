<%
Function impNoticias()

	DIM aNoticias, temp
	
	temp = ""
	
	'===== NOTÍCIAS =====
	sSQL = "SELECT TOP 2 noticia AS id, titulo, lead, data " & _ 
			"FROM	XCA_noticias " & _ 
			"WHERE	data<=getDate() AND (dataExpiracao>=getDate() OR dataExpiracao IS NULL) AND status='A' " & _ 
			"ORDER BY data DESC"
			
	aNoticias	= getArray(sSQL)
	if isArray(aNoticias) then
		for i = 0 to uBound(aNoticias, 2)
			temp =	temp &	"<div id=""noticia_centro"">" & _
							"<table width=""100%"" cellpadding=0 cellspacing=2 border=0 class=""texto2"">" & _
							"<tr>" & _
							"	<td>" & aNoticias(1,i) & "</td>" & _
							"</tr>" & _
							"<tr>" & _
							"	<td></td>" & _
							"</tr>" & _
							"<tr>" & _
							"	<td><a href=""noticias.asp?oper=V&noticia=" & aNoticias(0,i) & """ class=""lNoticia"">" & aNoticias(2,i) & "</a></td>" & _
							"</tr>" & _
							"</table>" & _
							"<div id=""noticia_titulo""> " & aNoticias(1,i) & "</div>" & _
							"<div id=""noticia_item"">" & _
							"<a href=""noticias.asp?oper=V&noticia=" & aNoticias(0,i) & """ class=""lNoticia"">" & aNoticias(2,i) & "</a>" & _
							"</div>" & _
							"</div>"
		next
	end if
	
	impNoticias = temp

End Function
%>