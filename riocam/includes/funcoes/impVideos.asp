<%
Function impVideos(origem)

	DIM aVideos, temp, ind, strInd, intTotal
	DIM strControle, qtVideo
	
	if uCase(trim(origem)) = "H" then
		qtVideo	= 2
	else
		qtVideo	= 2
	end if
	
	'===== VIDEOS =====
	sSQL =	"SELECT	v.modelo_video, v.titulo, v.nomeArquivo2, v.modelo, m.nome " & _
			"FROM			XCA_modelos_videos v " & _
			"	INNER JOIN	XCA_modelos m	ON v.modelo = m.modelo " & _
			"WHERE	v.status = 'A' " & _
			"	AND	m.status = 'A' " & _
			"ORDER BY v.modelo_video DESC;"
			
	aVideos	= getArray(sSQL)
	if isArray(aVideos) then
	
		intTotal = cInt(uBound(aVideos, 2))
		if intTotal > qtVideo then
			intTotal = qtVideo
		end if
	
		for i = 0 to intTotal
		
			temp =	temp &	"<div id=""menu_item"">" & _
							"<div id=""menu_foto"">" & _
							"<a href=""modelos.asp?oper=V&modelo=" & aVideos(3,i) & "&modelo_video=" & aVideos(0,i) & """>" & _
							"<img src=""modelos/" & aVideos(3,i) & "/video/" & aVideos(2,i) & """ width=62 height=72 border=0>" & _
							"</a>" & _
							"</div>" & _
							"<div id=""menu_texto"">" & _
							"<b>" & aVideos(1,i) & "</b><br>com " & aVideos(4,i) & _
							"</div>" & _
							"</div>"
		next
	end if
	
	impVideos = temp

End Function
%>