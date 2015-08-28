<%
Function exibeBanner(tipo)
	Dim temp, arrayDados, ind, largura, altura
	query = "SELECT	nomeArquivo FROM XCA_banners WHERE status='A' AND tipo='"&tipo&"' ORDER BY banner"
	arrayDados = criaArrayBD(query)
	
	If IsArray(arrayDados) Then
		Randomize
		ind = Cint(RND * UBound(arrayDados,2))
		If tipo = "F" Then largura = 490 : altura = 60 Else largura = 120 : altura = 60
		If len(arrayDados(0,ind)) > 0 Then
			If LCase(Right(Trim(arrayDados(0,ind)), 3)) = "swf" Then
				temp =	"<OBJECT classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000' " & _
						" codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,0,0' " & _
						" WIDTH='"&largura&"' HEIGHT='"&altura&"' id='banner' ALIGN=''> " & _
						" <PARAM NAME=movie VALUE='arquivos/banners/" & arrayDados(0,ind) & "'>" & _
						"	<PARAM NAME=quality VALUE=high>" & _
						"	<PARAM NAME=bgcolor VALUE=#FFFFFF>" & _
						"	<EMBED src='arquivos/banners/" & arrayDados(0,ind) & "' " & _
						"		quality=high bgcolor=#FFFFFF  WIDTH='"&largura&"' HEIGHT='"&altura&"' NAME='banner' ALIGN='' " & _
						"		TYPE='application/x-shockwave-flash' PLUGINSPAGE='http://www.macromedia.com/go/getflashplayer'>" & _
						"	</EMBED>" & _
						"</OBJECT>"
			Else
				nomeArquivo = "<img src='arquivos/banners/" & rs("nomeArquivo") & "' width='"&largura&"' height='"&altura&"' border=0>"
			End If
		Else
			temp = ""
		End If
	End If
	
	exibeBanner = temp
End Function
%>