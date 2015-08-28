<% option explicit %>
<!--#include file="includes/incConnectDB.asp" -->
<!--#include file="includes/geral.asp" -->
<%
'Response.Buffer = True
Response.AddHeader "cache-control", "private"
Response.AddHeader "pragma", "no-cache"
Response.ExpiresAbsolute = #January 1, 1990 00:00:01#
Response.Expires = 0
Session.LCID = 1046
	
	'***** VARIAVEIS GERAIS ******
	DIM oper, cmd, erro, msg_erro, msg, strRedir 
	DIM login, senha
	DIM i
	
	'***** VARIAVEIS DETALHE DA MODELO ******
	DIM modelo, aModelo
	DIM strNomeModelo, strDescModelo, strFotoModelo, strPessoalModelo
	DIM strTempSig, strTempElem
	
	'***** VARIAVEIS DO VÍDEO ******
	DIM modelo_video, codigoModelo, videoDown
	DIM tituloVideo, aVideos, videoModelo, tamanhoVideo
	
	'***** VARIAVEIS DA GALERIA ******
	DIM galeria, foto, gFotoModelo, aFotosModelos
	
	
	'***** RECUPERA AS VARIAVEIS ******
	oper			= cStr(uCase(trim(request("oper"))))
	cmd				= cStr(uCase(trim(request("cmd"))))
	
	modelo			= cStr((trim(request("modelo"))))
	modelo_video	= cStr((trim(request("modelo_video"))))
	galeria			= cStr((trim(request("galeria"))))
	foto			= cStr((trim(request("foto"))))




	'***** ABRE A CONEXÃO *****
	Call conectar() 


	'***** RECUPERA DETALHES DA MODELO SELECIONADA *****
	sSQL =	("SELECT	m.modelo, m.nome, m.nomeArquivo, m.descricao, " & _
			"			m.idade, m.altura, m.signo, m.elemento, m.esporte, m.esporte_eng, " & _
			"			m.olhos, m.cabelo, m.busto, m.cintura, m.quadril, " & _
			"			m.olhos_eng, m.cabelo_eng, m.busto_eng, m.cintura_eng, m.quadril_eng, " & _
			"			m.sexo " & _
			"FROM			XCA_modelos m " & _
			"WHERE	m.status = 'A' " & _
			"	AND	m.modelo = " & modelo & " ")
			
	aModelo	= getArray(sSQL)
	
	if isArray(aModelo) then
	
		strNomeModelo		= aModelo(1,0)
		strDescModelo		= aModelo(3,0)
		strPessoalModelo	= ""
		if len(trim(aModelo(4,0))) > 0 then
			strPessoalModelo	= strPessoalModelo & "<tr><td class=""inter_tabela_texto1"">Idade: </td><td class=""inter_tabela_texto2"">" & aModelo(4,0) & "</td><td>"
		end if
		if len(trim(aModelo(5,0))) > 0 then
			strPessoalModelo	= strPessoalModelo & "<tr><td class=""inter_tabela_texto1"">Altura: </td><td class=""inter_tabela_texto2"">" & aModelo(5,0) & "</td><td>"
		end if
		if len(trim(aModelo(10,0))) > 0 then
			strPessoalModelo	= strPessoalModelo & "<tr><td class=""inter_tabela_texto1"">Olhos: </td><td class=""inter_tabela_texto2"">" & aModelo(10,0) & "</td><td>"
		end if
		if len(trim(aModelo(11,0))) > 0 then
			strPessoalModelo	= strPessoalModelo & "<tr><td class=""inter_tabela_texto1"">Cabelos: </td><td class=""inter_tabela_texto2"">" & aModelo(11,0) & "</td><td>"
		end if
		if len(trim(aModelo(12,0))) > 0 then
			strPessoalModelo	= strPessoalModelo & "<tr><td class=""inter_tabela_texto1"">Busto: </td><td class=""inter_tabela_texto2"">" & aModelo(12,0) & "</td><td>"
		end if
		if len(trim(aModelo(13,0))) > 0 then
			strPessoalModelo	= strPessoalModelo & "<tr><td class=""inter_tabela_texto1"">Cintura: </td><td class=""inter_tabela_texto2"">" & aModelo(13,0) & "</td><td>"
		end if
		if len(trim(aModelo(14,0))) > 0 then
			strPessoalModelo	= strPessoalModelo & "<tr><td class=""inter_tabela_texto1"">Quadril: </td><td class=""inter_tabela_texto2"">" & aModelo(14,0) & "</td><td>"
		end if
		if len(trim(aModelo(6,0))) > 0 then
			strTempSig			= "mod_det_signo_" & uCase(cStr(trim(aModelo(6,0))))
			strPessoalModelo	= strPessoalModelo & "<tr><td class=""inter_tabela_texto1"">Signo: </td><td class=""inter_tabela_texto2"">" & id_getText(strTempSig) & "</td><td>"
		end if
		if len(trim(aModelo(7,0))) > 0 then
			strTempElem			= "mod_det_elem_" & uCase(cStr(trim(aModelo(7,0))))
			strPessoalModelo	= strPessoalModelo & "<tr><td class=""inter_tabela_texto1"">Elemento: </td><td class=""inter_tabela_texto2"">" & id_getText(strTempElem) & "</td><td>"
		end if
		if len(trim(aModelo(8,0))) > 0 then
			strPessoalModelo	= strPessoalModelo & "<tr><td class=""inter_tabela_texto1"">Esporte: </td><td class=""inter_tabela_texto2"">" & aModelo(8,0) & "</td><td>"
		end if
		strFotoModelo		= "<img src=""modelos/" & aModelo(0,i) & "/" & aModelo(2,i) & """ width=275 height=260 border=0>"
		
		'FAZ O CONTROLE DE ACESSO AS MODELOS
		addAcessoModelo modelo
		
	end if
	
	if oper = "V" then
	
		'***** RECUPERA DETALHES DO VÍDEO SELECIONADO DA MODELO PARA EXIBIR *****
		sSQL =	("SELECT	m.modelo, m.nome, m.nomeArquivo2, " & _
				"			v.titulo, v.nomeArquivo, v.tamanhoArquivo, " & _
				"			m.descricao " & _
				"FROM	(		XCA_modelos m " & _
				"	INNER JOIN	XCA_modelos_videos v	ON m.modelo = v.modelo) " & _
				"WHERE	m.status = 'A' " & _
				"	AND	v.status = 'A' " & _
				"	AND	v.modelo_video = " & modelo_video & " ")

		aVideos	= getArray(sSQL)

		if isArray(aVideos) then

			codigoModelo	= aVideos(0,0)
			tamanhoVideo	= converteBytes(aVideos(5,0)) & " (" & aVideos(5,0) & " bytes)"
			videoModelo		= Application("XCA_APP_HOME") & "modelos/" & codigoModelo & "/video/" & aVideos(4,0)
'			videoModelo		= "http://www.rioeroticam.com/demo/modelos/1/video/vid01.flv"
			videoDown		= aVideos(4,0)

		end if
		
	elseif oper = "G" then
	
		'***** RECUPERA DETALHES DA GALERIA SELECIONADO DA MODELO PARA EXIBIR *****
		sSQL =	("SELECT	mg.modelo, mg.modelo_galeria, gf.modelo_galeria_foto, gf.nomeArquivo, gf.tamanhoArquivo " & _
				"FROM	(		XCA_modelos_galerias mg " & _
				"	INNER JOIN	XCA_modelos_galerias_fotos gf	ON mg.modelo_galeria = gf.modelo_galeria) " & _
				"WHERE	mg.status = 'A' " & _
				"	AND	gf.status = 'A' " & _
				"	AND	gf.modelo_galeria_foto = " & foto & " ")

		aFotosModelos = getArray(sSQL)

		if isArray(aFotosModelos) then

			gFotoModelo		= (Application("XCA_APP_HOME") & "modelos/" & aFotosModelos(0,0) & "/galeria/" & aFotosModelos(3,0))
			gFotoModelo		= ("<img src=""" & gFotoModelo & """ width=555 border=0>")

		end if
	
	end if

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
	<head>
		<meta http-equiv="content-type" content="text/html;charset=iso-8859-1">
		<title><%=Application("XCA_APP_TITLE")%></title>
		<link href="includes/box.css"			type="text/css" rel="stylesheet" media="screen">
		<link href="includes/geral.css"			type="text/css" rel="stylesheet" media="screen">
		<link href="includes/formulario.css"	type="text/css" rel="stylesheet" media="screen">
		<script src="includes/funcoes.js" language="Javascript"></script>
	</head>
	<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" bottommargin="0">

	<!-- FORMULÁRIO GERAL DA PÁGINA -->
	<form id="geral" name="geral" action="<%=Request.ServerVariables("SCRIPT_NAME")%>" method="post">
	<input type="Hidden" name="oper" value="<%=oper%>">
	<input type="Hidden" name="cmd">
	<input type="Hidden" name="modelo" value="<%=modelo%>">
	<input type="Hidden" name="modelo_video" value="<%=modelo_video%>">
	<input type="Hidden" name="galeria" value="<%=galeria%>">
	<input type="Hidden" name="foto" value="<%=foto%>">
	
	
		<!-- UM BOX (DIV) PARA ENGLOBAR TODOS -->
		<div id="principal">

			<!-- BOX PÁGINA PARA CONTER TODO O LAYOUT -->
			<div id="pagina">

				<!-- CABECALHO DA PÁGINA -->
				<!--#include file="cabecalho.asp" -->
				
<%
				if len(msg_erro) > 0 then
%>
					<div id="msg_erro"><%=msg_erro%></div>
<%
				end if
%>

				<!-- BOX DE CONTEÚDO DA PÁGINA -->
				<div id="hm_conteudo">
				

					<!-- BOX DE MENU PRINCIPAL DA PÁGINA -->
					<div id="hm_conteudo_centro">
					
<%
				if oper = "G" then
%>
						<div id="detalhe_modelo_texto2"><%=gFotoModelo%></div>
<%
				elseif oper = "V" then
%>
						<div id="detalhe_modelo_foto">
							<table width="100%" cellpadding=0 cellspacing=0 border=0>
							<tr>
								<td align="center">
<!--
									<object classid="clsid:d27cdb6e-ae6d-11cf-96b8-444553540000" codebase="http://fpdownload.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=8,0,0,0" width="275" height="260" id="vplayer8" align="middle">
										<param name="allowScriptAccess" value="sameDomain" />
										<param name="movie" value="vplayer8.swf?video=<%=videoModelo%>" />
										<param name="menu" value="false" />
										<param name="quality" value="medium" />
										<param name="wmode" value="transparent" />
										<param name="bgcolor" value="#111111" />
										<param name="FlashVars" value="video=<%=videoModelo%>" />
										<embed src="vplayer8.swf?video=<%=videoModelo%>" quality="high" bgcolor="#333333" width="275" height="260" name="vplayer8" align="middle" allowScriptAccess="sameDomain" type="application/x-shockwave-flash" pluginspage="http://www.macromedia.com/go/getflashplayer" />
									</object>
-->
									<script language="javaScript">
										criaFlashVideo("vplayer8.swf", "video=<%=videoModelo%>", "275", "260");
									</script>
								</td>
							</tr>
							</table>
						</div>
						
						<div id="detalhe_modelo_texto">
							<table width="100%" cellpadding=5 cellspacing=1 border=0>
							<tr>
								<td colspan=2 class="inter_tabela_titulo"><%=strNomeModelo%></td>
							</tr>
							<%=strPessoalModelo%>
							</table>
						</div>
<%
				else
%>
						<div id="detalhe_modelo_foto"><%=strFotoModelo%></div>
						
						<div id="detalhe_modelo_texto">
							<table width="100%" cellpadding=5 cellspacing=1 border=0>
							<tr>
								<td colspan=2 class="inter_tabela_titulo"><%=strNomeModelo%></td>
							</tr>
							<%=strPessoalModelo%>
							</table>
						</div>
<%
				end if
%>
						<div id="detalhe_modelo_texto2">
							<table width="100%" cellpadding=5 cellspacing=1 border=0>
							<tr>
								<td align="right" class="inter_tabela_texto1"><%=retLinkConfer(modelo)%>&nbsp;</td>
							</tr>
							</table>
						</div>

						
						<div id="centro_espaco"></div>
						
						<div id="detalhe_modelo_galeria">
							<table width="100%" cellpadding=5 cellspacing=1 border=0>
							<tr>
								<td class="inter_tabela_texto1"><%=id_getText("mod_det_tit_1")%>:</td>
							</tr>
							<tr>
								<td class="inter_tabela_texto2"><%=impDMGaleria(modelo)%></td>
							</tr>
							</table>
						</div>
						
						<div id="detalhe_modelo_video">
							<table width="100%" cellpadding=5 cellspacing=1 border=0>
							<tr>
								<td class="inter_tabela_texto1"><%=id_getText("mod_det_tit_2")%>:</td>
							</tr>
							<tr>
								<td class="inter_tabela_texto2"><%=impDMVideos(modelo)%></td>
							</tr>
							</table>
						</div>

						<div id="centro_espaco"></div>
						
						<div id="detalhe_modelo_texto2">
							<table width="100%" cellpadding=5 cellspacing=1 border=0>
							<tr>
								<td class="inter_tabela_texto1"><%=id_getText("mod_det_tit_3")%>:</td>
							</tr>
							<tr>
								<td class="inter_tabela_texto2"><%=strDescModelo%></td>
							</tr>
							</table>
						</div>

					</div>

				</div>
				<!-- FIM BOX DE CONTEÚDO DA PÁGINA -->
				
				<!-- BOX DE MENU PRINCIPAL DA PÁGINA -->
				<!--#include file="menu.asp" -->
				
				<!-- BOX DE RODAPÉ DA PÁGINA -->
				<!--#include file="rodape.asp" -->

			</div>
			<!-- FIM BOX PÁGINA PARA CONTER TODO O LAYOUT -->

		</div>
		<!-- FIM BOX (DIV) PARA ENGLOBAR TODOS -->


	</form>

	</body>
</html>
<%
Function impDMGaleria(dmgModelo)

	DIM aModelos, temp, ind, strInd, intTotal
	DIM strControle, qtModelo, strLead
	
	'===== MODELOS E GALERIAS =====
	if verLogado then
		sSQL =	"SELECT	m.modelo, mg.modelo_galeria, mg.titulo, gf.nomeArquivo2, gf.modelo_galeria_foto " & _
				"FROM	(		XCA_modelos m " & _
				"	INNER JOIN	XCA_modelos_galerias mg		ON m.modelo = mg.modelo) " & _
				"	INNER JOIN	XCA_modelos_galerias_fotos gf		ON mg.modelo_galeria = gf.modelo_galeria " & _
				"WHERE	m.status = 'A' " & _
				"	AND	mg.status = 'A' " & _
				"	AND	gf.status = 'A' " & _
				"	AND	m.modelo = " & dmgModelo & " " & _
				"ORDER BY	mg.modelo_galeria, gf.modelo_galeria_foto;"
	else
		sSQL =	"SELECT	top 1 m.modelo, mg.modelo_galeria, mg.titulo, gf.nomeArquivo2, gf.modelo_galeria_foto " & _
				"FROM	(		XCA_modelos m " & _
				"	INNER JOIN	XCA_modelos_galerias mg		ON m.modelo = mg.modelo) " & _
				"	INNER JOIN	XCA_modelos_galerias_fotos gf		ON mg.modelo_galeria = gf.modelo_galeria " & _
				"WHERE	m.status = 'A' " & _
				"	AND	mg.status = 'A' " & _
				"	AND	gf.status = 'A' " & _
				"	AND	m.modelo = " & dmgModelo & " " & _
				"ORDER BY	mg.modelo_galeria, gf.modelo_galeria_foto;"
	end if
			
	aModelos	= getArray(sSQL)
	if isArray(aModelos) then
	
		for i = 0 to uBound(aModelos, 2)
		
			temp =	temp &	"<div id=""menu_foto"">" & _
							"<a href=""modelos.asp?oper=G&modelo="&dmgModelo&"&galeria=" & aModelos(1,i) & "&foto=" & aModelos(4,i) & """>" & _
							"<img src=""modelos/" & aModelos(0,i) & "/galeria/" & aModelos(3,i) & """ width=62 height=72 border=0>" & _
							"</a>" & _
							"</div>"
		next
	end if
	
	impDMGaleria = temp

End Function



Function impDMVideos(dmvModelo)

	DIM aVideos, temp, ind, strInd, intTotal
	DIM strControle, qtVideo
	
	'===== VIDEOS =====
	if verLogado then
		sSQL =	"SELECT	v.modelo_video, v.titulo, v.nomeArquivo2, v.modelo " & _
				"FROM			XCA_modelos m " & _
				"	INNER JOIN	XCA_modelos_videos v	ON m.modelo = v.modelo " & _
				"WHERE	m.status = 'A' " & _
				"	AND	v.status = 'A' " & _
				"	AND	m.modelo = " & dmvModelo & " " & _
				"ORDER BY v.modelo_video;"
	else
		sSQL =	"SELECT	top 1 v.modelo_video, v.titulo, v.nomeArquivo2, v.modelo " & _
				"FROM			XCA_modelos m " & _
				"	INNER JOIN	XCA_modelos_videos v	ON m.modelo = v.modelo " & _
				"WHERE	m.status = 'A' " & _
				"	AND	v.status = 'A' " & _
				"	AND	m.modelo = " & dmvModelo & " " & _
				"ORDER BY v.modelo_video;"
	end if
			
	aVideos	= getArray(sSQL)
	if isArray(aVideos) then
	
		for i = 0 to uBound(aVideos, 2)
		
			temp =	temp &	"<div id=""menu_foto"">" & _
							"<a href=""modelos.asp?oper=V&modelo=" & dmvModelo & "&modelo_video=" & aVideos(0,i) & """>" & _
							"<img src=""modelos/" & aVideos(3,i) & "/video/" & aVideos(2,i) & """ width=62 height=72 border=0>" & _
							"</a>" & _
							"</div>"
		next
	end if
	
	impDMVideos = temp

End Function
%>