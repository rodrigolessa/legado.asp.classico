<% option explicit %>
<!--#include file="includes/incConnectDB.asp" -->
<!--#include file="includes/geral.asp" -->
<%
'Response.Buffer = False
Response.AddHeader "cache-control", "private"
Response.AddHeader "pragma", "no-cache"
Response.ExpiresAbsolute = #January 1, 1990 00:00:01#
Response.Expires=0
Session.LCID = 1046
	
	'***** VARIAVEIS GERAIS ******
	DIM oper, cmd, erro, msg_erro, msg, strRedir
	DIM login, senha
	DIM i
	DIM noticia

	'***** ABRE A CONEXÃO *****
	Call conectar()
	
	if len(trim(request("noticia"))) > 0 then
		noticia = trim(request("noticia"))
	else
		noticia = "0"
	end if
	
	if len(trim(request("oper"))) > 0 then
		oper = trim(request("oper"))
	else
		oper = "L"
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
	<input type="Hidden" name="oper"	value="">
	<input type="Hidden" name="cmd"		value="">


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
					<div id="hm_conteudo_centro" style="width:400px;margin-left:30px;margin-top:10px;">
						
						<div id="centro_espaco"></div>
						
						<% if oper = "V" then response.write verNoticia(noticia) else response.write verOutras() %>
						
						<div id="centro_espaco"></div>

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
Function verNoticia(codigo)

	DIM aNoticias, temp, ind, strInd, intTotal
	
	'===== NOTICIAS =====
		
	if Len(codigo) < 1 then
		temp = "	AND n.data<=getDate() AND (n.dataExpiracao>=getDate() OR n.dataExpiracao IS NULL) "
	else
		temp = "	AND n.noticia=" & codigo & " " 
	end if
	
	sSQL =	"SELECT TOP 1	n.noticia AS id, " & _
			"				n.titulo, " & _
			"				n.data, " & _
			"				n.lead, " & _
			"				n.texto " & _
			"	FROM	XCA_noticias n " & _
			"	WHERE	n.status='A' " & temp & " "
			
			
	aNoticias	= getArray(sSQL)
	if isArray(aNoticias) then
	
			temp =	"<div id=""texto_interno"">" & _
					"	<p class=""inter_titulo"">" & aNoticias(1,0) & "</p> " & _
					"	<p class=""inter_texto"" align=""right"">" & formatDate(aNoticias(2,0), "DD/MM/AAAA") & "</p> " & _
					"	<p class=""inter_texto"">" & aNoticias(3,0) & "</p> " & _
					"	<p class=""inter_texto"">" & aNoticias(4,0) & "</p> " & _
					"	<p class=""inter_texto"" align=""right"">[ <a href=""noticias.asp?oper=L"" class=""lInterno"">outras notícias</a> ]</p> " & _
					"</div>"
					
	else
	
			temp =	"<div id=""texto_interno"">" & _
					"	<p class=""inter_titulo"">Notícia não encontrada...</p> " & _
					"</div>"
					
	end if
	
	vernoticia = temp

End Function



Function verOutras()

	DIM aOutras, temp, ind, strInd, intTotal, i
	
	'===== OUTRAS =====
		
	sSQL =	"SELECT 		n.noticia AS id, " & _
			"				n.titulo, " & _
			"				n.data, " & _
			"				n.lead, " & _
			"				n.texto " & _
			"	FROM	XCA_noticias n " & _
			"	WHERE	n.status='A' " & _
			"		AND	n.data<=getDate() " & _
			"		AND (n.dataExpiracao>=getDate() OR n.dataExpiracao IS NULL) " & _
			"	ORDER BY n.noticia DESC " 
			
	aOutras	= getArray(sSQL)
	if isArray(aOutras) then
	
		temp =				"<div id=""texto_interno"">" & _
							"	<p class=""inter_titulo"">Outras notícias</p> "
								
	
		for i=0 to uBound(aOutras, 2)
	
			temp = temp &	"	<p class=""inter_texto""><nobr>"&formatDate(aOutras(2,0), "DD/MM/AA hh:mm")&" -&nbsp;</nobr><a href='noticias.asp?oper=V&noticia="&aOutras(0,0)&"' target='_self' class=lInterno>"&aOutras(1,0)&"</a></p> "
					
		next
		
			temp = temp &	"</div>"
		
	else
					
			temp =	"<div id=""texto_interno"">" & _
					"	<p class=""inter_titulo"">Não existem notícias cadastradas!</p> " & _
					"</div>"
					
	end if
	
	verOutras = temp

End Function
%>