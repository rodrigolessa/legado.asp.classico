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
	
	
	'***** RECUPERA AS VARIAVEIS ******
	oper	= cStr(uCase(trim(request("oper"))))
	cmd		= cStr(uCase(trim(request("cmd"))))


	'***** ABRE A CONEXÃO *****
	Call conectar()
	

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
	<input type="Hidden" name="oper">
	<input type="Hidden" name="cmd">

	
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
					
						<%=impTodasModelos()%>
					
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
Function impTodasModelos()

	DIM aModelos, temp, ind, strInd, intTotal
	DIM strControle, qtModelo
	
	'===== MODELOS =====
	sSQL =	"SELECT	modelo AS id, nomeArquivo2, dataCadastro, nome, nomeArquivo, pais " & _
			"FROM	XCA_modelos " & _
			"WHERE	status = 'A' " & _
			"ORDER BY dataCadastro DESC"
			
	aModelos	= getArray(sSQL)
	if isArray(aModelos) then
	
		for i = 0 to UBound(aModelos,2)
		
			if i = 0  then
				temp =	temp &	"<div id=""todas_modelos_img"">" & _
								"<div id=""todas_modelos_nome"">" & aModelos(3,i) & "<span class=""todas_modelos_texto2""><br>" & aModelos(5,i) & "<br><br>" & retLinkConfer(aModelos(0,i)) & "</span></div>" & _
								"<a href=""modelos.asp?modelo=" & aModelos(0,i) & """>" & _
								"<img src=""modelos/" & aModelos(0,i) & "/" & aModelos(4,i) & """ width=238 height=255 border=0>" & _
								"</a>" & _
								"</div>"
			else
				temp =	temp &	"<div id=""todas_modelos_img"">" & _
								"<div id=""todas_modelos_nome"">" & aModelos(3,i) & "<span class=""todas_modelos_texto2""><br>" & aModelos(5,i) & "<br><br>" & retLinkConfer(aModelos(0,i)) & "</span></div>" & _
								"<a href=""modelos.asp?modelo=" & aModelos(0,i) & """>" & _
								"<img src=""modelos/" & aModelos(0,i) & "/" & aModelos(1,i) & """ width=62 height=72 border=0>" & _
								"</a>" & _
								"</div>"
			end if
			
		next
		
	end if
	
	impTodasModelos = temp

End Function
%>