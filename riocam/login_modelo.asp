<% option explicit %>
<!--#include file="includes/incConnectDB.asp" -->
<!--#include file="includes/geral.asp" -->
<%
'Response.Buffer = True
Response.AddHeader "cache-control", "private"
Response.AddHeader "pragma", "no-cache"
Response.ExpiresAbsolute = #January 1, 1990 00:00:01#
Response.Expires=0
Session.LCID = 1046
	
	'***** VARIAVEIS GERAIS ******
	DIM oper, cmd, erro, msg_erro, msg, strRedir 
	DIM login, senha
	DIM i, login_modelo, modelo, senha_modelo, modelo_nome

	
	'***** ABRE A CONEXÃO *****
	Call conectar()
	
	
	'***** RECUPERA AS VARIAVEIS ******
	oper		= cStr(uCase(trim(request("oper"))))
	cmd			= cStr(uCase(trim(request("cmd"))))
	
	msg_erro	= cStr(trim(request("msg_erro")))
	
	select case msg_erro
	case "NLOG"
		msg_erro = " - " & id_getText("msg_erro_01") & " <br>"
	end select

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
					<div id="hm_conteudo_centro" style="text-align:center;">
					
						<div id="centro_espaco"></div>
						
						<!-- BOX DO FORMULÁRIO DE LOGIN -->
						<!--#include file="form_login_modelo.asp" -->
					
						<div id="centro_espaco"></div>
						<div id="centro_espaco"></div>

					</div>


				</div>
				<!-- FIM BOX DE CONTEÚDO DA PÁGINA -->
				
				<!-- BOX DE RODAPÉ DA PÁGINA -->
				<!--#include file="rodape.asp" -->

			</div>
			<!-- FIM BOX PÁGINA PARA CONTER TODO O LAYOUT -->

		</div>
		<!-- FIM BOX (DIV) PARA ENGLOBAR TODOS -->


	</form>

	</body>
</html>