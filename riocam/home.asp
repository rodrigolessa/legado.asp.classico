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
	DIM i

	
	'***** ABRE A CONEX�O *****
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

	<!-- FORMUL�RIO GERAL DA P�GINA -->
	<form id="geral" name="geral" action="<%=Request.ServerVariables("SCRIPT_NAME")%>" method="post">
	<input type="Hidden" name="oper">
	<input type="Hidden" name="cmd">


		<!-- UM BOX (DIV) PARA ENGLOBAR TODOS -->
		<div id="principal">

			<!-- BOX P�GINA PARA CONTER TODO O LAYOUT -->
			<div id="pagina">

				<!-- CABECALHO DA P�GINA -->
				<!--#include file="cabecalho.asp" -->
				
<%
				if len(msg_erro) > 0 then
%>
					<div id="msg_erro"><%=msg_erro%></div>
<%
				end if
%>

				<!-- BOX DE CONTE�DO DA P�GINA -->
				<div id="hm_conteudo">
				

					<!-- BOX DE MENU PRINCIPAL DA P�GINA -->
					<div id="hm_conteudo_centro">
					
						<%=impModHome("H")%>
						
						<div id="chat_home_centro">
							<script language="javaScript">
								criaFlash("banners/banner_free.swf", "270", "293");
							</script>
						</div>
						
						<div id="centro_espaco"></div>
						
						<div id="galeria_centro">
							<div id="galeria_titulo"><%=id_getText("home_tit_01")%></div>
							<%=impGaleria("H")%>
						</div>
						
						<div id="galeria_centro">
							<div id="galeria_titulo"><%=id_getText("home_tit_02")%></div>
							<%=impVideos("H")%>
						</div>
						
						<div id="free_centro">
							<div id="free_titulo"><%=id_getText("home_tit_03")%></div>
							<script language="javaScript">
								criaFlash("banners/banner_modelos.swf", "160", "257");
							</script>
						</div>

					</div>


				</div>
				<!-- FIM BOX DE CONTE�DO DA P�GINA -->
				
				<!-- BOX DE MENU PRINCIPAL DA P�GINA -->
				<!--#include file="menu.asp" -->

				<!-- BOX DE RODAP� DA P�GINA -->
				<!--#include file="rodape.asp" -->

			</div>
			<!-- FIM BOX P�GINA PARA CONTER TODO O LAYOUT -->

		</div>
		<!-- FIM BOX (DIV) PARA ENGLOBAR TODOS -->


	</form>

	</body>
</html>