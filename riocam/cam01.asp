<% option explicit %>
<!--#include file="includes/incVerifLogado.asp" -->
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
	DIM oper, cmd, erro, msg, strRedir 
	DIM i

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
	<input type="Hidden" name="oper"	value="">
	<input type="Hidden" name="cmd"		value="">


		<!-- UM BOX (DIV) PARA ENGLOBAR TODOS -->
		<div id="principal">

			<!-- BOX PÁGINA PARA CONTER TODO O LAYOUT -->
			<div id="pagina">

				<!-- CABECALHO DA PÁGINA -->
				<!--#include file="cabecalho.asp" -->

				<!-- BOX DE CONTEÚDO DA PÁGINA -->
				<div id="hm_conteudo">

					<!-- BOX DO FLASH PARA EXIBIR A CAMERA DAS MODELOS -->
					<div id="hm_conteudo_video">
					
						<div id="video_centro">
							<div id="video_titulo" style="text-align:left;width:700px;">Categoria da sala</div>
							<script language="javaScript">
								criaFlash("cam01.swf", "700", "300")
							</script>
						</div>
						
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