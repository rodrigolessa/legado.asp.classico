<% Option Explicit %>
<!--#include file="includes/incVerifLogado.asp"-->
<!--#include file="includes/incConnectDB.asp"-->
<!--#include file="includes/geral_gerencia.asp"-->
<!--#include file="includes/incVars.asp"-->
<% 
Response.Buffer = True
Response.AddHeader "cache-control", "private"
Response.AddHeader "pragma", "no-cache"
Response.ExpiresAbsolute = #January 1, 1990 00:00:01#
Response.Expires = 0
Session.LCID = 1046

DIM query
DIM oper, cmd, erro, erro2
DIM msg_erro, msg_conf
DIM s, i

DIM strRedir, msgTitulo, m, msgLink

strRedir = "pagina=" & Request("pagina") & "&ordemField=" & Request("ordemField") & "&ordemType=" & Request("ordemType") & "&fCampo=" & Request("fCampo") & "&fTexto=" & Request("fTexto") & "&fSecoes=" & Request("fSecoes") & "&fFoto=" & Request("fFoto") & "&fStatus=" & Request("fStatus") 

msgTitulo	= request("t")
m	= request("m")
msgLink		= request("v")

If msgTitulo	= "" Then msgTitulo	= "Erro"
If m = "" Then m = "Ocorreu um erro no aplicativo favor tentar mais tarde"
If msgLink		= "" Then msgLink	= "javascript:history.back()"

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
		<title><%=Application("XCA_APP_TITLE")%></title>
		<link href="includes/box.css" type="text/css" rel="stylesheet" media="screen">
		<link href="includes/geral.css" type="text/css" rel="stylesheet" media="screen">
		<link href="includes/formulario.css" type="text/css" rel="stylesheet" media="screen">
		<script src="includes/geral_gerencia.js" language="Javascript"></script>
</head>

	<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" bottommargin="0">
	
	<!-- FORMULÁRIO GERAL DA PÁGINA -->
	<form id="geral" name="geral" action="<%=Request.ServerVariables("SCRIPT_NAME")%>" method="post">
	
	
		<!-- UM BOX (DIV) PARA ENGLOBAR TODOS -->
		<div id="principal">

			<!-- CABECALHO E MENU PRICIPAL DA PÁGINA -->
			<!--#Include file="cabecalho.asp"-->


			<!-- DIV PARA ENGLOBAR TODO O CONTEÚDO -->
			<div id="janela">

				<div id="titulo"><%=msgTitulo%></div>


				<div id="ferramenta">
					<!-- BOTOES E FILTROS -->
					
					<!-- BOTOES E FILTROS -->
				</div>
				

				<div id="conteudo">
				
				
					<table width="400" align="center" border="0" cellpadding="0" cellspacing="10">
					<tr>
						<td class="texto3"><%=m%></td>
					</tr>
					<tr>
						<td align="right"><a href="<%=msgLink%>" class="link1">« voltar</a></td>
					</tr>
					</table>
					
					
				</div>


			</div>


		</div>
		
	</form>

	</body>
</html>