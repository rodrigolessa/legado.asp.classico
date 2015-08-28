<% option explicit %>
<!--#include file="includes/incConnectDB.asp" -->
<!--#include file="includes/geral.asp" -->
<!--#Include file="includes/incEnviaEmail.asp"-->
<%
'Response.Buffer = False
Response.AddHeader "cache-control", "private"
Response.AddHeader "pragma", "no-cache"
Response.ExpiresAbsolute = #January 1, 1990 00:00:01#
Response.Expires = 0
Session.LCID = 1046
	
	'***** VARIAVEIS GERAIS ******
	DIM oper, cmd, erro, msg_erro, msg, strRedir
	DIM login, senha
	DIM i
	
	'***** VARIAVEIS E-MAIL ******
	DIM	usuario, nome, email, bairro, cidade
	DIM	estado, pais, telefone
	DIM nomeFrom, emailFrom, nomeTo, emailTo
	DIM	strSubject, strBody, strMsg
	
	Const maxNome		= 60
	Const maxEmail		= 60
	Const maxBairro		= 30
	Const maxCidade		= 30
	Const maxTelefone	= 15

	
	'***** ABRE A CONEXÃO *****
	Call conectar()

	
	'***** RECUPERA AS VARIAVEIS ******
	if len(trim(request("oper"))) > 0 then
		oper = cStr(uCase(trim(request("oper"))))
	else
		oper = "L"
	end if
	
	if len(trim(request("cmd"))) > 0 then
		cmd = cStr(uCase(trim(request("cmd"))))
	else
		cmd = ""
	end if
	
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
		
		<script language="JavaScript">
		//Variáveis Globais

		//Funções gerais
			function abreJanela(a,w,h) {
				var b
				b = b + "toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars="+s+",resizable=no,copyhistory=no";
				b = b + ",width="+w+",height="+h+",top="+((screen.height/2)-(h/2))+",left="+((screen.width/2)-(w/2));
				window.open(a,"_blank",b);
			}

			function validarEmail(valor) {
				if (/^\w+([\.-]?\w+)*@\w+([\.-]?\w+)*(\.\w{2,3})+$/.test(valor)){
					return true
				} else {
					return false;
				}
			}

			function enviaFaleConosco() {
				var f0 = document.forms[0];
				var fok = false;
				if (validarEmail(f0.email.value) == true) fok = true;
				if (fok == true) {	
		//			f0.oper.value="F";	
					f0.cmd.value="E";
					f0.action="contato.asp";
					f0.submit();
				} else {
					alert ("<%=id_getText("msg_erro_05")%>");
				}
			}
		</script>
	</head>
	<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" bottommargin="0">

	<!-- FORMULÁRIO GERAL DA PÁGINA -->
	<form id="geral" name="geral" action="<%=Request.ServerVariables("SCRIPT_NAME")%>" method="post">
	<input type="Hidden" name="oper"		value="">
	<input type="Hidden" name="cmd"			value="">
	<input type="Hidden" name="strSubject"	value="Contato de visitante">


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
						
						<script language="javaScript">
							criaFlash("assine.swf", "750", "400");
						</script>
						
						<div id="centro_espaco"></div>

					</div>
					<!-- FIM BOX DE CONTEÚDO DA PÁGINA -->

				</div>
				<!-- FIM BOX DE CONTEÚDO DA PÁGINA -->
				
				<!-- BOX DE MENU PRINCIPAL DA PÁGINA -->
				<!-- # include file="menu.asp" -->

				<!-- BOX DE RODAPÉ DA PÁGINA -->
				<!--#include file="rodape.asp" -->

			</div>
			<!-- FIM BOX PÁGINA PARA CONTER TODO O LAYOUT -->

		</div>
		<!-- FIM BOX (DIV) PARA ENGLOBAR TODOS -->


	</form>

	</body>
</html>