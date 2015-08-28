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
	
	Const maxLogin = 20
	Const maxSenha = 10
	Const maxNome = 100
	Const maxCidade = 30
	Const maxEstado = 2
	Const maxPais = 30
	Const maxEmail = 60
	Const maxTelefone = 16
	Const maxIdade = 2
	Const maxAltura = 4
	Const maxEsporte = 100
	Const maxOlhos = 50
	Const maxCabelo = 50
	Const maxBusto = 50
	Const maxCintura = 50
	Const maxQuadril = 50
	Const maxStatus = 1
	Const maxSala = 10

	
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
			
			function incluir() {
				document.forms[0].oper.value="I";
				document.forms[0].submit();
			}
			
			function editarLinha(a) {
				document.forms[0].oper.value="E";
				document.forms[0].modelo.value=a;
				document.forms[0].submit();
			}
			
			function salvar() {
				var f = document.forms[0];
					f.cmd.value="S";
					f.action="<%= Request.ServerVariables("SCRIPT_NAME") %>";
					f.submit();
			}

			function cadFoto() {
				var w			= 320;
				var h			= 360;
				var parametros	= "toolbar=no,location=no,directories=no,status=yes,menubar=no,scrollbars=no,resizable=no,copyhistory=no,width="+w+",height="+h+",top="+((screen.height / 2) - (h / 2))+",left="+((screen.width / 2) - (w / 2));
				var f 			= document.forms[0];
				getOption();
				if (f.modelo.value>0) {
					var sURL		= "dlgFotoMod.asp?modelo="+f.modelo.value;
					var	fotoMod		= window.open(sURL, "fotoMod", parametros);
						fotoMod.focus();
				} else {
					alert("É preciso selecionar um registro!");
				}
				return false;
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
						
						<p class="inter_titulo"><%=id_getText("cadastro_modelo_txt_01")%></p>
						
						<div id="centro_espaco"></div>
						
						<div id="conteudo">
							<% If oper = "I" OR oper = "E" Then %>
							<!-- #include file="cadastro_modelo_incluir_editar.asp"-->
							<% ElseIf oper = "F" Then %>
							<!-- #include file="cadastro_modelo_fim.asp"-->
							<% Else %>
							<!-- #include file="cadastro_modelo_inicio.asp"-->
							<% End If %>
						</div>

					</div>
					<!-- FIM BOX DE CONTEÚDO DA PÁGINA -->

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