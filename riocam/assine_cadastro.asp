<% Option Explicit %>
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
	DIM msg_erro
	Dim query
	Dim usuario, login, senha, nome, razao, cnpj, endereco, bairro, cidade, estado
	Dim cep, pais, email, telefone, tipo, dataCadastro, malaDir, status, senhaConf
	DIM aMaxUsu
	Dim oper, cmd, erro, erro2
	Dim s, i, codRemocao

	'VARIAVEIS DO CADASTRO INSTITUICAO
	Dim professor, nomeInstituicao, endereco2, bairro2, cidade2, estado2, cep2, telefone2
	Dim sNomeInstituicao

	'VARIAVEIS DE FILRO
	Dim filtro, fTexto, fCampo, fTipo, fmalaDir, fStatus
	Dim filtroUsuario

	'VARIAVEIS DE LISTA
	Dim branco, nBG
	Dim pagina, strPaginas, rCount
	Dim ordem, ordemField, ordemType
	Dim linhasDetalhe

	Const nomeTabela	= "XCA_usuarios"
	Const lenCodRemocao	= 10
	
	'***** ABRE A CONEX�O *****
	Call conectar()

	
	'***** RECUPERA AS VARIAVEIS ******
	if len(trim(request("oper"))) > 0 then
		oper = cStr(uCase(trim(request("oper"))))
	else
		oper = "I"
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
		//Vari�veis Globais

		//Fun��es gerais
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
			
			function salvar() {
				var f = document.forms[0]
				if (validarEmail(f.email.value) == true) {
					f.cmd.value="S";
					f.action="<%= Request.ServerVariables("SCRIPT_NAME") %>";
					f.submit();
				} else {
					alert ("O E-mail digitado est� incorreto!");
				}
			}


			function excluir() {
				var f=document.forms[0];
				var intDels=0;
				for (var i = 0; i < f.elements.length; i++) {
					if ((f.elements[i].name=="ausuario") && (f.elements[i].checked)) {
						intDels++;
					}
				}
				if (intDels>0) {
					if (confirm("Voc� est� prestes a apagar " + intDels + " registros.\n Tem certeza que deseja continuar?")) {
						document.forms[0].action="usuarios_excluir.asp";
						document.forms[0].submit();
					}
				} else {
					alert("� preciso selecionar ao menos um registro!")
				}
			}
			
		</script>
	</head>
	<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" bottommargin="0">

	<!-- FORMUL�RIO GERAL DA P�GINA -->
	<form id="geral" name="geral" action="<%=Request.ServerVariables("SCRIPT_NAME")%>" method="post">
	

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

						<div id="centro_espaco"></div>
						
						<p class="inter_titulo"><%=id_getText("cadastro_txt_17")%></p>
						
						<% If oper = "I" OR oper = "E" Then %>
						<!-- #include file="assine_cadastro_incluir.asp"-->
						<% ElseIf oper = "F" OR oper = "G" then %>
						<!-- #include file="assine_cadastro_finalizar.asp"-->
						<% End If %>

					</div>
					<!-- FIM BOX DE CONTE�DO DA P�GINA -->

				</div>
				<!-- FIM BOX DE CONTE�DO DA P�GINA -->
				
				<!-- BOX DE MENU PRINCIPAL DA P�GINA -->
\				<!-- # include file="menu.asp" -->

				<!-- BOX DE RODAP� DA P�GINA -->
				<!--#include file="rodape.asp" -->

			</div>
			<!-- FIM BOX P�GINA PARA CONTER TODO O LAYOUT -->

		</div>
		<!-- FIM BOX (DIV) PARA ENGLOBAR TODOS -->


	</form>

	</body>
</html>