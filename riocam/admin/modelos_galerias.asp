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
DIM codmodelo_galeria, dirmodelo_galeria, dirModRaiz, dirModGal, dirModVid
DIM modelo, modelo_galeria, login, senha, nome, titulo, titulo_eng, modelo_galeria_categ
DIM descricao, descricao_eng, pais, tipo, data, status
DIM olhos, cabelo, busto, cintura, quadril, dirModeloGaleria
DIM olhos_eng, cabelo_eng, busto_eng, cintura_eng, quadril_eng
DIM addFotos
DIM oper, cmd, erro, erro2
DIM msg_erro
DIM s, i, codRemocao

'VARIAVEIS DE FILRO
Dim filtro, fTexto, fCampo, fTipo, fmalaDir, fStatus
Dim filtroLista

'VARIAVEIS DE LISTA
Dim branco, nBG
Dim pagina, strPaginas, rCount
Dim ordem, ordemField, ordemType
Dim linhasDetalhe

Const nomeTabela = "XCA_modelos_galerias"
Const lenCodRemocao = 10

Call Conectar()

oper = Request("oper")
if oper = "" then
	oper = "L"
end if

If len(request("linhasDetalhe")) < 1 Then linhasDetalhe = 10 Else linhasDetalhe = request("linhasDetalhe")

'DEFINE O FILTRO DA PÁGINA
fTexto		= Request("fTexto") 
fCampo		= Request("fCampo")
fMalaDir	= Request("fMalaDir")
fStatus		= Request("fStatus")

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
	<title><%=Application("XCA_APP_TITLE")%></title>
	<link href="includes/box.css" type="text/css" rel="stylesheet" media="screen">
	<link href="includes/geral.css" type="text/css" rel="stylesheet" media="screen">
	<link href="includes/formulario.css" type="text/css" rel="stylesheet" media="screen">
	<script src="includes/geral_gerencia.js" language="Javascript"></script>

	<script language="JavaScript" type="text/javascript">
		function aplicarFiltro() {
			document.forms[0].oper.value="L";
			document.forms[0].submit();
		}

		function excluirFiltro() {
			document.forms[0].oper.value="L";
			document.forms[0].fTexto.value="";
			document.forms[0].fCampo.value="";
			document.forms[0].fMalaDir.value="";
			document.forms[0].fStatus.value="";
			document.forms[0].submit();
		}

		function mudarLinhasDetalhes(a) {
			document.forms[0].oper.value="L";
			document.forms[0].linhasDetalhe.value=a;
			document.forms[0].submit();
		}

		function navegar(pag) {
			var nPags = parseInt(document.forms[0].nPaginas.value)
			if (pag < 1) { alert("Você já está na primeira página!"); return true }
			if (pag > nPags) { alert("Você já está na última página!"); return true }
			document.forms[0].oper.value="L";
			document.forms[0].pagina.value=pag;
			document.forms[0].submit();
		}

		function ordenar(f, t) {
			document.forms[0].oper.value="L";
			document.forms[0].ordemField.value=f;
			document.forms[0].ordemType.value=t;
			document.forms[0].submit();
		}

		function listar() {
			document.forms[0].oper.value="L";
			document.forms[0].submit();
		}

		function incluir() {
			document.forms[0].oper.value="I";
			document.forms[0].submit();
		}

		function getOption() {
			var f = document.forms[0]
			f.modelo_galeria.value = 0;
			if (f.registroSel.length != undefined) {
				for (var i = 0; i < f.registroSel.length; i++) {
					if (f.registroSel[i].checked) { f.modelo_galeria.value = f.registroSel[i].value }
				}
			} else {
			f.modelo_galeria.value = f.registroSel.value
			}
		}

		function editar() {
			var f = document.forms[0];
			getOption();
			if (f.modelo_galeria.value>0) {
				f.oper.value="E";
				f.action="<%= Request.ServerVariables("SCRIPT_NAME") %>";
				f.submit();
			} else {
				alert("É preciso selecionar um registro!")
			}
		}

		function editarLinha(a) {
			document.forms[0].oper.value="E";
			document.forms[0].modelo_galeria.value=a;
			document.forms[0].submit();
		}

		function salvar() {
			var f = document.forms[0];
				f.cmd.value="S";
				f.action="<%= Request.ServerVariables("SCRIPT_NAME") %>";
				f.submit();
		}


		function excluir() {
			var f=document.forms[0];
			var intDels=0;
			for (var i = 0; i < f.elements.length; i++) {
				if ((f.elements[i].name=="amodelo_galeria") && (f.elements[i].checked)) {
					intDels++;
				}
			}
			if (intDels>0) {
				if (confirm("Você está prestes a apagar " + intDels + " registros.\n Tem certeza que deseja continuar?")) {
					document.forms[0].action="modelos_galerias_excluir.asp";
					document.forms[0].submit();
				}
			} else {
				alert("É preciso selecionar ao menos um registro!")
			}
		}
		
		function addFoto(nModeloGaleria) {
			var w			= 360;
			var h			= 300;
			var parametros	= "toolbar=no,location=no,directories=no,status=yes,menubar=no,scrollbars=yes,resizable=no,copyhistory=no,width="+w+",height="+h+",top="+((screen.height / 2) - (h / 2))+",left="+((screen.width / 2) - (w / 2));
			var sURL		= "dlgFotoGaleria.asp?modelo_galeria="+nModeloGaleria;
			var	fotoMod		= window.open(sURL, "fotoGaleria", parametros);
				fotoMod.focus();
		}

	</script>
	
</head>

	<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" bottommargin="0">
	
	<!-- FORMULÁRIO GERAL DA PÁGINA -->
	<form id="geral" name="geral" action="<%=Request.ServerVariables("SCRIPT_NAME")%>" method="post">
	
<!--
	<input type="hidden" name="oper" value="<%=oper%>">
	<input type="hidden" name="cmd">
	
	<input type="Hidden" name="pagina" value="<%=pagina%>">
	
	<input type="Hidden" name="ordemField" value="<%=ordemField%>">
	<input type="Hidden" name="ordemType" value="<%=ordemType%>">
-->	
	
		<!-- UM BOX (DIV) PARA ENGLOBAR TODOS -->
		<div id="principal">

			<!-- CABECALHO E MENU PRICIPAL DA PÁGINA -->
			<!--#Include file="cabecalho.asp"-->

			<!-- DIV PARA ENGLOBAR TODO O CONTEÚDO -->
			<div id="janela">

				<div id="titulo">Cadastro de galerias de fotos</div>

				<div id="ferramenta">
					<!-- BOTOES E FILTROS -->
					<%
						If oper = "I" OR oper = "E" Then
					%>					
						<img src="<%=dirImgsGerencia%>ico_sep.gif"			width=2		height=18	border=0>
						<img src="<%=dirImgsGerencia%>bt_novo_off.gif"		width=20	height=15	border=0	alt="Incluir">
						<img src="<%=dirImgsGerencia%>bt_editar_off.gif"	width=20	height=15	border=0	alt="Editar">
						<a href="#" onclick="salvar();	return false;"><img src="<%=dirImgsGerencia%>bt_salvar_on.gif"	width=20	height=15	border=0	alt="Salvar"></a>
						<a href="#" onclick="listar();	return false;"><img src="<%=dirImgsGerencia%>bt_abrir_on.gif"	width=20	height=15	border=0	alt="Listar"></a>
						<img src="<%=dirImgsGerencia%>bt_excluir_off.gif"	width=20	height=15	border=0	alt="Excluir">
						<img src="<%=dirImgsGerencia%>ico_sep.gif"			width=2		height=18	border=0>
					<%
						Else
					%>
						<img src="<%=dirImgsGerencia%>ico_sep.gif"			width=2		height=18	border=0>
						<a href="#" onclick="incluir();	return false;"><img src="<%=dirImgsGerencia%>bt_novo_on.gif"	width=20	height=15	border=0	alt="Incluir"></a>
						<a href="#" onclick="editar();	return false;"><img src="<%=dirImgsGerencia%>bt_editar_on.gif"	width=20	height=15	border=0	alt="Editar"></a>
						<img src="<%=dirImgsGerencia%>bt_salvar_off.gif"	width=20	height=15	border=0	alt="Salvar">
						<img src="<%=dirImgsGerencia%>bt_abrir_off.gif"		width=20	height=15	border=0	alt="Listar">
						<a href="#" onclick="excluir();	return false;"><img src="<%=dirImgsGerencia%>bt_excluir_on.gif"	width=20	height=15	border=0	alt="Excluir"></a>
						<img src="<%=dirImgsGerencia%>ico_sep.gif"			width=2		height=18	border=0>
						linhas:
						<select name="linhasDetalhe" onchange="JavaScript: mudarLinhasDetalhes(this.value)" class="sel">
							<option value="5"  <% If linhasDetalhe="5"  Then Response.Write " SELECTED " %>>5</option>
							<option value="10" <% If linhasDetalhe="10" or linhasDetalhe="9" Then Response.Write " SELECTED " %>>10</option>
							<option value="15" <% If linhasDetalhe="15" Then Response.Write " SELECTED " %>>15</option>
							<option value="20" <% If linhasDetalhe="20" Then Response.Write " SELECTED " %>>20</option>
							<option value="25" <% If linhasDetalhe="25" Then Response.Write " SELECTED " %>>25</option>
							<option value="30" <% If linhasDetalhe="30" Then Response.Write " SELECTED " %>>30</option>
							<option value="40" <% If linhasDetalhe="40" Then Response.Write " SELECTED " %>>40</option>
							<option value="50" <% If linhasDetalhe="50" Then Response.Write " SELECTED " %>>50</option>
						</select>
						<img src="<%=dirImgsGerencia%>ico_sep.gif"			width=2		height=18	border=0>
						
						
						filtro:
						<input type="Text" size="10" name="fTexto" class="cx" value="<%=fTexto%>">
						<select name="fCampo" size="1" class="sel">
							<option value="nome"		<% If fCampo = "nome"		Then Response.Write " SELECTED " %>>Nome
							<option value="titulo"		<% If fCampo = "titulo"		Then Response.Write " SELECTED " %>>Título
						</select>
						<select name="fStatus" size="1" class="sel">
							<option value="">Status
							<option value="A" <% If fStatus = "A" Then Response.Write " SELECTED " %>>Ativo
							<option value="P" <% If fStatus = "P" Then Response.Write " SELECTED " %>>Pendente
							<option value="I" <% If fStatus = "I" Then Response.Write " SELECTED " %>>Inativo
						</select>
						<input type="image" src="<%=dirImgsGerencia%>bt_filtro_aplicar.gif" value="Aplicar" onclick="JavaScript: aplicarFiltro();" style="width:20px; height:15px">
						<input type="image" src="<%=dirImgsGerencia%>bt_filtro_excluir.gif" value="Excluir" onclick="JavaScript: excluirFiltro();" style="width:20px; height:15px">
						
						
					<%
						End If
					%>
					<!-- BOTOES E FILTROS -->
				</div>


			<%	if len(trim(msg_erro)) > 0 then	%>
				<div id="msg_erro"><%=msg_erro%></div>
			<%	end if	%>


				<div id="conteudo">
					<% If oper = "I" OR oper = "E" Then %>
					<!-- #include file="modelos_galerias_incluir_editar.asp"-->
					<% Else %>
					<!-- #include file="modelos_galerias_lista.asp"-->
					<% End If %>
				</div>


			</div>


		</div>
		
	</form>

	</body>
</html>