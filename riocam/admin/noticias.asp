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
Response.Expires=0
Session.LCID = 1046

DIM msg_erro
Dim query
Dim noticia, titulo, lead, texto, imgG, status, notExpira
Dim sTitulo, sLead, sTexto
Dim data, hora, dataExpiracao, horaExpiracao, sData, sDataExpiracao
Dim oper, cmd, erro
Dim s, i

'variaveis de filro
Dim filtro, fTexto, fCampo, fStatus

'variaveis de lista
Dim branco, nBG
Dim pagina, strPaginas, rCount
Dim ordem, ordemField, ordemType
Dim linhasDetalhe

Const nomeTabela = "XCA_noticias"
oper = Request("oper")
If oper="" Then oper = "L"

If len(request("linhasDetalhe")) < 1 Then linhasDetalhe = 9 Else linhasDetalhe = request("linhasDetalhe")

'Define o filtro da p�gina
fTexto = Request("fTexto") 
fCampo = Request("fCampo")
fStatus = Request("fStatus")


Call Conectar()


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
//			document.forms[0].fMalaDir.value="";
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
			if (pag < 1) { alert('Voc� j� est� na primeira p�gina!'); return true }
			if (pag > nPags) { alert('Voc� j� est� na �ltima p�gina!'); return true }
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
			f.noticia.value = 0;
			if (f.registroSel.length != undefined) {
				for (var i = 0; i < f.registroSel.length; i++) {
					if (f.registroSel[i].checked) { f.noticia.value = f.registroSel[i].value }
				}
			} else {
			f.noticia.value = f.registroSel.value
			}
		}

		function editar() {
			var f = document.forms[0]
			getOption();
			if (f.noticia.value>0) {
				f.oper.value="E";
				f.action="<%= Request.ServerVariables("SCRIPT_NAME") %>";
				f.submit();
			} else {
				alert("� preciso selecionar um registro!")
			}
		}

		function editarLinha(a) {
			document.forms[0].oper.value="E";
			document.forms[0].noticia.value=a;
			document.forms[0].submit();
		}

		function salvar() {
			document.forms[0].cmd.value="S";
			document.forms[0].submit();
		}

		function excluir() {
			var f=document.forms[0];
			var intDels=0;
			for (var i = 0; i < f.elements.length; i++) {
				if ((f.elements[i].name=="anoticia") && (f.elements[i].checked)) {
					intDels++;
				}
			}
			if (intDels>0) {
				if (confirm("Voc� est� prestes a apagar " + intDels + " registros.\n Tem certeza que deseja continuar?")) {
					document.forms[0].action="noticias_excluir.asp";
					document.forms[0].submit();
				}
			} else {
				alert('� preciso selecionar ao menos um registro!')
			}
		}

		function altImagem(cmd, campo, chave, movimento) {
			if (document.forms[0].noticia.value>0) {
				//verifico se � na inclusao ou se � na edi��o do registro.
				var sURL = "dlgAltImagem.asp?cmd=" + cmd + "&nomeTabela=<%=nomeTabela%>&nomeCampoChave=noticia&nomeCampo=" + campo + "&chave=" + chave + "&movimento=" + movimento;
				var sName = "wDlgAlterarImagem";
				var features = "left=300,top=200,width=140,height=250";
			  var _WImg=window.open(sURL, sName, features);
				  _WImg.focus();
			} else {
				if (confirm("� preciso salvar a noticia antes.\nDeseja salvar agora?")) {
					document.forms[0].cmd.value="SF";
					document.forms[0].submit();
				}
			}
		}

		function altFoto(c, a, tipo, n) {
			//verifico se � na inclusao ou se � na edi��o do registro.
			var sURL = "dlgAltFoto.asp?cmd=" + c + "&nomeTabela=<%=nomeTabela%>&nomeCampo=noticia&chave=" + a + "&movimento=" + n + "&tipo=" + tipo;
			var sName = "wDlgAlterarFoto";
			var features = "left=300,top=200,width=180,height=280";
		  var _W=window.open(sURL, sName, features);
			  _W.focus();
		}
	</script>
	
</head>


	<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" bottommargin="0">
	
	<!-- FORMUL�RIO GERAL DA P�GINA -->
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

			<!-- CABECALHO E MENU PRICIPAL DA P�GINA -->
			<!--#Include file="cabecalho.asp"-->


			<!-- DIV PARA ENGLOBAR TODO O CONTE�DO -->
			<div id="janela">

				<div id="titulo">Cadastro de not�cias</div>

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
							<option value="titulo"		<% If fCampo = "titulo"		Then Response.Write " SELECTED " %>>T�tulo
							<option value="lead"		<% If fCampo = "lead"		Then Response.Write " SELECTED " %>>Lead
						</select>
						<select name="fStatus" size="1" class="sel">
							<option value="">Status
							<option value="A" <% If fStatus = "A" Then Response.Write " SELECTED " %>>Ativo
							<option value="P" <% If fStatus = "P" Then Response.Write " SELECTED " %>>Pendente
							<option value="I" <% If fStatus = "I" Then Response.Write " SELECTED " %>>Inativo
						</select>
						<input type="image" src="<%=dirImgsGerencia%>bt_filtro_aplicar.gif" value="Aplicar" onclick="JavaScript: aplicarFiltro();" style="width:20px; height:15px" alt="Aplicar Filtro">
						<input type="image" src="<%=dirImgsGerencia%>bt_filtro_excluir.gif" value="Excluir" onclick="JavaScript: excluirFiltro();" style="width:20px; height:15px" alt="Excluir Filtro">
						
						
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
					<!-- #include file="noticias_incluir_editar.asp"-->
					<% Else %>
					<!-- #include file="noticias_lista.asp"-->
					<% End If %>
				</div>


			</div>


		</div>
		
	</form>

	</body>
</html>