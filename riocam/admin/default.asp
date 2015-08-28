<% Option Explicit %>
<!--#include file="includes/incConnectDB.asp"-->
<%
Response.Buffer = True
Response.AddHeader "cache-control", "private"
Response.AddHeader "pragma", "no-cache"
Response.ExpiresAbsolute = #January 1, 1990 00:00:01#
Response.Expires=0
Session.LCID = 1046

DIM erro, msg_erro
DIM login, senha
DIM nUsu, sUsu
			
If Request("oper") <> "" Then
	If Request("oper") = "G" Then
		login = trim(request("login"))
		senha = trim(request("senha"))
		If login<>"" And senha<>"" Then
			sSQL = "SELECT	usuario, login, email, status " & _
					"FROM	XCA_usuariosADM " & _
					"WHERE	login = '" & Replace(login, "'", "''") & "' " & _
					"	AND	senha = '" & Replace(senha, "'", "''") & "' "
			Call conectar()
			rs.open sSQL, conn
			If Not rs.EOF Then
				Session("XCA_LOGADO_ADM") = "TRUE"
				Session("XCA_USUARIO_ADM") = rs("usuario")
				Session("XCA_LOGIN") = trim(rs("login"))
				Session("XCA_EMAIL") = trim(rs("email"))
			Else
				Session("XCA_LOGADO_ADM") = "FALSE"
				Session("XCA_USUARIO_ADM") = 0
				Session("XCA_LOGIN") = ""
				Session("XCA_EMAIL") = ""
				msg_erro = "Usuário ou senha inválido!"
			End If
		Else
			msg_erro = "Usuário ou senha inválido!"
		End If
	Else
		'logoff
		Session("XCA_LOGADO_ADM") = "FALSE"
		Session("XCA_USUARIO_ADM") = 0
		Session("XCA_LOGIN") = ""
		Session("XCA_EMAIL") = ""
	End If
End If
 %>
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
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
	<input type="Hidden" name="oper"	value="G">
	<input type="Hidden" name="cmd"		value="">
	
		<!-- UM BOX (DIV) PARA ENGLOBAR TODOS -->
		<div id="principal">

			<!-- CABECALHO E MENU PRICIPAL DA PÁGINA -->
			<!--#Include file="cabecalho.asp"-->


			<!-- DIV PARA ENGLOBAR TODO O CONTEÚDO -->
			<div id="janela">

				<div id="titulo">Login do Sistema</div>

				<div id="ferramenta">
				</div>

				<div id="faixa"></div>

			<%	if len(msg_erro) > 0 then	%>
				<div id="msg_erro"><%=msg_erro%></div>
			<%	end if	%>

				<div id="conteudo"><!-- #include file="login.asp"--></div>

			</div>


		</div>
		
	</form>

	</body>
</html>