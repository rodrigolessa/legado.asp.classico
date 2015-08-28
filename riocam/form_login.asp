<%
	'***** VERIFICA SE O USUARIO DESEJA EFETUAR O LOGIN *****
	If len(oper) > 0 Then

		If oper = "G" Then

			login = trim(request("lg_usuario"))
			senha = trim(request("lg_senha"))

			If login <> "" And senha <> "" Then

				sSQL = "SELECT	usuario, login, email, status " & _
						"FROM	XCA_usuarios " & _
						"WHERE	login = '" & replace(login, "'", "''") & "' " & _
						"	AND	senha = '" & replace(senha, "'", "''") & "' " & _
						"	AND	status = 'A' "
				call conectar()

				rs.open sSQL, conn

				if not rs.EOF then
					session("XCA_LOGADO_SITE")		= "TRUE"
					session("XCA_USUARIO")			= rs("usuario")
					session("XCA_USUARIO_LOGIN")	= trim(rs("login"))
					session("XCA_USUARIO_EMAIL")	= trim(rs("email"))
					'FAZ O CONTROLE DE ACESSO DOS USUÁRIOS
					addAcessoUsuario rs("usuario")
				else
					session("XCA_LOGADO_SITE")		= "FALSE"
					session("XCA_USUARIO")			= 0
					session("XCA_USUARIO_LOGIN")	= ""
					session("XCA_USUARIO_EMAIL")	= ""
					msg_erro	= " - Usuário ou senha inválido! <br>"
				end if

			Else

				msg_erro		= " - Usuário ou senha inválido! <br>"

			End If

		Else

			'LOGOFF
'			session("XCA_LOGADO_SITE") = "FALSE"
'			session("XCA_USUARIO") = 0
'			session("XCA_USUARIO_LOGIN") = ""
'			session("XCA_USUARIO_EMAIL") = ""

		End If

	End If


	If session("XCA_LOGADO_SITE") <> "TRUE" Then
%>
	<div id="titulo_login"><%=id_getText("login_txt_01")%></div>
	<div id="form_login">
		<%=id_getText("login_txt_02")%>
		<input type="text" id="lg_usuario" name="lg_usuario" value="" class="lg_tx">
		
		<%=id_getText("login_txt_03")%>
		<input type="password" id="lg_senha" name="lg_senha" value="" class="lg_tx2">
		
		<input type="button" id="lg_botao" name="lg_botao" value="<%=id_getText("login_txt_04")%>" class="lg_bt" onCLick="javaScript: logarSite();">
	</div>
<%
	Else
		response.write	"<div id=""titulo_login"">" & _
						uCase(Left(session("XCA_USUARIO_LOGIN"), 1)) & _
						Right(session("XCA_USUARIO_LOGIN"), Len(session("XCA_USUARIO_LOGIN"))-1) & ", " & _
						id_getText("login_txt_05") & _
						"</div>"
	End If
%>