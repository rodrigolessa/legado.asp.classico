<%
	'***** VERIFICA SE A MODELO DESEJA EFETUAR O LOGIN *****
	If len(oper) > 0 Then

		If oper = "G" Then

			login_modelo = trim(request("lg_usuario_modelo"))
			senha_modelo = trim(request("lg_senha_modelo"))

			If login_modelo <> "" And senha_modelo <> "" Then

				sSQL = "SELECT	modelo, login, email, status, sala " & _
						"FROM	XCA_modelos " & _
						"WHERE	login = '" & replace(login_modelo, "'", "''") & "' " & _
						"	AND	senha = '" & replace(senha_modelo, "'", "''") & "' " & _
						"	AND	status = 'A' "
				call conectar()

				rs.open sSQL, conn

				if not rs.EOF then
					session("XCA_MODELO_LOGADO_SITE")	= "TRUE"
					session("XCA_USUARIO_MODELO")		= rs("modelo")
					session("XCA_USUARIO_MODELO_LOGIN")	= lCase(trim(rs("login")))
					session("XCA_USUARIO_MODELO_EMAIL")	= lCase(trim(rs("email")))
					session("XCA_USUARIO_MODELO_SALA")	= lCase(trim(rs("sala")))
				else
					session("XCA_MODELO_LOGADO_SITE")	= "FALSE"
					session("XCA_USUARIO_MODELO")		= 0
					session("XCA_USUARIO_MODELO_LOGIN")	= ""
					session("XCA_USUARIO_MODELO_EMAIL")	= ""
					session("XCA_USUARIO_MODELO_SALA")	= ""
					msg_erro = " - Usuário ou senha inválido! <br>"
				end if

			Else
				msg_erro = " - Usuário ou senha inválido! <br>"
			End If

		Else
		
			'LOGOFF
			'session("XCA_LOGADO_SITE") = "FALSE"
			'session("XCA_USUARIO") = 0
			'session("XCA_USUARIO_LOGIN") = ""
			'session("XCA_USUARIO_EMAIL") = ""
			
		End If

	End If
	

	if len(msg_erro) > 0 then
%>
		<div id="msg_erro"><%=msg_erro%></div>
<%
	end if


	If session("XCA_MODELO_LOGADO_SITE") <> "TRUE" Then
%>
	<div id="titulo_login" style="text-align:left;width:200px;"><%=id_getText("login_txt_01")%></div>
	<div id="form_login" style="width:200px;">
		<%=id_getText("login_txt_02")%>
		<input type="text" id="lg_usuario_modelo" name="lg_usuario_modelo" value="" class="lg_tx">
		
		<%=id_getText("login_txt_03")%>
		<input type="password" id="lg_senha_modelo" name="lg_senha_modelo" value="" class="lg_tx2">
		
		<input type="button" id="lg_botao" name="lg_botao" value="<%=id_getText("login_txt_04")%>" class="lg_bt" onCLick="javaScript: logarSite();">
	</div>
<%
	Else
	
		response.redirect("camx.asp?oper=M&modelo="&session("XCA_USUARIO_MODELO")&"&modelo_login="&session("XCA_USUARIO_MODELO_LOGIN")&"&sala="&session("XCA_USUARIO_MODELO_SALA"))
		
'		response.write	"<div id=""titulo_login"" style=""text-align:left;width:200px;"">" & id_getText("login_txt_01") & "</div>" & _
'						"<div id=""form_logado"" style=""width:200px;"">" & _
'						uCase(Left(session("XCA_USUARIO_MODELO_LOGIN"), 1)) & _
'						Right(session("XCA_USUARIO_MODELO_LOGIN"), Len(session("XCA_USUARIO_MODELO_LOGIN"))-1) & ",<br>" & _
'						id_getText("login_txt_05") & _
'						"</div>"
	End If
%>