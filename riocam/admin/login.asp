<%
	If Session("XCA_LOGADO_ADM") <> "TRUE" Then
%>


<div id="label">Usuário:</div>
<div id="campo"><input type="text" name="login" value="<%=login%>" class="cx"></div>
<div id="label">Senha:</div>
<div id="campo"><input type="password" name="senha" class="cx"></div>
<div id="botao"><input type="submit" name="btLogar" value="Logar" class="bt"></div>
<div id="quebra" />
<div id="label_small">Área destinada para os usuários administradores deste site...</div>


<%
	Else
			nUsu = Application("XCA_UserCount") - 1
			If nUsu > 1 Then 
					sUsu = nUsu & " sessões de usuários estão ativas"
			Else
				If nUsu = 1 Then
					sUsu = " 1 sessão esta ativa"
				Else
					sUsu = "nenhuma sessão de usuário esta ativa"
				End If
			End If
			Response.Write	UCase(Left(Session("XCA_LOGIN"), 1)) & _
							Right(Session("XCA_LOGIN"), Len(Session("XCA_LOGIN"))-1) & ",<br><br>" & _
							"Seja bem vindo ao site administrativo " & _
							Application("XCA_APP_TITLE_SHORT") & "!<br><br>" & _
							"No momento " & sUsu & "<br><br>"
	End If
%>
