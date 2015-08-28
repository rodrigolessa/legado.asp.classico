<% 
Const maxusuario	= 4
Const maxlogin		= 20
Const maxsenha		= 20
Const maxemail		= 60
Const maxstatus		= 1


'VERIFICA SE É A PRIMEIRA VEZ
If cmd = "" Then
	'VERIFICA SE É A PRIMEIRA VEZ DA EDIÇÃO OU DA INCLUSÃO
	If oper = "E" Then
		'EDITAR
		usuario = Request("usuario")
		query = "SELECT usuario, login, senha, email, status " & _
						"	FROM " & nomeTabela & "	" & _
						"	WHERE usuario=" & usuario & ""
		rs.Open query, conn
		If Not rs.EOF Then
			usuario	= Trim(rs("usuario"))
			login	= Trim(rs("login"))
			senha	= Trim(rs("senha"))
			email	= Trim(rs("email"))
			status	= Trim(rs("status"))
		Else
			erro = "É preciso informar um usuário para poder edita-lo"
		End If
		rs.close
	Else
		'INCLUIR
		oper	= "I"
	End If
Else
	'SEGUNDA VEZ NO FORMULÁRIO, VALIDA DADOS PARA SALVAR
	usuario	= Request("usuario")
	login	= Request("login")
	senha	= Request("senha")
	email	= Request("email")
	status	= Request("status")	
	
	If Len(login) < 1 Or Len(login) > maxlogin Then erro = " - O login deve ser informado e deve conter no máximo " & maxlogin & " caracteres<br>"
	If Len(senha) < 1 Or Len(senha) > maxsenha Then erro = " - A senha deve ser informada e deve conter no máximo " & maxSenha & " caracteres<br>"
	'PREPARA CAMPO PARA SALVAR
	If erro = "" Then
		login = Replace(login, "'", "''")
		If oper = "I" Then
			query = "SELECT usuario FROM " & nomeTabela & " WHERE login = '" & login & "'"
			rs.Open query, conn
			If Not rs.EOF Then erro = erro & "Já existe um usuário com este login"
			rs.close

			If erro= "" Then
				query = "INSERT INTO " & nomeTabela & " (login, senha, email, status) VALUES " & _
						"	('" & login & "', '" & senha & "', '" & email & "', '" & status & "')"
				conn.execute query
				usuario	= ""
				login	= ""
				senha	= ""
				email	= ""
				status	= ""
				oper	= "I"
				cmd		= ""
				%>
				<script language="JavaScript">
					alert("Usuário cadastrado com sucesso!!")
				</script>
				<%
			End If
		Else
			'grava edicao
			query = "SELECT usuario FROM " & nomeTabela & " WHERE login = '" & login & "' AND usuario <> " & usuario & ""
			rs.Open query, conn
			If Not rs.EOF Then erro = "Já existe um usuário com este login<br>"
			rs.close
			If erro= "" Then
				query = "UPDATE " & nomeTabela & " SET " & _
						"	login = '" & login & "', " & _
						"	senha = '" & senha & "', " & _
						"	email = '" & email & "', " & _
						"	status = '" & status & "' " & _
						"WHERE usuario=" & usuario & ""
				conn.Execute query
				Response.Redirect "usuariosadm.asp?" & strRedir
				Response.End
			End If
		End If
	End If
End If
%>
<table width="100%" height="100%" border="0">
	<tr>
		<td valign="top">
			<table width="100%" cellpadding="3" cellspacing="0" border="0">
				<tr>
					<td>
						<div class="erro" align="right">* preenchimento obrigatório</div>
						<table border="0" cellpadding="0" cellspacing="5" class="texto2">
							

<%		
		If oper = "E" Then
			Response.Write "<input type='hidden' name='usuario' value='" & usuario & "'>"
		End If
%> 

							<% If erro <> "" Then %>
								<tr>
									<td class="erro">&nbsp;ERRO:</td>
									<td >&nbsp;&nbsp;</td>
									<td class="erro"><%=erro%></td>
								</tr>
							<% End If %>
							<tr>
								<td><span class="erro"><sup>*</sup></span>&nbsp;Login:</td>
								<td>&nbsp;&nbsp;</td>
								<td><input type="text" name="login" class="cx" size="50" maxlength="<%=maxlogin%>" value="<%=login%>"></td>
							</tr>
							<tr>
								<td><span class="erro"><sup>*</sup></span>&nbsp;Senha:</td>
								<td>&nbsp;&nbsp;</td>
								<td><input type="password" name="senha" class="cx" size="50" maxlength="<%=maxsenha%>" value="<%=senha%>"></td>
							</tr>
							<tr>
								<td>&nbsp;&nbsp;E-mail:</td>
								<td>&nbsp;&nbsp;</td>
								<td><input type="text" name="email" class="cx" size="50" maxlength="<%=maxemail%>" value="<%=email%>"></td>
							</tr>
							<tr>
								<td><span class="erro"><sup>*</sup></span>&nbsp;status:</td>
								<td>&nbsp;&nbsp;</td>
								<td><% Call criaCombostatus(status) %></td>
							</tr>
							<%
								If oper="I" Then
							%>
							<script language="javascript1.2">
								document.forms[0].login.focus()
							</script>
							<%
								End If
							%>
						</table>
					</td>
				</tr>
			</table>
		</td>
	</tr>
</table>
<%
Sub criaCombostatus(status)
		s = "<select name='status' class='sel'>" & vbCrLf
		s = s & "<option value='A' " 
		If status = "A" Then s = s & " SELECTED "
		s = s & ">Ativo" & vbCrLf
		s = s & "<option value='I' " 
		If status = "I" Then s = s & " SELECTED "
		s = s & ">Inativo" & vbCrLf
		s = s & "</select>" & vbCrLf
		Response.Write s
End Sub

Sub criaComboGrupo(grupo)
	Dim rs, query
	Dim s
	query = "Select grupo, nome From XCA_grupos Order By nome"
	Set rs = Server.CreateObject("ADODB.RecordSet")
	rs.ActiveConnection = conn
	rs.Open query
	If Not rs.Eof Then
		s = "<select name='grupo' class='sel'>" & vbCrLf & _
				"<option value=''> ..." & vbCrLf
		While NOt rs.Eof
			s = s & "<option value='" & Trim(rs("grupo")) & "' " 
			If Not IsNull(grupo) Then
				If Cstr(grupo) = Cstr(rs("grupo")) Then s = s & " SELECTED "
			End If
			s = s & ">" & Trim(rs("nome")) & vbCrLf
			rs.MoveNext
		Wend
		s = s & "</select>" & vbCrLf
		Response.Write s
	Else
		Response.Write "Nenhum grupo cadastrado"
	End If
	rs.close
	Set rs = Nothing
End Sub
%>