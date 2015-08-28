<% 
DIM strRedir
CONST maxLogin		= 30
CONST maxSenha		= 20
CONST maxNome		= 60
CONST maxEndereco	= 90
CONST maxBairro		= 30
CONST maxCidade		= 30
CONST maxEstado		= 2
CONST maxCep		= 10
CONST maxPais		= 15
CONST maxEmail		= 60
CONST maxTelefone	= 40
CONST maxStatus		= 1

strRedir = "pagina=" & request("pagina") & "&ordemField=" & request("ordemField") & "&ordemType=" & request("ordemType") & "&fCampo=" & request("fCampo") & "&fTexto=" & request("fTexto") & "&fStatus=" & request("fStatus") & "&fMalaDir=" & request("fMalaDir") 
'RESPONSE.WRITE STRREDIR
'VERIFICA SE É A PRIMEIRA VEZ
If request("cmd")="" Then
	'VERIFICA SE É A PRIMEIRA VEZ DA EDIÇÃO OU DA INCLUSÃO
	If oper = "E" Then
		'EDITAR
		usuario = request("usuario")
		query =	("SELECT usuario, login, senha, nome, endereco, bairro, cidade, estado, " & _
						"			cep, pais, email, telefone, dataCadastro, malaDir, status " & _
						"	FROM " & nomeTabela & " WHERE usuario=" & usuario & "")
		rs.Open query, conn
		If Not rs.EOF Then
			usuario			= trim(rs("usuario"))
			login			= trim(rs("login"))
			senha			= trim(rs("senha"))
			nome			= trim(rs("nome"))
			endereco		= trim(rs("endereco"))
			bairro			= trim(rs("bairro"))
			cidade			= trim(rs("cidade"))
			estado			= uCase(trim(rs("estado")))
			cep				= trim(rs("cep"))
			pais			= trim(rs("pais"))
			email			= lCase(trim(rs("email")))
			telefone		= trim(rs("telefone"))
			dataCadastro	= FormatDate(rs("dataCadastro"), "DD/MM/AA hh:mm")
			malaDir			= uCase(trim(rs("malaDir")))
			status			= uCase(trim(rs("status")))
		Else
			rs.close
			erro = "É preciso informar um cliente para poder edita-lo"
		End If
	Else
		'INCLUIR
		pais	= "Brasil"
		malaDir	= "S"
		status	= "P"
	End If
Else
	'SEGUNDA VEZ NO FORMULÁRIO, VALIDA DADOS PARA SALVAR
	usuario			= trim(request("usuario"))
	login			= trim(request("login"))
	senha			= trim(request("senha"))
	senhaConf		= trim(request("senhaConf"))
	nome			= trim(request("nome"))
	endereco		= trim(request("endereco"))
	bairro			= trim(request("bairro"))
	cidade			= trim(request("cidade"))
	estado			= trim(request("estado"))
	cep				= trim(request("cep"))
	pais			= trim(request("pais"))
	email			= trim(request("email"))
	telefone		= trim(request("telefone"))
	malaDir			= uCase(trim(request("malaDir")))
	status			= uCase(trim(request("status")))
	If malaDir <> "S" And malaDir <> "N" Then
		malaDir = "S"
	end if
	If status <> "A" And status <> "I" Then
		status = "P"
	end if
	'FAZ A VALIDAÇÃO DOS CAMPOS DO FORMULÁRIO
	If Len(login) = 0 Or Len(login) > maxLogin Then erro = " - O login deve ser informado e deve conter no máximo " & maxLogin & " caracteres<br>"
	If Len(senha) = 0 Or Len(senha) > maxSenha Then erro = erro & " - A senha deve ser informada e deve conter no máximo " & maxSenha & " caracteres<br>"
	If cStr(senha) <> cStr(senhaConf) Then erro = erro & " - A senha digitada não confere com a senha da confirmação<br>"
	If Len(nome) = 0 Or Len(nome) > maxNome Then erro = " - O nome deve ser informado e deve conter no máximo " & maxNome & " caracteres<br>"
	If Len(endereco) > maxEndereco Then erro = erro & " - O campo endereço deve conter no máximo " & maxEndereco & " caracteres<br>"
	If Len(bairro) > maxBairro Then erro = erro & " - O campo bairro deve conter no máximo " & maxBairro & " caracteres<br>"
	If Len(cidade) > maxCidade Then erro = erro & " - O campo cidade deve conter no máximo " & maxCidade & " caracteres<br>"
	If Len(estado) > maxestado Then erro = erro & " - O campo estado deve conter no máximo " & maxEstado & " caracteres<br>"
	If Len(cep) > maxCep Then erro = erro & " - O campo cep deve conter no máximo " & maxCep & " caracteres<br>"
	If Len(pais) > maxPais Then erro = erro & " - O campo pais deve conter no máximo " & maxPais & " caracteres<br>"
	If Len(email) = 0 Or Len(email) > maxEmail Then erro = erro & " - O e-mail deve ser informado e deve conter no máximo " & maxEmail & " caracteres<br>"
	If Len(telefone) > maxTelefone Then erro = erro & " - O campo telefone deve conter no máximo " & maxTelefone & " caracteres<br>"
	
	'PREPARA CAMPO PARA SALVAR
	If erro = "" And erro2 = "" Then
		email = lCase(replace(email, "'", ""))
		If oper = "I" Then
			'INCLUI USUARIO
			
			'VERIFICA DUPLICIDADE DE E-MAIL
			query = "SELECT usuario FROM " & nomeTabela & " WHERE email='" & email & "'"
			rs.Open query, conn
			If Not rs.EOF Then erro = erro & " - Já existe um usuário cadastrado com este email<br>"
			rs.close
			
			'VERIFICA DUPLICIDADE DE LOGIN
			query = "SELECT usuario FROM " & nomeTabela & " WHERE login='" & login & "'"
			rs.Open query, conn
			If Not rs.EOF Then erro = erro & " - Já existe um usuário cadastrado com este login<br>"
			rs.close
			
			If erro = "" Then
				Call preparaCampos()
				query =	("INSERT INTO " & nomeTabela & " (login, senha, nome, endereco, bairro, cidade, estado, " & _
								"			cep, pais, email, telefone, dataCadastro, malaDir, status) VALUES " & _
								"	('" & login & "', '" & senha & "', '" & nome & "', " & _
								"	 '" & endereco & "', '" & bairro & "', '" & cidade & "', '"& estado &"', '" & cep & "', " & _
								"	 '" & pais & "', '" & email & "', '" & telefone & "', getDate(), '" & malaDir & "', '" & status & "')")
				conn.Execute query 
				
				login = ""
				senha = ""
				senhaConf = ""
				nome = ""
				endereco = ""
				bairro = ""
				cidade = ""
				estado = ""
				cep = ""
				pais = ""
				email = ""
				telefone = ""
				tipo = ""
				malaDir = "S"
				status = "P"
				
				sSQL = "SELECT MAX(usuario) FROM XCA_usuarios"
				aMaxUsu = getArray(sSQL)
				if isArray(aMaxUsu) then
					usuario = aMaxUsu(0,0)
				end if
				
				Response.Redirect "assine_cadastro.asp?oper=F&usuario=" & usuario & "&" & strRedir
				Response.End
			End If
			
		Else
			'GRAVA EDICAO
			
			'VERIFICA DUPLICIDADE DE E-MAIL
			query = "SELECT usuario FROM " & nomeTabela & " WHERE email='" & email & "' AND usuario <> " & usuario
			rs.Open query, conn
			If Not rs.EOF Then erro = " - Já existe um cliente cadastrado com este email<br>"
			rs.close
			
			'VERIFICA DUPLICIDADE DE LOGIN
			query = "SELECT usuario FROM " & nomeTabela & " WHERE login='" & login & "' AND usuario <> " & usuario
			rs.Open query, conn
			If Not rs.EOF Then erro = " - Já existe um cliente cadastrado com este login<br>"
			rs.close
			
			If erro = "" Then
				Call preparaCampos()
				query =		("UPDATE " & nomeTabela & " SET " & _
								"login = '" & login & "', " & _
								"senha = '" & senha & "', " & _
								"nome = '" & nome & "', " & _
								"endereco = '" & endereco & "', " & _
								"bairro = '" & bairro & "', " & _
								"cidade = '" & cidade & "', " & _
								"estado = '" & estado & "', " & _
								"cep = '" & cep & "', " & _
								"pais = '" & pais & "', " & _
								"email = '" & email & "', " & _
								"telefone = '" & telefone & "', " & _
								"malaDir = '" & malaDir & "', " & _
								"status = '" & status & "' " & _
								"WHERE usuario = " & usuario)
				conn.Execute query
				
				Response.Redirect "home.asp?msg_erro=Atualização realizada com sucesso!&" & strRedir
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
						<div class="erro" align="right"><%=id_getText("cadastro_txt_14")%></div>
						<table border="0" cellpadding="0" cellspacing="5" class="texto2">
							
							<input type="hidden" name="oper" value="<%=oper%>">
							<input type="hidden" name="cmd">
							<input type="hidden" name="usuario"		value="<%=usuario%>">
							<input type="Hidden" name="pagina"		value="<%=request("pagina")%>">
							<input type="Hidden" name="ordemField"	value="<%=request("ordemField")%>">
							<input type="Hidden" name="ordemType"	value="<%=request("ordemType")%>">
							<input type="Hidden" name="fCampo"		value="<%=request("fCampo")%>">
							<input type="Hidden" name="fTexto"		value="<%=request("fTexto")%>">
							<input type="Hidden" name="fMalaDir"	value="<%=request("fMalaDir")%>">
							<input type="Hidden" name="fStatus"		value="<%=request("fStatus")%>">
							
							<% If erro <> "" Then %>
								<tr>
									<td class="erro">&nbsp;ERRO:</td>
									<td>&nbsp;&nbsp;</td>
									<td class="erro"><%=erro%></td>
								</tr>
							<% End If %>
							
							<tr>
								<td><span class="erro"><sup>*</sup></span>&nbsp;<%=id_getText("cadastro_txt_01")%></td>
								<td>&nbsp;&nbsp;</td>
								<td><input type="text" name="login" class="cx" size="50" maxlength="<%=maxLogin%>" value="<%=login%>"></td>
							</tr>
							<tr>
								<td><span class="erro"><sup>*</sup></span>&nbsp;<%=id_getText("cadastro_txt_02")%></td>
								<td>&nbsp;&nbsp;</td>
								<td><input type="password" name="senha" class="cx" size="50" maxlength="<%=maxSenha%>" value="<%=senha%>"></td>
							</tr>
							<tr>
								<td><span class="erro"><sup>*</sup></span>&nbsp;<%=id_getText("cadastro_txt_03")%></td>
								<td>&nbsp;&nbsp;</td>
								<td><input type="password" name="senhaConf" class="cx" size="50" maxlength="<%=maxSenha%>" value="<%=senhaConf%>"></td>
							</tr>
							<tr>
								<td><span class="erro"><sup>*</sup></span>&nbsp;<%=id_getText("cadastro_txt_04")%></td>
								<td>&nbsp;&nbsp;</td>
								<td><input type="text" name="nome" class="cx" size="50" maxlength="<%=maxNome%>" value="<%=nome%>"></td>
							</tr>
							<tr>
								<td>&nbsp;&nbsp;<%=id_getText("cadastro_txt_05")%></td>
								<td>&nbsp;&nbsp;</td>
								<td><input type="text" name="endereco" class="cx" size="50" maxlength="<%=maxEndereco%>" value="<%=endereco%>"></td>
							</tr>
							<tr>
								<td>&nbsp;&nbsp;<%=id_getText("cadastro_txt_06")%></td>
								<td>&nbsp;&nbsp;</td>
								<td><input type="text" name="bairro" class="cx" size="50" maxlength="<%=maxBairro%>" value="<%=bairro%>"></td>
							</tr>
							<tr>
								<td>&nbsp;&nbsp;<%=id_getText("cadastro_txt_07")%></td>
								<td>&nbsp;&nbsp;</td>
								<td><input type="text" name="cidade" class="cx" size="50" maxlength="<%=maxCidade%>" value="<%=cidade%>"></td>
							</tr>
							<tr>
								<td>&nbsp;&nbsp;<%=id_getText("cadastro_txt_08")%></td>
								<td>&nbsp;&nbsp;</td>
								<td><select name="estado" class="sel">
										<option value="OU" <% If estado="OU" Then Response.Write(" SELECTED ") %>>Outro</option>
										<option value="AC" <% If estado="AC" Then Response.Write(" SELECTED ") %>>AC</option>
										<option value="AL" <% If estado="AL" Then Response.Write(" SELECTED ") %>>AL</option>
										<option value="AM" <% If estado="AM" Then Response.Write(" SELECTED ") %>>AM</option>
										<option value="AP" <% If estado="AP" Then Response.Write(" SELECTED ") %>>AP</option>
										<option value="BA" <% If estado="BA" Then Response.Write(" SELECTED ") %>>BA</option>
										<option value="CE" <% If estado="CE" Then Response.Write(" SELECTED ") %>>CE</option>
										<option value="DF" <% If estado="DF" Then Response.Write(" SELECTED ") %>>DF</option>
										<option value="ES" <% If estado="ES" Then Response.Write(" SELECTED ") %>>ES</option>
										<option value="GO" <% If estado="GO" Then Response.Write(" SELECTED ") %>>GO</option>
										<option value="MA" <% If estado="MA" Then Response.Write(" SELECTED ") %>>MA</option>
										<option value="MG" <% If estado="MG" Then Response.Write(" SELECTED ") %>>MG</option>
										<option value="MS" <% If estado="MS" Then Response.Write(" SELECTED ") %>>MS</option>
										<option value="MT" <% If estado="MT" Then Response.Write(" SELECTED ") %>>MT</option>
										<option value="PA" <% If estado="PA" Then Response.Write(" SELECTED ") %>>PA</option>
										<option value="PB" <% If estado="PB" Then Response.Write(" SELECTED ") %>>PB</option>
										<option value="PE" <% If estado="PE" Then Response.Write(" SELECTED ") %>>PE</option>
										<option value="PI" <% If estado="PI" Then Response.Write(" SELECTED ") %>>PI</option>
										<option value="PR" <% If estado="PR" Then Response.Write(" SELECTED ") %>>PR</option>
										<option value="RJ" <% If estado="RJ" Then Response.Write(" SELECTED ") %>>RJ</option>
										<option value="RN" <% If estado="RN" Then Response.Write(" SELECTED ") %>>RN</option>
										<option value="RO" <% If estado="RO" Then Response.Write(" SELECTED ") %>>RO</option>
										<option value="RR" <% If estado="RR" Then Response.Write(" SELECTED ") %>>RR</option>
										<option value="RS" <% If estado="RS" Then Response.Write(" SELECTED ") %>>RS</option>
										<option value="SC" <% If estado="SC" Then Response.Write(" SELECTED ") %>>SC</option>
										<option value="SE" <% If estado="SE" Then Response.Write(" SELECTED ") %>>SE</option>
										<option value="SP" <% If estado="SP" Then Response.Write(" SELECTED ") %>>SP</option>
										<option value="TO" <% If estado="TO" Then Response.Write(" SELECTED ") %>>TO</option>
									</select><%=id_getText("cadastro_txt_09")%>
								</td>
							</tr>
							<tr>
								<td>&nbsp;&nbsp;<%=id_getText("cadastro_txt_10")%></td>
								<td>&nbsp;&nbsp;</td>
								<td><input type="text" name="pais" class="cx" size="50" maxlength="<%=maxPais%>" value="<%=pais%>"></td>
							</tr>
							<tr>
								<td><span class="erro"><sup>*</sup></span>&nbsp;<%=id_getText("cadastro_txt_11")%></td>
								<td>&nbsp;&nbsp;</td>
								<td><input type="text" name="email" class="cx" size="50" maxlength="<%=maxEmail%>" value="<%=email%>"></td>
							</tr>
							<tr>
								<td>&nbsp;&nbsp;<%=id_getText("cadastro_txt_12")%></td>
								<td>&nbsp;&nbsp;</td>
								<td><input type="text" name="telefone" class="cx" size="50" maxlength="<%=maxTelefone%>" value="<%=telefone%>"></td>
							</tr>
							<tr>
								<td><span class="erro"><sup>*</sup></span>&nbsp;<%=id_getText("cadastro_txt_13")%></td>
								<td>&nbsp;&nbsp;</td>
								<td><% Call criaComboMalaDir(malaDir) %></td>
							</tr>
							<tr>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
							</tr>
							<tr>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td>
									<input type="button"	onclick="salvar(); return false"	value="<%=id_getText("cadastro_txt_15")%>" class="bt">
							<%
								If oper = "E" Then
							%>
									<input type="button"	onclick="excluir(); return false"	value="<%=id_getText("cadastro_txt_16")%>" class="bt">
							<%
								end if
							%>
								</td>
							</tr>
						</table>
						
				<%
					If oper = "I" Then
				%>
				<script language="javascript1.2">
					document.forms[0].nome.focus()
				</script>
				<%
					End If
				%>
			</table>
		</td>
	</tr>
</table>
<%
Sub preparaCampos()
		login = replace(login, "'", "''")
		senha = replace(senha, "'", "''")
		nome = replace(nome, "'", "''")
		endereco = replace(endereco, "'", "''")
		bairro = replace(bairro, "'", "''")
		cidade = replace(cidade, "'", "''")
		pais = replace(pais, "'", "''")
		email = lCase(replace(email, "'", "''"))
		telefone = replace(telefone, "'", "''")
		status = uCase(status)
End Sub

Sub criaCombostatus(status)
		s = "<select name='status' class='cx'>" & vbCrLf
		s = s & "<option value='A' " 
		If status = "A" Then s = s & " SELECTED "
		s = s & ">Ativo" & vbCrLf
		s = s & "<option value='P' " 
		If status = "P" Then s = s & " SELECTED "
		s = s & ">Pendente" & vbCrLf
		s = s & "<option value='I' " 
		If status = "I" Then s = s & " SELECTED "
		s = s & ">Inativo" & vbCrLf
		s = s & "</select>" & vbCrLf
		Response.Write s
End Sub

Sub criaComboMalaDir(malaDir)
		s = "<select name='malaDir' class='cx'>" & vbCrLf
		s = s & "<option value='S' " 
		If malaDir = "S" Then s = s & " SELECTED "
		s = s & ">Sim" & vbCrLf
		s = s & "<option value='N' " 
		If malaDir = "N" Then s = s & " SELECTED "
		s = s & ">Não" & vbCrLf
		s = s & "</select>" & vbCrLf
		Response.Write s
End Sub
%>