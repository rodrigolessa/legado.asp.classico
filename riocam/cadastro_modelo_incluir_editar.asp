<%

'***** DESCRICAO DOS SIGNOS *****
'
'01 - Aries
'02 - Touro
'03 - Gêmeos
'04 - Câncer
'05 - Leão
'06 - Virgem
'07 - Libra
'08 - Escorpião
'09 - Sargitário
'10 - Capricórnio
'11 - Aquário
'12 - Peixes
'
'Fogo		-	Terra		-	Ar		-	Água
'Áries		-	Touro		-	Gêmeos	-	Câncer
'Leão		-	Virgem		-	Libra	-	Escorpião
'Sagitário	-	Capricórnio	-	Aquário	-	Peixes
 
' 
'*********************************

strRedir = "pagina=" & Request("pagina") & "&ordemField=" & Request("ordemField") & "&ordemType=" & Request("ordemType") & "&fCampo=" & Request("fCampo") & "&fTexto=" & Request("fTexto") & "&fStatus=" & Request("fStatus") & "&fSexo=" & Request("fSexo")

'VERIFICA SE É A PRIMEIRA VEZ
If request("cmd") = "" Then
	'VERIFICA SE É A PRIMEIRA VEZ DA EDIÇÃO OU DA INCLUSÃO
	If oper = "E" Then
		'EDITAR
		modelo	= Request("modelo")
		
		query =	"SELECT	modelo, login, senha, nome, cidade, estado, pais, " & _
				"		idade, altura, signo, elemento, esporte, esporte_eng, descricao, descricao_eng, " & _
				"		email, telefone, dataCadastro, status, sexo, modelo_categ, " & _
				"		olhos, cabelo, busto, cintura, quadril, " & _
				"		olhos_eng, cabelo_eng, busto_eng, cintura_eng, quadril_eng, " & _
				"		sala " & _
				"FROM " & nomeTabela & " WHERE modelo = " & modelo & ""
				
		rs.Open query, conn
		
		If Not rs.EOF Then
			modelo			= trim(rs("modelo"))
			login			= trim(rs("login"))
			senha			= trim(rs("senha"))
			nome			= trim(rs("nome"))
			cidade			= trim(rs("cidade"))
			estado			= uCase(trim(rs("estado")))
			pais			= trim(rs("pais"))
			email			= lCase(trim(rs("email")))
			telefone		= trim(rs("telefone"))
			idade			= trim(rs("idade"))
			altura			= trim(rs("altura"))
			signo			= trim(rs("signo"))
			elemento		= trim(rs("elemento"))
			esporte			= trim(rs("esporte"))
			esporte_eng		= trim(rs("esporte_eng"))
			olhos			= trim(rs("olhos"))
			olhos_eng		= trim(rs("olhos_eng"))
			cabelo			= trim(rs("cabelo"))
			cabelo_eng		= trim(rs("cabelo_eng"))
			busto			= trim(rs("busto"))
			busto_eng		= trim(rs("busto_eng"))
			cintura			= trim(rs("cintura"))
			cintura_eng		= trim(rs("cintura_eng"))
			quadril			= trim(rs("quadril"))
			quadril_eng		= trim(rs("quadril_eng"))
			descricao		= trim(rs("descricao"))
			descricao_eng	= trim(rs("descricao_eng"))
			sexo			= trim(rs("sexo"))
			modelo_categ	= trim(rs("modelo_categ"))
			dataCadastro	= FormatDate(rs("dataCadastro"), "DD/MM/AA hh:mm")
			status			= trim(rs("status"))
			sala			= trim(rs("sala"))
		Else
			rs.close
			erro = "É preciso informar um cliente para poder edita-lo"
		End If
		
	Else
		'INCLUIR
		oper	= "I"
		pais	= "Brasil"
		status	= "P"
	End If
	
Else

	'SEGUNDA VEZ NO FORMULÁRIO, VALIDA DADOS PARA SALVAR
	modelo			= Request("modelo")
	login			= Request("login")
	senha			= Request("senha")
	nome			= Request("nome")
	cidade			= Request("cidade")
	estado			= Request("estado")
	pais			= Request("pais")
	email			= Request("email")
	telefone		= request("telefone")
	idade			= trim(request("idade"))
	altura			= trim(request("altura"))
	signo			= trim(request("signo"))
	elemento		= trim(request("elemento"))
	esporte			= trim(request("esporte"))
	esporte_eng		= trim(request("esporte_eng"))
	olhos			= trim(request("olhos"))
	olhos_eng		= trim(request("olhos_eng"))
	cabelo			= trim(request("cabelo"))
	cabelo_eng		= trim(request("cabelo_eng"))
	busto			= trim(request("busto"))
	busto_eng		= trim(request("busto_eng"))
	cintura			= trim(request("cintura"))
	cintura_eng		= trim(request("cintura_eng"))
	quadril			= trim(request("quadril"))
	quadril_eng		= trim(request("quadril_eng"))
	descricao		= trim(request("descricao"))
	descricao_eng	= trim(request("descricao_eng"))
	sexo			= trim(request("sexo"))
	modelo_categ	= trim(request("modelo_categ"))
	status			= request("status")
	sala			= request("sala")
	If status <> "A" And status <> "I" Then
		status = "P"
	end if
	
	If Len(login) = 0 Or Len(login) > maxLogin Then erro = " - O login deve ser informado e deve conter no máximo " & maxLogin & " caracteres<br>"
	If Len(senha) = 0 Or Len(senha) > maxSenha Then erro = erro & " - A senha deve ser informada e deve conter no máximo " & maxSenha & " caracteres<br>"
	If Len(nome) = 0 Or Len(nome) > maxNome Then erro = " - O campo nome deve ser informado e deve conter no máximo " & maxNome & " caracteres<br>"
	If Len(email) = 0 Or Len(email) > maxEmail Then erro = erro & " - O campo e-mail deve ser informado e deve conter no máximo " & maxEmail & " caracteres<br>"
	If Len(cidade) > maxCidade Then erro = erro & " - O campo cidade deve conter no máximo " & maxCidade & " caracteres<br>"
	If Len(estado) > maxestado Then erro = erro & " - O campo estado deve conter no máximo " & maxEstado & " caracteres<br>"
	If Len(pais) > maxPais Then erro = erro & " - O campo pais deve conter no máximo " & maxPais & " caracteres<br>"
	If Len(telefone) > maxTelefone Then erro = erro & " - O campo telefone deve conter no máximo " & maxTelefone & " caracteres<br>"
	If Len(idade) = 0 or Len(idade) > maxIdade Then erro = erro & " - O campo idade deve ser informado e deve conter no máximo " & maxIdade & " caracteres<br>"
	
	'PREPARA CAMPO PARA SALVAR
	If erro = "" And erro2 = "" Then
	
		email	= lCase(replace(email, "'", ""))
		login	= replace(login, "'", "")
		
		If oper = "I" Then
		
			'*************************************************
			'INCLUI MODELO
			'*************************************************
			
			ON ERROR RESUME NEXT
			
			'INICIA TRANSAÇÃO DO BANCO DE DADOS
			conn.BeginTrans
			
			query = "SELECT modelo FROM " & nomeTabela & " WHERE email = '" & email & "'"
			rs.Open query, conn
			If Not rs.EOF Then
				erro = erro & " - Já existe um modelo cadastrado com este email<br>"
			end if
			rs.close
			
			query = "SELECT modelo FROM " & nomeTabela & " WHERE login = '" & login & "'"
			rs.Open query, conn
			If Not rs.EOF Then
				erro = erro & " - Já existe um modelo cadastrado com este login<br>"
			end if
			rs.close
			
			If erro = "" Then
			
				Call preparaCampos()
				
				query =	"INSERT INTO " & nomeTabela & " " & _
						"	(	login, senha, nome, cidade, estado, modelo_categ, " & _
						"		idade, altura, signo, elemento, esporte, descricao, " & _
						"		esporte_eng, descricao_eng, " & _
						"		olhos, cabelo, busto, cintura, quadril, " & _
						"		olhos_eng, cabelo_eng, busto_eng, cintura_eng, quadril_eng, " & _
						"		pais, email, telefone, dataCadastro, status, sexo, sala) " & _
						"	VALUES " & _
						"	(	'" & login & "', '" & senha & "', '" & nome & "', '" & cidade & "', '"& estado &"', "& modelo_categ &", " & _
						"		'" & idade & "', '" & altura & "', '" & signo & "', '" & elemento & "', '" & esporte & "', '" & descricao & "', " & _
						"		'" & esporte_eng & "', '" & descricao_eng & "', " & _
						"		'" & olhos & "', '" & cabelo & "', '" & busto & "', '" & cintura & "', '" & quadril & "', " & _
						"		'" & olhos_eng & "', '" & cabelo_eng & "', '" & busto_eng & "', '" & cintura_eng & "', '" & quadril_eng & "', " & _
						"		'" & pais & "', '" & email & "', '" & telefone & "', getDate(), '" & status & "', '" & sexo & "', '" & sala & "')"

				conn.Execute query
				
				If conn.Errors.Count = 0 Then

					conn.CommitTrans

					'LIMPA TODOS OS CAMPOS DO FORMULÁRIO
					login			= ""
					senha			= ""
					nome			= ""
					cidade			= ""
					estado			= ""
					pais			= "Brasil"
					email			= ""
					telefone		= ""
					tipo			= ""
					idade			= ""
					sexo			= "F"
					altura			= ""
					signo			= ""
					elemento		= ""
					esporte			= ""
					esporte_eng		= ""
					olhos			= ""
					olhos_eng		= ""
					cabelo			= ""
					cabelo_eng		= ""
					busto			= ""
					busto_eng		= ""
					cintura			= ""
					cintura_eng		= ""
					quadril			= ""
					quadril_eng		= ""
					descricao		= ""
					descricao_eng	= ""
					modelo_categ	= ""
					status			= "P"
					sala			= ""
					
					oper			= "I"
					cmd				= ""

					%>
					<script language="JavaScript">
						alert("Modelo cadastrado com sucesso!!")
					</script>
					<%

					'CRIA UM DIRETÓRIO PARA A MODELO, COM O IDENTIFICADOR DA TABELA
					sSQL = "SELECT MAX(modelo) AS codModelo FROM " & nomeTabela & " "
					rs.Open sSQL, conn
					If Not rs.EOF Then
						codModelo = rs("codModelo")
					end if
					rs.close

					dirModelo = Application("XCA_APP_SERVERDIR_BASE") & "\modelos\" & codModelo

					if not existeDiretorio(dirModelo) then
						dirModRaiz	= criaDiretorio(dirModelo)
						dirModGal	= dirModRaiz & "\galeria"
						dirModVid	= dirModRaiz & "\video"
						dirModGal	= criaDiretorio(dirModGal)
						dirModVid	= criaDiretorio(dirModVid)
					end if

				Else

					conn.RollbackTrans
					erro =	" - Ocorreu erro na inclusão de um registro, o cadastro não foi efetuado!<br>" & _
							"<blockquote>" & _
							"<b>Erro: </b>" & ERR.number & "<br>" & _
							"<b>Descrição do erro: </b>" & ERR.Description & "<br>" & _
							"</blockquote>"
							'"<b>SQL: </b>" & query & _

				End If
				
			Else
			
				conn.RollbackTrans
	
			End If

			ON ERROR GOTO 0
			
		Else
		
			'************************************************************
			'GRAVA EDICAO
			'************************************************************
			
			ON ERROR RESUME NEXT
			
			'INICIA TRANSAÇÃO DO BANCO DE DADOS
			conn.BeginTrans
			
			query = "SELECT modelo FROM " & nomeTabela & " WHERE email = '" & email & "' AND modelo <> " & modelo
			rs.Open query, conn
			If Not rs.EOF Then
				erro = erro & " - Já existe um modelo cadastrado com este email<br>"
			end if
			rs.close
			
			query = "SELECT modelo FROM " & nomeTabela & " WHERE login = '" & login & "' AND modelo <> " & modelo
			rs.Open query, conn
			If Not rs.EOF Then
				erro = erro & " - Já existe um modelo cadastrado com este login<br>"
			end if
			rs.close
			
			
			If erro = "" Then
			
				Call preparaCampos()
				
				query =	"UPDATE " & nomeTabela & " SET " & _
						" login			='" & login & "', " & _
						" senha			='" & senha & "', " & _
						" nome			='" & nome & "', " & _
						" cidade		='" & cidade & "', " & _
						" estado		='" & estado & "', " & _
						" pais			='" & pais & "', " & _
						" email			='" & email & "', " & _
						" telefone		='" & telefone & "', " & _
						" idade			='" & idade & "', " & _
						" signo			='" & signo & "', " & _
						" elemento		='" & elemento & "', " & _
						" altura		='" & altura & "', " & _
						" esporte		='" & esporte & "', " & _
						" esporte_eng	='" & esporte_eng & "', " & _
						" olhos			='" & olhos & "', " & _
						" olhos_eng		='" & olhos_eng & "', " & _
						" cabelo		='" & cabelo & "', " & _
						" cabelo_eng	='" & cabelo_eng & "', " & _
						" busto			='" & busto & "', " & _
						" busto_eng		='" & busto_eng & "', " & _
						" cintura		='" & cintura & "', " & _
						" cintura_eng	='" & cintura_eng & "', " & _
						" quadril		='" & quadril & "', " & _
						" quadril_eng	='" & quadril_eng & "', " & _
						" descricao		='" & descricao & "', " & _
						" descricao_eng	='" & descricao_eng & "', " & _
						" sexo			='" & sexo & "', " & _
						" sala			='" & sala & "', " & _
						" modelo_categ	= " & modelo_categ & ", " & _
						" status		='" & status & "' " & _
						"WHERE modelo	= " & modelo
						
				conn.Execute query
				
				If conn.Errors.Count = 0 Then

					conn.CommitTrans

					response.redirect "modelos.asp?modelo=" & modelo & "&" & strRedir
					response.end

				Else

					conn.RollbackTrans
					erro = " - Ocorreu erro na alteração de um registro, a atualização não foi efetuada!<br>"

				End If
				
			Else
			
				conn.RollbackTrans
				
			End If

			ON ERROR GOTO 0
	
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
						<table border="0" cellpadding="0" cellspacing="2" class="texto2">
							
							<input type="hidden" name="oper"		value="<%=oper%>">
							<input type="hidden" name="cmd">
							<input type="hidden" name="modelo"		value="<%=modelo%>">
							<input type="Hidden" name="pagina"		value="<%=Request("pagina")%>">
							<input type="Hidden" name="ordemField"	value="<%=Request("ordemField")%>">
							<input type="Hidden" name="ordemType"	value="<%=Request("ordemType")%>">
							<input type="Hidden" name="fCampo"		value="<%=Request("fCampo")%>">
							<input type="Hidden" name="fTexto"		value="<%=Request("fTexto")%>">
							<input type="Hidden" name="fStatus"		value="<%=Request("fStatus")%>">
							<input type="Hidden" name="fSexo"		value="<%=Request("fSexo")%>">
							
							<% If erro <> "" Then %>
								<tr>
									<td class="erro" valign="top">&nbsp;ERRO:</td>
									<td >&nbsp;&nbsp;</td>
									<td class="erro"><%=erro%></td>
								</tr>
							<% End If %>
							<tr>
								<td><span class="erro"><sup>*</sup></span>&nbsp;Login:</td>
								<td>&nbsp;&nbsp;</td>
								<td><input type="text" name="login" class="cx" size="20" maxlength="<%=maxLogin%>" value="<%=login%>"></td>
							</tr>
							<tr>
								<td><span class="erro"><sup>*</sup></span>&nbsp;Senha:</td>
								<td>&nbsp;&nbsp;</td>
								<td><input type="text" name="senha" class="cx" size="20" maxlength="<%=maxSenha%>" value="<%=senha%>"></td>
							</tr>
							<tr>
								<td colspan=3>&nbsp;&nbsp;</td>
							</tr>
							<tr>
								<td><span class="erro"><sup>*</sup></span>&nbsp;Nome:</td>
								<td>&nbsp;&nbsp;</td>
								<td><input type="text" name="nome" class="cx" size="50" maxlength="<%=maxNome%>" value="<%=nome%>"></td>
							</tr>
							<tr>
								<td><span class="erro"><sup>*</sup></span>&nbsp;E-mail:</td>
								<td>&nbsp;&nbsp;</td>
								<td><input type="text" name="email" class="cx" size="50" maxlength="<%=maxEmail%>" value="<%=email%>"></td>
							</tr>
							<tr>
								<td><span class="erro"><sup>*</sup></span>&nbsp;Idade:</td>
								<td>&nbsp;&nbsp;</td>
								<td>
									<table border="0" cellpadding="0" cellspacing="0" class="texto2">
									<tr>
										<td><input type="text" name="idade" class="cx" size="5" maxlength="<%=maxIdade%>" value="<%=idade%>"></td>
										<td>&nbsp;</td>
										<td>&nbsp;Altura:</td>
										<td>&nbsp;</td>
										<td><input type="text" name="altura" class="cx" size="5" maxlength="<%=maxAltura%>" value="<%=altura%>"></td>
										<td>&nbsp;</td>
										<td>&nbsp;Sexo:</td>
										<td>&nbsp;</td>
										<td>
											<select name="sexo" class="sel">
												<option value="F" <% If sexo = "F" Then Response.Write(" SELECTED ") %>>Feminino</option>
												<option value="M" <% If sexo = "M" Then Response.Write(" SELECTED ") %>>Masculino</option>
												<option value="T" <% If sexo = "T" Then Response.Write(" SELECTED ") %>>Transexual</option>
											</select>
										</td>
									</tr>
									</table>
								</td>
							</tr>
							<tr>
								<td>&nbsp;&nbsp;Cidade:</td>
								<td>&nbsp;&nbsp;</td>
								<td>
									<table border="0" cellpadding="0" cellspacing="0" class="texto2">
									<tr>
										<td><input type="text" name="cidade" class="cx" size="25" maxlength="<%=maxCidade%>" value="<%=cidade%>"></td>
										<td>&nbsp;&nbsp;</td>
										<td>&nbsp;&nbsp;Estado:</td>
										<td>&nbsp;&nbsp;</td>
										<td><select name="estado" class="sel">
												<option value="OU" <% If estado="OU" Then Response.Write(" SELECTED ") %>>-</option>
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
											</select>
										</td>
									</tr>
									</table>
								</td>
							</tr>
							<tr>
								<td>&nbsp;&nbsp;País:</td>
								<td>&nbsp;&nbsp;</td>
								<td><input type="text" name="pais" class="cx" size="25" maxlength="<%=maxPais%>" value="<%=pais%>"></td>
							</tr>
							<tr>
								<td>&nbsp;&nbsp;Telefone:</td>
								<td>&nbsp;&nbsp;</td>
								<td><input type="text" name="telefone" class="cx" size="25" maxlength="<%=maxTelefone%>" value="<%=telefone%>"></td>
							</tr>
							<tr>
								<td>&nbsp;&nbsp;Categoria:</td>
								<td>&nbsp;&nbsp;</td>
								<td><%=criaComboCateg(modelo_categ)%></td>
							</tr>
							<tr>
								<td>&nbsp;&nbsp;Signo:</td>
								<td>&nbsp;&nbsp;</td>
								<td>
									<table border="0" cellpadding="0" cellspacing="0" class="texto2">
									<tr>
										<td>
											<select name="signo" class="sel">
												<option value="1"  <% If signo = "1"  Then Response.Write(" SELECTED ") %>>Aries</option>
												<option value="2"  <% If signo = "2"  Then Response.Write(" SELECTED ") %>>Touro</option>
												<option value="3"  <% If signo = "3"  Then Response.Write(" SELECTED ") %>>Gêmeos</option>
												<option value="4"  <% If signo = "4"  Then Response.Write(" SELECTED ") %>>Câncer</option>
												<option value="5"  <% If signo = "5"  Then Response.Write(" SELECTED ") %>>Leão</option>
												<option value="6"  <% If signo = "6"  Then Response.Write(" SELECTED ") %>>Virgem</option>
												<option value="7"  <% If signo = "7"  Then Response.Write(" SELECTED ") %>>Libra</option>
												<option value="8"  <% If signo = "8"  Then Response.Write(" SELECTED ") %>>Escorpião</option>
												<option value="9"  <% If signo = "9"  Then Response.Write(" SELECTED ") %>>Sargitário</option>
												<option value="10" <% If signo = "10" Then Response.Write(" SELECTED ") %>>Capricórnio</option>
												<option value="11" <% If signo = "11" Then Response.Write(" SELECTED ") %>>Aquário</option>
												<option value="12" <% If signo = "12" Then Response.Write(" SELECTED ") %>>Peixes</option>
											</select>
										</td>
										<td>&nbsp;</td>
										<td>&nbsp;Elemento:</td>
										<td>&nbsp;</td>
										<td>
											<select name="elemento" class="sel">
												<option value="T" <% If elemento = "T" Then Response.Write(" SELECTED ") %>>Terra</option>
												<option value="F" <% If elemento = "F" Then Response.Write(" SELECTED ") %>>Fogo</option>
												<option value="R" <% If elemento = "R" Then Response.Write(" SELECTED ") %>>Ar</option>
												<option value="A" <% If elemento = "A" Then Response.Write(" SELECTED ") %>>Água</option>
											</select>
										</td>
									</tr>
									</table>
								</td>
							</tr>
							
							
							<tr>
								<td>&nbsp;&nbsp;Olhos:</td>
								<td>&nbsp;&nbsp;</td>
								<td>
									<table border="0" cellpadding="0" cellspacing="0" class="texto2">
									<tr>
										<td><input type="text" name="olhos" class="cx" size="25" maxlength="<%=maxOlhos%>" value="<%=olhos%>"></td>
										<td>&nbsp;&nbsp;</td>
										<td>&nbsp;&nbsp;(em inglês):</td>
										<td>&nbsp;&nbsp;</td>
										<td><input type="text" name="olhos_eng" class="cx" size="25" maxlength="<%=maxOlhos%>" value="<%=olhos_eng%>"></td>
									</tr>
									</table>
								</td>
							</tr>
							
							<tr>
								<td>&nbsp;&nbsp;Cabelo:</td>
								<td>&nbsp;&nbsp;</td>
								<td>
									<table border="0" cellpadding="0" cellspacing="0" class="texto2">
									<tr>
										<td><input type="text" name="cabelo" class="cx" size="25" maxlength="<%=maxCabelo%>" value="<%=cabelo%>"></td>
										<td>&nbsp;&nbsp;</td>
										<td>&nbsp;&nbsp;(em inglês):</td>
										<td>&nbsp;&nbsp;</td>
										<td><input type="text" name="cabelo_eng" class="cx" size="25" maxlength="<%=maxCabelo%>" value="<%=cabelo_eng%>"></td>
									</tr>
									</table>
								</td>
							</tr>
							
							<tr>
								<td>&nbsp;&nbsp;Busto:</td>
								<td>&nbsp;&nbsp;</td>
								<td>
									<table border="0" cellpadding="0" cellspacing="0" class="texto2">
									<tr>
										<td><input type="text" name="busto" class="cx" size="25" maxlength="<%=maxBusto%>" value="<%=busto%>"></td>
										<td>&nbsp;&nbsp;</td>
										<td>&nbsp;&nbsp;(em inglês):</td>
										<td>&nbsp;&nbsp;</td>
										<td><input type="text" name="busto_eng" class="cx" size="25" maxlength="<%=maxBusto%>" value="<%=busto_eng%>"></td>
									</tr>
									</table>
								</td>
							</tr>
							
							<tr>
								<td>&nbsp;&nbsp;Cintura:</td>
								<td>&nbsp;&nbsp;</td>
								<td>
									<table border="0" cellpadding="0" cellspacing="0" class="texto2">
									<tr>
										<td><input type="text" name="cintura" class="cx" size="25" maxlength="<%=maxCintura%>" value="<%=cintura%>"></td>
										<td>&nbsp;&nbsp;</td>
										<td>&nbsp;&nbsp;(em inglês):</td>
										<td>&nbsp;&nbsp;</td>
										<td><input type="text" name="cintura_eng" class="cx" size="25" maxlength="<%=maxCintura%>" value="<%=cintura_eng%>"></td>
									</tr>
									</table>
								</td>
							</tr>
							
							<tr>
								<td>&nbsp;&nbsp;Quadril:</td>
								<td>&nbsp;&nbsp;</td>
								<td>
									<table border="0" cellpadding="0" cellspacing="0" class="texto2">
									<tr>
										<td><input type="text" name="quadril" class="cx" size="25" maxlength="<%=maxQuadril%>" value="<%=quadril%>"></td>
										<td>&nbsp;&nbsp;</td>
										<td>&nbsp;&nbsp;(em inglês):</td>
										<td>&nbsp;&nbsp;</td>
										<td><input type="text" name="quadril_eng" class="cx" size="25" maxlength="<%=maxQuadril%>" value="<%=quadril_eng%>"></td>
									</tr>
									</table>
								</td>
							</tr>
							
							
							<tr>
								<td>&nbsp;&nbsp;Esporte:</td>
								<td>&nbsp;&nbsp;</td>
								<td><input type="text" name="esporte" class="cx" size="95" maxlength="<%=maxEsporte%>" value="<%=esporte%>" style="width:500;"></td>
							</tr>
							<tr>
								<td>&nbsp;&nbsp;Esporte (inglês):</td>
								<td>&nbsp;&nbsp;</td>
								<td><input type="text" name="esporte_eng" class="cx" size="95" maxlength="<%=maxEsporte%>" value="<%=esporte_eng%>" style="width:500;"></td>
							</tr>
							
							<tr>
								<td valign="top">&nbsp;&nbsp;Descrição:</td>
								<td>&nbsp;&nbsp;</td>
								<td>
									<table cellpadding="0" cellspacing="0" border="0">
										<tr>
											<td><textarea cols="80" rows="12" name="descricao" class="tx" style="width:500; height:100"><%=descricao%></textarea></td>
											<td valign="top">&nbsp;</td>
											<script language="javascript1.2">
												editor_generate("descricao");
											</script>
										</tr>
									</table>
								</td>
							</tr>
							
							<tr>
								<td valign="top">&nbsp;&nbsp;Descrição (inglês):</td>
								<td>&nbsp;&nbsp;</td>
								<td>
									<table cellpadding="0" cellspacing="0" border="0">
										<tr>
											<td><textarea cols="80" rows="12" name="descricao_eng" class="tx" style="width:500; height:100"><%=descricao_eng%></textarea></td>
											<td valign="top">&nbsp;</td>
											<script language="javascript1.2">
												editor_generate("descricao_eng");
											</script>
										</tr>
									</table>
								</td>
							</tr>
							
							<tr>
								<td>&nbsp;&nbsp;Sala de VídeoChat:</td>
								<td>&nbsp;&nbsp;</td>
								<td><input type="text" name="sala" class="cx" size="15" maxlength="<%=maxSala%>" value="<%=sala%>" style="width:100;"></td>
							</tr>
							
							<tr>
								<td><span class="erro"><sup>*</sup></span>&nbsp;status:</td>
								<td>&nbsp;&nbsp;</td>
								<td><% Call criaComboStatus(status) %></td>
							</tr>
						</table>
						
				<% If oper = "I" Then %>
				<script language="javascript1.2">
					document.forms[0].nome.focus()
				</script>
				<% End If%>
			</table>
		</td>
	</tr>
</table>
<%
Sub preparaCampos()
	login			= replace(login, "'", "´")
	senha			= replace(senha, "'", "´")
	nome			= replace(nome, "'", "´")
	cidade			= replace(cidade, "'", "´")
	pais			= replace(pais, "'", "´")
	email			= lCase(replace(email, "'", "´"))
	telefone		= replace(telefone, "'", "´")
	idade			= replace(idade, "'", "´")
	sexo			= uCase(replace(sexo, "'", "´"))
	altura			= replace(altura, "'", "´")
	signo			= replace(signo, "'", "´")
	elemento		= replace(elemento, "'", "´")
	esporte			= replace(esporte, "'", "´")
	esporte_eng		= replace(esporte_eng, "'", "´")
	olhos			= replace(olhos, "'", "´")
	olhos_eng		= replace(olhos_eng, "'", "´")
	cabelo			= replace(cabelo, "'", "´")
	cabelo_eng		= replace(cabelo_eng, "'", "´")
	busto			= replace(busto, "'", "´")
	busto_eng		= replace(busto_eng, "'", "´")
	cintura			= replace(cintura, "'", "´")
	cintura_eng		= replace(cintura_eng, "'", "´")
	quadril			= replace(quadril, "'", "´")
	quadril_eng		= replace(quadril_eng, "'", "´")
	descricao		= replace(descricao, "'", "´")
	descricao_eng	= replace(descricao_eng, "'", "´")
	status			= uCase(status)
	sala			= lCase(replace(sala, "'", "´"))
End Sub

Sub criaComboStatus(status)
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


'FUNÇÃO PARA CRIAS A COMBO DE CATEGORIAS DAS MODELOS
function criaComboCateg(ccModeloCateg)
	DIM rs2, sSQL2, temp
	SET rs2	= CreateObject("ADODB.Recordset")
	
	ccModeloCateg = trim(ccModeloCateg)
	if len(ccModeloCateg) > 0 then
		if isNumeric(ccModeloCateg) then
			ccModeloCateg = cInt(ccModeloCateg)
		else
			ccModeloCateg = 0
		end if
	else
		ccModeloCateg = 0
	end if
	
	temp = ""

	sSQL2 = "SELECT modelo_categ, descricao FROM XCA_modelos_categs WHERE status = 'A' ORDER BY descricao "
	rs2.open sSQL2, conn

	if not rs2.EOF then

		temp = temp &	"<SELECT name=modelo_categ class=""sel"">"
						'"<option value=""""> </option>"
		while not rs2.EOF 
			temp = temp & "<option value="""&rs2("modelo_categ")&""" "
			if rs2("modelo_categ") = ccModeloCateg then temp = temp & "SELECTED"
			temp = temp & ">"&rs2("descricao")&"</option>"
			rs2.MoveNext
		wend
		temp = temp &	"</SELECT>"
		
	end if
	
		rs2.close
	set rs2 = nothing

	criaComboCateg = temp
end function
%>