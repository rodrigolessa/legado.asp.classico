<% 
Dim strRedir
Const maxDescricao = 100

erro = ""

strRedir = "pagina=" & Request("pagina") & "&ordemField=" & Request("ordemField") & "&ordemType=" & Request("ordemType")

'VERIFICA SE � A PRIMEIRA VEZ
If Request("cmd") = "" Then

	'VERIFICA SE � A PRIMEIRA VEZ DA EDI��O OU DA INCLUS�O
	If oper = "E" Then
	
		'EDITAR
		modelo_categ = request("modelo_categ")
		
		query = "SELECT modelo_categ, descricao, descricao_eng, status FROM " & nomeTabela & " WHERE modelo_categ = " & modelo_categ
		rs.Open query, conn
		If Not rs.EOF Then
			descricao		= Trim(rs("descricao"))
			descricao_eng	= Trim(rs("descricao_eng"))
			status			= UCase(rs("status"))
		Else
			erro = "� preciso informar uma categoria para poder edita-la"
		End If
		rs.close
		
	Else
		'INCLUIR
	End If
	
Else
	'SEGUNDA VEZ NO FORMUL�RIO, VALIDA DADOS PARA SALVAR
	modelo_categ	= Request("modelo_categ")
	descricao		= Request("descricao")
	descricao_eng	= Request("descricao_eng")
	status			= Request("status")
	If Len(descricao) < 1 Or Len(descricao) > maxDescricao Then erro = "-O campo descricao deve ser informado e deve conter no m�ximo " & maxDescricao & " caracteres<br>"
	If Len(descricao_eng) < 1 Or Len(descricao_eng) > maxDescricao Then erro = "-O campo descricao (ingl�s) deve ser informado e deve conter no m�ximo " & maxDescricao & " caracteres<br>"
	
	'PREPARA CAMPO PARA SALVAR
	If erro = "" Then
	
		descricao		= replace(trim(descricao), "'", "''")
		descricao_eng	= replace(trim(descricao_eng), "'", "''")
		status			= uCase(status)
		
		If oper = "I" Then
		
			'*************************************************
			'INCLUI CATEGORIAS DE MODELOS
			'*************************************************
			
'			ON ERROR RESUME NEXT
			
			'INICIA TRANSA��O DO BANCO DE DADOS
			conn.BeginTrans
		
			query = "SELECT modelo_categ FROM " & nomeTabela & " WHERE descricao = '" & descricao & "' "
			rs.Open query, conn
			If Not rs.EOF Then
				erro = erro & "J� existe uma categoria com esta descri��o"
			end if
			rs.close
			
			If erro = "" Then
				sSQL =	"INSERT INTO " & nomeTabela & " (descricao, descricao_eng, status) VALUES " & _
						"	('" & descricao & "', '" & descricao_eng & "', '" & status & "')"
				conn.execute(sSQL)
			End If
			
			If conn.Errors.Count = 0 Then
			
				conn.CommitTrans
				erro			= ""
				
				'LIMPA TODOS OS CAMPOS DO FORMUL�RIO
				descricao		= ""
				descricao_eng	= ""
				status			= "I"
				
				%>
				<script language="JavaScript">
					alert("Categoria cadastrada com sucesso!!")
				</script>
				<%

			Else
			
				conn.RollbackTrans
				erro = "Ocorreu erro na inclus�o de um registro, o cadastro n�o foi efetuado!"
				
			End If

'			ON ERROR GOTO 0
			
		Else
		
			'************************************************************
			'GRAVA EDICAO
			'************************************************************
			
'			ON ERROR RESUME NEXT
			
			'INICIA TRANSA��O DO BANCO DE DADOS
			conn.BeginTrans
			
			query = "SELECT modelo_categ FROM " & nomeTabela & " WHERE descricao = '" & descricao & "' AND modelo_categ <> " & modelo_categ
			rs.Open query, conn
			If Not rs.EOF Then erro = "J� existe uma categoria cadastrada com esta descri��o<br>"
			rs.close
			If erro = "" Then
				query =	"UPDATE " & nomeTabela & " SET " & _
						"		descricao		= '" & descricao & "', " & _
						"		descricao_eng	= '" & descricao_eng & "', " & _
						"		status			= '" & status & "' " & _
						" WHERE modelo_categ = " & modelo_categ & ""
				conn.execute query
			End If
			
			If conn.Errors.Count = 0 Then
			
				conn.CommitTrans
				
				erro = ""
				
				response.redirect "modelos_categs.asp?modelo_categ=" & modelo_categ & "&" & strRedir
				response.end
				
			Else
			
				conn.RollbackTrans
				erro = "Ocorreu erro na altera��o de um registro, a atualiza��o n�o foi efetuada!"
				
			End If

'			ON ERROR GOTO 0
			
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
						<div class="erro" align="right">* preenchimento obrigat�rio</div>
						<table border="0" cellpadding="0" cellspacing="5" class="texto2">
						
							<input type="hidden" name="oper"			value="<%=oper%>">
							<input type="hidden" name="cmd">
							<input type="Hidden" name="pagina"			value="<%=Request("pagina")%>">
							<input type="Hidden" name="ordemField"		value="<%=Request("ordemField")%>">
							<input type="Hidden" name="ordemType"		value="<%=Request("ordemType")%>">
							<input type="hidden" name="modelo_categ"	value="<%=modelo_categ%>">
							
							<% If erro <> "" Then %>
								<tr>
									<td class="erro">&nbsp;ERRO:</td>
									<td >&nbsp;&nbsp;</td>
									<td class="erro"><%=erro%></td>
								</tr>
							<% End If %>
							<tr>
								<td><span class="erro"><sup>*</sup></span>&nbsp;Descri��o:</td>
								<td>&nbsp;&nbsp;</td>
								<td>
									<input type="text" name="descricao" maxlength="<%=maxDescricao%>" class="cx" size="50" value="<%=descricao%>">
								</td>
							</tr>
							<tr>
								<td><span class="erro"><sup>*</sup></span>&nbsp;Descri��o (ingl�s):</td>
								<td>&nbsp;&nbsp;</td>
								<td>
									<input type="text" name="descricao_eng" maxlength="<%=maxDescricao%>" class="cx" size="50" value="<%=descricao_eng%>">
								</td>
							</tr>
							<tr>
								<td>&nbsp;&nbsp;Status:</td>
								<td>&nbsp;&nbsp;</td>
								<td><% Call criaComboStatus(status) %></td>
							</tr>

							</form>
							
							<%
								If oper="I" Then
							%>
							<script language="javascript1.2">
								document.forms[0].descricao.focus()
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