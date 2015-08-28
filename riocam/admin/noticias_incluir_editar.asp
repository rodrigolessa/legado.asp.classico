<% 
Dim strRedir
Const maxTitulo = 100
Const maxLead = 200

strRedir = "pagina=" & Request("pagina") & "&ordemField=" & Request("ordemField") & "&ordemType=" & Request("ordemType") & "&fCampo=" & Request("fCampo") & "&fTexto=" & Request("fTexto")  & "&fStatus=" & Request("fStatus")

'VERIFICA SE É A PRIMEIRA VEZ
If Request("cmd")="" Then
	'VERIFICA SE É A PRIMEIRA VEZ DA EDIÇÃO OU DA INCLUSÃO
	If oper = "E" Then
		'EDITAR
		noticia = Request("noticia")
		query = "SELECT noticia, titulo, lead, texto, data, dataExpiracao, status FROM " & nomeTabela & " WHERE noticia=" & noticia 
		rs.Open query, conn
		If Not rs.EOF Then
			titulo = Trim(rs("titulo"))
			lead = Trim(rs("lead"))
			texto = Trim(rs("texto"))
			data = formatDate(rs("data"), "DD/MM/AA")
			If Not IsNull(rs("dataExpiracao")) Then
				dataExpiracao = formatDate(rs("dataExpiracao"), "DD/MM/AA")
				notExpira = 0
			Else
				dataExpiracao = ""
				notExpira = 1
			End If
			status = UCase(rs("status"))
		Else
 			erro = "É preciso informar uma noticia para poder edita-la"
		End If
		rs.close
	Else
		'INCLUIR
		data = Now()
		data = formatDate(data, "DD/MM/AA")
		dataExpiracao = DateAdd("m", 1, Date())
		dataExpiracao = formatDate(dataExpiracao, "DD/MM/AA")
		notExpira = 0
	End If
Else
	'SEGUNDA VEZ NO FORMULÁRIO, VALIDA DADOS PARA SALVAR
	noticia = Request("noticia")
	titulo = Request("titulo")
	lead = Request("lead")
	texto = Request("texto")
	data = Request("data")
	If data = "" Then data = formatDate(Now(), "DD/MM/AA")
	notExpira = Request("notExpira")
	If notExpira<>"" Then 
		notExpira = 1
		dataExpiracao = Null
	Else 
		notExpira = 0
		dataExpiracao = Request("dataExpiracao")
		If dataExpiracao <> "" Then horaExpiracao = Request("horaExpiracao") Else dataExpiracao = Null
	End If
	status = Request("status")
	If Len(titulo) < 1 Or Len(titulo) > maxTitulo Then erro = erro & "-O campo título deve ser informado e deve conter no máximo " & maxTitulo & " caracteres<br>"
	If Len(lead) > maxLead Then erro = erro & " - O campo lead deve conter no máximo " & maxLead & " caracteres<br>"
	If Not IsDate(data) Then erro = erro & " - Data inválida<br>"
	If Not IsNull(dataExpiracao) Then
		If Not IsDate(dataExpiracao) Then 
			erro = erro & " - Data de expiração inválida<br>"
		Else
			If Cdate(data) > Cdate(dataExpiracao) Then erro = erro & "-Data de expiração menor que data da noticia<br>"
		End If
	End If
	'PREPARA CAMPO PARA SALVAR
	If erro="" Then
		Call preparaCampos()
		If oper="I" Then
				Application.Lock
				query = "INSERT INTO " & nomeTabela & " (titulo, lead, texto, data, dataExpiracao, status) VALUES " & _
								"	('" & sTitulo & "', '" & sLead & "', '" & sTexto & "', " & sData & ", " & _
										sDataExpiracao & ", '" & status & "')"
				conn.Execute query
				Application.UnLock
				If Request("cmd") = "S" Then
					titulo = ""
					lead = ""
					texto = ""
					status = "A"
					Response.Write "<script language='JavaScript'> alert('noticia cadastrado com sucesso!') </script>"
				Else
					query = "Select Max(noticia) From " & nomeTabela
					rs.open query, conn
					If not rs.Eof Then noticia = rs(0)
					rs.close
					oper = "E"
				End If
		Else
			'GRAVA EDICAO
				query = "UPDATE " & nomeTabela & " SET " & _
								"		titulo = '" & sTitulo & "', " & _
								"		lead = '" & sLead & "', " & _
								"		texto = '" & sTexto & "', " & _
								"		data = " & sData & ", " & _
								"		dataExpiracao = " & sDataExpiracao & ", " & _
								"		status = '" & status & "' " & _
								" WHERE noticia = " & noticia & ""
				conn.Execute query
				Response.Redirect "noticias.asp?noticia=" & noticia & "&" & strRedir
				Response.End
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
							<table border="0" cellpadding="0" cellspacing="0" class="texto2">
								
									<input type="hidden" name="oper" value="<%=oper%>">
									<input type="hidden" name="cmd">
									<input type='hidden' name='noticia' value='<%= noticia %>'>
									<input type='Hidden' name='pagina' value='<%=Request("pagina")%>'>
									<input type='Hidden' name='ordemField' value='<%=Request("ordemField")%>'>
									<input type='Hidden' name='ordemType' value='<%=Request("ordemType")%>'>
									<input type='Hidden' name='fCampo' value='<%=Request("fCampo")%>'>
									<input type='Hidden' name='fTexto' value='<%=Request("fTexto")%>'>
									<input type='Hidden' name='fStatus' value='<%=Request("fStatus")%>'>
								<% If erro <> "" Then %>
									<tr>
										<td class="erro">&nbsp;ERRO:</td>
										<td>&nbsp;&nbsp;</td>
										<td class="erro"><%=erro%></td>
									</tr>
								<% End If %>
								<tr>
									<td><span class="erro"><sup>*</sup></span>&nbsp;Data de início:</td>
									<td>&nbsp;&nbsp;</td>
									<td valign="middle">
										<table cellpadding="0" cellspacing="2" border="0">
										<tr>
											<td><input type="text" name="data" maxlength="10" class="cx" size="12" value="<%=data%>" OnFocus="SetarEvento(this,10,'I')"></td>
											<td><a href="#" onclick="openDlgCalendario('data', '<%=data%>'); return(false)"><img src="<%=dirImgsGerencia%>ico_calendario.gif" border=0></a></td>
										</tr>
										</table>
									</td>
								</tr>
								<tr>
									<td><span class="erro"><sup>*</sup></span>&nbsp;Data de expiração:</td>
									<td>&nbsp;&nbsp;</td>
									<td valign="middle">
										<table cellpadding="0" cellspacing="2" border="0">
										<tr>
											<td><input type="text" name="dataExpiracao" maxlength="10" class="cx" size="12" value="<%=dataExpiracao%>" OnFocus="SetarEvento(this,10,'I')"></td>
											<td><a href="#" onclick="openDlgCalendario('dataExpiracao', '<%=dataExpiracao%>'); return(false)"><img src="<%=dirImgsGerencia%>ico_calendario.gif" border=0></a></td>
											<td>&nbsp;<span class="texto2">noticia não expira</span><input type="checkbox" name="notExpira" value="1" <% If notExpira=1 Then Response.Write " CHECKED " %>></td>
										</tr>
										</table>
									</td>
								</tr>
								<tr>
									<td><span class="erro"><sup>*</sup></span>&nbsp;Título:</td>
									<td>&nbsp;&nbsp;</td>
									<td><input type="text" name="titulo" maxlength="<%=maxTitulo%>" class="cx" size="97" value='<%= replace(titulo, """", "&quot;") %>'></td>
								</tr>
								<tr>
									<td valign="top">&nbsp;&nbsp;Lead:</td>
									<td>&nbsp;&nbsp;</td>
									<td valign="top">
										<table cellpadding="0" cellspacing="0" border="0">
											<tr>
												<td><textarea cols="80" rows="12" name="lead" class="tx"><%=lead%></textarea></td>
												<td valign="top">&nbsp;</td>
											</tr>
										</table>
									</td>
								</tr>
								<tr>
									<td valign="top">&nbsp;&nbsp;Texto:</td>
									<td>&nbsp;&nbsp;</td>
									<td>
										<table cellpadding="0" cellspacing="0" border="0">
											<tr>
												<td><textarea cols="80" rows="12" name="texto" class="tx" style="width:500; height:100"><%=texto%></textarea></td>
												<td valign="top">&nbsp;</td>
												<script language="javascript1.2">
													editor_generate('texto');
												</script>
											</tr>
										</table>
									</td>
								</tr>
								<tr>
									<td>&nbsp;&nbsp;Status:</td>
									<td>&nbsp;&nbsp;</td>
									<td><% Call criaComboStatus(status) %></td>
								</tr>
								
							</table>
					</td>
				</tr>
			</table>
		</td>
	</tr>
</table>
<%
Sub preparaCampos()
		sTitulo = replace(titulo, "'", "''")
		sLead = replace(lead, "'", "''")
		sTexto = replace(texto, "'", "''")
		If Not IsNull(data) Then 
			'A funcao convert é utilizada para compatibilidade com qualquer configuracao regional
			'do sistema ou do banco de dados, o 120 é utilizado para informar que o formato da data
			'é ODBC (24h) sem milisegundos para a funcao com milisegundos deve-se utilizar 121
			sData = "'" & Year(data) & "-" & Month(data) & "-" & Day(data) & "'"
		Else 
			sData = "NULL"
		End If
		If Not IsNull(dataExpiracao) Then
			sDataExpiracao = "'" & Year(dataExpiracao) & "-" & Month(dataExpiracao) & "-" & Day(dataExpiracao) & "'"
		Else 
			sDataExpiracao = "NULL"
		End If
		status = UCase(status)
End Sub
%>